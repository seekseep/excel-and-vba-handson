#!/usr/bin/env python3
"""
.bas（標準モジュール）から、Excel のマクロ有効ブック（.xlsm）に入れる vbaProject.bin を作る。

Excel が無くても作れるよう、Microsoft の公開仕様どおりに組み立てる:
  - [MS-OVBA] Office VBA File Format Structure … PROJECT / dir / モジュールのストリームと、その圧縮
  - [MS-CFB]  Compound File Binary File Format … それらを 1 つのファイルにまとめる入れ物

コンパイル済みのキャッシュ（p-code）は入れず、_VBA_PROJECT の Version を 0xFFFF にする。
こうすると Excel は開いたときにソースから作り直す（Office どうしでも使われている方法）。

使い方（build_workbooks.py から呼ぶ）:
    bin = build_vba_project(modules=[("FirstMacro", source_text)], sheet_code_names=["Sheet1"])
"""

import random
import struct
import uuid

CODEPAGE = 932  # 日本語の Excel が読むコードページ（ソースもこれで保存する）
ENCODING = "cp932"


# ---------------------------------------------------------------------------
# MS-OVBA 2.4.1 圧縮
# ---------------------------------------------------------------------------

def _copy_token_help(difference):
    bit_count = max((difference - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit_count
    maximum_length = length_mask + 3
    return bit_count, length_mask, maximum_length


def _compress_chunk(data, start, end):
    """data[start:end]（最大 4096 バイト）を 1 つの圧縮チャンクにする。"""
    out = bytearray()
    pos = start
    while pos < end:
        flag_index = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if pos >= end:
                break
            # このチャンクの中で、いちばん長く一致する過去の位置を探す（3 バイト以上）
            best_len = 0
            best_off = 0
            difference = pos - start
            if difference > 0:
                _, _, max_len = _copy_token_help(difference)
                limit = min(max_len, end - pos)
                for cand in range(pos - 1, start - 1, -1):
                    length = 0
                    while length < limit and data[cand + length] == data[pos + length]:
                        length += 1
                    if length > best_len:
                        best_len = length
                        best_off = pos - cand
                        if length == limit:
                            break
            if best_len >= 3:
                bit_count, _, _ = _copy_token_help(difference)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flag |= 1 << bit
                pos += best_len
            else:
                out.append(data[pos])
                pos += 1
        out[flag_index] = flag
    return bytes(out)


def compress(data):
    """MS-OVBA の CompressedContainer。"""
    out = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        end = min(start + 4096, len(data))
        body = _compress_chunk(data, start, end)
        if len(body) > 4096 and end - start == 4096:
            # 圧縮すると大きくなるときは、圧縮しないチャンクにする（ちょうど 4096 バイトのときだけ許される）
            header = (4096 + 2 - 3) | (0b011 << 12)
            out += struct.pack("<H", header) + data[start:end]
        else:
            header = (len(body) + 2 - 3) | (0b011 << 12) | (1 << 15)
            out += struct.pack("<H", header) + body
    return bytes(out)


def decompress(data):
    """compress の確かめ用（MS-OVBA 2.4.1.3.1）。"""
    assert data[0] == 1
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        compressed = header >> 15
        chunk_end = pos + size
        pos += 2
        chunk_start = len(out)
        if not compressed:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end:
            flag = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if flag & (1 << bit):
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    difference = len(out) - chunk_start
                    bit_count, length_mask, _ = _copy_token_help(difference)
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


# ---------------------------------------------------------------------------
# MS-OVBA 2.3 ストリームの中身
# ---------------------------------------------------------------------------

def _rec(rec_id, payload):
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _dir_stream(project_name, modules):
    """modules = [(name, is_document)]。圧縮前の dir ストリーム（MS-OVBA 2.3.4.2）。"""
    b = bytearray()
    b += _rec(0x0001, struct.pack("<I", 0x00000001))  # PROJECTSYSKIND: 32 ビット Windows
    b += _rec(0x0002, struct.pack("<I", 0x00000409))  # PROJECTLCID: 仕様で 0x409 と決まっている
    b += _rec(0x0014, struct.pack("<I", 0x00000409))  # PROJECTLCIDINVOKE: 同上
    b += _rec(0x0003, struct.pack("<H", CODEPAGE))    # PROJECTCODEPAGE
    b += _rec(0x0004, project_name.encode(ENCODING))  # PROJECTNAME
    b += _rec(0x0005, b"") + _rec(0x0040, b"")        # PROJECTDOCSTRING
    b += _rec(0x0006, b"") + _rec(0x003D, b"")        # PROJECTHELPFILEPATH
    b += _rec(0x0007, struct.pack("<I", 0))           # PROJECTHELPCONTEXT
    b += _rec(0x0008, struct.pack("<I", 0))           # PROJECTLIBFLAGS
    # PROJECTVERSION は Size が 4 なのに中身が 6 バイト（仕様どおり）
    b += struct.pack("<HIIH", 0x0009, 4, 1, 0)
    b += _rec(0x000C, b"") + _rec(0x003C, b"")        # PROJECTCONSTANTS

    # 参照設定: stdole（OLE Automation）。VBA と Excel のライブラリは暗黙に参照される
    ref_name = "stdole"
    libid = "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation"
    b += _rec(0x0016, ref_name.encode(ENCODING)) + _rec(0x003E, ref_name.encode("utf-16-le"))
    libid_bytes = libid.encode(ENCODING)
    b += _rec(0x000D, struct.pack("<I", len(libid_bytes)) + libid_bytes + struct.pack("<IH", 0, 0))

    b += _rec(0x000F, struct.pack("<H", len(modules)))  # PROJECTMODULES
    b += _rec(0x0013, struct.pack("<H", 0xFFFF))        # PROJECTCOOKIE
    for name, is_document in modules:
        mb = name.encode(ENCODING)
        mu = name.encode("utf-16-le")
        b += _rec(0x0019, mb) + _rec(0x0047, mu)         # MODULENAME / MODULENAMEUNICODE
        b += _rec(0x001A, mb) + _rec(0x0032, mu)         # MODULESTREAMNAME
        b += _rec(0x001C, b"") + _rec(0x0048, b"")       # MODULEDOCSTRING
        b += _rec(0x0031, struct.pack("<I", 0))          # MODULEOFFSET: キャッシュ無しなのでソースは先頭から
        b += _rec(0x001E, struct.pack("<I", 0))          # MODULEHELPCONTEXT
        b += _rec(0x002C, struct.pack("<H", 0xFFFF))     # MODULECOOKIE
        b += struct.pack("<HI", 0x0022 if is_document else 0x0021, 0)  # MODULETYPE
        b += struct.pack("<HI", 0x002B, 0)               # 終端
    b += struct.pack("<HI", 0x0010, 0)                   # dir の終端
    return bytes(b)


def _encrypt(project_id, data):
    """PROJECT の CMG / DPB / GC に使うデータの難読化（MS-OVBA 2.4.3.2）。"""
    seed = random.randint(0, 255)
    version = 2
    proj_key = sum(project_id.encode(ENCODING)) & 0xFF
    out = bytearray([seed, version ^ seed, proj_key ^ seed])
    unencrypted_1 = proj_key
    encrypted_1 = out[2]
    encrypted_2 = out[1]
    for _ in range((seed & 6) // 2):
        temp = random.randint(0, 255)
        enc = ((encrypted_2 + unencrypted_1) & 0xFF) ^ temp
        out.append(enc)
        encrypted_2, encrypted_1, unencrypted_1 = encrypted_1, enc, temp
    for byte in struct.pack("<I", len(data)) + data:
        enc = byte ^ ((encrypted_2 + unencrypted_1) & 0xFF)
        out.append(enc)
        encrypted_2, encrypted_1, unencrypted_1 = encrypted_1, enc, byte
    return out.hex().upper()


def _project_stream(project_id, documents, std_modules):
    cmg = _encrypt(project_id, struct.pack("<I", 0))  # 保護なし
    dpb = _encrypt(project_id, bytes([0x00]))         # パスワードなし
    gc = _encrypt(project_id, bytes([0xFF]))          # 表示する
    lines = [f'ID="{project_id}"']
    lines += [f"Document={d}/&H00000000" for d in documents]
    lines += [f"Module={m}" for m in std_modules]
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'CMG="{cmg}"',
        f'DPB="{dpb}"',
        f'GC="{gc}"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    lines += [f"{n}=0, 0, 0, 0, C" for n in documents + std_modules]
    return ("\r\n".join(lines) + "\r\n").encode(ENCODING)


def _project_wm(names):
    b = bytearray()
    for n in names:
        b += n.encode(ENCODING) + b"\x00" + n.encode("utf-16-le") + b"\x00\x00"
    b += b"\x00\x00"
    return bytes(b)


DOC_ATTRS = (
    'Attribute VB_Base = "0{{{base}}}"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = True\r\n"
    "Attribute VB_Exposed = True\r\n"
    "Attribute VB_TemplateDerived = False\r\n"
    "Attribute VB_Customizable = True\r\n"
)
WORKBOOK_BASE = "00020819-0000-0000-C000-000000000046"
SHEET_BASE = "00020820-0000-0000-C000-000000000046"


def _module_source(name, body, base=None):
    head = f'Attribute VB_Name = "{name}"\r\n'
    if base:
        head += DOC_ATTRS.format(base=base)
    text = head + body.replace("\r\n", "\n").replace("\n", "\r\n")
    if not text.endswith("\r\n"):
        text += "\r\n"
    return text.encode(ENCODING)


# ---------------------------------------------------------------------------
# MS-CFB 入れ物（バージョン 3、512 バイトのセクター）
# ---------------------------------------------------------------------------

SECTOR = 512
MINI = 64
MINI_CUTOFF = 4096
FREESECT, ENDOFCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


class _Entry:
    def __init__(self, name, kind, data=b"", children=None):
        self.name = name
        self.kind = kind  # 1 storage, 2 stream, 5 root
        self.data = data
        self.children = children or []
        self.left = self.right = self.child = NOSTREAM
        self.color = 1
        self.start = ENDOFCHAIN
        self.size = 0


def _sort_key(entry):
    return (len(entry.name), entry.name.upper())


def _build_tree(entries, ids):
    """兄弟を名前順の平衡二分木にし、赤黒木の条件を満たす色を付ける。根の番号を返す。"""
    entries = sorted(entries, key=_sort_key)
    n = len(entries)
    if n == 0:
        return NOSTREAM
    height = n.bit_length()
    full = (1 << height) - 1 == n

    def build(lo, hi, depth):
        if lo > hi:
            return NOSTREAM
        mid = (lo + hi) // 2
        e = entries[mid]
        e.left = build(lo, mid - 1, depth + 1)
        e.right = build(mid + 1, hi, depth + 1)
        e.color = 0 if (not full and depth == height - 1) else 1  # 最下段だけ赤（0）
        return ids[id(e)]

    return build(0, n - 1, 0)


def _cfb(root):
    # 平らな一覧（0 番が Root Entry）
    flat = []

    def walk(e):
        flat.append(e)
        for c in e.children:
            walk(c)

    walk(root)
    ids = {id(e): i for i, e in enumerate(flat)}
    for e in flat:
        if e.kind in (1, 5):
            e.child = _build_tree(e.children, ids)

    streams = [e for e in flat if e.kind == 2]
    small = [e for e in streams if len(e.data) < MINI_CUTOFF]
    big = [e for e in streams if len(e.data) >= MINI_CUTOFF]

    # ミニストリーム（小さいストリームを 64 バイト単位で詰めたもの）とミニ FAT
    mini_data = bytearray()
    minifat = []
    for e in small:
        e.size = len(e.data)
        if not e.data:
            e.start = ENDOFCHAIN
            continue
        count = -(-len(e.data) // MINI)
        e.start = len(mini_data) // MINI
        mini_data += e.data + b"\x00" * (count * MINI - len(e.data))
        for k in range(count):
            minifat.append(e.start + k + 1 if k < count - 1 else ENDOFCHAIN)

    def sectors(n):
        return -(-n // SECTOR)

    dir_bytes_len = len(flat) * 128
    n_dir = sectors(dir_bytes_len)
    n_minifat = sectors(len(minifat) * 4)
    n_mini = sectors(len(mini_data))
    n_big = [sectors(len(e.data)) for e in big]
    n_data = n_dir + n_minifat + n_mini + sum(n_big)
    n_fat = 1
    while n_fat * (SECTOR // 4) < n_data + n_fat:
        n_fat += 1
    assert n_fat <= 109, "DIFAT が要る大きさは想定していない"

    fat = []
    layout = {}

    def alloc(key, count):
        if count == 0:
            layout[key] = ENDOFCHAIN
            return
        first = len(fat)
        for k in range(count):
            fat.append(first + k + 1 if k < count - 1 else ENDOFCHAIN)
        layout[key] = first

    fat_first = len(fat)
    fat += [FATSECT] * n_fat
    alloc("dir", n_dir)
    alloc("minifat", n_minifat)
    alloc("mini", n_mini)
    for i, e in enumerate(big):
        alloc(("big", i), n_big[i])
        e.start = layout[("big", i)]
        e.size = len(e.data)
    fat += [FREESECT] * (n_fat * (SECTOR // 4) - len(fat))

    root.start = layout["mini"]
    root.size = len(mini_data)

    # ディレクトリ
    dir_bytes = bytearray()
    for e in flat:
        name = e.name.encode("utf-16-le") + b"\x00\x00"
        assert len(name) <= 64
        dir_bytes += name + b"\x00" * (64 - len(name))
        dir_bytes += struct.pack("<HBBIII", len(name), e.kind, e.color, e.left, e.right, e.child)
        dir_bytes += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        dir_bytes += struct.pack("<IQ", e.start if e.kind != 1 else 0, e.size if e.kind != 1 else 0)
    while len(dir_bytes) % SECTOR:
        # 空きのディレクトリ項目
        dir_bytes += b"\x00" * 64 + struct.pack("<HBBIII", 0, 0, 0, NOSTREAM, NOSTREAM, NOSTREAM)
        dir_bytes += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16 + struct.pack("<IQ", 0, 0)

    def pad(b):
        return bytes(b) + b"\x00" * (sectors(len(b)) * SECTOR - len(b))

    body = bytearray()
    body += b"".join(struct.pack("<I", v) for v in fat)
    body += dir_bytes
    body += pad(b"".join(struct.pack("<I", v) for v in minifat) + struct.pack("<I", FREESECT) * 0) if minifat else b""
    if minifat:
        # ミニ FAT の残りは FREESECT で埋める
        used = len(minifat) * 4
        body[-(n_minifat * SECTOR - used):] = struct.pack("<I", FREESECT) * ((n_minifat * SECTOR - used) // 4)
    body += pad(mini_data)
    for e in big:
        body += pad(e.data)

    difat = [fat_first + k for k in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header = bytearray()
    header += bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\x00" * 6
    header += struct.pack("<IIIIIIIII", 0, n_fat, layout["dir"], 0, MINI_CUTOFF,
                          layout["minifat"] if n_minifat else ENDOFCHAIN, n_minifat, ENDOFCHAIN, 0)
    header += b"".join(struct.pack("<I", v) for v in difat)
    assert len(header) == SECTOR
    return bytes(header + body)


# ---------------------------------------------------------------------------
# まとめ
# ---------------------------------------------------------------------------

def build_vba_project(modules, sheet_code_names, workbook_code_name="ThisWorkbook"):
    """
    modules: [(モジュール名, ソース)]。ソースは .bas の中身（Attribute VB_Name 行は無くてよい）。
    sheet_code_names: ブックのシートのコード名（Sheet1, Sheet2, …）。シートの数だけ要る。
    """
    project_id = "{" + str(uuid.uuid4()).upper() + "}"
    documents = [workbook_code_name] + list(sheet_code_names)
    std_names = [m for m, _ in modules]

    vba_children = [
        _Entry("_VBA_PROJECT", 2, b"\xcc\x61\xff\xff\x00\x00\x00"),
        _Entry("dir", 2, compress(_dir_stream("VBAProject", [(d, True) for d in documents] + [(m, False) for m in std_names]))),
    ]
    vba_children.append(_Entry(workbook_code_name, 2, compress(_module_source(workbook_code_name, "", WORKBOOK_BASE))))
    for s in sheet_code_names:
        vba_children.append(_Entry(s, 2, compress(_module_source(s, "", SHEET_BASE))))
    for name, src in modules:
        body = "\n".join(l for l in src.replace("\r\n", "\n").split("\n") if not l.startswith("Attribute VB_Name"))
        vba_children.append(_Entry(name, 2, compress(_module_source(name, body))))

    root = _Entry("Root Entry", 5, children=[
        _Entry("PROJECT", 2, _project_stream(project_id, documents, std_names)),
        _Entry("PROJECTwm", 2, _project_wm(documents + std_names)),
        _Entry("VBA", 1, children=vba_children),
    ])
    return _cfb(root)


if __name__ == "__main__":
    sample = bytes(range(256)) * 20 + b"Attribute VB_Name = \"A\"\r\n" * 50
    assert decompress(compress(sample)) == sample
    print("compress round-trip ok")
