#!/usr/bin/env python3
"""
完成版のサンプルのスプレッドシート（sections/spreadsheets.json の done）に、その節の完成コード
（example/gas/*.gs）を Apps Script として入れる。受講者が「コピーを作成」すると、Apps Script も一緒にコピーされる。

  - 初回は clasp でスプレッドシートにスクリプトを作り、そのスクリプト ID を spreadsheets.json の script に書く
  - 2 回目からは同じスクリプトに push し直す（example/gas/ を直したら、もう一度実行する）

使い方（保守者のみ）:
    npx @google/clasp login          # 最初の 1 回だけ。done のスプレッドシートを編集できるアカウントで
    python3 tools/push_gas.py        # 全節
    python3 tools/push_gas.py 05-invoice-button/03-pdf-and-button   # 節を絞る

clasp を使うには https://script.google.com/home/usersettings で「Google Apps Script API」をオンにしておく。
"""

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SECTIONS = ROOT / "sections"
TABLE = SECTIONS / "spreadsheets.json"

MANIFEST = {
    "timeZone": "Asia/Tokyo",
    "dependencies": {},
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
}


def clasp(*args, cwd):
    subprocess.run(["npx", "-y", "@google/clasp", *args], cwd=cwd, check=True)


def push(key, entry):
    gas_files = sorted((SECTIONS / key / "example" / "gas").glob("*.gs"))
    if not gas_files:
        print(f"[push_gas] {key}: example/gas/*.gs が無いので飛ばす")
        return
    title = key.split("/")[-1]
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        if entry.get("script"):
            (work / ".clasp.json").write_text(json.dumps({"scriptId": entry["script"], "rootDir": "."}))
        else:
            # --type を渡すと clasp は --parentId を無視して新しいスプレッドシートを作るので、渡さない
            clasp("create", "--title", title, "--parentId", entry["done"], "--rootDir", ".", cwd=work)
            entry["script"] = json.loads((work / ".clasp.json").read_text())["scriptId"]
        # clasp create が作ったファイルは使わず、完成コードと manifest だけにする
        for f in work.iterdir():
            if f.name != ".clasp.json":
                f.unlink() if f.is_file() else shutil.rmtree(f)
        for f in gas_files:
            shutil.copy(f, work / f.name)
        (work / "appsscript.json").write_text(json.dumps(MANIFEST, indent=2) + "\n")
        clasp("push", "--force", cwd=work)
    print(f"[push_gas] {key}: {', '.join(f.name for f in gas_files)} -> {entry['script']}")


def main():
    table = json.loads(TABLE.read_text())
    keys = sys.argv[1:] or [k for k, v in table.items() if v.get("done")]
    try:
        for key in keys:
            push(key, table[key])
    finally:
        # 途中で失敗しても、作ったスクリプトの ID は残す（次に同じスプレッドシートへもう 1 つ作らないように）
        TABLE.write_text(json.dumps(table, ensure_ascii=False, indent=2) + "\n")


if __name__ == "__main__":
    main()
