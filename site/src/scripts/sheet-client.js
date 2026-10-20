// `::::sheetdemo` / `::sheetcode`（libs/sheets.mjs）のシート表示の動き。
//   - セルを押すと、名前ボックスと数式バーにそのセルの番地と中身を出す
//   - `::::sheetdemo` の「▶ 実行」で data-after の値に切り替え、「↺ リセット」で戻す
//   - msgbox を書いたデモは、実行でメッセージの小窓を出し、「OK」で閉じる
//   - macro を付けたデモは、Excel は「マクロ」の一覧から選んで「実行」、スプレッドシートは
//     Apps Script エディタのタブで関数を選んで「▶ 実行」したときだけ結果を出す（実際の操作の順番をなぞる）
//   - スプレッドシートは 2 つのブラウザタブ（.gbrowser）で、実行するとスプレッドシートのタブに結果が出る
//   - ブックに置いたボタン（.sheet__button）を押すと、そのボタンのマクロを実行したことにする
//   - スプレッドシートのメニューバーの独自メニュー（.sheet__menu）の項目も、ボタンと同じく動く
//   - `::::formulademo` は「次へ ▶」で式を 1 つずつ入れ、式のセルを選ぶと読んでいるセルを色の枠で囲む
// コードは動かさない。何が起きるかのイメージを見せるだけ。
// 本文より先に <head> で読まれるので、クリックは document で受ける。
(function () {
  function showCell(cell) {
    var sheet = cell.closest('.sheet');
    sheet.querySelectorAll('.sheet__cell.is-selected').forEach(function (c) {
      c.classList.remove('is-selected');
    });
    cell.classList.add('is-selected');
    sheet.querySelector('.sheet__namebox').textContent = cell.dataset.ref;
    var formula = sheet.querySelector('.sheet__formula');
    // ::::formulademo の式は、参照に色を付けたもの（ビルド時に作った HTML）を出す。
    if (cell.dataset.fhtml && cell.classList.contains('is-fset')) formula.innerHTML = cell.dataset.fhtml;
    else formula.textContent = cell.dataset.raw !== undefined ? cell.dataset.raw : cell.textContent;
    drawRefs(sheet, cell);
  }

  // ---- ::::formulademo（関数を 1 つずつ入れる） ----

  function cellAt(sheet, ref) {
    return sheet.querySelector('.sheet__grid:not([hidden]) .sheet__cell[data-ref="' + ref + '"]');
  }

  // 選んだセルの式が読んでいる範囲を、Excel の式の編集中と同じく色の枠で囲む。
  // 枠はグリッドの中に重ねるので、スクロールしてもセルと一緒に動く。
  function drawRefs(sheet, cell) {
    sheet.querySelectorAll('.fdemo__box').forEach(function (b) {
      b.remove();
    });
    sheet.querySelectorAll('.sheet__cell.is-fmatch').forEach(function (c) {
      c.classList.remove('is-fmatch');
    });
    if (!cell || !cell.dataset.frefs || !cell.classList.contains('is-fset')) return;
    var grid = cell.closest('.sheet__grid');
    var g = grid.getBoundingClientRect();
    cell.dataset.frefs.split(' ').forEach(function (item) {
      var parts = item.split('|');
      var ends = parts[0].split(':');
      var a = cellAt(sheet, ends[0]);
      var b = cellAt(sheet, ends[1] || ends[0]);
      if (!a || !b) return;
      var r1 = a.getBoundingClientRect();
      var r2 = b.getBoundingClientRect();
      var box = document.createElement('div');
      box.className = 'fdemo__box fdemo__box--' + parts[1];
      box.style.left = r1.left - g.left - grid.clientLeft + grid.scrollLeft + 'px';
      box.style.top = r1.top - g.top - grid.clientTop + grid.scrollTop + 'px';
      box.style.width = r2.right - r1.left + 'px';
      box.style.height = r2.bottom - r1.top + 'px';
      grid.appendChild(box);
    });
    // SUMIF で条件に合って足された行。
    (cell.dataset.fmatch || '').split(' ').forEach(function (ref) {
      var c = ref && cellAt(sheet, ref);
      if (c) c.classList.add('is-fmatch');
    });
  }

  // セルがグリッドの見えている範囲の外にあれば、見える位置までスクロールする（ページは動かさない）。
  function scrollCellIntoGrid(cell) {
    var grid = cell.closest('.sheet__grid');
    var g = grid.getBoundingClientRect();
    var r = cell.getBoundingClientRect();
    // 見出しの行・列（sticky）に隠れないよう、その分だけ余白を取る。
    var head = 28;
    if (r.top < g.top + head) grid.scrollTop -= g.top + head - r.top;
    else if (r.bottom > g.bottom) grid.scrollTop += r.bottom - g.bottom + 8;
    if (r.left < g.left + 42) grid.scrollLeft -= g.left + 42 - r.left;
    else if (r.right > g.right) grid.scrollLeft += r.right - g.right + 8;
  }

  // k 番目の手順まで入れた状態にする（0 は何も入れていない状態）。
  function fdemoGo(demo, k) {
    var n = Number(demo.dataset.fsteps);
    k = Math.max(0, Math.min(n, k));
    var prev = Number(demo.dataset.fstep || 0);
    demo.dataset.fstep = String(k);
    demo.querySelectorAll('.sheet__cell[data-fstep]').forEach(function (cell) {
      var s = Number(cell.dataset.fstep);
      var on = s <= k;
      cell.textContent = on ? cell.dataset.fshow : '';
      if (on) cell.dataset.raw = cell.dataset.fraw;
      else delete cell.dataset.raw;
      cell.classList.toggle('is-fset', on);
      cell.classList.remove('is-fflash');
      // 今の手順で入ったセルを一瞬光らせる（進んだときだけ）。
      if (on && s === k && k > prev) {
        void cell.offsetWidth;
        cell.classList.add('is-fflash');
      }
    });
    demo.querySelectorAll('.fdemo__steps li').forEach(function (li) {
      var j = Number(li.dataset.fstep);
      li.classList.toggle('is-done', j < k);
      li.classList.toggle('is-current', j === k);
    });
    var say = null;
    demo.querySelectorAll('.fdemo__say').forEach(function (d) {
      var on = Number(d.dataset.fstep) === k;
      d.hidden = !on;
      if (on) say = d;
    });
    demo.querySelector('.fdemo__prev').disabled = k === 0;
    demo.querySelector('.fdemo__next').disabled = k === n;
    var sheet = demo.querySelector('.sheet');
    var focus = say && say.dataset.focus ? cellAt(sheet, say.dataset.focus) : null;
    if (focus) {
      showCell(focus);
      scrollCellIntoGrid(focus);
      // スクロールで位置が変わるので、枠は描き直す。
      drawRefs(sheet, focus);
    } else {
      sheet.querySelectorAll('.sheet__cell.is-selected').forEach(function (c) {
        c.classList.remove('is-selected');
      });
      sheet.querySelector('.sheet__namebox').textContent = '';
      sheet.querySelector('.sheet__formula').textContent = '';
      drawRefs(sheet, null);
    }
  }

  // 全画面などで大きさが変わったら、選んでいるセルの枠を描き直す。
  window.addEventListener('resize', function () {
    document.querySelectorAll('.fdemo .sheet__cell.is-selected').forEach(function (cell) {
      drawRefs(cell.closest('.sheet'), cell);
    });
  });

  function run(demo) {
    demo.querySelectorAll('.sheet__cell[data-after]').forEach(function (cell) {
      cell.dataset.before = cell.textContent;
      cell.dataset.beforeRaw = cell.dataset.raw !== undefined ? cell.dataset.raw : '';
      cell.dataset.beforeNumber = cell.classList.contains('is-number') ? '1' : '';
      cell.textContent = cell.dataset.after;
      if (cell.dataset.afterRaw !== cell.dataset.after) cell.dataset.raw = cell.dataset.afterRaw;
      else delete cell.dataset.raw;
      if (cell.hasAttribute('data-after-number')) cell.classList.add('is-number');
      // リセットで元に戻すセルの印（見た目は変えない）。
      cell.classList.add('is-changed');
    });
    // ログの欄と、実行で増えるシート（最初に足したものを選んだ状態にする）。
    var log = demo.querySelector('.sheetdemo__log');
    if (log) log.hidden = false;
    var added = demo.querySelectorAll('.sheet__tab--added');
    var openAdded = demo.hasAttribute('data-open-added');
    var current = demo.querySelector('.sheet__tab.is-active');
    if (current) demo.dataset.startSheet = current.dataset.sheet;
    added.forEach(function (t) {
      t.hidden = false;
    });
    if (added.length && openAdded) showSheet(demo.querySelector('.sheet'), added[0].dataset.sheet);
    var dialog = demo.querySelector('.sheetdemo__dialog');
    if (dialog) {
      dialog.hidden = false;
      // 小窓は Excel 用とスプレッドシート用の 2 つがあり、CSS で片方を隠している。
      // 見えているほうの OK にフォーカスを当てる（隠れた要素にはフォーカスが当たらない）。
      var platform = document.documentElement.dataset.platform === 'sheets' ? 'sheets' : 'excel';
      var ok = dialog.querySelector('.sheetdemo__dialog-box--' + platform + ' .sheetdemo__dialog-ok');
      if (ok) ok.focus({ preventScroll: true });
    }
    setRunDisabled(demo, true);
    demo.dataset.ran = '';
    guideStep(demo, 'done');
    refresh(demo);
  }

  // 実行ボタンは、単純な「▶ 実行」か、マクロ選択の流れ（マクロ／関数の実行）のどちらか。
  function setRunDisabled(demo, disabled) {
    demo.querySelectorAll('.sheetdemo__run, .sheetdemo__macro-open, .sheetdemo__gas-run').forEach(function (b) {
      b.disabled = disabled;
    });
  }

  // ---- Excel の手順の案内（「開発」→「マクロ」→ 選んで「実行」） ----
  // 今の手順を案内の欄で太字にし、次に押すところを光らせる。

  var STEPS = ['tab', 'macro', 'run', 'done'];

  function guideStep(demo, step) {
    var guide = demo.querySelector('.sheetdemo__guide');
    if (!guide) return;
    var i = STEPS.indexOf(step);
    guide.querySelectorAll('li').forEach(function (li) {
      var j = STEPS.indexOf(li.dataset.step);
      li.classList.toggle('is-done', j < i);
      li.classList.toggle('is-current', j === i);
    });
    demo.querySelectorAll('.is-guide').forEach(function (el) {
      el.classList.remove('is-guide');
    });
    var target = null;
    if (step === 'tab') target = demo.querySelector('.sheetdemo__dev-tab');
    else if (step === 'macro') target = demo.querySelector('.sheetdemo__macro-open');
    else if (step === 'run') {
      var expected = demo.querySelector('.sheetdemo__bar').dataset.macro;
      var selected = demo.querySelector('.sheetdemo__macro-item.is-selected');
      target =
        selected && selected.dataset.name === expected
          ? demo.querySelector('.sheetdemo__macros-run')
          : demo.querySelector('.sheetdemo__macro-item[data-name="' + expected + '"]');
    }
    if (target) target.classList.add('is-guide');
  }

  // リボンのタブを切り替える（「開発」を開くと「マクロ」が出る）。
  function showDevTab(demo, on) {
    var dev = demo.querySelector('.sheetdemo__dev-tab');
    if (!dev) return;
    dev.classList.toggle('is-active', on);
    demo.querySelector('.sheetdemo__home-tab').classList.toggle('is-active', !on);
    demo.querySelector('.sheetdemo__home-body').hidden = on;
    demo.querySelector('.sheetdemo__dev-body').hidden = !on;
  }

  // Excel と同じく、開いた時点で一覧の先頭（名前順）を選び、「マクロ名」の欄にも入れておく。
  function openMacros(demo) {
    var picker = demo.querySelector('.sheetdemo__macros');
    var first = picker.querySelector('.sheetdemo__macro-item');
    selectMacro(demo, first);
    picker.hidden = false;
    guideStep(demo, 'run');
  }

  // 一覧の項目を選ぶ（item が null なら何も選ばない）。「マクロ名」の欄も合わせる。
  function selectMacro(demo, item, keepName) {
    demo.querySelectorAll('.sheetdemo__macro-item').forEach(function (b) {
      b.classList.toggle('is-selected', b === item);
      b.setAttribute('aria-selected', b === item ? 'true' : 'false');
    });
    var name = demo.querySelector('.sheetdemo__macros-name');
    if (name && !keepName) name.value = item ? item.dataset.name : '';
    // 一覧に無い名前のときは、Excel と同じく「実行」を押せない。
    demo.querySelector('.sheetdemo__macros-run').disabled = !item;
  }

  // 「マクロ名」の欄に打った名前と同じマクロを、一覧で選ぶ（大文字・小文字は区別しない）。
  function matchMacroName(demo) {
    var value = demo.querySelector('.sheetdemo__macros-name').value.trim().toLowerCase();
    var found = null;
    demo.querySelectorAll('.sheetdemo__macro-item').forEach(function (b) {
      if (!found && b.dataset.name.toLowerCase() === value) found = b;
    });
    selectMacro(demo, found, true);
    guideStep(demo, 'run');
  }

  function runSelectedMacro(demo) {
    var selected = demo.querySelector('.sheetdemo__macro-item.is-selected');
    if (!selected) return;
    demo.querySelector('.sheetdemo__macros').hidden = true;
    runMacro(demo, selected.dataset.name);
    // 違うマクロを選んだときは、もう一度「マクロ」からやり直してもらう。
    if (!('ran' in demo.dataset)) guideStep(demo, 'macro');
  }

  // 選んだマクロが、このデモで再現しているマクロなら結果を出す（実行済みなら何もしない）。
  function runMacro(demo, name) {
    var expected = demo.querySelector('.sheetdemo__bar').dataset.macro;
    if (name === expected && !('ran' in demo.dataset)) run(demo);
  }

  // シート上のボタンとメニューの項目。このデモで再現しているマクロなら実行し、ほかは吹き出しの文言を返す。
  function pressMacro(demo, macro) {
    var bar = demo.querySelector('.sheetdemo__bar');
    var expected = bar && bar.dataset.macro;
    if (expected && macro === expected) {
      runMacro(demo, macro);
      return null;
    }
    // macro の無いデモ（全体の結果を見せるもの）は、どのボタン・メニュー項目からでも動かす。
    if (!expected && bar && bar.hasAttribute('data-any-run')) {
      if (!('ran' in demo.dataset)) run(demo);
      return null;
    }
    if (!expected) return 'このデモは上の「▶ 実行」で動かします';
    // 見えているほう（Excel はボタン、スプレッドシートはメニューの項目）の文字で案内する。
    var sheets = document.documentElement.dataset.platform === 'sheets';
    var selector = (sheets ? '.sheet__menu-item' : '.sheet__button') + '[data-macro="' + expected + '"]';
    var target = demo.querySelector(selector);
    return 'このデモで動くのは' + (target ? '「' + target.textContent + '」' : expected) + 'です';
  }

  function pressButton(button) {
    var demo = button.closest('.sheetdemo');
    if (!demo) return;
    var text = pressMacro(demo, button.dataset.macro);
    // 表の右端に置いたボタンでも見切れないよう、ボタンと同じ幅で折り返す。
    if (text) showHint(button.parentElement, text, button.style.width);
  }

  function pressMenuItem(item) {
    var menu = item.closest('.sheet__menu');
    closeMenu(menu, true);
    var demo = item.closest('.sheetdemo');
    if (!demo) return;
    var text = pressMacro(demo, item.dataset.macro);
    // メニューは画面の右端に寄りやすいので、吹き出しはメニューバーの右下に出す（見切れないように）。
    if (text) showHint(menu.closest('.sheet__menubar'), text);
  }

  function showHint(host, text, width) {
    var old = host.querySelector('.sheet__button-hint');
    if (old) old.remove();
    var hint = document.createElement('span');
    hint.className = 'sheet__button-hint';
    hint.setAttribute('role', 'status');
    hint.textContent = text;
    if (width) hint.style.width = width;
    host.appendChild(hint);
    // 下に並んだボタンより手前に出す。
    host.classList.add('is-hinting');
    setTimeout(function () {
      hint.remove();
      if (!host.querySelector('.sheet__button-hint')) host.classList.remove('is-hinting');
    }, 2500);
  }

  // ---- スプレッドシートの独自メニュー（onOpen で足すメニュー） ----

  function openMenu(menu) {
    var old = menu.closest('.sheet__menubar').querySelector('.sheet__button-hint');
    if (old) old.remove();
    menu.classList.add('is-open');
    menu.querySelector('.sheet__menu-open').setAttribute('aria-expanded', 'true');
    menu.querySelector('.sheet__menu-list').hidden = false;
  }

  function closeMenu(menu, focusOpener) {
    var opener = menu.querySelector('.sheet__menu-open');
    menu.classList.remove('is-open');
    opener.setAttribute('aria-expanded', 'false');
    menu.querySelector('.sheet__menu-list').hidden = true;
    if (focusOpener) opener.focus({ preventScroll: true });
  }

  function reset(demo) {
    var dialog = demo.querySelector('.sheetdemo__dialog');
    if (dialog) dialog.hidden = true;
    var log = demo.querySelector('.sheetdemo__log');
    if (log) log.hidden = true;
    if (demo.dataset.startSheet) showSheet(demo.querySelector('.sheet'), demo.dataset.startSheet);
    demo.querySelectorAll('.sheet__tab--added').forEach(function (t) {
      t.hidden = true;
    });
    demo.querySelectorAll('.sheet__cell.is-changed').forEach(function (cell) {
      cell.textContent = cell.dataset.before;
      if (cell.dataset.beforeRaw) cell.dataset.raw = cell.dataset.beforeRaw;
      else delete cell.dataset.raw;
      if (!cell.dataset.beforeNumber) cell.classList.remove('is-number');
      cell.classList.remove('is-changed');
    });
    var picker = demo.querySelector('.sheetdemo__macros');
    if (picker) picker.hidden = true;
    setRunDisabled(demo, false);
    delete demo.dataset.ran;
    showDevTab(demo, false);
    guideStep(demo, 'tab');
    gasReset(demo);
    refresh(demo);
  }

  // ---- スプレッドシートの 2 つのブラウザタブ（スプレッドシート ／ Apps Script エディタ） ----
  // エディタのタブで関数を選んで「▶ 実行」→ スプレッドシートのタブで結果を見る、の順に案内する。
  // Excel の案内（is-guide）とは別の印（is-gguide）で光らせる。guideStep が is-guide を全部消すため。

  var GAS_STEPS = ['select', 'run', 'look', 'done'];

  function gasStep(demo, step) {
    var guide = demo.querySelector('.sheetdemo__guide--sheets');
    if (!guide) return;
    demo.dataset.gasStep = step;
    var i = GAS_STEPS.indexOf(step);
    guide.querySelectorAll('li').forEach(function (li) {
      var j = GAS_STEPS.indexOf(li.dataset.step);
      li.classList.toggle('is-done', j < i);
      li.classList.toggle('is-current', j === i);
    });
    demo.querySelectorAll('.is-gguide').forEach(function (el) {
      el.classList.remove('is-gguide');
    });
    var target = null;
    if (step === 'select') target = demo.querySelector('.sheetdemo__gas-select');
    else if (step === 'run') target = demo.querySelector('.sheetdemo__gas-run');
    else if (step === 'look') target = demo.querySelector('.gbrowser__tab[data-pane="sheet"]');
    if (target) target.classList.add('is-gguide');
  }

  // 選択欄が、このデモで再現している関数を指していれば「実行」へ、違えば「選ぶ」へ。
  function gasSyncSelect(demo) {
    var step = demo.dataset.gasStep;
    if (step !== 'select' && step !== 'run') return;
    var bar = demo.querySelector('.sheetdemo__bar');
    var expected = bar && bar.dataset.macro ? gasNameOf(bar.dataset.macro) : '';
    var select = demo.querySelector('.sheetdemo__gas-select');
    gasStep(demo, select.value === expected ? 'run' : 'select');
  }

  function gasNameOf(vbaName) {
    return vbaName.charAt(0).toLowerCase() + vbaName.slice(1);
  }

  function showPane(demo, name) {
    var browser = demo.querySelector('.gbrowser');
    if (!browser) return;
    browser.dataset.pane = name;
    browser.querySelectorAll('.gbrowser__tab').forEach(function (t) {
      var on = t.dataset.pane === name;
      t.classList.toggle('is-active', on);
      t.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
    if (name !== 'sheet') return;
    if (demo.dataset.gasStep === 'look') gasStep(demo, 'done');
    // 隠れていた alert の OK には、タブを開いたときにフォーカスを当てる。
    var ok = demo.querySelector('.sheetdemo__dialog:not([hidden]) .sheetdemo__dialog-box--sheets .sheetdemo__dialog-ok');
    if (ok) ok.focus({ preventScroll: true });
  }

  function gasTime() {
    var d = new Date();
    return [d.getHours(), d.getMinutes(), d.getSeconds()]
      .map(function (n) {
        return (n < 10 ? '0' : '') + n;
      })
      .join(':');
  }

  // 実行ログに「お知らせ」の行（実行開始／実行完了）を足す。first なら情報の行より前に置く。
  function gasNotice(demo, text, first) {
    var body = demo.querySelector('.gbrowser__log-body');
    var row = document.createElement('div');
    row.className = 'gbrowser__log-row gbrowser__log-row--notice';
    [gasTime(), 'お知らせ', text].forEach(function (t, i) {
      var span = document.createElement('span');
      span.className = ['gbrowser__log-time', 'gbrowser__log-kind', 'gbrowser__log-text'][i];
      span.textContent = t;
      row.appendChild(span);
    });
    if (first) body.insertBefore(row, body.firstChild);
    else body.appendChild(row);
    body.scrollTop = body.scrollHeight;
  }

  function gasRun(demo) {
    var log = demo.querySelector('.gbrowser__log');
    var select = demo.querySelector('.sheetdemo__gas-select');
    var bar = demo.querySelector('.sheetdemo__bar');
    if (!log) {
      runMacro(demo, select.value);
      return;
    }
    gasClearLog(demo);
    log.hidden = false;
    gasNotice(demo, '実行開始', true);
    var expected = bar && bar.dataset.macro;
    // 違う関数を選んだときは、何も変わらずに終わる（このデモで再現しているのは 1 つだけ）。
    if (!expected || select.value !== gasNameOf(expected)) {
      gasNotice(demo, '実行完了');
      return;
    }
    run(demo);
    var time = gasTime();
    demo.querySelectorAll('.gbrowser__log-row--info').forEach(function (row) {
      row.querySelector('.gbrowser__log-time').textContent = time;
      row.hidden = false;
    });
    // alert を出す関数は、OK を押すまで終わらない（本物も「実行開始」のまま待つ）。
    if (demo.querySelector('.sheetdemo__dialog')) demo.dataset.gasWaiting = '1';
    else gasNotice(demo, '実行完了');
    // ログだけを出す関数は、結果がエディタのタブにあるのでタブを切り替える案内は出さない。
    var onSheet = demo.querySelector('.sheetdemo__dialog') || demo.querySelector('.sheet__cell.is-changed') || demo.querySelector('.sheet__tab--added');
    demo.querySelector('.gbrowser__notice').hidden = !onSheet;
    gasStep(demo, onSheet ? 'look' : 'done');
  }

  function gasFinish(demo) {
    delete demo.dataset.gasWaiting;
    gasNotice(demo, '実行完了');
  }

  function gasClearLog(demo) {
    demo.querySelectorAll('.gbrowser__log-row--notice').forEach(function (row) {
      row.remove();
    });
    demo.querySelectorAll('.gbrowser__log-row--info').forEach(function (row) {
      row.hidden = true;
    });
  }

  function gasReset(demo) {
    if (!demo.querySelector('.gbrowser')) return;
    gasClearLog(demo);
    delete demo.dataset.gasWaiting;
    demo.querySelector('.gbrowser__log').hidden = true;
    demo.querySelector('.gbrowser__notice').hidden = true;
    showPane(demo, 'editor');
    demo.dataset.gasStep = 'select';
    gasSyncSelect(demo);
  }

  // 関数を選ぶと、その関数の行までコードを送る（本物は送らないが、どこを動かすのかを見せるため）。
  function gasScrollTo(demo) {
    var select = demo.querySelector('.sheetdemo__gas-select');
    var code = demo.querySelector('.gbrowser__code');
    if (!select || !code) return;
    var line = code.querySelector('.gbrowser__code-line[data-function="' + select.value + '"]');
    code.querySelectorAll('.gbrowser__code-line.is-current').forEach(function (l) {
      l.classList.remove('is-current');
    });
    if (!line) return;
    line.classList.add('is-current');
    code.scrollTop = Math.max(0, line.offsetTop - code.offsetTop - 8);
  }

  document.addEventListener('change', function (e) {
    var target = e.target;
    if (!(target instanceof Element) || !target.classList.contains('sheetdemo__gas-select')) return;
    var demo = target.closest('.sheetdemo');
    gasScrollTo(demo);
    gasSyncSelect(demo);
  });

  // ページを開いたときの案内（先頭の関数がデモの関数なら、もう「実行」から）。
  function initGas() {
    document.querySelectorAll('.sheetdemo .gbrowser').forEach(function (browser) {
      var demo = browser.closest('.sheetdemo');
      if (demo.dataset.gasStep) return;
      demo.dataset.gasStep = 'select';
      gasSyncSelect(demo);
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initGas);
  else initGas();
  // Astro のページ遷移（View Transitions）を使っていても、新しいページで案内を付け直す。
  document.addEventListener('astro:page-load', initGas);

  // 選んでいるセルがあれば、数式バーを今の中身に合わせる。
  function refresh(demo) {
    var selected = demo.querySelector('.sheet__cell.is-selected');
    if (selected) showCell(selected);
  }

  document.addEventListener('click', function (e) {
    var target = e.target;
    if (!(target instanceof Element)) return;
    // ページ内のクリックをすべて受けるので、シートの外ならすぐ抜ける。
    if (!target.closest('.sheet, .sheetdemo')) return;
    var opener = target.closest('.sheet__menu-open');
    if (opener) {
      var menu = opener.closest('.sheet__menu');
      if (menu.classList.contains('is-open')) closeMenu(menu);
      else openMenu(menu);
      return;
    }
    var menuItem = target.closest('.sheet__menu-item');
    if (menuItem) {
      pressMenuItem(menuItem);
      return;
    }
    var button = target.closest('.sheet__button');
    if (button) {
      pressButton(button);
      return;
    }
    var cell = target.closest('.sheet__cell');
    if (cell) {
      showCell(cell);
      return;
    }
    var demo = target.closest('.sheetdemo');
    if (demo && demo.classList.contains('fdemo')) {
      var k = Number(demo.dataset.fstep || 0);
      var step = target.closest('.fdemo__steps li');
      if (target.closest('.fdemo__next')) fdemoGo(demo, k + 1);
      else if (target.closest('.fdemo__prev')) fdemoGo(demo, k - 1);
      else if (target.closest('.sheetdemo__reset')) fdemoGo(demo, 0);
      else if (step) fdemoGo(demo, Number(step.dataset.fstep));
      return;
    }
    if (target.closest('.sheetdemo__dev-tab')) {
      showDevTab(demo, true);
      if (!('ran' in demo.dataset)) guideStep(demo, 'macro');
      return;
    }
    if (target.closest('.sheetdemo__macro-open')) {
      openMacros(demo);
      return;
    }
    var item = target.closest('.sheetdemo__macro-item');
    if (item) {
      selectMacro(demo, item);
      guideStep(demo, 'run');
      return;
    }
    if (target.closest('.sheetdemo__macros-cancel')) {
      demo.querySelector('.sheetdemo__macros').hidden = true;
      guideStep(demo, 'macro');
      return;
    }
    if (target.closest('.sheetdemo__macros-run')) {
      runSelectedMacro(demo);
      return;
    }
    if (target.closest('.sheetdemo__gas-run')) {
      gasRun(demo);
      return;
    }
    var pane = target.closest('.gbrowser__tab');
    if (pane) {
      showPane(demo, pane.dataset.pane);
      return;
    }
    var ok = target.closest('.sheetdemo__dialog-ok');
    if (ok) {
      ok.closest('.sheetdemo__dialog').hidden = true;
      // alert を閉じると、エディタで待っていた実行が終わる。
      if (demo.dataset.gasWaiting) gasFinish(demo);
      return;
    }
    if (target.closest('.sheetdemo__run')) run(demo);
    else if (target.closest('.sheetdemo__reset')) reset(demo);
  });
  // 開いているメニューは、その外側を押すか Esc で閉じる。
  document.addEventListener('click', function (e) {
    var target = e.target;
    document.querySelectorAll('.sheet__menu.is-open').forEach(function (menu) {
      if (!(target instanceof Node) || !menu.contains(target)) closeMenu(menu);
    });
  });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    document.querySelectorAll('.sheet__menu.is-open').forEach(function (menu) {
      closeMenu(menu, true);
    });
  });

  // 「マクロ名」の欄に名前を打つと、一覧の同じ名前のマクロを選ぶ。Enter で実行、Esc で閉じる（Excel と同じ）。
  document.addEventListener('input', function (e) {
    var target = e.target;
    if (!(target instanceof Element) || !target.classList.contains('sheetdemo__macros-name')) return;
    matchMacroName(target.closest('.sheetdemo'));
  });
  document.addEventListener('keydown', function (e) {
    var target = e.target;
    if (!(target instanceof Element) || !target.closest('.sheetdemo__macros')) return;
    var demo = target.closest('.sheetdemo');
    if (e.key === 'Enter' && !target.closest('button')) {
      e.preventDefault();
      runSelectedMacro(demo);
    } else if (e.key === 'Escape') {
      demo.querySelector('.sheetdemo__macros').hidden = true;
      guideStep(demo, 'macro');
    }
  });

  // Excel と同じく、一覧のマクロ名をダブルクリックしても実行できる。
  document.addEventListener('dblclick', function (e) {
    var target = e.target;
    if (!(target instanceof Element)) return;
    var item = target.closest('.sheetdemo__macro-item');
    if (!item) return;
    var demo = item.closest('.sheetdemo');
    demo.querySelector('.sheetdemo__macros').hidden = true;
    runMacro(demo, item.dataset.name);
  });
  // シートの見出しを押したら、そのシートを出す（ブックの全シートは描いてあり、選んだもの以外は hidden）。
  function showSheet(sheet, name) {
    if (!sheet) return;
    sheet.querySelectorAll('.sheet__grid[data-sheet]').forEach(function (g) {
      g.hidden = g.dataset.sheet !== name;
    });
    sheet.querySelectorAll('.sheet__tab[data-sheet]').forEach(function (t) {
      t.classList.toggle('is-active', t.dataset.sheet === name);
    });
    sheet.querySelectorAll('.sheet__cell.is-selected').forEach(function (c) {
      c.classList.remove('is-selected');
    });
    sheet.querySelector('.sheet__namebox').textContent = '';
    sheet.querySelector('.sheet__formula').textContent = '';
  }

  document.addEventListener('click', function (e) {
    var target = e.target;
    if (!(target instanceof Element)) return;
    var tab = target.closest('.sheet__tab[data-sheet]');
    if (tab) {
      showSheet(tab.closest('.sheet'), tab.dataset.sheet);
      return;
    }
    var full = target.closest('.sheetdemo__fullscreen');
    if (full) toggleFullscreen(full.closest('.sheetdemo'), full);
  });

  // 下端の帯をドラッグして、シートの表示を縦に広げる（全シート共通の高さにする）。
  document.addEventListener('pointerdown', function (e) {
    var target = e.target;
    if (!(target instanceof Element) || !target.classList.contains('sheet__resizer')) return;
    var sheet = target.closest('.sheet');
    var visible = sheet.querySelector('.sheet__grid:not([hidden])');
    if (!visible) return;
    e.preventDefault();
    var startY = e.clientY;
    var startH = visible.getBoundingClientRect().height;
    target.setPointerCapture(e.pointerId);
    function move(ev) {
      var h = Math.max(120, startH + ev.clientY - startY);
      sheet.style.setProperty('--sheet-grid-height', h + 'px');
      sheet.classList.add('is-resized');
    }
    function up() {
      target.removeEventListener('pointermove', move);
      target.removeEventListener('pointerup', up);
    }
    target.addEventListener('pointermove', move);
    target.addEventListener('pointerup', up);
  });

  // 全画面。Fullscreen API が使えないブラウザ（iPhone の Safari など）は、画面に固定する表示で代わりにする。
  function toggleFullscreen(demo, button) {
    var on = !(document.fullscreenElement === demo || demo.classList.contains('is-fullscreen'));
    if (on && demo.requestFullscreen) {
      demo.requestFullscreen().catch(function () {
        demo.classList.add('is-fullscreen');
      });
    } else if (!on && document.fullscreenElement === demo) {
      document.exitFullscreen();
    } else {
      demo.classList.toggle('is-fullscreen', on);
    }
    button.setAttribute('aria-pressed', on ? 'true' : 'false');
    button.textContent = on ? '✕ 全画面を終了' : '⛶ 全画面';
  }

  document.addEventListener('fullscreenchange', function () {
    document.querySelectorAll('.sheetdemo__fullscreen[aria-pressed="true"]').forEach(function (b) {
      var demo = b.closest('.sheetdemo');
      if (document.fullscreenElement !== demo && !demo.classList.contains('is-fullscreen')) {
        b.setAttribute('aria-pressed', 'false');
        b.textContent = '⛶ 全画面';
      }
    });
  });
})();
