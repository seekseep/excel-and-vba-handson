// `::::sheetdemo` / `::sheetcode`（libs/sheets.mjs）のシート表示の動き。
//   - セルを押すと、名前ボックスと数式バーにそのセルの番地と中身を出す
//   - `::::sheetdemo` の「▶ 実行」で data-after の値に切り替え、「↺ リセット」で戻す
//   - msgbox を書いたデモは、実行でメッセージの小窓を出し、「OK」で閉じる
//   - macro を付けたデモは、Excel は「マクロ」の一覧から選んで「実行」、スプレッドシートは
//     関数の選択欄で選んで「▶ 実行」したときだけ結果を出す（実際の操作の順番をなぞる）
//   - ブックに置いたボタン（.sheet__button）を押すと、そのボタンのマクロを実行したことにする
//   - スプレッドシートのメニューバーの独自メニュー（.sheet__menu）の項目も、ボタンと同じく動く
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
    sheet.querySelector('.sheet__formula').textContent =
      cell.dataset.raw !== undefined ? cell.dataset.raw : cell.textContent;
  }

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
    demo.querySelector('.sheetdemo__reset').disabled = false;
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

  // 開いた時点では何も選ばない。「実行」はいつでも押せて、選ばずに押したら赤字で知らせる。
  function openMacros(demo) {
    var picker = demo.querySelector('.sheetdemo__macros');
    picker.querySelectorAll('.sheetdemo__macro-item').forEach(function (b) {
      b.classList.remove('is-selected');
    });
    picker.querySelector('.sheetdemo__macros-error').hidden = true;
    picker.hidden = false;
    guideStep(demo, 'run');
  }

  // 選んだマクロが、このデモで再現しているマクロなら結果を出す（実行済みなら何もしない）。
  function runMacro(demo, name) {
    var expected = demo.querySelector('.sheetdemo__bar').dataset.macro;
    var reset = demo.querySelector('.sheetdemo__reset');
    if (name === expected && reset && reset.disabled) run(demo);
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
      if (demo.querySelector('.sheetdemo__reset').disabled) run(demo);
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
    demo.querySelector('.sheetdemo__reset').disabled = true;
    showDevTab(demo, false);
    guideStep(demo, 'tab');
    refresh(demo);
  }

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
    if (target.closest('.sheetdemo__dev-tab')) {
      showDevTab(demo, true);
      if (demo.querySelector('.sheetdemo__reset').disabled) guideStep(demo, 'macro');
      return;
    }
    if (target.closest('.sheetdemo__macro-open')) {
      openMacros(demo);
      return;
    }
    var item = target.closest('.sheetdemo__macro-item');
    if (item) {
      item.closest('.sheetdemo__macros-list').querySelectorAll('.sheetdemo__macro-item').forEach(function (b) {
        b.classList.toggle('is-selected', b === item);
      });
      demo.querySelector('.sheetdemo__macros-error').hidden = true;
      guideStep(demo, 'run');
      return;
    }
    if (target.closest('.sheetdemo__macros-cancel')) {
      demo.querySelector('.sheetdemo__macros').hidden = true;
      guideStep(demo, 'macro');
      return;
    }
    if (target.closest('.sheetdemo__macros-run')) {
      var selected = demo.querySelector('.sheetdemo__macro-item.is-selected');
      if (!selected) {
        demo.querySelector('.sheetdemo__macros-error').hidden = false;
        return;
      }
      demo.querySelector('.sheetdemo__macros').hidden = true;
      runMacro(demo, selected.dataset.name);
      // 違うマクロを選んだときは、もう一度「マクロ」からやり直してもらう。
      if (demo.querySelector('.sheetdemo__reset').disabled) guideStep(demo, 'macro');
      return;
    }
    if (target.closest('.sheetdemo__gas-run')) {
      runMacro(demo, demo.querySelector('.sheetdemo__gas-select').value);
      return;
    }
    var ok = target.closest('.sheetdemo__dialog-ok');
    if (ok) {
      ok.closest('.sheetdemo__dialog').hidden = true;
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
