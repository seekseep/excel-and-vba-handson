// `::::sheetdemo` / `::sheetcode`（libs/sheets.mjs）のシート表示の動き。
//   - セルを押すと、名前ボックスと数式バーにそのセルの番地と中身を出す
//   - `::::sheetdemo` の「▶ 実行」で data-after の値に切り替え、「↺ リセット」で戻す
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
    demo.querySelectorAll('.sheet__cell[data-after]').forEach(function (cell, i) {
      cell.dataset.before = cell.textContent;
      cell.dataset.beforeRaw = cell.dataset.raw !== undefined ? cell.dataset.raw : '';
      cell.dataset.beforeNumber = cell.classList.contains('is-number') ? '1' : '';
      cell.textContent = cell.dataset.after;
      if (cell.dataset.afterRaw !== cell.dataset.after) cell.dataset.raw = cell.dataset.afterRaw;
      else delete cell.dataset.raw;
      if (cell.hasAttribute('data-after-number')) cell.classList.add('is-number');
      // 上から順に光らせて、書き込まれた順番が分かるようにする。
      cell.style.animationDelay = Math.min(i * 40, 1200) + 'ms';
      cell.classList.add('is-changed');
    });
    demo.querySelector('.sheetdemo__run').disabled = true;
    demo.querySelector('.sheetdemo__reset').disabled = false;
    refresh(demo);
  }

  function reset(demo) {
    demo.querySelectorAll('.sheet__cell.is-changed').forEach(function (cell) {
      cell.textContent = cell.dataset.before;
      if (cell.dataset.beforeRaw) cell.dataset.raw = cell.dataset.beforeRaw;
      else delete cell.dataset.raw;
      if (!cell.dataset.beforeNumber) cell.classList.remove('is-number');
      cell.classList.remove('is-changed');
      cell.style.animationDelay = '';
    });
    demo.querySelector('.sheetdemo__run').disabled = false;
    demo.querySelector('.sheetdemo__reset').disabled = true;
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
    var cell = target.closest('.sheet__cell');
    if (cell) {
      showCell(cell);
      return;
    }
    var button = target.closest('.sheetdemo__run, .sheetdemo__reset');
    if (!button) return;
    var demo = button.closest('.sheetdemo');
    if (button.classList.contains('sheetdemo__run')) run(demo);
    else reset(demo);
  });
})();
