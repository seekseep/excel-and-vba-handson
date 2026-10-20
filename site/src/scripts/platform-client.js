// `::::platform`（remark-platform.mjs）の Excel ／ スプレッドシート切り替え。
// 選んだ環境は <html data-platform="..."> に持ち、表示の出し分けは platform.css に任せる。
// <head> で読み込まれるので、属性だけは本文の描画前に付けてちらつきを防ぐ。
(function () {
  var KEY = 'handson-platform';
  var DIR = { excel: 'vba/', sheets: 'gas/' };

  function load() {
    try {
      return localStorage.getItem(KEY) === 'sheets' ? 'sheets' : 'excel';
    } catch (e) {
      return 'excel';
    }
  }

  // `::codeview` のパネルも、選んだ環境のファイルを開いた状態にそろえる。
  // vba/SalesSummary.bas を開いていたら gas/SalesSummary.gs のように、同じ名前の相方を優先する。
  function syncEditors(platform) {
    document.querySelectorAll('.editor').forEach(function (editor) {
      var active = editor.querySelector('.editor__file.is-active');
      if (active && active.dataset.file.indexOf(DIR[platform]) === 0) return;
      var files = Array.prototype.slice.call(editor.querySelectorAll('.editor__file'));
      var inDir = files.filter(function (btn) {
        return btn.dataset.file.indexOf(DIR[platform]) === 0;
      });
      if (inDir.length === 0) return;
      var base = active ? baseName(active.dataset.file) : '';
      var pair = inDir.filter(function (btn) {
        return baseName(btn.dataset.file) === base;
      })[0];
      (pair || inDir[0]).click();
    });
  }

  function baseName(file) {
    return file.split('/').pop().replace(/\.[^.]+$/, '');
  }

  function apply(platform, save) {
    document.documentElement.dataset.platform = platform;
    if (save) {
      try {
        localStorage.setItem(KEY, platform);
      } catch (e) {
        // 保存できなくてもこのページ内の切り替えは効く。
      }
    }
    syncEditors(platform);
  }

  document.documentElement.dataset.platform = load();

  function setup() {
    document.querySelectorAll('.platform__tab').forEach(function (tab) {
      tab.addEventListener('click', function () {
        // 押したタブの位置が画面上で動かないよう、切り替え前後のずれを補正する。
        var before = tab.getBoundingClientRect().top;
        apply(tab.dataset.platform, true);
        window.scrollBy(0, tab.getBoundingClientRect().top - before);
      });
    });
    apply(load(), false);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', setup);
  } else {
    setup();
  }
})();
