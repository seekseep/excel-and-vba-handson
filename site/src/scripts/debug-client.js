// `::::debugdemo`（libs/debugdemo.mjs）の動き。
//   - 「次へ ▶」「◀ 前へ」「↺ 最初から」と、手順の案内の番号で、その手順のあとの状態にする
//   - 次に押すところ（ブレークポイントを置く行の帯、F5 ／ F8 ／ ■ などのボタン）を光らせ、押すと 1 手進む
//   - 状態（ブレークポイント・止まっている行・変数・ログ）はビルド時に計算して data-ddemo に入っている
// コードは動かさない。本文より先に <head> で読まれるので、クリックは document で受ける。
(function () {
  function dataOf(demo) {
    if (!demo._ddemo) demo._ddemo = JSON.parse(demo.dataset.ddemo);
    return demo._ddemo;
  }

  function el(tag, className, text) {
    var e = document.createElement(tag);
    if (className) e.className = className;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  // 止まっている行が欄の外にあれば、その行が見える位置までコードの欄だけをスクロールする。
  function scrollLineIntoView(code, line) {
    var top = line.offsetTop - code.offsetTop;
    if (top < code.scrollTop + 8) code.scrollTop = Math.max(0, top - 8);
    else if (top + line.offsetHeight > code.scrollTop + code.clientHeight - 8) code.scrollTop = top + line.offsetHeight - code.clientHeight + 8;
  }

  function renderApp(app, s, platform) {
    app.querySelectorAll('.ddemo__line').forEach(function (line) {
      var i = Number(line.dataset.line);
      line.classList.toggle('is-bp', s.bps.indexOf(i) !== -1);
      line.classList.toggle('is-current', s.cur === i);
    });
    var current = app.querySelector('.ddemo__line.is-current');
    if (current) scrollLineIntoView(app.querySelector('.ddemo__code'), current);
    app.classList.toggle('is-paused', s.paused);
    var mode = app.querySelector('.ddemo__mode');
    if (mode) mode.hidden = !s.paused;

    if (platform === 'excel') {
      var tbody = app.querySelector('.ddemo__vars tbody');
      tbody.textContent = '';
      s.vars.forEach(function (v) {
        var tr = el('tr', v[3] ? 'is-changed' : '');
        tr.appendChild(el('td', '', v[0]));
        tr.appendChild(el('td', '', v[1]));
        tr.appendChild(el('td', '', v[2]));
        tbody.appendChild(tr);
      });
      app.querySelector('.ddemo__vars').hidden = !s.paused;
      app.querySelector('.ddemo__empty').hidden = s.paused;
    } else {
      app.querySelector('.ddemo__debugger').hidden = !s.paused;
      var list = app.querySelector('.ddemo__gvars');
      list.textContent = '';
      s.vars.forEach(function (v) {
        var li = el('li', v[3] ? 'is-changed' : '');
        li.appendChild(el('span', 'ddemo__gvar-name', v[0] + ':'));
        li.appendChild(el('span', 'ddemo__gvar-value', v[1]));
        list.appendChild(li);
      });
    }

    var log = app.querySelector('.ddemo__log');
    if (log) {
      var body = log.querySelector('.ddemo__log-body');
      body.textContent = '';
      s.log.forEach(function (l) {
        body.appendChild(el('div', 'ddemo__log-line', l));
      });
      log.hidden = s.log.length === 0;
      body.scrollTop = body.scrollHeight;
    }
  }

  // 次の手順で押すところを光らせる（最後の手順のあとは何も光らせない）。
  function markNext(demo, next) {
    demo.querySelectorAll('.ddemo .is-guide').forEach(function (e) {
      e.classList.remove('is-guide');
    });
    if (!next) return;
    demo.querySelectorAll('.ddemo__app').forEach(function (app) {
      var platform = app.dataset.dplatform;
      var target = null;
      if (next.act === 'break') {
        var line = next.line[platform];
        if (line !== null) target = app.querySelector('.ddemo__gutter[data-line="' + line + '"]');
      } else {
        app.querySelectorAll('[data-act]').forEach(function (b) {
          if (!target && b.dataset.act.split(' ').indexOf(next.act) !== -1) target = b;
        });
      }
      if (target) target.classList.add('is-guide');
    });
  }

  function go(demo, k) {
    var data = dataOf(demo);
    var n = data.plan.length;
    k = Math.max(0, Math.min(n, k));
    demo.dataset.dstep = String(k);
    var state = data.states[k];
    demo.querySelectorAll('.ddemo__app').forEach(function (app) {
      renderApp(app, state[app.dataset.dplatform], app.dataset.dplatform);
    });
    markNext(demo, data.plan[k]);
    demo.querySelectorAll('.ddemo__steps li').forEach(function (li) {
      var j = Number(li.dataset.dstep);
      li.classList.toggle('is-done', j < k);
      li.classList.toggle('is-current', j === k);
    });
    demo.querySelectorAll('.ddemo__say').forEach(function (d) {
      d.hidden = Number(d.dataset.dstep) !== k;
    });
    demo.querySelector('.ddemo__prev').disabled = k === 0;
    demo.querySelector('.ddemo__next').disabled = k === n;
  }

  document.addEventListener('click', function (e) {
    var target = e.target;
    if (!(target instanceof Element)) return;
    var demo = target.closest('.ddemo');
    if (!demo) return;
    var k = Number(demo.dataset.dstep || 0);
    if (target.closest('.ddemo__next')) go(demo, k + 1);
    else if (target.closest('.ddemo__prev')) go(demo, k - 1);
    else if (target.closest('.ddemo__reset')) go(demo, 0);
    else if (target.closest('.ddemo__steps li')) go(demo, Number(target.closest('.ddemo__steps li').dataset.dstep));
    // 光っているところ（次の手順）を押したときだけ進む。ほかのボタンや行は押しても何も起きない。
    else if (target.closest('.ddemo__gutter.is-guide, [data-act].is-guide')) go(demo, k + 1);
  });

  // 本物と同じキーでも進められる（F5 ／ F8 ／ F9。コードの欄などデモの中にフォーカスがあるときだけ）。
  var KEYS = { F5: ['run', 'continue'], F8: ['step'], F9: ['break'] };
  document.addEventListener('keydown', function (e) {
    var target = e.target;
    if (!(target instanceof Element) || !KEYS[e.key]) return;
    var demo = target.closest('.ddemo');
    if (!demo) return;
    var data = dataOf(demo);
    var k = Number(demo.dataset.dstep || 0);
    var next = data.plan[k];
    if (!next || KEYS[e.key].indexOf(next.act) === -1) return;
    e.preventDefault();
    go(demo, k + 1);
  });

  function init() {
    document.querySelectorAll('.ddemo').forEach(function (demo) {
      go(demo, Number(demo.dataset.dstep || 0));
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
  document.addEventListener('astro:page-load', init);
})();
