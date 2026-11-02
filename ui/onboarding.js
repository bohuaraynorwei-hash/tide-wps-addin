/**
 * Tide onboarding — a five-step chat bubble walkthrough.
 *
 * Deliberately decoupled from tide.js: it drives the pane the same way a user
 * would (clicking tabs), so a refactor inside the app cannot break the tour.
 * The bubble is positioned from the target's rect and clamped to the viewport;
 * the target gets a ring so the eye lands on the right control.
 */
(function () {
  'use strict';

  var SEEN_KEY = 'tide.tourSeen';
  var VERSION = '1';   // bump to re-show the tour after a big UI change

  var STEPS = [
    {
      target: '.brand',
      place: 'bottom',
      text: '欢迎使用 Tide。跟着气泡走一遍，30 秒摸清关键操作；随时可以点右上角的 ? 重看。'
    },
    {
      before: 'matches',
      target: '#btnScan',
      place: 'bottom',
      text: '第一步是「扫描当前表」：列出所有含术语的行，点任意一行会跳回原文单元格。'
    },
    {
      before: 'translate',
      target: 'button[data-tab="translate"]',
      place: 'bottom',
      text: '「翻译」页逐条处理。点「写入并下一条」时，旧译文会自动存进 EN (old)，可随时回退。'
    },
    {
      target: '#entryDraft',
      place: 'top',
      text: '译文写在这里。表格里已有文字，用「读取选区→原文」拉进来，不用手动复制。'
    },
    {
      before: 'review',
      target: 'button[data-tab="review"]',
      place: 'bottom',
      text: '「审校」页运行质检：占位符、数字、术语、长度都会查一遍，点问题行直接跳到对应单元格。'
    },
    {
      before: 'settings',
      target: '#colSource',
      place: 'top',
      text: '列映射在设置里。表头写了 Source Text 与 EN 时，Tide 会自动识别，不需要手填。'
    },
    {
      close: true,
      text: '就这些。术语表在设置里导入，记忆库支持 TMX；点右上角 ? 可以再看一次。'
    }
  ];

  var layer = null;
  var bubble = null;
  var ring = null;
  var index = 0;

  function el(tag, cls, text) {
    var node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text) node.textContent = text;
    return node;
  }

  function showTab(name) {
    var button = document.querySelector('button[data-tab="' + name + '"]');
    if (button && !button.classList.contains('on')) button.click();
  }

  function build() {
    if (layer) return;
    layer = el('div', 'tour-layer');
    ring = el('div', 'tour-ring');
    bubble = el('div', 'tour-bubble');

    var head = el('div', 'tour-head');
    var avatar = el('div', 'tour-avatar');
    avatar.innerHTML = '<svg viewBox="0 0 24 24"><path d="M7 8.5h10" stroke="#fff" stroke-width="2.4" stroke-linecap="round"/><path d="M12 9.6v6.6" stroke="#fff" stroke-width="2.4" stroke-linecap="round"/></svg>';
    head.appendChild(avatar);
    head.appendChild(el('b', '', 'Tide 引导'));
    head.appendChild(el('span', 'tour-count', ''));
    bubble.appendChild(head);

    bubble.appendChild(el('p', 'tour-text', ''));

    var foot = el('div', 'tour-foot');
    var skip = el('button', 'skip', '跳过');
    skip.addEventListener('click', function () { finish(); });
    foot.appendChild(skip);
    foot.appendChild(el('div', 'tour-dots', ''));

    var prev = el('button', 'tour-btn ghost', '上一步');
    prev.addEventListener('click', function () { go(index - 1); });
    foot.appendChild(prev);

    var next = el('button', 'tour-btn primary', '下一步');
    next.addEventListener('click', function () {
      if (index >= STEPS.length - 1) finish();
      else go(index + 1);
    });
    foot.appendChild(next);
    bubble.appendChild(foot);

    bubble.appendChild(el('div', 'tail'));
    layer.appendChild(ring);
    layer.appendChild(bubble);
    document.body.appendChild(layer);

    document.addEventListener('keydown', function (event) {
      if (!layer || layer.hidden) return;
      if (event.key === 'Escape') finish();
      if (event.key === 'ArrowRight') next.click();
      if (event.key === 'ArrowLeft' && index > 0) go(index - 1);
    });
    window.addEventListener('resize', function () { if (layer && !layer.hidden) place(STEPS[index]); });
  }

  function place(step) {
    var target = step.target ? document.querySelector(step.target) : null;
    var vw = window.innerWidth;
    var vh = window.innerHeight;
    var bw = 266;
    var bh = bubble.offsetHeight || 150;

    if (!target) {
      ring.style.display = 'none';
      bubble.style.left = Math.max(12, (vw - bw) / 2) + 'px';
      bubble.style.top = Math.max(24, (vh - bh) / 2) + 'px';
      var tail = bubble.querySelector('.tail');
      if (tail) tail.style.display = 'none';
      return;
    }

    var rect = target.getBoundingClientRect();
    ring.style.display = 'block';
    ring.style.left = (rect.left - 4) + 'px';
    ring.style.top = (rect.top - 4) + 'px';
    ring.style.width = (rect.width + 8) + 'px';
    ring.style.height = (rect.height + 8) + 'px';

    var left = Math.min(Math.max(8, rect.left + rect.width / 2 - bw / 2), vw - bw - 8);
    var below = rect.bottom + 12;
    var above = rect.top - bh - 12;
    var top = step.place === 'top' && above > 8 ? above : (below + bh < vh ? below : Math.max(8, above));

    bubble.style.left = left + 'px';
    bubble.style.top = top + 'px';

    var tail = bubble.querySelector('.tail');
    if (tail) {
      tail.style.display = 'block';
      var tailX = Math.min(Math.max(16, rect.left + rect.width / 2 - left - 6), bw - 24);
      tail.style.left = tailX + 'px';
      if (top > rect.top) { tail.style.top = '-5px'; tail.style.bottom = 'auto'; }
      else { tail.style.top = 'auto'; tail.style.bottom = '-5px'; }
    }
  }

  function render() {
    var step = STEPS[index];
    bubble.querySelector('.tour-text').textContent = step.text;
    bubble.querySelector('.tour-count').textContent = step.close ? '' : (index + 1) + ' / ' + (STEPS.length - 1);

    var dots = bubble.querySelector('.tour-dots');
    dots.textContent = '';
    for (var i = 0; i < STEPS.length - 1; i += 1) {
      var dot = el('i', i === index ? 'on' : '');
      dots.appendChild(dot);
    }

    var buttons = bubble.querySelectorAll('.tour-btn');
    buttons[0].style.visibility = index === 0 ? 'hidden' : 'visible';
    buttons[1].textContent = index === STEPS.length - 1 ? '完成' : '下一步';

    place(step);
    requestAnimationFrame(function () { place(step); });
  }

  function go(next) {
    if (!layer) return;
    if (next < 0 || next >= STEPS.length) return;
    index = next;
    var step = STEPS[index];
    if (step.before) showTab(step.before);
    // Let the tab switch settle before measuring the new target.
    window.setTimeout(render, 60);
  }

  function finish() {
    if (!layer) return;
    layer.style.display = 'none';
    try { window.localStorage.setItem(SEEN_KEY, VERSION); } catch (err) { /* storage may be blocked */ }
  }

  function start(force) {
    build();
    index = 0;
    layer.style.display = 'block';
    if (STEPS[0].before) showTab(STEPS[0].before);
    window.setTimeout(render, 80);
    if (force) { /* replay is explicit, nothing else to do */ }
  }

  function boot() {
    var open = document.getElementById('btnTour');
    if (open) open.addEventListener('click', function () { start(true); });
    var seen = null;
    try { seen = window.localStorage.getItem(SEEN_KEY); } catch (err) { seen = null; }
    if (seen !== VERSION) window.setTimeout(function () { start(false); }, 900);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
