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
  var VERSION = '2';   // bump to re-show the tour after a big UI change

  // One step per visible tab: the guide follows the tab strip, so every
  // function is introduced in the place it lives.
  var TAB_TIPS = [
    ['matches', '命中', '扫描当前表，列出所有含术语的行。点任意一行，会跳回它所在的原文单元格。'],
    ['translate', '翻译', '逐条处理的主场。「写入并下一条」会把旧译文自动存进 EN (old)，之后可一键回退。'],
    ['keys', 'Key 定位', '按表头里的「完整Key」找词条，支持精确与模糊匹配，点结果直接跳转。'],
    ['find', '搜索', '跨工作表搜原文或译文，支持正则；结果同样可以点击定位。'],
    ['review', '审校', '修订的接受/拒绝、条目评论，以及质检与 CSV 导出都在这里。'],
    ['tm', '记忆库', '导入 TMX 或 Excel 作为记忆，看复用率，并把当前表已有的译文合并进来。'],
    ['grid', '网格', '表格式的逐条编辑，可按「未译 / 有 QA / 有修订」筛选，就地改译文。'],
    ['project', '项目', '一眼看完整本工作簿：每张表的条目数、已译比例、复用率与问题数。'],
    ['terms', '术语', '术语表检索与采用。只显示当前目标语言的术语，切换语言后自动换库。'],
    ['history', '记录', '加密归档与历史，可导出留档；口令只存在本机。'],
    ['sync', '同步', '术语表的拉取与推送，需要在本机运行同步 helper。'],
    ['settings', '设置', '目标语言、列映射、项目档案与术语表导入，都集中在这一页。']
  ];

  var STEPS = [{
    target: '.brand',
    place: 'bottom',
    i18n: 'tour.welcome',
    zh: '欢迎使用 Tide。每个标签页都有一句自己的说明，跟着气泡走一遍就能认全。随时点右上角 ? 重看。'
  }].concat(TAB_TIPS.map(function (row) {
    return {
      before: row[0],
      target: 'button[data-tab="' + row[0] + '"]',
      place: 'bottom',
      i18n: 'tour.' + row[0],
      zh: row[1] + '：' + row[2]
    };
  })).concat([{
    close: true,
    i18n: 'tour.end',
    zh: '就这些。术语表在设置里导入，记忆库支持 TMX；点右上角 ? 可以再看一次。'
  }]);

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
    avatar.innerHTML = '<svg viewBox="0 0 24 24">' +
      '<rect x="5.5" y="4" width="13" height="2.8" rx="1.4" fill="#2f8be6"/>' +
      '<rect x="6" y="13.5" width="3" height="5.5" rx="1.3" fill="#63b0ff"/>' +
      '<rect x="10.5" y="9.5" width="3" height="9.5" rx="1.3" fill="#2f8be6"/>' +
      '<rect x="15" y="12" width="3" height="7" rx="1.3" fill="#0a56c9"/></svg>';
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

    var next = el('button', 'tour-btn primary', 'Next');
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
    bubble.querySelector('.tour-text').textContent = step.zh;
    bubble.querySelector('.tour-count').textContent = step.close ? '' : (index + 1) + ' / ' + STEPS.length;

    var dots = bubble.querySelector('.tour-dots');
    dots.textContent = '';
    if (STEPS.length > 9) dots.classList.add('dense');
    for (var i = 0; i < STEPS.length; i += 1) {
      var dot = el('i', i === index ? 'on' : '');
      dots.appendChild(dot);
    }

    var buttons = bubble.querySelectorAll('.tour-btn');
    buttons[0].style.visibility = index === 0 ? 'hidden' : 'visible';
    // Short, language-neutral labels: they fit one line at any bubble width.
    buttons[1].textContent = index === STEPS.length - 1 ? 'Done' : 'Next';

    place(step);
    requestAnimationFrame(function () { place(step); });
  }

  function go(next) {
    if (!layer) return;
    if (next < 0 || next >= STEPS.length) return;
    index = next;
    var step = STEPS[index];
    if (step.before) {
      showTab(step.before);
      // The strip scrolls; make sure the tab this step points at is visible.
      var button = document.querySelector('button[data-tab="' + step.before + '"]');
      if (button && button.scrollIntoView) {
        try { button.scrollIntoView({ inline: 'nearest', block: 'nearest' }); } catch (err) { /* older engine */ }
      }
    }
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
