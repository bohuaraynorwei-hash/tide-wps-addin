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
  var VERSION = '4';   // bump to re-show the tour after a big UI change

  // One step per visible tab: the guide follows the tab strip, so every
  // function is introduced in the place it lives.
  var TAB_TIPS = [
    ['tasks', '我的任务', '登录 NAS 后，这里按项目列出分配给你的任务。点条目会跳到原文，并在翻译页同步选中它；展开后可以逐条定位或提交。'],
    ['translate', '翻译', '逐条处理的主场：「写入」只写当前译文单元格，「写入并下一条」写完后继续；原文框右侧可定位译文或原文。'],
    ['terms', '术语', '登录后会自动同步被指派项目的全量术语表；多项目时在顶部选择项目，切换项目就切换术语库。'],
    ['matches', '命中', '扫描当前表，列出所有含术语的行。点任意一行，会跳回它所在的原文单元格。'],
    ['keys', 'Key 定位', '按表头里的「完整Key」找词条，支持精确与模糊匹配，点结果直接跳转。'],
    ['find', '搜索', '跨工作表搜原文或译文，支持正则；结果同样可以点击定位。'],
    ['review', '审校', '修订的接受/拒绝、条目评论，以及质检与 CSV 导出都在这里。'],
    ['tm', '记忆库', '导入 TMX 或 Excel 作为记忆，看复用率；「导出修改记忆」只导出当前文档中由插件产生的修改。'],
    ['grid', '网格', '表格式的逐条编辑，可按「未译 / 有 QA / 有修订」筛选，就地改译文。'],
    ['history', '记录', '加密归档与历史，可导出留档；口令只存在本机。'],
    ['sync', '同步', '术语表的拉取与推送，需要在本机运行同步 helper。'],
    ['settings', '设置', '目标语言、列映射、项目档案与术语表导入，都集中在这一页。']
  ];

  // Feature steps are attached to the tab that owns them, so the tour still
  // follows the navigation instead of dumping every new button at the end.
  var FEATURE_STEPS = {
    translate: [{
      before: 'translate',
      target: '#btnWriteDraft',
      place: 'left',
      zh: '「写入」只把当前草稿写进译文单元格，不跳到下一条；需要连续处理时用下面的「写入并下一条」。'
    }],
    find: [{
      before: 'find',
      target: '#btnRepPreview',
      place: 'top',
      zh: '下面的「批量替换」只改译文列：预览里不会出现原文，执行后源文一个字都不会动。'
    }],
    terms: [{
      before: 'terms',
      target: '#termProject',
      place: 'bottom',
      zh: '这里选择要查看的项目术语表；切换项目会立即换库，不需要重新登录。'
    }],
    tm: [{
      before: 'tm',
      target: '#btnTmExportChanges',
      place: 'bottom',
      zh: '「导出修改记忆」把当前文档中由插件改过的条目导出为 TMX，换设备后可重新导入。'
    }],
    review: [{
      before: 'review',
      target: '#btnPunctQA',
      place: 'top',
      zh: '质检分三个按钮：全部、标点符号、占位符。占位符（{0}、颜色代码、换行符）只有点「占位符」时才检查。'
    }]
  };

  var STEPS = [{
    target: '.brand',
    place: 'bottom',
    zh: '欢迎使用 Tide。这里把 WPS 插件和 NAS 平台串起来：登录、接任务、翻译、术语和提交都在一条流程里。随时点右上角 ? 重看。'
  }, {
    target: '#topNas',
    place: 'bottom',
    zh: '顶部 NAS 是常驻入口：随时回到登录页，或打开 NAS 操作台。账号登录状态会保存在本设备。'
  }, {
    before: 'nas',
    target: '#btnNasConsole',
    place: 'bottom',
    zh: 'NAS 页面只保留登录和网页跳转。打开操作台后，项目、成员、术语、历史与账号管理都在 NAS 平台完成。'
  }];

  TAB_TIPS.forEach(function (row) {
    STEPS.push({
      before: row[0],
      target: 'button[data-tab="' + row[0] + '"]',
      place: 'bottom',
      zh: row[1] + '：' + row[2]
    });
    (FEATURE_STEPS[row[0]] || []).forEach(function (step) { STEPS.push(step); });
  });

  STEPS.push({
    before: 'settings',
    target: '#btnLang',
    place: 'bottom',
    zh: '标题栏的地球按钮切换界面语言：中文 / Tiếng Việt / Русский / Português (Brasil)。目标语言仍在「设置」里选。'
  });
  STEPS.push({
    close: true,
    zh: '就这些。建议顺序：先在 NAS 建项目/指派任务，再回 WPS 登录，从「我的任务」开始；术语表会自动同步。点右上角 ? 可以再看一次。'
  });

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
    if (name === 'nas') {
      var top = document.getElementById('topNas');
      if (top && !top.classList.contains('on')) top.click();
      return;
    }
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
    avatar.innerHTML = '<svg viewBox="0 0 24 24"><defs><linearGradient id="av" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#3366ff"/><stop offset=".52" stop-color="#7a5cff"/><stop offset="1" stop-color="#ff5c9e"/></linearGradient></defs>' +
      '<rect x="5.5" y="4.2" width="13" height="2.8" rx="1.4" fill="url(#av)"/>' +
      '<rect x="6" y="14" width="3" height="5.4" rx="1.4" fill="url(#av)"/>' +
      '<rect x="10.5" y="10" width="3" height="9.4" rx="1.4" fill="url(#av)"/>' +
      '<rect x="15" y="12.6" width="3" height="6.8" rx="1.4" fill="url(#av)"/></svg>';
    head.appendChild(avatar);
    head.appendChild(el('b', '', 'Tide 引导'));
    head.appendChild(el('span', 'tour-count', ''));
    bubble.appendChild(head);

    bubble.appendChild(el('p', 'tour-text', ''));

    var foot = el('div', 'tour-foot');
    var skip = el('button', 'skip', 'Skip');
    skip.addEventListener('click', function () { finish(); });
    foot.appendChild(skip);
    foot.appendChild(el('div', 'tour-dots', ''));

    var prev = el('button', 'tour-btn ghost', 'Back');
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
