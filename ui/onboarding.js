/**
 * Tide page guide: the ? button opens a page picker, then a short walkthrough
 * scoped to that page. Each page only teaches its primary controls.
 */
(function () {
  'use strict';

  var SEEN_KEY = 'tide.tourSeen';
  var VERSION = '12';

  var ORDER = ['tasks', 'translate', 'terms', 'matches', 'keys', 'find', 'review', 'versions', 'tm', 'grid', 'history', 'sync', 'settings', 'nas'];

  var GUIDES = {
    tasks: {
      name: '我的任务',
      steps: [
        { target: '#taskList', place: 'bottom', text: '登录 NAS 后，这里列出分配给你的任务。' },
        { target: '#btnTasksRefresh', place: 'bottom', text: '任务有变化时点刷新；点条目会跳到原文并同步翻译页。' }
      ]
    },
    translate: {
      name: '翻译',
      steps: [
        { target: '#entrySource', place: 'bottom', text: '原文和术语提示都在这里；点术语标签可填入草稿。' },
        { target: '#btnColumnChoice', place: 'bottom', text: '识别不准时点这里，手工指定源语列和目标语列。' },
        { target: '#btnWriteDraft', place: 'left', text: '「写入」只写当前条；连续处理用「写入并下一条」。' },
        { target: '#btnExportSessionEdits', place: 'top', text: '底部导出完整表格或修改记录；只保留本次修改行。' }
      ]
    },
    terms: {
      name: '术语',
      steps: [
        { target: '#termSourceBox', place: 'bottom', text: '先看当前术语来源：NAS 项目、缓存、本地或内置。' },
        { target: '#btnSwitchTermSource', place: 'top', text: '未登录时可切换 NAS 缓存和本地术语；登录后 NAS 优先。' },
        { target: '#termSearch', place: 'bottom', text: '按中文或英文片段查术语，可直接定位到命中行。' }
      ]
    },
    matches: {
      name: '命中',
      steps: [
        { target: '#btnScan', place: 'bottom', text: '扫描当前表，找出包含术语的行。' },
        { target: '#matchTable', place: 'bottom', text: '点任意命中行，直接跳回原文单元格。' }
      ]
    },
    keys: {
      name: 'Key 定位',
      steps: [
        { target: '#keyQuery', place: 'bottom', text: '输入完整 Key 或编号，支持精确和模糊匹配。' },
        { target: '#btnKeySearch', place: 'bottom', text: '点结果即可跳到对应词条。' }
      ]
    },
    find: {
      name: '搜索',
      steps: [
        { target: '#findScopeTabs', place: 'bottom', text: '两个选项卡：选工作表、填行号；两项同时满足才限定范围。' },
        { target: '#btnFindJumpSource', place: 'top', text: '跳转到原文只定位，不搜索、不替换。' },
        { target: '#btnRepPreview', place: 'top', text: '批量替换只改译文列；先预览，再执行。' }
      ]
    },
    review: {
      name: '审校',
      steps: [
        { target: '#btnPunctQA', place: 'top', text: '质检按全部、标点、占位符分开检查。' },
        { target: '#reviewTable', place: 'bottom', text: '在这里接受或拒绝修订，并查看条目评论。' }
      ]
    },
    versions: {
      name: '版本',
      steps: [
        { target: '#versionKind', place: 'bottom', text: '原文和译文版本分开选择，互不混淆。' },
        { target: '#versionEntries', place: 'bottom', text: '展开词条，可把单个词条回滚到任意历史值。' }
      ]
    },
    tm: {
      name: '记忆库',
      steps: [
        { target: '#btnTmExportChanges', place: 'bottom', text: '导出当前文档中由插件产生的修改记忆。' },
        { target: '#tmQuery', place: 'bottom', text: '按源文或译文片段检索旧译。' }
      ]
    },
    grid: {
      name: '表格编辑',
      steps: [
        { target: '#gridFilter', place: 'bottom', text: '按未译、QA 问题或未处理修订筛选条目。' },
        { target: '#gridTable', place: 'bottom', text: '直接改译文，失焦后写回并记录修订。' }
      ]
    },
    history: {
      name: '记录',
      steps: [
        { target: '#btnUnlock', place: 'bottom', text: '用本机口令解锁加密归档。' },
        { target: '#historyTable', place: 'bottom', text: '这里保留每次扫描的旧值和新值，可导出留档。' }
      ]
    },
    sync: {
      name: '同步',
      steps: [
        { target: '#btnPull', place: 'bottom', text: '拉取需要本机同步 helper；默认关闭上传。' },
        { target: '#btnPush', place: 'bottom', text: '上传会把本机数据发到 GitHub，启用前确认授权范围。' }
      ]
    },
    settings: {
      name: '设置',
      steps: [
        { target: '#targetLang', place: 'bottom', text: '先设目标语言，它影响列识别和术语筛选。' },
        { target: '#colSource', place: 'bottom', text: '列映射可手工改；「识别/选择列」可重新检测。' },
        { target: '#btnSave', place: 'top', text: '改完点保存；列映射会用于当前表。' }
      ]
    },
    nas: {
      name: 'NAS',
      steps: [
        { target: '#btnNasConsole', place: 'bottom', text: '打开 NAS 操作台，项目、成员、术语和历史都在平台完成。' }
      ]
    }
  };

  var layer = null;
  var ring = null;
  var bubble = null;
  var menu = null;
  var page = null;
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

  function markSeen() {
    try { window.localStorage.setItem(SEEN_KEY, VERSION); } catch (err) { /* storage blocked */ }
  }

  function buildMenu() {
    if (menu) return;
    menu = el('div', 'guide-menu');
    var panel = el('div', 'guide-panel');
    var head = el('div', 'guide-head');
    head.appendChild(el('b', '', '页面引导'));
    var close = el('button', 'guide-close', '×');
    close.type = 'button';
    close.addEventListener('click', closeMenu);
    head.appendChild(close);
    panel.appendChild(head);
    panel.appendChild(el('p', 'guide-intro', '选择一个页面，查看该页主要操作。'));
    var grid = el('div', 'guide-grid');
    ORDER.forEach(function (key) {
      var guide = GUIDES[key];
      if (!guide) return;
      var btn = el('button', 'guide-item', guide.name);
      btn.type = 'button';
      btn.addEventListener('click', function () { selectPage(key); });
      grid.appendChild(btn);
    });
    panel.appendChild(grid);
    var foot = el('div', 'guide-foot');
    foot.appendChild(el('span', '', '首次使用建议先看「我的任务」和「翻译」。'));
    var done = el('button', 'guide-done', '关闭');
    done.type = 'button';
    done.addEventListener('click', closeMenu);
    foot.appendChild(done);
    panel.appendChild(foot);
    menu.appendChild(panel);
    menu.addEventListener('click', function (event) {
      if (event.target === menu) closeMenu();
    });
    document.body.appendChild(menu);
  }

  function openMenu() {
    buildMenu();
    if (layer) layer.style.display = 'none';
    menu.classList.add('on');
  }

  function closeMenu() {
    if (menu) menu.classList.remove('on');
    markSeen();
  }

  function buildTour() {
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
    var switchPage = el('button', 'tour-link', '换页');
    switchPage.type = 'button';
    switchPage.addEventListener('click', openMenu);
    head.appendChild(switchPage);
    bubble.appendChild(head);
    bubble.appendChild(el('p', 'tour-text', ''));
    var foot = el('div', 'tour-foot');
    var close = el('button', 'skip', '关闭');
    close.type = 'button';
    close.addEventListener('click', finish);
    foot.appendChild(close);
    foot.appendChild(el('div', 'tour-dots', ''));
    var prev = el('button', 'tour-btn ghost', '上一步');
    prev.type = 'button';
    prev.addEventListener('click', function () { go(index - 1); });
    foot.appendChild(prev);
    var next = el('button', 'tour-btn primary', '下一步');
    next.type = 'button';
    next.addEventListener('click', function () {
      if (index >= GUIDES[page].steps.length - 1) finish();
      else go(index + 1);
    });
    foot.appendChild(next);
    bubble.appendChild(foot);
    bubble.appendChild(el('div', 'tail'));
    layer.appendChild(ring);
    layer.appendChild(bubble);
    document.body.appendChild(layer);
    document.addEventListener('keydown', function (event) {
      if (!layer || layer.style.display === 'none') return;
      if (event.key === 'Escape') finish();
    });
    window.addEventListener('resize', function () {
      if (layer && layer.style.display !== 'none') place(GUIDES[page].steps[index]);
    });
  }

  function selectPage(key) {
    if (!GUIDES[key]) return;
    page = key;
    index = 0;
    if (menu) menu.classList.remove('on');
    buildTour();
    layer.style.display = 'block';
    showTab(key);
    window.setTimeout(render, 80);
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
    var guide = GUIDES[page];
    var step = guide.steps[index];
    bubble.querySelector('.tour-text').textContent = step.text;
    bubble.querySelector('.tour-count').textContent = guide.name + ' ' + (index + 1) + '/' + guide.steps.length;
    var dots = bubble.querySelector('.tour-dots');
    dots.textContent = '';
    for (var i = 0; i < guide.steps.length; i += 1) dots.appendChild(el('i', i === index ? 'on' : ''));
    var buttons = bubble.querySelectorAll('.tour-btn');
    buttons[0].style.visibility = index === 0 ? 'hidden' : 'visible';
    buttons[1].textContent = index === guide.steps.length - 1 ? '完成' : '下一步';
    place(step);
    requestAnimationFrame(function () { place(step); });
  }

  function go(next) {
    if (!page) return;
    var steps = GUIDES[page].steps;
    if (next < 0 || next >= steps.length) return;
    index = next;
    render();
  }

  function finish() {
    if (layer) layer.style.display = 'none';
    markSeen();
  }

  function boot() {
    var open = document.getElementById('btnTour');
    if (open) open.addEventListener('click', openMenu);
    var seen = null;
    try { seen = window.localStorage.getItem(SEEN_KEY); } catch (err) { seen = null; }
    if (seen !== VERSION) window.setTimeout(openMenu, 900);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
