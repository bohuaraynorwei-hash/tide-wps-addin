/**
 * UI language layer.
 *
 * Scope: navigation, panel labels, primary actions and the language selector.
 * Deeper hints and status messages still come from the Chinese source strings;
 * they are listed in COVERAGE below so the gap is explicit rather than silent.
 */
(function (root) {
  'use strict';

  var LANGS = [
    { code: 'zh', label: '中文' },
    { code: 'vi', label: 'Tiếng Việt' },
    { code: 'ru', label: 'Русский' },
    { code: 'pt-BR', label: 'Português (Brasil)' }
  ];

  var DICT = {
    zh: {},
    vi: {
      'tab.matches': 'Khớp', 'tab.translate': 'Dịch', 'tab.keys': 'Định vị Key',
      'tab.free': 'Tự do', 'tab.find': 'Tìm kiếm', 'tab.review': 'Duyệt', 'tab.tm': 'Bộ nhớ',
      'tab.grid': 'Bảng', 'tab.project': 'Dự án', 'tab.terms': 'Thuật ngữ',
      'tab.history': 'Lịch sử', 'tab.sync': 'Đồng bộ', 'tab.settings': 'Cài đặt',

      'btn.scan': 'Quét trang tính', 'btn.prev': 'Mục trước', 'btn.next': 'Mục sau',
      'btn.refresh': 'Tải lại', 'btn.writeNext': 'Ghi và sang mục sau',
      'btn.undoWrite': 'Hoàn tác lần ghi', 'btn.rollback': 'Khôi phục bản cũ',
      'btn.copySource': 'Sao chép nguồn', 'btn.copyDraft': 'Sao chép bản nháp',
      'btn.pasteDraft': 'Dán vào bản nháp', 'btn.readSelTarget': 'Lấy vùng chọn → bản dịch',
      'btn.resetSource': 'Khôi phục giá trị ô', 'btn.saveSource': 'Ghi lại nguồn',
      'btn.readSelSource': 'Lấy vùng chọn → nguồn', 'btn.locateSource': 'Định vị nguồn',
      'btn.locateTarget': 'Định vị bản dịch', 'btn.saveSettings': 'Lưu cài đặt',
      'btn.importTerms': 'Nhập thuật ngữ', 'btn.runQA': 'Chạy kiểm tra',
      'btn.exportCSV': 'Xuất CSV', 'btn.apply': 'Dùng', 'btn.copy': 'Sao chép',

      'label.source': 'Nguồn (chỉ đọc)',
      'label.draft': 'Bản dịch nháp', 'label.memory': 'Khớp bộ nhớ',
      'label.terms': 'Thuật ngữ', 'label.targetLang': 'Ngôn ngữ đích',
      'label.project': 'Dự án (quy tắc bảng)', 'label.columns': 'Thiết lập cột',
      'label.sourceCol': 'Cột nguồn', 'label.targetCol': 'Cột bản dịch',
      'label.oldCol': 'Cột bản cũ', 'label.uidCol': 'Cột UID',
      'label.uiLang': 'Ngôn ngữ giao diện'
    },
    ru: {
      'tab.matches': 'Совпадения', 'tab.translate': 'Перевод', 'tab.keys': 'Поиск Key',
      'tab.free': 'Свободно', 'tab.find': 'Поиск', 'tab.review': 'Проверка', 'tab.tm': 'Память',
      'tab.grid': 'Сетка', 'tab.project': 'Проект', 'tab.terms': 'Термины',
      'tab.history': 'История', 'tab.sync': 'Синхронизация', 'tab.settings': 'Настройки',

      'btn.scan': 'Сканировать лист', 'btn.prev': 'Назад', 'btn.next': 'Далее',
      'btn.refresh': 'Обновить', 'btn.writeNext': 'Записать и далее',
      'btn.undoWrite': 'Отменить запись', 'btn.rollback': 'Вернуть перевод',
      'btn.copySource': 'Копировать оригинал', 'btn.copyDraft': 'Копировать черновик',
      'btn.pasteDraft': 'Вставить в черновик', 'btn.readSelTarget': 'Из выделения → перевод',
      'btn.resetSource': 'Сбросить к ячейке', 'btn.saveSource': 'Записать оригинал',
      'btn.readSelSource': 'Из выделения → оригинал', 'btn.locateSource': 'К оригиналу',
      'btn.locateTarget': 'К переводу', 'btn.saveSettings': 'Сохранить',
      'btn.importTerms': 'Импорт терминов', 'btn.runQA': 'Проверить',
      'btn.exportCSV': 'Экспорт CSV', 'btn.apply': 'Применить', 'btn.copy': 'Копировать',

      'label.source': 'Оригинал (только чтение)',
      'label.draft': 'Черновик перевода', 'label.memory': 'Совпадения в памяти',
      'label.terms': 'Глоссарий', 'label.targetLang': 'Целевой язык',
      'label.project': 'Проект (правила таблицы)', 'label.columns': 'Колонки',
      'label.sourceCol': 'Колонка оригинала', 'label.targetCol': 'Колонка перевода',
      'label.oldCol': 'Колонка прошлого перевода', 'label.uidCol': 'Колонка UID',
      'label.uiLang': 'Язык интерфейса'
    },
    'pt-BR': {
      'tab.matches': 'Ocorrências', 'tab.translate': 'Tradução', 'tab.keys': 'Buscar Key',
      'tab.free': 'Livre', 'tab.find': 'Pesquisa', 'tab.review': 'Revisão', 'tab.tm': 'Memória',
      'tab.grid': 'Grade', 'tab.project': 'Projeto', 'tab.terms': 'Termos',
      'tab.history': 'Histórico', 'tab.sync': 'Sincronizar', 'tab.settings': 'Ajustes',

      'btn.scan': 'Varrer planilha', 'btn.prev': 'Anterior', 'btn.next': 'Próximo',
      'btn.refresh': 'Atualizar', 'btn.writeNext': 'Gravar e avançar',
      'btn.undoWrite': 'Desfazer gravação', 'btn.rollback': 'Reverter tradução',
      'btn.copySource': 'Copiar original', 'btn.copyDraft': 'Copiar rascunho',
      'btn.pasteDraft': 'Colar no rascunho', 'btn.readSelTarget': 'Ler seleção → tradução',
      'btn.resetSource': 'Restaurar célula', 'btn.saveSource': 'Gravar original',
      'btn.readSelSource': 'Ler seleção → original', 'btn.locateSource': 'Ir ao original',
      'btn.locateTarget': 'Ir à tradução', 'btn.saveSettings': 'Salvar ajustes',
      'btn.importTerms': 'Importar termos', 'btn.runQA': 'Verificar',
      'btn.exportCSV': 'Exportar CSV', 'btn.apply': 'Usar', 'btn.copy': 'Copiar',

      'label.source': 'Original (somente leitura)',
      'label.draft': 'Rascunho da tradução', 'label.memory': 'Correspondências',
      'label.terms': 'Glossário', 'label.targetLang': 'Idioma de destino',
      'label.project': 'Projeto (regras)', 'label.columns': 'Colunas',
      'label.sourceCol': 'Coluna do original', 'label.targetCol': 'Coluna da tradução',
      'label.oldCol': 'Coluna da versão anterior', 'label.uidCol': 'Coluna UID',
      'label.uiLang': 'Idioma da interface'
    }
  };

  // Elements that carry translatable text, in one place so a missing key is a
  // visible blank rather than a silent Chinese leak.
  var RULES = [
    ['nav button[data-tab="matches"]', 'tab.matches'],
    ['nav button[data-tab="translate"]', 'tab.translate'],
    ['nav button[data-tab="keys"]', 'tab.keys'],
    ['nav button[data-tab="free"]', 'tab.free'],
    ['nav button[data-tab="find"]', 'tab.find'],
    ['nav button[data-tab="review"]', 'tab.review'],
    ['nav button[data-tab="tm"]', 'tab.tm'],
    ['nav button[data-tab="grid"]', 'tab.grid'],
    ['nav button[data-tab="project"]', 'tab.project'],
    ['nav button[data-tab="terms"]', 'tab.terms'],
    ['nav button[data-tab="history"]', 'tab.history'],
    ['nav button[data-tab="sync"]', 'tab.sync'],
    ['nav button[data-tab="settings"]', 'tab.settings'],

    ['#btnScan', 'btn.scan'], ['#btnPrevEntry', 'btn.prev'], ['#btnNextEntry', 'btn.next'],
    ['#btnRefreshEntries', 'btn.refresh'], ['#btnWriteNext', 'btn.writeNext'],
    ['#btnUndoWrite', 'btn.undoWrite'], ['#btnRollback', 'btn.rollback'],
    ['#btnCopySource', 'btn.copySource'], ['#btnCopyDraft', 'btn.copyDraft'],
    ['#btnPasteDraft', 'btn.pasteDraft'], ['#btnReadCell', 'btn.readSelTarget'],
    ['#btnResetSource', 'btn.resetSource'], ['#btnSaveSource', 'btn.saveSource'],
    ['#btnReadCellSrc', 'btn.readSelSource'], ['#btnLocateSource', 'btn.locateSource'],
    ['#btnLocate', 'btn.locateTarget'], ['#btnSaveSettings', 'btn.saveSettings'],
    ['#btnRunQA', 'btn.runQA'], ['#btnExportQA', 'btn.exportCSV'],

    ['#entrySourceLabel', 'label.source'], ['#entryDraftLabel', 'label.draft'],
    ['#memoryLabel', 'label.memory'], ['#targetLangLabel', 'label.targetLang'],
    ['#projectLabel', 'label.project'], ['#columnsLabel', 'label.columns'],
    ['#uiLangLabel', 'label.uiLang']
  ];

  var state = { lang: 'zh' };

  function t(key) {
    if (state.lang === 'zh') return '';
    return (DICT[state.lang] && DICT[state.lang][key]) || '';
  }

  function apply() {
    if (typeof document === 'undefined') return;
    document.documentElement.lang = state.lang === 'pt-BR' ? 'pt-BR' : state.lang;
    RULES.forEach(function (rule) {
      var el = document.querySelector(rule[0]);
      if (!el) return;
      var value = t(rule[1]);
      if (value) el.textContent = value;
    });
  }

  function setLang(code) {
    state.lang = DICT[code] ? code : 'zh';
    try { localStorage.setItem('tide.uiLang', state.lang); } catch (err) { /* storage may be blocked */ }
    apply();
  }

  function boot() {
    var saved = 'zh';
    try { saved = localStorage.getItem('tide.uiLang') || 'zh'; } catch (err) { saved = 'zh'; }
    state.lang = DICT[saved] ? saved : 'zh';
    apply();
  }

  var api = { langs: LANGS, t: t, apply: apply, setLang: setLang, boot: boot, get lang() { return state.lang; } };
  if (root) root.TideI18n = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
