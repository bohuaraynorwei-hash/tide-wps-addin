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
      'tab.tasks': 'Nhiệm vụ', 'tab.matches': 'Khớp', 'tab.translate': 'Dịch', 'tab.keys': 'Định vị Key',
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
      'label.uiLang': 'Ngôn ngữ giao diện', 'tour.skip': 'Bỏ qua', 'tour.next': 'Tiếp', 'tour.done': 'Xong',
      'tour.settings': 'Ngôn ngữ đích, ánh xạ cột, hồ sơ dự án và nhập thuật ngữ.',
      'tab.versions': 'Phiên bản', 'btn.columnChoice': 'Chọn cột',
      'btn.columnDetect': 'Nhận diện', 'btn.close': 'Đóng',
      'label.columnManual': 'Chỉ định cột nguồn / đích',
      'label.oldColShort': 'Bản dịch cũ', 'hint.columnManual': 'Nhập chữ cột; gợi ý chỉ điền, không tự ghi đè.',
      'btn.columnApply': 'Áp dụng cột', 'btn.exportFull': 'Xuất bảng đầy đủ',
      'btn.exportLog': 'Xuất nhật ký sửa', 'label.findScope': 'Phạm vi (trang tính + dòng)',
      'find.tabSheet': '1 Trang tính', 'find.tabRow': '2 Dòng',
      'find.rowPlaceholder': 'Nhập số dòng, ví dụ 128', 'btn.findJump': 'Đến bản gốc',
      'btn.findClear': 'Xóa phạm vi', 'find.allSheets': 'Tất cả trang tính',
      'hint.findScope': 'Hai điều kiện cùng có hiệu lực; chỉ nhập một điều kiện sẽ không tìm.',
      'hint.version': 'Bản gốc và bản dịch được lưu riêng; chỉ khôi phục mục hiện tại.',
      'version.target': 'Bản dịch', 'version.source': 'Bản gốc',
      'version.filter': 'Lọc nguồn / dòng / ô', 'term.currentSource': 'Nguồn thuật ngữ hiện tại',
      'term.nasProject': 'Dự án NAS', 'term.nasCache': 'Bộ nhớ đệm NAS',
      'term.local': 'Thuật ngữ cục bộ', 'term.bundled': 'Thuật ngữ tích hợp', 'term.notLoaded': 'Chưa tải',
      'term.bundledNote': 'đi kèm tiện ích', 'term.onDevice': 'máy này',
      'version.initial': 'Ban đầu', 'version.rollback': 'Khôi phục',
      'find.scopeActive': 'Định vị: {sheet} · dòng {row}', 'find.scopeNone': 'Chưa giới hạn: tìm toàn bộ workbook',
      'find.searchScoped': 'Tìm trong phạm vi', 'find.searchAll': 'Tìm toàn bộ workbook',
      'guide.menuTitle': 'Hướng dẫn theo trang', 'guide.intro': 'Chọn một trang để xem thao tác chính.',
      'guide.close': 'Đóng', 'guide.foot': 'Lần đầu nên xem “Nhiệm vụ” và “Dịch”.',
      'guide.switch': 'Đổi trang', 'guide.back': 'Quay lại', 'guide.next': 'Tiếp', 'guide.finish': 'Xong', 'guide.tourClose': 'Đóng',
      'guide.page.tasks': 'Nhiệm vụ', 'guide.page.translate': 'Dịch', 'guide.page.terms': 'Thuật ngữ',
      'guide.page.matches': 'Khớp', 'guide.page.keys': 'Định vị Key', 'guide.page.find': 'Tìm kiếm',
      'guide.page.review': 'Duyệt', 'guide.page.versions': 'Phiên bản', 'guide.page.tm': 'Bộ nhớ',
      'guide.page.grid': 'Bảng', 'guide.page.history': 'Lịch sử', 'guide.page.sync': 'Đồng bộ',
      'guide.page.settings': 'Cài đặt', 'guide.page.nas': 'NAS',
      "guide.tasks.0": "Sau khi đăng nhập NAS, các nhiệm vụ được giao hiển thị tại đây.",
      "guide.tasks.1": "Khi nhiệm vụ thay đổi, bấm tải lại; bấm mục để đến bản gốc và đồng bộ trang dịch.",
      "guide.translate.0": "Bản gốc và gợi ý thuật ngữ ở đây; bấm nhãn thuật ngữ để điền bản nháp.",
      "guide.translate.1": "Khi nhận diện sai, bấm đây để chọn cột nguồn và đích thủ công.",
      "guide.translate.2": "“Ghi” chỉ ghi mục hiện tại; dùng “Ghi và sang mục sau” để làm liên tục.",
      "guide.translate.3": "Cuối trang xuất bảng đầy đủ hoặc nhật ký sửa; chỉ giữ các dòng đã sửa.",
      "guide.terms.0": "Xem nguồn thuật ngữ hiện tại: dự án NAS, bộ nhớ đệm, cục bộ hoặc tích hợp.",
      "guide.terms.1": "Khi chưa đăng nhập, có thể chuyển giữa bộ nhớ đệm NAS và thuật ngữ cục bộ; sau khi đăng nhập, NAS được ưu tiên.",
      "guide.terms.2": "Tìm thuật ngữ theo đoạn tiếng Trung hoặc tiếng Anh; có thể đến dòng khớp.",
      "guide.matches.0": "Quét trang tính hiện tại để tìm các dòng có thuật ngữ.",
      "guide.matches.1": "Bấm dòng khớp để quay lại ô bản gốc.",
      "guide.keys.0": "Nhập Key đầy đủ hoặc số; hỗ trợ khớp chính xác và gần đúng.",
      "guide.keys.1": "Bấm kết quả để đến mục tương ứng.",
      "guide.find.0": "Hai tab: chọn trang tính, nhập dòng; chỉ giới hạn khi cả hai cùng có.",
      "guide.find.1": "Đến bản gốc chỉ định vị, không tìm hoặc thay thế.",
      "guide.find.2": "Thay thế hàng loạt chỉ sửa cột bản dịch; xem trước rồi thực hiện.",
      "guide.review.0": "Kiểm tra theo tất cả, dấu câu, chỗ giữ chỗ.",
      "guide.review.1": "Chấp nhận hoặc từ chối sửa đổi và xem bình luận tại đây.",
      "guide.versions.0": "Chọn riêng phiên bản bản gốc và bản dịch.",
      "guide.versions.1": "Mở mục để khôi phục một mục về bất kỳ giá trị lịch sử.",
      "guide.tm.0": "Xuất bộ nhớ thay đổi do plugin tạo trong tài liệu hiện tại.",
      "guide.tm.1": "Tìm bản dịch cũ theo đoạn nguồn hoặc bản dịch.",
      "guide.grid.0": "Lọc theo chưa dịch, lỗi QA hoặc sửa đổi chưa xử lý.",
      "guide.grid.1": "Sửa bản dịch trực tiếp; khi rời ô sẽ ghi lại và lưu sửa đổi.",
      "guide.history.0": "Mở kho lưu trữ mã hóa bằng mật khẩu trên máy.",
      "guide.history.1": "Lưu giá trị cũ và mới mỗi lần quét; có thể xuất.",
      "guide.sync.0": "Kéo cần helper đồng bộ trên máy; tải lên mặc định tắt.",
      "guide.sync.1": "Tải lên sẽ gửi dữ liệu máy lên GitHub; xác nhận phạm vi trước.",
      "guide.settings.0": "Đặt ngôn ngữ đích trước; ảnh hưởng nhận diện cột và thuật ngữ.",
      "guide.settings.1": "Có thể sửa ánh xạ cột; “Chọn cột” để nhận diện lại.",
      "guide.settings.2": "Lưu sau khi sửa; ánh xạ dùng cho trang hiện tại.",
      "guide.nas.0": "Mở bảng điều khiển NAS; dự án, thành viên, thuật ngữ và lịch sử ở đó.",
      "guide.nextPage": "Trang sau",
      'version.entriesTarget': 'mục có phiên bản dịch', 'version.entriesSource': 'mục có phiên bản gốc',
      'version.emptyTarget': 'Chưa có phiên bản bản dịch. Ghi bản dịch qua plugin để tạo lịch sử.',
      'version.emptySource': 'Chưa có phiên bản bản gốc. Chỉ ghi khi dùng “Ghi lại nguồn”.',
      'version.meta': '{count} phiên bản · hiện tại: {value}', 'version.emptyValue': 'trống',
      'version.before': 'Trước: ', 'version.value': 'Giá trị: ',
      'task.loginEmpty': 'Hãy đăng nhập NAS để xem nhiệm vụ được giao.',
      'task.loading': 'Đang tải nhiệm vụ…', 'task.hint': 'Nhấn nhiệm vụ để định vị bản gốc.',
      'task.none': 'Hiện chưa có nhiệm vụ được giao.', 'task.loadFailed': 'Không tải được nhiệm vụ: ',
      'task.count': '{count} nhiệm vụ', 'task.empty': 'Chưa có nhiệm vụ được giao.',
      'task.selectedItems': 'mục đã chọn', 'task.due': 'Hạn', 'task.nextSegment': 'Đoạn tiếp: ',
      'task.locate': 'Định vị', 'task.done': 'Hoàn tất', 'task.submit': 'Gửi',
      'task.items': 'Mục', 'task.loadingPage': 'Đang tải…', 'task.prevPage': 'Trang trước',
      'task.nextPage': 'Trang sau', 'task.hintLogin': 'Sau khi đăng nhập NAS, nhiệm vụ được giao sẽ hiện tại đây.',
      'task.login': 'Đăng nhập NAS', 'status.new': 'Chờ dịch', 'status.draft': 'Bản nháp',
      'status.translated': 'Đã dịch', 'status.reviewed': 'Đã duyệt',
      'term.items': 'mục', 'term.cachedAt': 'lưu lúc', 'term.synced': 'đã đồng bộ',
      'term.projectGlossary': 'Thuật ngữ dự án', 'term.imported': 'nhập trên máy',
      'find.scopeNeedBoth': 'Cần chọn trang tính và nhập dòng',
    },
    ru: {
      'tab.tasks': 'Задачи', 'tab.matches': 'Совпадения', 'tab.translate': 'Перевод', 'tab.keys': 'Поиск Key',
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
      'label.uiLang': 'Язык интерфейса', 'tour.skip': 'Пропустить', 'tour.next': 'Далее', 'tour.done': 'Готово',
      'tour.settings': 'Целевой язык, колонки, профили проекта и импорт глоссария.',
      'tab.versions': 'Версии', 'btn.columnChoice': 'Колонки',
      'btn.columnDetect': 'Распознать', 'btn.close': 'Закрыть',
      'label.columnManual': 'Колонки оригинала / перевода',
      'label.oldColShort': 'Старый перевод', 'hint.columnManual': 'Введите буквы колонок; подсказки только заполняют поля.',
      'btn.columnApply': 'Применить', 'btn.exportFull': 'Экспорт полной таблицы',
      'btn.exportLog': 'Журнал правок', 'label.findScope': 'Диапазон (лист + строка)',
      'find.tabSheet': '1 Лист', 'find.tabRow': '2 Строка',
      'find.rowPlaceholder': 'Номер строки, например 128', 'btn.findJump': 'К оригиналу',
      'btn.findClear': 'Сбросить диапазон', 'find.allSheets': 'Все листы',
      'hint.findScope': 'Оба условия действуют вместе; при одном условии поиск не начнётся.',
      'hint.version': 'Оригинал и перевод хранятся отдельно; откат затрагивает только текущую запись.',
      'version.target': 'Перевод', 'version.source': 'Оригинал',
      'version.filter': 'Фильтр оригинала / строки / ячейки', 'term.currentSource': 'Текущий источник терминов',
      'term.nasProject': 'Проект NAS', 'term.nasCache': 'Кэш NAS',
      'term.local': 'Локальный глоссарий', 'term.bundled': 'Встроенный глоссарий', 'term.notLoaded': 'Не загружено',
      'term.bundledNote': 'входит в плагин', 'term.onDevice': 'на этом устройстве',
      'version.initial': 'Начальное', 'version.rollback': 'Вернуть',
      'find.scopeActive': 'Диапазон: {sheet} · строка {row}', 'find.scopeNone': 'Без ограничений: вся книга',
      'find.searchScoped': 'Искать в диапазоне', 'find.searchAll': 'Искать во всей книге',
      'guide.menuTitle': 'Обучение по страницам', 'guide.intro': 'Выберите страницу с основными действиями.',
      'guide.close': 'Закрыть', 'guide.foot': 'Сначала посмотрите «Задачи» и «Перевод».',
      'guide.switch': 'Сменить', 'guide.back': 'Назад', 'guide.next': 'Далее', 'guide.finish': 'Готово', 'guide.tourClose': 'Закрыть',
      'guide.page.tasks': 'Задачи', 'guide.page.translate': 'Перевод', 'guide.page.terms': 'Термины',
      'guide.page.matches': 'Совпадения', 'guide.page.keys': 'Поиск Key', 'guide.page.find': 'Поиск',
      'guide.page.review': 'Проверка', 'guide.page.versions': 'Версии', 'guide.page.tm': 'Память',
      'guide.page.grid': 'Сетка', 'guide.page.history': 'История', 'guide.page.sync': 'Синхронизация',
      'guide.page.settings': 'Настройки', 'guide.page.nas': 'NAS',
      "guide.tasks.0": "После входа в NAS здесь показаны назначенные задачи.",
      "guide.tasks.1": "При изменениях нажмите обновить; нажмите запись, чтобы перейти к оригиналу и синхронизировать перевод.",
      "guide.translate.0": "Оригинал и подсказки терминов здесь; нажмите термин, чтобы вставить его в черновик.",
      "guide.translate.1": "При ошибке распознавания укажите колонки оригинала и перевода вручную.",
      "guide.translate.2": "«Записать» пишет только текущую запись; для серии — «Записать и далее».",
      "guide.translate.3": "Внизу — экспорт полной таблицы или журнала правок; только изменённые строки.",
      "guide.terms.0": "Сначала проверьте источник терминов: проект NAS, кэш, локальный или встроенный.",
      "guide.terms.1": "Без входа можно переключать кэш NAS и локальный глоссарий; после входа приоритет у NAS.",
      "guide.terms.2": "Ищите термин по фрагменту китайского или английского; можно перейти к строке.",
      "guide.matches.0": "Сканируйте лист, чтобы найти строки с терминами.",
      "guide.matches.1": "Нажмите строку совпадения, чтобы перейти к ячейке оригинала.",
      "guide.keys.0": "Введите полный Key или номер; точное и нечёткое совпадение.",
      "guide.keys.1": "Нажмите результат, чтобы перейти к записи.",
      "guide.find.0": "Две вкладки: лист и строка; диапазон работает только при обоих условиях.",
      "guide.find.1": "Переход к оригиналу только позиционирует, не ищет и не заменяет.",
      "guide.find.2": "Массовая замена меняет только колонку перевода; сначала предпросмотр.",
      "guide.review.0": "Проверка: всё, пунктуация, плейсхолдеры.",
      "guide.review.1": "Принимайте или отклоняйте правки и смотрите комментарии.",
      "guide.versions.0": "Версии оригинала и перевода выбираются отдельно.",
      "guide.versions.1": "Разверните запись, чтобы вернуть её к любому историческому значению.",
      "guide.tm.0": "Экспорт памяти изменений, созданных плагином.",
      "guide.tm.1": "Ищите старый перевод по фрагменту оригинала или перевода.",
      "guide.grid.0": "Фильтр: непереведённые, ошибки QA, необработанные правки.",
      "guide.grid.1": "Правьте перевод; при уходе из ячейки запись сохраняется.",
      "guide.history.0": "Откройте зашифрованный архив локальным паролем.",
      "guide.history.1": "Хранит старые и новые значения после сканирования; можно экспортировать.",
      "guide.sync.0": "Для загрузки нужен локальный helper; отправка по умолчанию выключена.",
      "guide.sync.1": "Отправка передаёт данные на GitHub; подтвердите объём.",
      "guide.settings.0": "Сначала целевой язык: влияет на колонки и глоссарий.",
      "guide.settings.1": "Колонки можно менять; «Колонки» — распознать заново.",
      "guide.settings.2": "Сохраните; колонки применятся к текущему листу.",
      "guide.nas.0": "Откройте консоль NAS: проекты, участники, термины и история.",
      "guide.nextPage": "Далее",
      'version.entriesTarget': 'записей с версиями перевода', 'version.entriesSource': 'записей с версиями оригинала',
      'version.emptyTarget': 'Версий перевода нет. Запишите перевод через плагин.',
      'version.emptySource': 'Версий оригинала нет. Они появляются только после «Записать оригинал».',
      'version.meta': '{count} версий · текущая: {value}', 'version.emptyValue': 'пусто',
      'version.before': 'До: ', 'version.value': 'Значение: ',
      'task.loginEmpty': 'Войдите в NAS, чтобы увидеть назначенные задачи.',
      'task.loading': 'Загрузка задач…', 'task.hint': 'Нажмите задачу, чтобы перейти к оригиналу.',
      'task.none': 'Назначенных задач пока нет.', 'task.loadFailed': 'Не удалось загрузить задачи: ',
      'task.count': '{count} задач', 'task.empty': 'Назначенных задач пока нет.',
      'task.selectedItems': 'выбранные записи', 'task.due': 'Срок', 'task.nextSegment': 'Далее: ',
      'task.locate': 'Перейти', 'task.done': 'Готово', 'task.submit': 'Отправить',
      'task.items': 'Записи', 'task.loadingPage': 'Загрузка…', 'task.prevPage': 'Назад',
      'task.nextPage': 'Далее', 'task.hintLogin': 'После входа в NAS здесь появятся назначенные задачи.',
      'task.login': 'Войти в NAS', 'status.new': 'К переводу', 'status.draft': 'Черновик',
      'status.translated': 'Переведено', 'status.reviewed': 'Проверено',
      'term.items': 'зап.', 'term.cachedAt': 'кэш', 'term.synced': 'синхронизировано',
      'term.projectGlossary': 'Глоссарий проекта', 'term.imported': 'локальный импорт',
      'find.scopeNeedBoth': 'Выберите лист и введите строку',
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
      'label.uiLang': 'Idioma da interface', 'tour.skip': 'Pular', 'tour.next': 'Avançar', 'tour.done': 'Concluir',
      'tour.settings': 'Idioma de destino, colunas, perfis de projeto e importação de glossário.',
      'tab.versions': 'Versões', 'btn.columnChoice': 'Colunas',
      'btn.columnDetect': 'Detectar', 'btn.close': 'Fechar',
      'label.columnManual': 'Definir colunas original / tradução',
      'label.oldColShort': 'Tradução antiga', 'hint.columnManual': 'Digite as colunas; as sugestões apenas preenchem, sem sobrescrever.',
      'btn.columnApply': 'Aplicar colunas', 'btn.exportFull': 'Exportar tabela completa',
      'btn.exportLog': 'Registro de edições', 'label.findScope': 'Escopo (planilha + linha)',
      'find.tabSheet': '1 Planilha', 'find.tabRow': '2 Linha',
      'find.rowPlaceholder': 'Número da linha, ex. 128', 'btn.findJump': 'Ir ao original',
      'btn.findClear': 'Limpar escopo', 'find.allSheets': 'Todas as planilhas',
      'hint.findScope': 'As duas condições valem juntas; com apenas uma, a busca não começa.',
      'hint.version': 'Original e tradução ficam separados; a reversão afeta só o item atual.',
      'version.target': 'Tradução', 'version.source': 'Original',
      'version.filter': 'Filtrar original / linha / célula', 'term.currentSource': 'Fonte atual de termos',
      'term.nasProject': 'Projeto NAS', 'term.nasCache': 'Cache NAS',
      'term.local': 'Glossário local', 'term.bundled': 'Glossário integrado', 'term.notLoaded': 'Não carregado',
      'term.bundledNote': 'incluído no plugin', 'term.onDevice': 'neste dispositivo',
      'version.initial': 'Inicial', 'version.rollback': 'Reverter',
      'find.scopeActive': 'Escopo: {sheet} · linha {row}', 'find.scopeNone': 'Sem limite: toda a pasta',
      'find.searchScoped': 'Buscar no escopo', 'find.searchAll': 'Buscar em toda a pasta',
      'guide.menuTitle': 'Guia por página', 'guide.intro': 'Escolha uma página para ver as ações principais.',
      'guide.close': 'Fechar', 'guide.foot': 'Comece por “Tarefas” e “Tradução”.',
      'guide.switch': 'Trocar', 'guide.back': 'Voltar', 'guide.next': 'Avançar', 'guide.finish': 'Concluir', 'guide.tourClose': 'Fechar',
      'guide.page.tasks': 'Tarefas', 'guide.page.translate': 'Tradução', 'guide.page.terms': 'Termos',
      'guide.page.matches': 'Ocorrências', 'guide.page.keys': 'Buscar Key', 'guide.page.find': 'Pesquisa',
      'guide.page.review': 'Revisão', 'guide.page.versions': 'Versões', 'guide.page.tm': 'Memória',
      'guide.page.grid': 'Grade', 'guide.page.history': 'Histórico', 'guide.page.sync': 'Sincronizar',
      'guide.page.settings': 'Ajustes', 'guide.page.nas': 'NAS',
      "guide.tasks.0": "Após entrar no NAS, as tarefas atribuídas aparecem aqui.",
      "guide.tasks.1": "Se as tarefas mudarem, atualize; clique no item para ir ao original e sincronizar a tradução.",
      "guide.translate.0": "Original e sugestões de termos aqui; clique no termo para preencher o rascunho.",
      "guide.translate.1": "Se a detecção errar, defina as colunas de original e tradução manualmente.",
      "guide.translate.2": "“Gravar” grava só o item atual; use “Gravar e avançar” para continuar.",
      "guide.translate.3": "No rodapé, exporte a tabela completa ou o registro; somente as linhas alteradas.",
      "guide.terms.0": "Veja a fonte atual de termos: projeto NAS, cache, local ou integrado.",
      "guide.terms.1": "Sem login, alterne entre cache NAS e glossário local; após login, o NAS tem prioridade.",
      "guide.terms.2": "Busque termos por trecho em chinês ou inglês; vá direto à linha.",
      "guide.matches.0": "Varra a planilha atual para encontrar linhas com termos.",
      "guide.matches.1": "Clique na linha encontrada para voltar à célula original.",
      "guide.keys.0": "Digite a Key completa ou o número; busca exata e aproximada.",
      "guide.keys.1": "Clique no resultado para ir ao item correspondente.",
      "guide.find.0": "Duas abas: planilha e linha; o escopo só vale com as duas.",
      "guide.find.1": "Ir ao original apenas posiciona; não busca nem substitui.",
      "guide.find.2": "Substituição em massa só na coluna de tradução; visualize antes.",
      "guide.review.0": "Verificação: tudo, pontuação, marcadores.",
      "guide.review.1": "Aceite ou rejeite revisões e veja comentários.",
      "guide.versions.0": "Escolha versões de original e tradução separadamente.",
      "guide.versions.1": "Expanda o item para reverter para qualquer valor histórico.",
      "guide.tm.0": "Exporte a memória de alterações criadas pelo plugin.",
      "guide.tm.1": "Pesquise traduções antigas por trecho de original ou tradução.",
      "guide.grid.0": "Filtre por não traduzido, QA ou revisões pendentes.",
      "guide.grid.1": "Edite a tradução; ao sair, grava e registra a revisão.",
      "guide.history.0": "Desbloqueie o arquivo criptografado com a senha local.",
      "guide.history.1": "Mantém valores antigos e novos de cada varredura; pode exportar.",
      "guide.sync.0": "Baixar exige helper local; envio desativado por padrão.",
      "guide.sync.1": "O envio manda dados ao GitHub; confirme o escopo.",
      "guide.settings.0": "Defina o idioma de destino primeiro; afeta colunas e termos.",
      "guide.settings.1": "O mapeamento pode ser alterado; “Colunas” detecta novamente.",
      "guide.settings.2": "Salve; o mapeamento vale para a planilha atual.",
      "guide.nas.0": "Abra o console NAS: projetos, membros, termos e histórico.",
      "guide.nextPage": "Próxima",
      'version.entriesTarget': 'itens com versões de tradução', 'version.entriesSource': 'itens com versões de original',
      'version.emptyTarget': 'Ainda não há versões de tradução. Grave pelo plugin para criar o histórico.',
      'version.emptySource': 'Ainda não há versões de original. Só são criadas ao usar “Gravar original”.',
      'version.meta': '{count} versões · atual: {value}', 'version.emptyValue': 'vazio',
      'version.before': 'Antes: ', 'version.value': 'Valor: ',
      'task.loginEmpty': 'Entre no NAS para ver as tarefas atribuídas.',
      'task.loading': 'Carregando tarefas…', 'task.hint': 'Clique na tarefa para localizar o original.',
      'task.none': 'Ainda não há tarefas atribuídas.', 'task.loadFailed': 'Falha ao carregar tarefas: ',
      'task.count': '{count} tarefas', 'task.empty': 'Ainda não há tarefas atribuídas.',
      'task.selectedItems': 'itens selecionados', 'task.due': 'Prazo', 'task.nextSegment': 'Próximo: ',
      'task.locate': 'Localizar', 'task.done': 'Concluído', 'task.submit': 'Enviar',
      'task.items': 'Itens', 'task.loadingPage': 'Carregando…', 'task.prevPage': 'Anterior',
      'task.nextPage': 'Próxima', 'task.hintLogin': 'Após entrar no NAS, as tarefas atribuídas aparecem aqui.',
      'task.login': 'Entrar no NAS', 'status.new': 'A traduzir', 'status.draft': 'Rascunho',
      'status.translated': 'Traduzido', 'status.reviewed': 'Revisado',
      'term.items': 'itens', 'term.cachedAt': 'cache em', 'term.synced': 'sincronizado',
      'term.projectGlossary': 'Glossário do projeto', 'term.imported': 'importado localmente',
      'find.scopeNeedBoth': 'Selecione a planilha e informe a linha',
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
    ['nav button[data-tab="tasks"]', 'tab.tasks'],
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

  var staticObserver = null;

  function applyStaticTexts() {
    if (typeof document === 'undefined') return;
    var maps = root.TideI18nStatic || {};
    var map = maps[state.lang] || {};
    var nodes = document.querySelectorAll('button:not([data-i18n]), label:not([data-i18n]), option:not([data-i18n]), th, h2, h3');
    nodes.forEach(function (el) {
      if (el.children.length) return;
      var current = String(el.textContent || '').trim();
      if (!current) return;
      var orig = el.dataset.i18nStaticOrig;
      var applied = el.dataset.i18nStaticApplied;
      if (!orig) {
        orig = current;
        el.dataset.i18nStaticOrig = orig;
      } else if (current !== orig && current !== applied) {
        // Dynamic text changed; treat the new source string as the base.
        orig = current;
        el.dataset.i18nStaticOrig = orig;
        el.dataset.i18nStaticApplied = '';
        applied = '';
      }
      if (state.lang === 'zh') {
        if (current !== orig) el.textContent = orig;
        return;
      }
      var value = map[orig];
      if (value && current !== value) {
        el.textContent = value;
        el.dataset.i18nStaticApplied = value;
      }
    });
  }

  function watchStaticTexts() {
    if (staticObserver || typeof document === 'undefined') return;
    staticObserver = new MutationObserver(function () { applyStaticTexts(); });
    staticObserver.observe(document.body, { childList: true, subtree: true });
  }

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
      if (typeof el.dataset.i18nOrig === 'undefined') {
        el.dataset.i18nOrig = el.textContent;
      }
      var value = t(rule[1]);
      el.textContent = value || el.dataset.i18nOrig;
    });
    document.querySelectorAll('[data-i18n]').forEach(function (el) {
      if (typeof el.dataset.i18nOrig === 'undefined') el.dataset.i18nOrig = el.textContent;
      var value = t(el.dataset.i18n);
      el.textContent = value || el.dataset.i18nOrig;
    });
    document.querySelectorAll('[data-i18n-ph]').forEach(function (el) {
      if (typeof el.dataset.i18nPhOrig === 'undefined') el.dataset.i18nPhOrig = el.getAttribute('placeholder') || '';
      var value = t(el.dataset.i18nPh);
      el.setAttribute('placeholder', value || el.dataset.i18nPhOrig);
    });
    if (typeof applyStaticTexts === 'function') applyStaticTexts();
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
    watchStaticTexts();
  }

  var api = { langs: LANGS, t: t, apply: apply, setLang: setLang, boot: boot, get lang() { return state.lang; } };
  if (root) root.TideI18n = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
