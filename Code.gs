/**
 *  СКРИПТ ДЛЯ GOOGLE ТАБЛИЦЫ — приём ответов из HTML-опроса
 *
 *  Что делает: каждую отправленную анкету добавляет одной строкой в таблицу.
 *  Как установить: см. README.md, раздел «Подключение таблицы» (5 минут, без кода).
 *
 *  ВАЖНО: названия полей f0, f1, ... должны совпадать с SHEET_HEADERS в index.html
 */

var SHEET_HEADERS = [
  'Дата и время', 'Имя', 'Возраст', 'Проживание в деревне',
  'Чего не хватает (главное)', 'Чего не хватает — уточнение',
  'Что нравится в сквере/мосту', 'Что хочется изменить',
  'Важные элементы (1-5)', 'Важные элементы сквера',
  'Мост безопасен?', 'Опасения по мосту', 'Опасения — уточнение',
  'Помощь проекту', 'Чем можете помочь', 'Контакт',
  'Пожелания', 'Согласие'
];

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
           .setFontWeight('bold')
           .setBackground('#e8f5e9');
      sheet.autoResizeColumns(1, SHEET_HEADERS.length);
    }

    var row = SHEET_HEADERS.map(function (_, i) {
      return String(e.parameter['f' + i] || '').trim();
    });

    // защита от мусора: пустые анкеты не записываем
    var joined = row.join('');
    if (joined === '' || joined === 'Согласие') return;

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Ответы приходят как f0..f17 — разворачиваем в читаемые заголовки */
function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var out = [
    ['Поля анкеты:'].concat(SHEET_HEADERS),
    ['Ключи:'].concat(SHEET_HEADERS.map(function (_, i) { return 'f' + i; })),
    [],
    sheet.getDataRange().getValues()
  ];
  return ContentService
    .createTextOutput(out.map(function (r) { return r.join(' | '); }).join('\n'))
    .setMimeType(ContentService.MimeType.TEXT);
}