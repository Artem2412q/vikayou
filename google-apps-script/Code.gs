/**
 * Backend для сайта-приглашения.
 * 1) Добавляет подтверждение в привязанную Google Таблицу.
 * 2) Отправляет уведомление на почту владельца.
 *
 * Этот скрипт нужно вставить в Google Таблице:
 * Расширения → Apps Script.
 */

const NOTIFICATION_EMAIL = 'artemcepko69@gmail.com';
const ANSWERS_SHEET_NAME = 'Ответы';

const ALLOWED_TIMES = ['15:00', '15:30', '16:00', '16:30'];
const ALLOWED_CUISINES = [
  'Русская',
  'Немецкая',
  'Французская',
  'Итальянская',
  'Испанская',
  'Азиатская'
];

function doGet() {
  return jsonResponse_({
    ok: true,
    message: 'Сервис подтверждений работает.'
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const params = (e && e.parameter) || {};
    const name = clean_(params.name, 60) || 'Виктория';
    const date = clean_(params.date, 20);
    const time = clean_(params.time, 10);
    const cuisine = clean_(params.cuisine, 60);

    if (date !== '15.07.2026') {
      return jsonResponse_({ ok: false, error: 'Некорректная дата.' });
    }

    if (ALLOWED_TIMES.indexOf(time) === -1) {
      return jsonResponse_({ ok: false, error: 'Некорректное время.' });
    }

    if (ALLOWED_CUISINES.indexOf(cuisine) === -1) {
      return jsonResponse_({ ok: false, error: 'Некорректная кухня.' });
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Скрипт должен быть привязан к Google Таблице.');
    }

    const sheet = getOrCreateAnswersSheet_(spreadsheet);
    const confirmedAt = new Date();

    sheet.appendRow([
      confirmedAt,
      safeCell_(name),
      safeCell_(date),
      safeCell_(time),
      safeCell_(cuisine),
      'Подтверждено'
    ]);

    const timeZone = spreadsheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    const confirmedAtText = Utilities.formatDate(
      confirmedAt,
      timeZone,
      'dd.MM.yyyy HH:mm:ss'
    );

    const subject = '❤️ Виктория подтвердила свидание';
    const plainBody = [
      'Виктория подтвердила свидание ❤️',
      '',
      'Дата: ' + date,
      'Время: ' + time,
      'Кухня: ' + cuisine,
      '',
      'Подтверждено: ' + confirmedAtText
    ].join('\n');

    const htmlBody = [
      '<div style="font-family:Arial,sans-serif;line-height:1.6;color:#2b2430">',
      '<h2 style="margin:0 0 18px;color:#9a5b67">Виктория подтвердила свидание ❤️</h2>',
      '<p><b>📅 Дата:</b> ' + escapeHtml_(date) + '</p>',
      '<p><b>🕒 Время:</b> ' + escapeHtml_(time) + '</p>',
      '<p><b>🍽️ Кухня:</b> ' + escapeHtml_(cuisine) + '</p>',
      '<p style="margin-top:22px;color:#777;font-size:13px">Подтверждено: ' + escapeHtml_(confirmedAtText) + '</p>',
      '</div>'
    ].join('');

    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: 'Свидание с Викторией'
    });

    return jsonResponse_({ ok: true });
  } catch (error) {
    console.error(error);
    return jsonResponse_({
      ok: false,
      error: error && error.message ? error.message : 'Внутренняя ошибка.'
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseError) {
      // Блокировка могла не успеть установиться — это не требует действий.
    }
  }
}

function getOrCreateAnswersSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(ANSWERS_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(ANSWERS_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Подтверждено в',
      'Имя',
      'Дата',
      'Время',
      'Кухня',
      'Статус'
    ]);

    const header = sheet.getRange(1, 1, 1, 6);
    header.setFontWeight('bold');
    header.setBackground('#ead7a4');
    header.setFontColor('#332b25');

    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('dd.MM.yyyy HH:mm:ss');
    sheet.setColumnWidth(1, 170);
    sheet.setColumnWidth(2, 130);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 130);
  }

  return sheet;
}

function clean_(value, maxLength) {
  return String(value || '')
    .replace(/[<>]/g, '')
    .replace(/[\u0000-\u001F\u007F]/g, ' ')
    .trim()
    .slice(0, maxLength);
}

function safeCell_(value) {
  const text = String(value || '');
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
