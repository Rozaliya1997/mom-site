/**
 * Google Apps Script для приёма заявок с сайта Венеры Хафизовой
 * 
 * Что делает:
 *  1. Получает данные с формы сайта (POST-запрос)
 *  2. Сохраняет каждую заявку как новую строку в Google-таблице
 *  3. Отправляет email-уведомление на почту
 *
 * Настройка — см. файл GOOGLE-SHEETS-SETUP.md
 */

// ⚠️ ИЗМЕНИ ЭТУ СТРОКУ — впиши почту, на которую приходить уведомления:
const NOTIFICATION_EMAIL = 'hafizovavenera@mail.ru';

// Заголовки колонок в таблице (в этом же порядке заявки будут записываться)
const COLUMNS = [
  'Дата и время',
  'Имя',
  'Телефон',
  'Email',
  'Тип бизнеса',
  'Задача',
  'Удобное время'
];

/**
 * Главная функция — вызывается, когда форма отправляет данные.
 */
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Если в таблице ещё нет заголовков — добавим их и оформим
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUMNS);
      sheet.getRange(1, 1, 1, COLUMNS.length)
        .setFontWeight('bold')
        .setBackground('#2D4A3E')
        .setFontColor('#F7F2E8')
        .setVerticalAlignment('middle');
      sheet.setFrozenRows(1);
      // Подгоним ширину колонок под содержимое
      sheet.setColumnWidth(1, 160);  // дата
      sheet.setColumnWidth(2, 140);  // имя
      sheet.setColumnWidth(3, 140);  // телефон
      sheet.setColumnWidth(4, 200);  // email
      sheet.setColumnWidth(5, 220);  // тип бизнеса
      sheet.setColumnWidth(6, 320);  // задача
      sheet.setColumnWidth(7, 180);  // удобное время
      sheet.setRowHeight(1, 36);
    }

    // Honey-pot защита от ботов — если поле _honey заполнено, не сохраняем
    if (e.parameter._honey) {
      return jsonResponse({ status: 'ignored' });
    }

    const timestamp = new Date();
    const row = [
      timestamp,
      e.parameter['Имя'] || '',
      e.parameter['Телефон'] || '',
      e.parameter['Email'] || '',
      e.parameter['Тип бизнеса'] || '',
      e.parameter['Задача'] || '',
      e.parameter['Удобное время'] || ''
    ];
    sheet.appendRow(row);

    // Форматируем ячейку с датой
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('dd.MM.yyyy HH:mm');

    // Отправляем email-уведомление
    sendNotificationEmail(e.parameter, timestamp);

    return jsonResponse({ status: 'success' });
  } catch (error) {
    // Если что-то пошло не так — пишем в лог скрипта
    console.error('Ошибка обработки заявки:', error);
    return jsonResponse({ status: 'error', message: error.toString() });
  }
}

/**
 * Отправляет email с деталями заявки.
 */
function sendNotificationEmail(data, timestamp) {
  const name = data['Имя'] || 'Клиент';
  const subject = '🔔 Новая заявка с сайта — ' + name;

  // Красивое HTML-письмо
  const html = `
    <div style="font-family: -apple-system, Helvetica, Arial, sans-serif; max-width: 560px; margin: 0 auto; background: #F7F2E8; padding: 32px;">
      <div style="background: #FFFFFF; border-radius: 8px; padding: 32px; border: 1px solid #D9CFB8;">
        <div style="font-size: 13px; color: #2D4A3E; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 8px; font-weight: 500;">— Новая заявка</div>
        <h1 style="font-family: Georgia, serif; font-weight: 400; font-size: 28px; color: #1F1A12; margin: 0 0 24px 0; line-height: 1.2;">${escapeHtml(name)}</h1>

        <table style="width: 100%; border-collapse: collapse; font-size: 15px; color: #1F1A12;">
          ${row('Телефон', data['Телефон'])}
          ${row('Email', data['Email'])}
          ${row('Тип бизнеса', data['Тип бизнеса'])}
          ${row('Задача', data['Задача'])}
          ${row('Удобное время', data['Удобное время'])}
          ${row('Когда пришла заявка', Utilities.formatDate(timestamp, 'Europe/Moscow', 'dd.MM.yyyy в HH:mm'))}
        </table>

        ${data['Телефон'] ? `
          <div style="margin-top: 32px; padding-top: 24px; border-top: 1px solid #E8DFCB;">
            <a href="tel:${escapeHtml(data['Телефон'])}" style="display: inline-block; background: #2D4A3E; color: #F7F2E8; padding: 12px 24px; border-radius: 999px; text-decoration: none; font-weight: 500; margin-right: 8px;">Позвонить</a>
            <a href="https://wa.me/${escapeHtml((data['Телефон'] || '').replace(/\\D/g, ''))}" style="display: inline-block; background: transparent; color: #2D4A3E; padding: 12px 24px; border-radius: 999px; text-decoration: none; font-weight: 500; border: 1px solid #2D4A3E;">WhatsApp</a>
          </div>
        ` : ''}
      </div>

      <div style="text-align: center; margin-top: 24px; font-size: 12px; color: #6B5F4D;">
        Заявка отправлена с сайта · Все заявки сохраняются в Google Таблице
      </div>
    </div>
  `;

  // Простой текстовый вариант (для почтовиков, которые не отображают HTML)
  const text = [
    'Новая заявка с сайта',
    '',
    'Имя: ' + (data['Имя'] || '—'),
    'Телефон: ' + (data['Телефон'] || '—'),
    'Email: ' + (data['Email'] || '—'),
    'Тип бизнеса: ' + (data['Тип бизнеса'] || '—'),
    'Задача: ' + (data['Задача'] || '—'),
    'Удобное время: ' + (data['Удобное время'] || '—'),
    '',
    'Дата: ' + Utilities.formatDate(timestamp, 'Europe/Moscow', 'dd.MM.yyyy в HH:mm')
  ].join('\n');

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    body: text,
    htmlBody: html,
    name: 'Сайт Венеры Хафизовой'
  });
}

/**
 * Вспомогательные функции
 */
function row(label, value) {
  if (!value) return '';
  return `
    <tr>
      <td style="padding: 12px 0; border-bottom: 1px solid #E8DFCB; color: #6B5F4D; font-size: 13px; text-transform: uppercase; letter-spacing: 0.04em; vertical-align: top; width: 40%;">${escapeHtml(label)}</td>
      <td style="padding: 12px 0; border-bottom: 1px solid #E8DFCB; color: #1F1A12; vertical-align: top;">${escapeHtml(value)}</td>
    </tr>
  `;
}

function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Тестовая функция — можно запустить из меню скрипта (Run → testSubmit),
 * чтобы проверить, что и таблица заполняется, и письмо приходит.
 */
function testSubmit() {
  const fakeEvent = {
    parameter: {
      'Имя': 'Тестовый Клиент',
      'Телефон': '+7 999 123 45 67',
      'Email': 'test@example.com',
      'Тип бизнеса': 'ИП на УСН',
      'Задача': 'Тестовая отправка для проверки настройки',
      'Удобное время': 'Будни после 18:00'
    }
  };
  const result = doPost(fakeEvent);
  console.log('Результат:', result.getContent());
}
