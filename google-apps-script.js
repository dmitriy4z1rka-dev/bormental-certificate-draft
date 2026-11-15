/**
 * Приём заявок на подарочные сертификаты «Борменталь» в Google Таблицу.
 * Этот файл НЕ подключается к странице — его код вставляется в Google Apps Script.
 *
 * КАК ПОДКЛЮЧИТЬ
 * 1. Откройте Google Таблицу, куда будут падать заявки (или создайте новую: sheets.new).
 *    Заявки пишутся в первый лист. Если он пустой, скрипт сам добавит строку заголовков.
 * 2. В таблице: меню «Расширения» → «Apps Script».
 * 3. В редакторе удалите всё содержимое файла Code.gs и вставьте этот файл целиком.
 * 4. Проверьте NOTIFY_EMAIL ниже — на этот адрес придёт письмо о каждой заявке.
 * 5. Нажмите «Сохранить» (значок дискеты).
 * 6. «Начать развертывание» (Deploy) → «Новое развертывание».
 *    Значок шестерёнки → тип «Веб-приложение».
 *      • Описание: например, «Заявки на сертификаты».
 *      • Запуск от имени: «Я» (ваш аккаунт).
 *      • У кого есть доступ: «Все».
 *    Нажмите «Начать развертывание».
 * 7. Google попросит разрешения: «Предоставить доступ» → выберите свой аккаунт →
 *    «Дополнительные настройки» → «Перейти на страницу проекта (небезопасно)» → «Разрешить».
 *    Это ваш собственный скрипт; доступ нужен, чтобы писать в таблицу и отправлять письма.
 * 8. Скопируйте «URL веб-приложения» (заканчивается на /exec).
 * 9. В index.html вставьте его в константу в начале скрипта:
 *      const FORM_ENDPOINT = 'https://script.google.com/macros/s/…/exec';
 *    Задеплойте страницу.
 * 10. Проверка: откройте URL /exec в браузере — должно показаться {"ok":true,"status":"ready"}.
 *     Затем отправьте тестовую заявку со страницы: появится строка в таблице и придёт письмо.
 *
 * ЕСЛИ МЕНЯЕТЕ КОД ПОЗЖЕ
 * «Начать развертывание» → «Управление развертываниями» → карандаш → «Версия: новая версия» →
 * «Развернуть». Так адрес /exec останется прежним. «Новое развертывание» создаст новый адрес.
 *
 * ОГРАНИЧЕНИЯ
 * • MailApp на бесплатном аккаунте Google — около 100 писем в сутки.
 * • Письма от Google могут попасть в «Спам» у получателя: проверьте папку после первой заявки.
 */

const NOTIFY_EMAIL = 'oivanova19@yandex.ru'; // почта клиники с сайта борменталь56.рф
const TIMEZONE = 'Asia/Yekaterinburg';       // Оренбург, UTC+5
const MAX_LEN = 500;
const HEADERS = ['Дата и время', 'Имя', 'Телефон', 'Сертификат', 'Комментарий', 'Статус'];

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};

    // Ловушка для спам-ботов: человек это поле не видит и не заполняет.
    if (String(p.website || '').trim() !== '') {
      return json_({ ok: true });
    }

    const name = clean_(p.name);
    const phone = clean_(p.phone);
    const certificate = clean_(p.certificate);
    const comment = clean_(p.comment);

    if (!name || !phone) {
      return json_({ ok: false, error: 'Нет имени или телефона' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];
    const date = Utilities.formatDate(new Date(), TIMEZONE, 'dd.MM.yyyy HH:mm');

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
      sheet.appendRow([date, name, phone, certificate, comment, 'Новая']);
    } finally {
      lock.releaseLock();
    }

    // Письмо — после записи в таблицу: если почта не отправится, заявка всё равно сохранена.
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'Новая заявка на сертификат: ' + plain_(certificate),
        body: [
          'Новая заявка на подарочный сертификат',
          '',
          'Имя: ' + plain_(name),
          'Телефон: ' + plain_(phone),
          'Сертификат: ' + plain_(certificate),
          'Комментарий: ' + (plain_(comment) || '—'),
          'Получена: ' + date,
          '',
          'Таблица с заявками: ' + ss.getUrl()
        ].join('\n')
      });
    } catch (mailErr) {
      console.error('Письмо не отправлено: ' + mailErr);
    }

    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

// Проверка, что веб-приложение развернуто: откройте URL /exec в браузере.
function doGet() {
  return json_({ ok: true, status: 'ready' });
}

// Обрезает до MAX_LEN символов и экранирует апострофом значения,
// которые таблица могла бы принять за формулу (= + - @).
function clean_(value) {
  let v = String(value == null ? '' : value).replace(/[\r\n\t]+/g, ' ').trim();
  if (/^[=+\-@]/.test(v)) v = "'" + v;
  return v.slice(0, MAX_LEN);
}

// Для письма убираем экранирующий апостроф.
function plain_(value) {
  return String(value || '').replace(/^'/, '');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
