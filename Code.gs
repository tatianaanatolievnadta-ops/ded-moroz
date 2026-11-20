/**
 * Приём заявок с сайта «Мультик с Дедом Морозом».
 * Заявка → строка в Google Таблице, фото → папка на Google Диске, уведомление → Telegram.
 * Как подключить — см. файл «ЗАПУСК САЙТА.txt».
 */

// ================== НАСТРОЙКИ ==================
const TELEGRAM_TOKEN   = '';  // токен бота от @BotFather (можно оставить пустым — тогда без уведомлений)
const TELEGRAM_CHAT_ID = '';  // ваш chat id (запустите функцию getChatId, см. инструкцию)
const FOLDER_NAME      = 'Заказы — мультики Дед Мороз';
// ===============================================

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const tz = Session.getScriptTimeZone();
    const id = 'DM-' + Utilities.formatDate(new Date(), tz, 'MMdd-HHmmss');

    // фото в отдельную папку заказа
    const folder = getFolder_().createFolder(id + ' ' + String(d.childName || '').slice(0, 40));
    let firstKid = null;
    (d.files || []).slice(0, 6).forEach(function (f, i) {
      const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.type || 'image/jpeg', f.name || ('photo_' + (i + 1) + '.jpg'));
      const file = folder.createFile(blob);
      if (!firstKid && /ребенок|rebenok/i.test(f.name || '')) firstKid = file;
    });

    sheet.appendRow([
      new Date(), id, d.childName, d.age, d.packageTitle, d.style,
      d.price, d.discount, d.total, d.promo, d.wishes,
      d.contact, d.contactType, d.deadline, folder.getUrl(), 'Новая'
    ]);

    notify_(id, d, folder.getUrl(), firstKid);
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Приём заявок работает ✔');
}

// ---------- служебное ----------
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Заявки');
  if (!sh) {
    sh = ss.insertSheet('Заявки');
    sh.appendRow(['Дата', '№ заказа', 'Имя ребёнка', 'Возраст', 'Формат', 'Стиль',
      'Цена', 'Скидка %', 'К оплате', 'Промокод', 'Пожелания', 'Контакт', 'Куда прислать',
      'К дате', 'Папка с фото', 'Статус']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold').setBackground('#121212').setFontColor('#C9A45C');
    sh.getRange('P2:P1000').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Новая', 'Ждёт оплату', 'Оплачено', 'В работе', 'Отправлено', 'Отказ'], true).build());
  }
  return sh;
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function notify_(id, d, folderUrl, photo) {
  if (!TELEGRAM_TOKEN || !TELEGRAM_CHAT_ID) return;
  const text =
    '🎅 Новая заявка ' + id + '\n' +
    '👶 ' + (d.childName || '') + (d.age ? ', ' + d.age : '') + '\n' +
    '🎬 ' + (d.packageTitle || '') + ' · ' + (d.style || '') + '\n' +
    '💰 ' + d.total + ' ₽' + (d.discount ? ' (скидка ' + d.discount + '%)' : '') + '\n' +
    '📅 К дате: ' + (d.deadline || 'не указано') + '\n' +
    '📩 ' + (d.contact || '') + ' → ' + (d.contactType || '') + '\n\n' +
    '📝 ' + (d.wishes || '—') + '\n\n' +
    '📁 Фото: ' + folderUrl;
  const base = 'https://api.telegram.org/bot' + TELEGRAM_TOKEN;
  try {
    if (photo) {
      UrlFetchApp.fetch(base + '/sendPhoto', { method: 'post', muteHttpExceptions: true,
        payload: { chat_id: TELEGRAM_CHAT_ID, photo: photo.getBlob(), caption: text.slice(0, 1000) } });
    } else {
      UrlFetchApp.fetch(base + '/sendMessage', { method: 'post', muteHttpExceptions: true,
        payload: { chat_id: TELEGRAM_CHAT_ID, text: text } });
    }
  } catch (e) { /* уведомление не критично */ }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Запустите один раз после того, как написали своему боту любое сообщение. Chat id появится в «Журнале выполнения». */
function getChatId() {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_TOKEN + '/getUpdates', { muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText());
  (j.result || []).forEach(function (u) {
    const m = u.message || u.channel_post;
    if (m) Logger.log('chat id: ' + m.chat.id + '  (' + (m.chat.username || m.chat.title || m.chat.first_name) + ')');
  });
  if (!(j.result || []).length) Logger.log('Пусто. Напишите боту /start и запустите ещё раз.');
}

/** Проверка: создаёт тестовую заявку без фото. */
function testOrder() {
  const res = doPost({ postData: { contents: JSON.stringify({
    childName: 'Тест', age: '4 года', packageTitle: 'Мультик-история', style: '3D-мультик',
    price: 4490, discount: 10, total: 4040, promo: 'MOROZ10', wishes: 'Тестовая заявка',
    contact: '@test', contactType: 'Telegram', deadline: '', files: [] }) } });
  Logger.log(res.getContent());
}
