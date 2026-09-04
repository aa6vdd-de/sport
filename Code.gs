
const CONFIG = {
  SPREADSHEET_ID: '18j41BcbMz0Iptpum9v8zs8-6QunlQFLkYSn6rBLkQE8',
  ADMIN_EMAIL: 'S.shehri@ut.edu.sa',
  SHEET_NAME: 'ردود النموذج 1',
  UNIVERSITY_NAME: 'جامعة تبوك',
  DEPARTMENT_NAME: 'عمادة شؤون الطلاب - النشاط الرياضي'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('نظام إدارة حجوزات الأنشطة الرياضية')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  ensureAdminColumns_(sheet);
  return sheet;
}

function ensureAdminColumns_(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  const needed = ['الحالة', 'ملاحظات القرار', 'تاريخ القرار', 'التقرير النهائي', 'روابط الصور'];
  let changed = false;
  needed.forEach(name => {
    if (!headers.includes(name)) {
      headers.push(name);
      changed = true;
    }
  });
  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function listBookings() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];
  const headers = data[0];

  return data.slice(1).map((row, i) => {
    const obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx] || '');
    obj.__row = i + 2;
    if (!obj['الحالة']) obj['الحالة'] = 'جديد';
    return obj;
  }).reverse();
}

function updateBookingStatus(rowNumber, status, note) {
  const allowed = ['مقبول', 'مرفوض'];
  if (!allowed.includes(status)) throw new Error('حالة غير صالحة.');

  const sheet = getSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const statusCol = headers.indexOf('الحالة') + 1;
  const noteCol = headers.indexOf('ملاحظات القرار') + 1;
  const dateCol = headers.indexOf('تاريخ القرار') + 1;

  sheet.getRange(rowNumber, statusCol).setValue(status);
  sheet.getRange(rowNumber, noteCol).setValue(note || '');
  sheet.getRange(rowNumber, dateCol).setValue(new Date());

  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const record = {};
  headers.forEach((h, i) => record[h] = row[i] || '');

  sendDecisionEmail_(record, status, note || '');
  return {ok: true};
}

function sendDecisionEmail_(record, status, note) {
  const email = record['الإيميل'] || record['البريد الإلكتروني'] || record['Email'] || '';
  if (!email) return;

  const name = record['الاسم'] || 'المستفيد';
  const activity = record['اسم النشاط أو الفعالية'] || record['اسم النشاط او الفعالية'] || 'النشاط الرياضي';
  const date = record['التاريخ'] || '';
  const time = record['الوقت'] || '';
  const venue = record['مقر الحجز'] || '';

  const approved = status === 'مقبول';
  const subject = approved
    ? `تم قبول طلب حجز: ${activity}`
    : `تحديث طلب حجز: ${activity}`;

  let html = `
    <div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.9">
      <h2>${CONFIG.DEPARTMENT_NAME}</h2>
      <p>السلام عليكم ${name}،</p>
      <p>نفيدكم بأن طلب حجز <strong>${activity}</strong> قد تم <strong>${approved ? 'قبوله' : 'رفضه'}</strong>.</p>
      ${venue ? `<p><strong>المقر:</strong> ${venue}</p>` : ''}
      ${date ? `<p><strong>التاريخ:</strong> ${date}</p>` : ''}
      ${time ? `<p><strong>الوقت:</strong> ${time}</p>` : ''}
      ${note ? `<p><strong>ملاحظات:</strong> ${note}</p>` : ''}
      <p>مع تحيات ${CONFIG.UNIVERSITY_NAME}<br>${CONFIG.DEPARTMENT_NAME}</p>
    </div>`;

  GmailApp.sendEmail(email, subject, 'يرجى فتح الرسالة بصيغة HTML.', {
    htmlBody: html,
    name: CONFIG.DEPARTMENT_NAME
  });
}

function saveFinalReport(rowNumber, actualParticipants, summary, notes, imageLinks) {
  const sheet = getSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const reportCol = headers.indexOf('التقرير النهائي') + 1;
  const imagesCol = headers.indexOf('روابط الصور') + 1;
  const statusCol = headers.indexOf('الحالة') + 1;

  const report = [
    `عدد المشاركين الفعلي: ${actualParticipants || ''}`,
    `ملخص التنفيذ: ${summary || ''}`,
    `الملاحظات: ${notes || ''}`
  ].join('\n');

  sheet.getRange(rowNumber, reportCol).setValue(report);
  sheet.getRange(rowNumber, imagesCol).setValue((imageLinks || []).join('\n'));
  sheet.getRange(rowNumber, statusCol).setValue('مكتمل');

  return {ok: true};
}

function onFormSubmit(e) {
  try {
    const values = e.namedValues || {};
    const name = first_(values['الاسم']);
    const activity = first_(values['اسم النشاط أو الفعالية']) || first_(values['اسم النشاط او الفعالية']);
    const venue = first_(values['مقر الحجز']);
    const date = first_(values['التاريخ']);
    const time = first_(values['الوقت']);

    const subject = `طلب حجز رياضي جديد${activity ? ': ' + activity : ''}`;
    const html = `
      <div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.9">
        <h3>طلب حجز جديد</h3>
        <p><strong>الاسم:</strong> ${name || '-'}</p>
        <p><strong>النشاط/الفعالية:</strong> ${activity || '-'}</p>
        <p><strong>المقر:</strong> ${venue || '-'}</p>
        <p><strong>التاريخ:</strong> ${date || '-'}</p>
        <p><strong>الوقت:</strong> ${time || '-'}</p>
        <p>يمكن مراجعة الطلب وقبوله أو رفضه من لوحة إدارة الحجوزات.</p>
      </div>`;
    GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, 'يوجد طلب حجز رياضي جديد.', {
      htmlBody: html,
      name: CONFIG.DEPARTMENT_NAME
    });
  } catch (err) {
    console.error(err);
  }
}

function first_(v) {
  return Array.isArray(v) ? (v[0] || '') : (v || '');
}

function setupFormSubmitTrigger() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  return 'تم إنشاء تنبيه البريد للطلبات الجديدة.';
}
