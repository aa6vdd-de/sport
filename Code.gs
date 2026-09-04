const PROJECT_ID = "sport-d2cb8";
const COLLECTION = "bookings";
const ADMIN_EMAIL = "S.shehri@ut.edu.sa";

/**
 * أي رد جديد من Google Form يُرسل تلقائياً إلى Firestore.
 * هذه الدالة لا تُشغّل يدوياً.
 */
function onFormSubmit(e) {
  if (!e || !e.range) {
    throw new Error("لا تشغّل onFormSubmit يدوياً. استخدم setupTrigger مرة واحدة فقط.");
  }

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  sendRowToFirebase_(sheet, row);

  const data = getRowData_(sheet, row);
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: "طلب حجز نشاط رياضي جديد",
    htmlBody: `
      <div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.9">
        <h2>طلب حجز نشاط رياضي جديد</h2>
        <p><b>الاسم:</b> ${escapeHtml_(data.name || "-")}</p>
        <p><b>النشاط أو الفعالية:</b> ${escapeHtml_(data.activityName || "-")}</p>
        <p><b>البريد الإلكتروني:</b> ${escapeHtml_(data.email || "-")}</p>
        <p><b>رقم الجوال:</b> ${escapeHtml_(data.phone || "-")}</p>
        <p><b>عدد المشاركين:</b> ${escapeHtml_(data.participantsCount || "-")}</p>
        <p><b>الفئة:</b> ${escapeHtml_(data.category || "-")}</p>
        <p><b>مقر الحجز:</b> ${escapeHtml_(data.venue || "-")}</p>
        <p><b>طريقة الحجز:</b> ${escapeHtml_(data.bookingMethod || "-")}</p>
        <p><b>التاريخ:</b> ${escapeHtml_(data.date || "-")}</p>
        <p><b>الوقت:</b> ${escapeHtml_(data.time || "-")}</p>
      </div>`
  });
}

/**
 * شغّلها مرة واحدة فقط بعد تحديث هذا الكود لإصلاح السجلات القديمة.
 * بعد ذلك، الـ Trigger يتولى كل الطلبات الجديدة تلقائياً.
 */
function syncExistingRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  for (let row = 2; row <= lastRow; row++) {
    sendRowToFirebase_(sheet, row);
  }
  Logger.log("تم تحديث جميع الطلبات الموجودة في Firebase.");
}

/**
 * شغّلها مرة واحدة فقط لإنشاء Trigger تلقائي عند وصول رد جديد.
 */
function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "onFormSubmit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log("تم إنشاء Trigger بنجاح.");
}

/**
 * Web App endpoint:
 * - accept: قبول + إرسال بريد
 * - reject: رفض + سبب + إرسال بريد
 *
 * الموقع يستدعيه بصيغة JSONP لتجنب مشكلة CORS بين GitHub Pages وApps Script.
 */
function doGet(e) {
  const callback = safeCallback_(e && e.parameter && e.parameter.callback);
  try {
    const action = String((e && e.parameter && e.parameter.action) || "").trim();
    if (!["accept", "reject", "hide"].includes(action)) {
      return jsonp_(callback, {ok: true, message: "Sports booking API is running."});
    }

    const id = String(e.parameter.id || "").trim();
    if (!id) throw new Error("معرّف الطلب مفقود.");

    const reason = String(e.parameter.reason || "").trim();

    if (action === "hide") {
      hideBooking_(id);
      return jsonp_(callback, {ok: true, status: "hidden"});
    }

    if (action === "reject" && !reason) {
      throw new Error("سبب الرفض مطلوب.");
    }

    const result = processDecision_(id, action, reason);
    return jsonp_(callback, result);

  } catch (err) {
    return jsonp_(callback, {ok: false, error: String(err && err.message ? err.message : err)});
  }
}

function processDecision_(documentId, action, reason) {
  const record = getFirestoreDocument_(documentId);
  const email = fieldString_(record.fields, "email");
  const name = fieldString_(record.fields, "name") || "مقدم الطلب";
  const activity = fieldString_(record.fields, "activityName") || "النشاط أو الفعالية";
  const venue = fieldString_(record.fields, "venue");
  const date = fieldString_(record.fields, "date");
  const time = fieldString_(record.fields, "time");

  if (!email) {
    throw new Error("البريد الإلكتروني غير موجود في بيانات هذا الطلب. شغّل syncExistingRows مرة واحدة بعد تحديث الكود.");
  }

  const accepted = action === "accept";
  const status = accepted ? "مقبول" : "مرفوض";

  updateFirestoreDecision_(documentId, status, accepted ? "" : reason);

  const subject = accepted
    ? `تم قبول طلب حجز النشاط الرياضي - ${activity}`
    : `تم رفض طلب حجز النشاط الرياضي - ${activity}`;

  const decisionBlock = accepted
    ? `<div style="background:#e8f6ee;border-radius:10px;padding:12px;color:#176d45"><b>حالة الطلب: مقبول</b></div>`
    : `<div style="background:#fff0ef;border-radius:10px;padding:12px;color:#b33b34"><b>حالة الطلب: مرفوض</b><br><b>سبب الرفض:</b> ${escapeHtml_(reason)}</div>`;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: `
      <div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.9;max-width:650px;margin:auto">
        <h2 style="color:#0b6b45">إدارة النشاط الرياضي - عمادة شؤون الطلاب</h2>
        <p>السلام عليكم ${escapeHtml_(name)}،</p>
        <p>نفيدكم بتحديث حالة طلب حجز النشاط أو الفعالية الموضح أدناه:</p>
        ${decisionBlock}
        <p><b>اسم النشاط أو الفعالية:</b> ${escapeHtml_(activity)}</p>
        ${venue ? `<p><b>مقر الحجز:</b> ${escapeHtml_(venue)}</p>` : ""}
        ${date ? `<p><b>التاريخ:</b> ${escapeHtml_(date)}</p>` : ""}
        ${time ? `<p><b>الوقت:</b> ${escapeHtml_(time)}</p>` : ""}
        <p>مع تحيات<br><b>إدارة النشاط الرياضي<br>عمادة شؤون الطلاب - جامعة تبوك</b></p>
      </div>`
  });

  return {ok: true, status: status};
}

function sendRowToFirebase_(sheet, row) {
  const data = getRowData_(sheet, row);
  const documentId = "row_" + row;

  const url = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/${COLLECTION}/${encodeURIComponent(documentId)}`;

  let existingHidden = false;
  try {
    const existing = getFirestoreDocument_(documentId);
    existingHidden =
      existing &&
      existing.fields &&
      existing.fields.hidden &&
      existing.fields.hidden.booleanValue === true;
  } catch (_) {
    existingHidden = false;
  }

  const fields = {
    name: {stringValue: String(data.name || "")},
    email: {stringValue: String(data.email || "")},
    phone: {stringValue: String(data.phone || "")},
    activityName: {stringValue: String(data.activityName || "")},
    participantsCount: {stringValue: String(data.participantsCount || "")},
    category: {stringValue: String(data.category || "")},
    venue: {stringValue: String(data.venue || "")},
    bookingMethod: {stringValue: String(data.bookingMethod || "")},
    date: {stringValue: String(data.date || "")},
    time: {stringValue: String(data.time || "")},
    status: {stringValue: "جديد"},
    sheetRow: {integerValue: String(row)},
    createdAt: {timestampValue: new Date().toISOString()}
  };

  const response = UrlFetchApp.fetch(url, {
    method: "patch",
    contentType: "application/json",
    payload: JSON.stringify({fields: fields}),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`فشل رفع الصف ${row} إلى Firestore: ${response.getContentText()}`);
  }
}

/**
 * قراءة الأعمدة بمواقعها الثابتة A:J.
 * هذا يتجنب اختلاف كتابة "الأيميل/الإيميل" أو أي مسافات في عناوين Google Form.
 */
function getRowData_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, Math.max(10, sheet.getLastColumn())).getDisplayValues()[0];

  return {
    timestamp: values[0] || "",          // A
    name: values[1] || "",               // B
    email: values[2] || "",              // C
    phone: values[3] || "",              // D
    activityName: values[4] || "",       // E
    participantsCount: findParticipantsByHeader_(sheet, row),
    category: values[5] || "",           // F
    venue: values[6] || "",              // G
    bookingMethod: values[7] || "",      // H
    date: values[8] || "",               // I
    time: values[9] || ""                // J
  };
}

/**
 * عدد المشاركين قد يكون بعد العمود J في النموذج الحالي أو يضاف لاحقاً،
 * لذلك نبحث عنه بالعنوان.
 */
function findParticipantsByHeader_(sheet, row) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastColumn).getDisplayValues()[0];
  const values = sheet.getRange(row,1,1,lastColumn).getDisplayValues()[0];

  for (let i=0; i<headers.length; i++) {
    const h = normalizeArabic_(headers[i]);
    if (h.includes("عدد المشاركين") || h.includes("عدد المستفيدين")) {
      return values[i] || "";
    }
  }
  return "";
}


/**
 * إخفاء الطلب من الموقع فقط.
 * الطلب يبقى محفوظاً في Firestore وGoogle Sheet.
 */
function hideBooking_(documentId) {
  const base =
    `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/${COLLECTION}/${encodeURIComponent(documentId)}`;

  const url = base + "?updateMask.fieldPaths=hidden&updateMask.fieldPaths=hiddenAt";

  const body = {
    fields: {
      hidden: { booleanValue: true },
      hiddenAt: { timestampValue: new Date().toISOString() }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: "patch",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();

  if (responseCode < 200 || responseCode >= 300) {
    throw new Error(
      "تعذر إخفاء الطلب من الموقع: " + response.getContentText()
    );
  }
}

function getFirestoreDocument_(documentId) {
  const url = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/${COLLECTION}/${encodeURIComponent(documentId)}`;
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions:true});
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("تعذر قراءة الطلب من Firestore.");
  }
  return JSON.parse(response.getContentText());
}

function updateFirestoreDecision_(documentId, status, rejectionReason) {
  const base = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/${COLLECTION}/${encodeURIComponent(documentId)}`;
  const url = base
    + `?updateMask.fieldPaths=status`
    + `&updateMask.fieldPaths=rejectionReason`
    + `&updateMask.fieldPaths=decisionUpdatedAt`;

  const body = {
    fields: {
      status: {stringValue: status},
      rejectionReason: {stringValue: rejectionReason || ""},
      decisionUpdatedAt: {timestampValue: new Date().toISOString()}
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method:"patch",
    contentType:"application/json",
    payload:JSON.stringify(body),
    muteHttpExceptions:true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("تعذر تحديث حالة الطلب في Firestore.");
  }
}

function fieldString_(fields, key) {
  const f = fields && fields[key];
  if (!f) return "";
  return String(
    f.stringValue !== undefined ? f.stringValue :
    f.integerValue !== undefined ? f.integerValue :
    f.doubleValue !== undefined ? f.doubleValue : ""
  );
}

function safeCallback_(name) {
  const cb = String(name || "callback");
  return /^[A-Za-z_$][A-Za-z0-9_$.]*$/.test(cb) ? cb : "callback";
}

function jsonp_(callback, data) {
  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify(data)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function normalizeArabic_(s) {
  return String(s || "")
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/\s+/g, " ")
    .trim();
}

function escapeHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
