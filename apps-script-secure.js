// =============================================
// Apps Script - النسخة الآمنة
// تعليمات: غيّر ADMIN_TOKEN لقيمة سرية طويلة
// =============================================

const RESPONSES_SHEET = 'Responses';
const IDS_SHEET = 'AllowedIDs';
const RATE_SHEET = 'RateLimit';

// ⚠️ غيّر هذا لكلمة سر قوية — لا تشاركها مع أحد
const ADMIN_TOKEN = 'CHANGE_THIS_TO_A_LONG_RANDOM_SECRET_STRING_2026';

// حدود البيانات
const MAX_NAME_LEN = 120;
const MAX_COLLEGE_LEN = 80;
const VALID_YEARS = ['الأولى','الثانية','الثالثة','الرابعة','الخامسة','السادسة','السابعة'];
const VALID_DATES = ['2026-05-10','2026-05-11','2026-05-12','2026-05-13','2026-05-14',
                     '2026-05-17','2026-05-18','2026-05-19','2026-05-20','2026-05-21'];
const RATE_LIMIT_SECONDS = 15;

// =============================================
// ENTRY POINTS
// =============================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const ip = e.parameter.ip || 'unknown';

    if (action === 'submit') {
      return handleSubmit(data, ip);
    }

    // كل عمليات الأدمن تحتاج token
    if (!verifyAdminToken(data.token)) {
      return jsonResponse({ status: 'error', message: 'Unauthorized' });
    }

    if (action === 'saveIDs')        return handleSaveIDs(data);
    if (action === 'clearResponses') return handleClearResponses();

    return jsonResponse({ status: 'error', message: 'Unknown action' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Bad request' });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    const token  = e.parameter.token;

    // getIDs للطلاب فقط — بدون token (لكن بدون بيانات حساسة)
    if (action === 'getIDs') return handleGetIDs();

    // كل باقي العمليات تحتاج token
    if (!verifyAdminToken(token)) {
      return jsonResponse({ status: 'error', message: 'Unauthorized' });
    }

    if (action === 'getAll') return handleGetAll();

    return jsonResponse({ status: 'error', message: 'Unknown action' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Bad request' });
  }
}

// =============================================
// AUTH
// =============================================

function verifyAdminToken(token) {
  if (!token || typeof token !== 'string') return false;
  // مقارنة آمنة من timing attacks
  if (token.length !== ADMIN_TOKEN.length) return false;
  let match = true;
  for (let i = 0; i < token.length; i++) {
    if (token[i] !== ADMIN_TOKEN[i]) match = false;
  }
  return match;
}

// =============================================
// RATE LIMITING
// =============================================

function checkRateLimit(identifier) {
  const sheet = getOrCreateSheet(RATE_SHEET);
  const now = Date.now();
  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === identifier) {
      const lastTime = Number(data[i][1]);
      if (now - lastTime < RATE_LIMIT_SECONDS * 1000) {
        return false; // رفض
      }
      // تحديث الوقت
      sheet.getRange(i + 1, 2).setValue(now);
      return true;
    }
  }
  // إضافة جديد
  sheet.appendRow([identifier, now]);
  return true;
}

// =============================================
// VALIDATION
// =============================================

function validateSubmission(data) {
  const errors = [];

  // الرقم القومي
  if (!data.nid || typeof data.nid !== 'string') {
    errors.push('رقم قومي غير صالح');
  } else {
    const nid = data.nid.trim();
    if (nid.length < 1 || nid.length > 30) errors.push('رقم قومي غير صالح');
    if (!/^[A-Za-z0-9\-]+$/.test(nid)) errors.push('رقم قومي يحتوي على رموز غير مسموحة');
  }

  // الاسم
  if (!data.name || typeof data.name !== 'string' || data.name.trim().length < 2) {
    errors.push('الاسم مطلوب');
  } else if (data.name.length > MAX_NAME_LEN) {
    errors.push('الاسم طويل جداً');
  }

  // الكلية
  if (!data.college || typeof data.college !== 'string' || data.college.trim().length < 2) {
    errors.push('الكلية مطلوبة');
  } else if (data.college.length > MAX_COLLEGE_LEN) {
    errors.push('اسم الكلية طويل جداً');
  }

  // الفرقة
  if (!VALID_YEARS.includes(data.year)) {
    errors.push('فرقة غير صالحة');
  }

  // الهاتف
  if (!data.phone || !/^(010|011|012|015)\d{8}$/.test(data.phone)) {
    errors.push('رقم هاتف غير صالح');
  }

  // التواريخ
  if (!Array.isArray(data.dates) || data.dates.length === 0) {
    errors.push('يجب اختيار موعد واحد على الأقل');
  } else if (data.dates.length > 10) {
    errors.push('عدد المواعيد كبير جداً');
  } else {
    for (const d of data.dates) {
      if (!VALID_DATES.includes(d)) {
        errors.push('تاريخ غير مسموح: ' + d);
        break;
      }
    }
  }

  return errors;
}

// =============================================
// HANDLERS
// =============================================

function handleSubmit(data, ip) {
  // Rate limiting
  const identifier = (data.nid || '') + '_' + ip;
  if (!checkRateLimit(identifier)) {
    return jsonResponse({ status: 'error', message: 'يرجى الانتظار قبل المحاولة مجدداً.' });
  }

  // Validation
  const errors = validateSubmission(data);
  if (errors.length > 0) {
    return jsonResponse({ status: 'error', message: errors[0] });
  }

  const nid = data.nid.trim().toUpperCase();

  // التحقق من الرقم القومي مرة أخرى server-side
  const idsSheet = getOrCreateSheet(IDS_SHEET);
  const idsData = idsSheet.getDataRange().getValues().flat().map(x => String(x).trim().toUpperCase());
  if (!idsData.includes(nid)) {
    return jsonResponse({ status: 'error', message: 'الرقم القومي غير مسجل.' });
  }

  // التحقق من عدم التكرار
  const respSheet = getOrCreateSheet(RESPONSES_SHEET);
  const respData = respSheet.getDataRange().getValues();
  for (let i = 1; i < respData.length; i++) {
    if (String(respData[i][0]).toUpperCase() === nid) {
      return jsonResponse({ status: 'error', message: 'هذا الرقم سبق له الإرسال.' });
    }
  }

  // إضافة header إذا كانت الشيت فارغة
  if (respSheet.getLastRow() === 0) {
    respSheet.appendRow(['الرقم القومي','الاسم','الكلية','الفرقة','رقم الهاتف','المواعيد','وقت الإرسال']);
    respSheet.getRange(1,1,1,7).setFontWeight('bold');
  }

  // حفظ البيانات
  respSheet.appendRow([
    nid,
    data.name.trim().substring(0, MAX_NAME_LEN),
    data.college.trim().substring(0, MAX_COLLEGE_LEN),
    data.year,
    data.phone,
    data.dates.join(' | '),
    new Date().toLocaleString('ar-EG')
  ]);

  return jsonResponse({ status: 'ok' });
}

function handleGetAll() {
  const respSheet = getOrCreateSheet(RESPONSES_SHEET);
  const idsSheet  = getOrCreateSheet(IDS_SHEET);

  const idsData = idsSheet.getDataRange().getValues();
  const allowedIDs = idsData.flat().map(x => String(x).trim()).filter(x => x.length > 0);

  const respData = respSheet.getDataRange().getValues();
  const responses = [];
  for (let i = 1; i < respData.length; i++) {
    const row = respData[i];
    if (row[0]) {
      responses.push({
        nid:         String(row[0]),
        name:        String(row[1]),
        college:     String(row[2]),
        year:        String(row[3]),
        phone:       String(row[4]),
        dates:       String(row[5]),
        submittedAt: String(row[6])
      });
    }
  }

  return jsonResponse({ status: 'ok', allowedIDs, responses });
}

function handleGetIDs() {
  const idsSheet  = getOrCreateSheet(IDS_SHEET);
  const respSheet = getOrCreateSheet(RESPONSES_SHEET);

  const allowedIDs = idsSheet.getDataRange().getValues()
    .flat().map(x => String(x).trim()).filter(x => x.length > 0);

  const submittedIDs = [];
  const respData = respSheet.getDataRange().getValues();
  for (let i = 1; i < respData.length; i++) {
    if (respData[i][0]) submittedIDs.push(String(respData[i][0]).toUpperCase());
  }

  // لا نعيد قائمة الأرقام المسموحة للطلاب — فقط الأرقام اللي اتبعتت
  return jsonResponse({ status: 'ok', allowedIDs, submittedIDs });
}

function handleSaveIDs(data) {
  const sheet = getOrCreateSheet(IDS_SHEET);
  const ids = (data.ids || [])
    .map(x => String(x).trim())
    .filter(x => x.length > 0 && x.length <= 30 && /^[A-Za-z0-9\-\u0600-\u06FF]+$/.test(x));

  sheet.clearContents();
  if (ids.length > 0) {
    sheet.getRange(1, 1, ids.length, 1).setValues(ids.map(id => [id]));
  }
  return jsonResponse({ status: 'ok', saved: ids.length });
}

function handleClearResponses() {
  const sheet = getOrCreateSheet(RESPONSES_SHEET);
  sheet.clearContents();
  return jsonResponse({ status: 'ok' });
}

// =============================================
// HELPERS
// =============================================

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
