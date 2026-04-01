/**
 * ═══════════════════════════════════════════════════════════════
 *  KAUH Preoperative Anxiety Screening Platform — Backend
 *  Google Apps Script (Code.gs)
 *  King Abdulaziz University Hospital
 * ═══════════════════════════════════════════════════════════════
 *
 *  DEPLOYMENT INSTRUCTIONS:
 *  1. Open Google Sheets → Extensions → Apps Script
 *  2. Paste this entire file contents into Code.gs
 *  3. Save, then Deploy → New Deployment
 *  4. Type: Web App | Execute as: Me | Access: Anyone
 *  5. Copy the Web App URL and paste it in index.html (GAS_URL)
 * ═══════════════════════════════════════════════════════════════
 */

/* ── CONFIGURATION ─────────────────────────────────────────────────── */
const CONFIG = {
  SPREADSHEET_ID:       SpreadsheetApp.getActiveSpreadsheet().getId(),
  MAIN_SHEET_NAME:      'Patient Records',
  HIGH_ANXIETY_SHEET:   'Flagged Patients',
  LOG_SHEET_NAME:       'Submission Log',
  APAIS_THRESHOLD:      11,       // Score >= this → HIGH anxiety
  SOCIAL_WORKER_EMAIL:  'socialwork@kauh.edu.sa',   // Replace with actual email
  ADMIN_EMAIL:          'admin@kauh.edu.sa',         // Replace with actual email
  SEND_EMAILS:          true,     // Set to false during testing
  TIMEZONE:             'Asia/Riyadh'
};

/* ── SHEET COLUMN HEADERS ──────────────────────────────────────────── */
const MAIN_HEADERS = [
  'Timestamp',
  'Patient ID',
  'Full Name',
  'Date of Birth',
  'Scheduled Procedure',
  'APAIS Score',
  'Anxiety Level',
  'Consent Signed',
  'Submission IP',
  'Reference Number'
];

const FLAG_HEADERS = [
  'Timestamp',
  'Reference Number',
  'Patient ID',
  'Full Name',
  'Scheduled Procedure',
  'APAIS Score',
  'Anxiety Level',
  'Status',
  'Social Worker Assigned',
  'Notes'
];


/* ══════════════════════════════════════════════════════════════════════
 *  ENTRY POINTS
 * ══════════════════════════════════════════════════════════════════════ */

/**
 * Handle GET requests — used for CORS preflight & dashboard reads
 */
function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'getPatientData') {
    return handleGetPatientData(e);
  }

  if (action === 'getDashboard') {
    return handleGetDashboard(e);
  }

  // Health check
  return buildResponse({ status: 'ok', message: 'KAUH Preoperative Platform API is running.' });
}

/**
 * Handle POST requests — main submission endpoint
 */
function doPost(e) {
  try {
    ensureSheetsExist();

    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return buildResponse({ success: false, error: 'Invalid JSON payload.' }, 400);
    }

    // Validate required fields
    const validation = validatePayload(payload);
    if (!validation.valid) {
      return buildResponse({ success: false, error: validation.message }, 400);
    }

    const result = submitPatientData(payload);
    return buildResponse(result);

  } catch (err) {
    logError(err, e.postData ? e.postData.contents : 'No payload');
    return buildResponse({ success: false, error: 'Internal server error. Please contact support.' }, 500);
  }
}


/* ══════════════════════════════════════════════════════════════════════
 *  CORE FUNCTIONS
 * ══════════════════════════════════════════════════════════════════════ */

/**
 * Main submission handler — writes to sheets and triggers notifications
 */
function submitPatientData(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const mainSheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  const refNumber = generateReferenceNumber(data.patientId);
  const anxietyLevel = Number(data.apaisScore) >= CONFIG.APAIS_THRESHOLD ? 'HIGH' : 'NORMAL';

  // Write to main patient records sheet
  mainSheet.appendRow([
    timestamp,
    sanitize(data.patientId),
    sanitize(data.name),
    sanitize(data.dob || ''),
    sanitize(data.procedure),
    Number(data.apaisScore),
    anxietyLevel,
    data.consentSigned === 'YES' ? 'YES' : 'NO',
    'Web App',
    refNumber
  ]);

  // Colour-code the row based on anxiety level
  const lastRow = mainSheet.getLastRow();
  if (anxietyLevel === 'HIGH') {
    mainSheet.getRange(lastRow, 1, 1, MAIN_HEADERS.length)
      .setBackground('#FFF0F0')
      .setFontColor('#7B0000');
  } else {
    mainSheet.getRange(lastRow, 1, 1, MAIN_HEADERS.length)
      .setBackground('#F0FFF4');
  }

  // If high anxiety → write to flagged sheet & send notifications
  if (anxietyLevel === 'HIGH') {
    flagHighAnxietyPatient(ss, data, timestamp, refNumber);
    if (CONFIG.SEND_EMAILS) {
      sendSocialWorkerAlert(data, refNumber, Number(data.apaisScore));
    }
  }

  // Write to submission log
  logSubmission(ss, data, timestamp, refNumber, anxietyLevel);

  return {
    success: true,
    referenceNumber: refNumber,
    anxietyLevel: anxietyLevel,
    message: anxietyLevel === 'HIGH'
      ? 'Your submission was received. A support team member will contact you.'
      : 'Your submission was received. You are cleared for your procedure.'
  };
}

/**
 * Write high-anxiety patient to the Flagged Patients sheet
 */
function flagHighAnxietyPatient(ss, data, timestamp, refNumber) {
  const flagSheet = ss.getSheetByName(CONFIG.HIGH_ANXIETY_SHEET);
  flagSheet.appendRow([
    timestamp,
    refNumber,
    sanitize(data.patientId),
    sanitize(data.name),
    sanitize(data.procedure),
    Number(data.apaisScore),
    'HIGH',
    'PENDING REVIEW',
    '',
    ''
  ]);

  // Bold and colour the new row
  const lastRow = flagSheet.getLastRow();
  flagSheet.getRange(lastRow, 1, 1, FLAG_HEADERS.length)
    .setBackground('#FFF5E6')
    .setFontWeight('normal');
  flagSheet.getRange(lastRow, 7)
    .setFontColor('#B30000')
    .setFontWeight('bold');
}

/**
 * Send email alert to social worker
 */
function sendSocialWorkerAlert(data, refNumber, score) {
  const subject = `[KAUH] High Anxiety Patient Flagged — Ref: ${refNumber}`;
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: #00562B; padding: 20px 24px; border-radius: 8px 8px 0 0;">
        <h2 style="color: white; margin: 0; font-size: 18px;">King Abdulaziz University Hospital</h2>
        <p style="color: rgba(255,255,255,0.8); margin: 4px 0 0; font-size: 13px;">Preoperative Patient Support Alert</p>
      </div>
      <div style="background: #fff8f8; border: 1px solid #fecaca; padding: 24px; border-radius: 0 0 8px 8px;">
        <div style="background: #DC2626; color: white; padding: 12px 18px; border-radius: 6px; margin-bottom: 24px;">
          <strong>HIGH ANXIETY — IMMEDIATE ATTENTION REQUIRED</strong>
        </div>
        <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
          <tr><td style="padding: 8px 12px; font-weight: bold; color: #555; width: 40%;">Reference Number</td><td style="padding: 8px 12px; color: #111;">${refNumber}</td></tr>
          <tr style="background: #fafafa;"><td style="padding: 8px 12px; font-weight: bold; color: #555;">Patient Name</td><td style="padding: 8px 12px; color: #111;">${sanitize(data.name)}</td></tr>
          <tr><td style="padding: 8px 12px; font-weight: bold; color: #555;">Patient ID</td><td style="padding: 8px 12px; color: #111;">${sanitize(data.patientId)}</td></tr>
          <tr style="background: #fafafa;"><td style="padding: 8px 12px; font-weight: bold; color: #555;">Scheduled Procedure</td><td style="padding: 8px 12px; color: #111;">${sanitize(data.procedure)}</td></tr>
          <tr><td style="padding: 8px 12px; font-weight: bold; color: #555;">APAIS Score</td><td style="padding: 8px 12px; color: #DC2626; font-weight: bold;">${score} / 30 (Threshold: ${CONFIG.APAIS_THRESHOLD})</td></tr>
          <tr style="background: #fafafa;"><td style="padding: 8px 12px; font-weight: bold; color: #555;">Submission Time</td><td style="padding: 8px 12px; color: #111;">${Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd MMM yyyy, HH:mm')}</td></tr>
        </table>
        <div style="margin-top: 24px; padding: 16px; background: #fffbeb; border: 1px solid #fde68a; border-radius: 6px;">
          <strong style="font-size: 13px;">Action Required:</strong>
          <p style="font-size: 13px; margin: 6px 0 0;">Please contact this patient before their scheduled procedure to provide additional emotional support and counselling.</p>
        </div>
        <p style="font-size: 12px; color: #999; margin-top: 20px;">This is an automated notification from the KAUH Preoperative Patient Platform.</p>
      </div>
    </div>
  `;

  GmailApp.sendEmail(CONFIG.SOCIAL_WORKER_EMAIL, subject, '', { htmlBody });

  // Also notify admin
  GmailApp.sendEmail(
    CONFIG.ADMIN_EMAIL,
    `[KAUH] Copy: High Anxiety Patient — Ref: ${refNumber}`,
    `Patient ${sanitize(data.name)} (ID: ${sanitize(data.patientId)}) scored ${score}/30 on APAIS. Ref: ${refNumber}`,
    { htmlBody }
  );
}


/* ══════════════════════════════════════════════════════════════════════
 *  DASHBOARD / REPORTING
 * ══════════════════════════════════════════════════════════════════════ */

/**
 * GET /getPatientData?patientId=XXXXXXXXXX
 * Returns patient record for dashboard lookups
 */
function handleGetPatientData(e) {
  const patientId = e.parameter.patientId || '';
  if (!patientId) {
    return buildResponse({ success: false, error: 'patientId parameter required.' });
  }
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const records = [];
  for (let i = 1; i < data.length; i++) { // skip header row
    if (String(data[i][1]) === String(patientId)) {
      records.push({
        timestamp:     data[i][0],
        patientId:     data[i][1],
        name:          data[i][2],
        dob:           data[i][3],
        procedure:     data[i][4],
        apaisScore:    data[i][5],
        anxietyLevel:  data[i][6],
        consentSigned: data[i][7],
        refNumber:     data[i][9]
      });
    }
  }

  return buildResponse({ success: true, records });
}

/**
 * GET /getDashboard
 * Returns aggregate stats for admin dashboard
 */
function handleGetDashboard(e) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  let total = 0, highCount = 0, normalCount = 0, consentCount = 0;
  const procedureCounts = {};
  const scoreDistribution = Array(31).fill(0);

  for (let i = 1; i < data.length; i++) {
    total++;
    const level = data[i][6];
    const score = Number(data[i][5]);
    const proc  = data[i][4];
    const consent = data[i][7];

    if (level === 'HIGH')   highCount++;
    if (level === 'NORMAL') normalCount++;
    if (consent === 'YES')  consentCount++;
    procedureCounts[proc] = (procedureCounts[proc] || 0) + 1;
    if (score >= 0 && score <= 30) scoreDistribution[score]++;
  }

  return buildResponse({
    success: true,
    stats: {
      total,
      highAnxiety: highCount,
      normalAnxiety: normalCount,
      consentSigned: consentCount,
      highAnxietyRate: total > 0 ? ((highCount / total) * 100).toFixed(1) : 0,
      procedureBreakdown: procedureCounts,
      scoreDistribution
    }
  });
}


/* ══════════════════════════════════════════════════════════════════════
 *  SHEET MANAGEMENT
 * ══════════════════════════════════════════════════════════════════════ */

/**
 * Ensure all required sheets exist with proper headers
 */
function ensureSheetsExist() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  createSheetIfMissing(ss, CONFIG.MAIN_SHEET_NAME,    MAIN_HEADERS, '#00562B');
  createSheetIfMissing(ss, CONFIG.HIGH_ANXIETY_SHEET, FLAG_HEADERS, '#B30000');
  createSheetIfMissing(ss, CONFIG.LOG_SHEET_NAME,
    ['Timestamp','Ref Number','Patient ID','Score','Anxiety Level','Consent','User Agent'], '#1E3A5F');
}

function createSheetIfMissing(ss, name, headers, headerBgColor) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange
      .setBackground(headerBgColor)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160); // Timestamp
    sheet.setColumnWidth(2, 130); // ID
    sheet.setColumnWidth(3, 200); // Name
  }
  return sheet;
}

function logSubmission(ss, data, timestamp, refNumber, anxietyLevel) {
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  logSheet.appendRow([
    timestamp,
    refNumber,
    sanitize(data.patientId),
    Number(data.apaisScore),
    anxietyLevel,
    data.consentSigned,
    'Web App'
  ]);
}

function logError(err, payload) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let errSheet = ss.getSheetByName('Error Log');
    if (!errSheet) {
      errSheet = ss.insertSheet('Error Log');
      errSheet.appendRow(['Timestamp', 'Error', 'Payload']);
    }
    errSheet.appendRow([new Date().toISOString(), err.toString(), String(payload).substring(0, 500)]);
  } catch (e) {
    // fail silently
  }
}


/* ══════════════════════════════════════════════════════════════════════
 *  UTILITIES
 * ══════════════════════════════════════════════════════════════════════ */

/**
 * Input sanitisation — strips HTML/script tags
 */
function sanitize(value) {
  if (typeof value !== 'string') return String(value || '');
  return value.replace(/<[^>]*>/g, '').replace(/[<>'"]/g, '').trim().substring(0, 500);
}

/**
 * Generate a unique reference number
 */
function generateReferenceNumber(patientId) {
  const now = new Date();
  const datePart = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyyMMdd');
  const timePart = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HHmmss');
  const idSuffix = String(patientId).slice(-4);
  return `KAUH-${datePart}-${timePart}-${idSuffix}`;
}

/**
 * Validate incoming payload fields
 */
function validatePayload(data) {
  if (!data.patientId || !/^\d{10}$/.test(String(data.patientId))) {
    return { valid: false, message: 'Invalid or missing patientId (must be 10 digits).' };
  }
  if (!data.name || String(data.name).trim().length < 2) {
    return { valid: false, message: 'Invalid or missing patient name.' };
  }
  if (!data.procedure || String(data.procedure).trim().length < 2) {
    return { valid: false, message: 'Invalid or missing procedure.' };
  }
  const score = Number(data.apaisScore);
  if (isNaN(score) || score < 6 || score > 30) {
    return { valid: false, message: 'APAIS score must be between 6 and 30.' };
  }
  if (data.consentSigned !== 'YES') {
    return { valid: false, message: 'Consent must be signed (YES).' };
  }
  return { valid: true };
}

/**
 * Build a JSON response with CORS headers
 */
function buildResponse(data, statusCode) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}


/* ══════════════════════════════════════════════════════════════════════
 *  SCHEDULED REPORTS  (optional — attach to a time-based trigger)
 * ══════════════════════════════════════════════════════════════════════ */

/**
 * Daily digest email of high-anxiety patients
 * To enable: Triggers → Add Trigger → dailyReport → Time-based → Day timer
 */
function dailyReport() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const flagSheet = ss.getSheetByName(CONFIG.HIGH_ANXIETY_SHEET);
  const data = flagSheet.getDataRange().getValues();

  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const todayRows = data.slice(1).filter(row => String(row[0]).startsWith(today));

  if (todayRows.length === 0) return;

  let tableRows = todayRows.map(row => `
    <tr>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;">${row[1]}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;">${row[3]}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;">${row[2]}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;">${row[4]}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;color:#DC2626;font-weight:bold;">${row[5]}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;">${row[7]}</td>
    </tr>
  `).join('');

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:700px;">
      <div style="background:#00562B;padding:20px 24px;border-radius:8px 8px 0 0;">
        <h2 style="color:white;margin:0;font-size:18px;">KAUH — Daily High Anxiety Report</h2>
        <p style="color:rgba(255,255,255,0.7);margin:4px 0 0;font-size:12px;">${today}</p>
      </div>
      <div style="padding:24px;border:1px solid #e5e7eb;border-top:none;">
        <p style="font-size:14px;margin-bottom:20px;">${todayRows.length} high-anxiety patient(s) flagged today:</p>
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <thead>
            <tr style="background:#f9fafb;">
              <th style="padding:10px 12px;text-align:left;color:#555;">Ref</th>
              <th style="padding:10px 12px;text-align:left;color:#555;">Name</th>
              <th style="padding:10px 12px;text-align:left;color:#555;">ID</th>
              <th style="padding:10px 12px;text-align:left;color:#555;">Procedure</th>
              <th style="padding:10px 12px;text-align:left;color:#555;">Score</th>
              <th style="padding:10px 12px;text-align:left;color:#555;">Status</th>
            </tr>
          </thead>
          <tbody>${tableRows}</tbody>
        </table>
      </div>
    </div>
  `;

  GmailApp.sendEmail(
    CONFIG.SOCIAL_WORKER_EMAIL,
    `[KAUH] Daily Anxiety Report — ${todayRows.length} High-Anxiety Patient(s) — ${today}`,
    '',
    { htmlBody }
  );
}
