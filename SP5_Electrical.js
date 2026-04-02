// ═══════════════════════════════════════════════════════════════
// SP5 ELECTRICAL MANAGEMENT SYSTEM — Google Apps Script Backend
// File: Code.gs
// ═══════════════════════════════════════════════════════════════

// ── CONFIG: Apni Google Sheet ka ID yahan paste karo ────────────
const SHEET_ID = '1xwW3dLD13a3mwBwN_XAx_l0F4ndgUU8LBpYuScNoVfQ';
// Sheet ID milega URL mein: docs.google.com/spreadsheets/d/[THIS_PART]/edit

// ── Sheet names ──────────────────────────────────────────────────
const SHEETS = {
  MOTORS:      'MotorMaster',
  EMPLOYEES:   'EmployeeMaster',
  MAINTENANCE: 'MotorMaintenance',
  GREASING:    'MotorGreasing',
  REPAIR:      'MotorRepair',
  BREAKDOWN:   'MotorBreakdownHistory',
  LOCATION_H:  'MotorLocationHistory',
  TRANSFORMER: 'TransformerMaster',
  TR_MAINT:    'TransformerMaintenance',
  IOP:         'IOP_Data',
  TELEPHONE:   'TelephoneDirectory',
  SPARE:       'SpareMotorMaster',
  VFD:         'VFD_Details',
  VENDOR:      'VendorMaster',
  USERS:       'Users',
  SHIFT:       'ShiftMaster',
  AREA:        'AreaMaster',
  LOCATION:    'LocationMaster',
};

// ── Web App Entry Point ──────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SP5 Electrical Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ── HTML include helper ──────────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ════════════════════════════════════════════════════════════════
// AUTH — Login check
// ════════════════════════════════════════════════════════════════
function login(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const userIdx  = headers.indexOf('UserName');
    const passIdx  = headers.indexOf('Password');
    const roleIdx  = headers.indexOf('Role');
    const deptIdx  = headers.indexOf('Department');
    const statIdx  = headers.indexOf('Status');
    const idIdx    = headers.indexOf('UserID');

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[userIdx] === username &&
          row[passIdx] === password &&
          row[statIdx] === 'Active') {
        return {
          success: true,
          user: {
            id:    row[idIdx],
            name:  row[userIdx],
            role:  row[roleIdx],
            dept:  row[deptIdx]
          }
        };
      }
    }
    return { success: false, message: 'Invalid username or password' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// GENERIC SHEET READER — returns array of objects
// ════════════════════════════════════════════════════════════════
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, data: [], message: 'Sheet not found: ' + sheetName };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    const headers = data[0];
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, j) => {
        let val = data[i][j];
        if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd-MM-yyyy');
        obj[h] = val === null || val === undefined ? '' : val;
      });
      // Skip completely empty rows
      if (Object.values(obj).every(v => v === '')) continue;
      rows.push(obj);
    }
    return { success: true, data: rows };
  } catch(e) {
    return { success: false, data: [], message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// MOTORS
// ════════════════════════════════════════════════════════════════
function getMotors() { return getSheetData(SHEETS.MOTORS); }

function getMotorById(motorId) {
  const result = getSheetData(SHEETS.MOTORS);
  if (!result.success) return result;
  const motor = result.data.find(m => m.MotorID === motorId);
  return motor ? { success: true, data: motor } : { success: false, message: 'Motor not found' };
}

function updateMotorStatus(motorId, newStatus, updatedBy) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.MOTORS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('MotorID');
    const statusCol = headers.indexOf('Status');
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === motorId) {
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        logAction('Motor Status Update', motorId + ' → ' + newStatus, updatedBy);
        return { success: true };
      }
    }
    return { success: false, message: 'Motor not found' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// MAINTENANCE
// ════════════════════════════════════════════════════════════════
function getMaintenance() { return getSheetData(SHEETS.MAINTENANCE); }

function addMaintenance(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.MAINTENANCE);
    const id = 'MT' + new Date().getTime();
    // Calculate NextDueDate (180 days)
    const maintDate = new Date(data.MaintenanceDate);
    const nextDue = new Date(maintDate);
    nextDue.setDate(nextDue.getDate() + 180);
    const nextDueStr = Utilities.formatDate(nextDue, Session.getScriptTimeZone(), 'dd-MM-yyyy');

    sheet.appendRow([
      id,
      data.MotorID,
      data.Shift || '',
      data.MaintenanceType,
      data.MaintenanceDate,
      nextDueStr,
      data.DoneBy,
      data.Remarks || ''
    ]);
    // Send reminder email to supervisors
    sendMaintenanceDoneEmail(data.MotorID, data.MaintenanceType, data.DoneBy, nextDueStr);
    return { success: true, id: id, nextDueDate: nextDueStr };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// GREASING
// ════════════════════════════════════════════════════════════════
function getGreasing() { return getSheetData(SHEETS.GREASING); }

function addGreasing(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.GREASING);
    const id = 'GR' + new Date().getTime();

    // Auto frequency: HT=90, LT=120
    const motorResult = getMotorById(data.MotorID);
    let days = 120;
    if (motorResult.success && motorResult.data.MotorType === 'HT') days = 90;

    const grDate = new Date(data.GreasingDate);
    const nextDue = new Date(grDate);
    nextDue.setDate(nextDue.getDate() + days);
    const nextDueStr = Utilities.formatDate(nextDue, Session.getScriptTimeZone(), 'dd-MM-yyyy');

    sheet.appendRow([
      id,
      data.MotorID,
      data.GreasingDate,
      nextDueStr,
      data.GreaseType || 'Lithium EP2',
      data.GreaseQty_gm || '',
      data.DoneBy || '',
      data.Remarks || ''
    ]);
    return { success: true, id: id, nextDueDate: nextDueStr, daysFrequency: days };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// REPAIR
// ════════════════════════════════════════════════════════════════
function getRepairs() { return getSheetData(SHEETS.REPAIR); }

function addRepair(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPAIR);
    const id = 'REP' + new Date().getTime();
    sheet.appendRow([
      id,
      data.MotorID,
      data.MotorDescription || '',
      data.JobNo || '',
      data.VendorID || '',
      data.VendorName || '',
      data.SentDate,
      data.ExpectedReturnDate || '',
      '',
      data.FaultDescription || '',
      data.RepairType || 'Repair',
      '',
      data.DispatchedBy || '',
      data.ApprovedBy || '',
      'Dispatched',
      ''
    ]);
    // Update motor status to REPAIR
    updateMotorStatus(data.MotorID, 'REPAIR', data.DispatchedBy);
    // Send alert email
    sendRepairDispatchEmail(data);
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function markRepairReturned(repairId, returnDate, remarks, updatedBy) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPAIR);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('RepairID');
    const statusCol = headers.indexOf('RepairStatus');
    const returnCol = headers.indexOf('ActualReturnDate');
    const remarkCol = headers.indexOf('Return_Remarks');
    const motorCol  = headers.indexOf('MotorID');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === repairId) {
        sheet.getRange(i+1, statusCol+1).setValue('Returned');
        sheet.getRange(i+1, returnCol+1).setValue(returnDate);
        sheet.getRange(i+1, remarkCol+1).setValue(remarks);
        // Update motor status back to RUNNING
        updateMotorStatus(data[i][motorCol], 'RUNNING', updatedBy);
        // Notification
        const vendorCol = headers.indexOf('VendorName');
        sendMotorReturnedNotification(data[i][motorCol], data[i][vendorCol] || '', returnDate, remarks);
        return { success: true };
      }
    }
    return { success: false, message: 'Repair record not found' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// BREAKDOWN
// ════════════════════════════════════════════════════════════════
function getBreakdowns() { return getSheetData(SHEETS.BREAKDOWN); }

function addBreakdown(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.BREAKDOWN);
    const id = 'BRK' + new Date().getTime();

    // Auto downtime
    let downtime = '';
    if (data.StartTime && data.EndTime) {
      const start = new Date(data.StartTime);
      const end   = new Date(data.EndTime);
      downtime = ((end - start) / 3600000).toFixed(2);
    }

    sheet.appendRow([
      id,
      data.MotorID,
      data.BreakdownDate,
      data.StartTime || '',
      data.EndTime || '',
      downtime,
      data.Cause || '',
      data.Action || '',
      data.DoneBy || '',
      data.Shift || '',
      '',
      'Open',
      data.Remarks || ''
    ]);
    // Immediate alert email
    sendBreakdownAlertEmail(data, id, downtime);
    return { success: true, id: id, downtimeHours: downtime };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// IOP DATA — with confirmation required for edits
// ════════════════════════════════════════════════════════════════
function getIOPData(filterIOP, filterJB, filterType) {
  const result = getSheetData(SHEETS.IOP);
  if (!result.success) return result;
  let data = result.data;
  if (filterIOP)  data = data.filter(r => r.IOP  === filterIOP);
  if (filterJB)   data = data.filter(r => r.JB_No === filterJB);
  if (filterType) data = data.filter(r => r.IO_Type === filterType);
  return { success: true, data: data };
}

function updateIOPEntry(sno, updates, confirmedBy, confirmedByName) {
  try {
    if (!confirmedBy || !confirmedByName) {
      return { success: false, message: 'ConfirmedBy is required for IOP edits' };
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.IOP);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const snoCol  = headers.indexOf('SNo');
    const confCol = headers.indexOf('ConfirmedBy');
    const dateCol = headers.indexOf('LastModifiedDate');
    const byCol   = headers.indexOf('LastModifiedBy');
    const remCol  = headers.indexOf('Remarks');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][snoCol]) === String(sno)) {
        // Apply updates
        Object.keys(updates).forEach(key => {
          const col = headers.indexOf(key);
          if (col >= 0) sheet.getRange(i+1, col+1).setValue(updates[key]);
        });
        // Set audit fields
        sheet.getRange(i+1, confCol+1).setValue(confirmedBy);
        sheet.getRange(i+1, dateCol+1).setValue(new Date().toISOString());
        sheet.getRange(i+1, byCol+1).setValue(confirmedByName);
        return { success: true };
      }
    }
    return { success: false, message: 'IOP entry not found' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// TELEPHONE DIRECTORY
// ════════════════════════════════════════════════════════════════
function searchDirectory(query) {
  const result = getSheetData(SHEETS.TELEPHONE);
  if (!result.success) return result;
  if (!query) return result;
  const q = query.toLowerCase();
  const filtered = result.data.filter(r =>
    (r.Name && r.Name.toLowerCase().includes(q)) ||
    (r.Department && r.Department.toLowerCase().includes(q)) ||
    (r.Location && r.Location.toLowerCase().includes(q)) ||
    (r.Designation && r.Designation.toLowerCase().includes(q)) ||
    (r.OfficeExt && String(r.OfficeExt).includes(q))
  );
  return { success: true, data: filtered };
}

// ════════════════════════════════════════════════════════════════
// DASHBOARD — summary stats
// ════════════════════════════════════════════════════════════════
function getDashboardData() {
  try {
    const motors     = getSheetData(SHEETS.MOTORS).data || [];
    const maint      = getSheetData(SHEETS.MAINTENANCE).data || [];
    const greasing   = getSheetData(SHEETS.GREASING).data || [];
    const repairs    = getSheetData(SHEETS.REPAIR).data || [];
    const breakdowns = getSheetData(SHEETS.BREAKDOWN).data || [];

    const today = new Date();
    today.setHours(0,0,0,0);
    const in7 = new Date(today); in7.setDate(in7.getDate() + 7);

    function parseDate(s) {
      if (!s) return null;
      const p = s.split('-');
      if (p.length === 3) return new Date(p[2], p[1]-1, p[0]);
      return new Date(s);
    }

    return {
      success: true,
      stats: {
        totalMotors:    motors.length,
        running:        motors.filter(m => m.Status === 'RUNNING').length,
        underRepair:    motors.filter(m => m.Status === 'REPAIR').length,
        spare:          motors.filter(m => m.Status === 'SPARE').length,
        htMotors:       motors.filter(m => m.MotorType === 'HT').length,
        ltMotors:       motors.filter(m => m.MotorType === 'LT').length,
        maintOverdue:   maint.filter(m => { const d = parseDate(m.NextDueDate); return d && d < today; }).length,
        maintDueWeek:   maint.filter(m => { const d = parseDate(m.NextDueDate); return d && d >= today && d <= in7; }).length,
        greasOverdue:   greasing.filter(g => { const d = parseDate(g.NextDueDate); return d && d < today; }).length,
        greaseDueWeek:  greasing.filter(g => { const d = parseDate(g.NextDueDate); return d && d >= today && d <= in7; }).length,
        activeRepairs:  repairs.filter(r => r.RepairStatus === 'Dispatched').length,
        recentBreakdowns: breakdowns.slice(-5).reverse(),
        overdueMaintenanceList: maint
          .filter(m => { const d = parseDate(m.NextDueDate); return d && d < today; })
          .slice(0, 10),
        overdueGreasingList: greasing
          .filter(g => { const d = parseDate(g.NextDueDate); return d && d < today; })
          .slice(0, 10),
      }
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// OTHER GETTERS
// ════════════════════════════════════════════════════════════════
function getTransformers()  { return getSheetData(SHEETS.TRANSFORMER); }
function getSpareMotors()   { return getSheetData(SHEETS.SPARE); }
function getVFDDetails()    { return getSheetData(SHEETS.VFD); }
function getEmployees()     { return getSheetData(SHEETS.EMPLOYEES); }
function getVendors()       { return getSheetData(SHEETS.VENDOR); }
function getShifts()        { return getSheetData(SHEETS.SHIFT); }
function getAreas()         { return getSheetData(SHEETS.AREA); }
function getLocations()     { return getSheetData(SHEETS.LOCATION); }

function addTransformerMaintenance(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.TR_MAINT);
    const id = 'TM' + new Date().getTime();
    sheet.appendRow([id, data.TransformerID, data.MaintenanceType, data.MaintenanceDate,
      data.NextDueDate || '', data.OilLevel || '', data.OilTemp_C || '',
      data.WdgTemp_C || '', data.InsulationTest || '', data.DoneBy || '',
      data.Shift || '', data.Remarks || '']);
    return { success: true, id: id };
  } catch(e) { return { success: false, message: e.message }; }
}

// ════════════════════════════════════════════════════════════════
// TELEGRAM + EMAIL NOTIFICATIONS
// ════════════════════════════════════════════════════════════════

// ── STEP 1: Telegram Bot setup karo ─────────────────────────────
// 1. Telegram mein @BotFather ko message karo → /newbot → naam do
// 2. Bot token milega — niche paste karo
// 3. Apna SP5 group banao → bot ko admin banao
// 4. Koi bhi message karo group mein, phir browser mein open karo:
//    https://api.telegram.org/bot[TOKEN]/getUpdates
//    chat_id milega response mein — niche paste karo
const TELEGRAM_BOT_TOKEN = '8745256510:AAETkXbsZ0pFexHwB3bdSbt-KgVfVtIWJY0';  // e.g. '7123456789:AAFxxx...'
const TELEGRAM_CHAT_ID   = '-5210590157';   // e.g. '-1001234567890'

// ── STEP 2: Email addresses ──────────────────────────────────────
const NOTIFICATION_EMAILS = ['poshanm@gmail.com', 'hmaitray@gmail.com'];

// ── Core Telegram sender ─────────────────────────────────────────
function sendTelegram(message) {
  if (!TELEGRAM_BOT_TOKEN || TELEGRAM_BOT_TOKEN === 'YOUR_BOT_TOKEN_HERE') return;
  try {
    const url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: TELEGRAM_CHAT_ID,
        text: message,
        parse_mode: 'HTML'
      }),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log('Telegram error: ' + e.message); }
}

// ── Core Email sender ────────────────────────────────────────────
function sendEmail(subject, body) {
  try {
    NOTIFICATION_EMAILS.forEach(email => {
      MailApp.sendEmail({ to: email, subject: subject, body: body });
    });
  } catch(e) { Logger.log('Email error: ' + e.message); }
}

// ── 1. BREAKDOWN ALERT — Immediate ───────────────────────────────
function sendBreakdownAlertEmail(data, id, downtime) {
  const tgMsg =
`🚨 <b>BREAKDOWN ALERT — SP5 Electrical</b>

⚙️ Motor: <b>${data.MotorID}</b>
📅 Date: ${data.BreakdownDate}
⏰ Time: ${data.StartTime || 'Not specified'}
❌ Cause: <b>${data.Cause || 'Not specified'}</b>
⏱ Downtime: ${downtime ? downtime + ' hrs' : 'Ongoing'}
👤 Reported by: ${data.DoneBy || '-'}
🔧 Action: ${data.Action || 'Under investigation'}
🆔 ID: ${id}`;

  const emailBody =
`SP5 ELECTRICAL — BREAKDOWN ALERT
Motor ID   : ${data.MotorID}
Date       : ${data.BreakdownDate}
Cause      : ${data.Cause || 'Not specified'}
Start Time : ${data.StartTime || '-'}
Downtime   : ${downtime ? downtime + ' hours' : 'Ongoing'}
Reported By: ${data.DoneBy || '-'}
Action     : ${data.Action || 'Under investigation'}
ID         : ${id}`;

  sendTelegram(tgMsg);
  sendEmail('🚨 BREAKDOWN: Motor ' + data.MotorID + ' — ' + data.BreakdownDate, emailBody);
}

// ── 2. REPAIR DISPATCHED ─────────────────────────────────────────
function sendRepairDispatchEmail(data) {
  const tgMsg =
`🛠 <b>Motor Sent for Repair — SP5</b>

⚙️ Motor: <b>${data.MotorID}</b>
🏭 Vendor: ${data.VendorName || 'CWS'}
📅 Sent: ${data.SentDate}
🔧 Type: ${data.RepairType || 'Repair'}
❌ Fault: ${data.FaultDescription || '-'}
📋 Job No: ${data.JobNo || '-'}
👤 By: ${data.DispatchedBy || '-'}`;

  sendTelegram(tgMsg);
  sendEmail('🛠 Motor Dispatched: ' + data.MotorID + ' → ' + (data.VendorName||'CWS'),
    `Motor ${data.MotorID} dispatched.\nVendor: ${data.VendorName}\nFault: ${data.FaultDescription}\nBy: ${data.DispatchedBy}`);
}

// ── 3. MAINTENANCE DONE ──────────────────────────────────────────
function sendMaintenanceDoneEmail(motorId, type, doneBy, nextDue) {
  const tgMsg =
`✅ <b>Maintenance Done — SP5</b>

⚙️ Motor: <b>${motorId}</b>
🔧 Type: ${type}
👤 Done By: ${doneBy}
📅 Next Due: <b>${nextDue}</b>`;

  sendTelegram(tgMsg);
  sendEmail('✅ Maintenance Done: Motor ' + motorId,
    `Motor ${motorId} maintenance completed.\nType: ${type}\nDone By: ${doneBy}\nNext Due: ${nextDue}`);
}

// ── 4. MOTOR RETURNED FROM REPAIR ───────────────────────────────
function sendMotorReturnedNotification(motorId, vendorName, returnDate, remarks) {
  const tgMsg =
`✅ <b>Motor Returned from Repair — SP5</b>

⚙️ Motor: <b>${motorId}</b>
🏭 From: ${vendorName || 'Vendor'}
📅 Return Date: ${returnDate}
📝 Condition: ${remarks || 'OK'}
🟢 Status: RUNNING`;

  sendTelegram(tgMsg);
  sendEmail('✅ Motor Returned: ' + motorId,
    `Motor ${motorId} returned from repair.\nVendor: ${vendorName}\nReturn Date: ${returnDate}\nCondition: ${remarks}`);
}

// ════════════════════════════════════════════════════════════════
// SCHEDULED REMINDERS — Set this as Time-driven trigger
// Run daily at 8:00 AM
// ════════════════════════════════════════════════════════════════
function runDailyReminders() {
  const today = new Date(); today.setHours(0,0,0,0);
  const in7   = new Date(today); in7.setDate(in7.getDate() + 7);
  const in5   = new Date(today); in5.setDate(in5.getDate() + 5);

  function parseDate(s) {
    if (!s) return null;
    const p = String(s).split('-');
    if (p.length === 3) return new Date(p[2], p[1]-1, p[0]);
    return new Date(s);
  }

  // Maintenance overdue
  const maint = getSheetData(SHEETS.MAINTENANCE).data || [];
  const maintDue = maint.filter(m => { const d = parseDate(m.NextDueDate); return d && d <= in7; });
  if (maintDue.length > 0) {
    const list = maintDue.map(m => `• ${m.MotorID} — Due: ${m.NextDueDate} (${m.MaintenanceType})`).join('\n');
    const tgList = maintDue.map(m => `  • <b>${m.MotorID}</b> → ${m.NextDueDate} (${m.MaintenanceType})`).join('\n');
    sendTelegram(`🔔 <b>Maintenance Due This Week — SP5</b>\n\n${tgList}\n\n📋 Total: ${maintDue.length} motors`);
    sendEmail(`⚠ ${maintDue.length} Motors: Maintenance Due This Week`,
      `The following motors have maintenance due within 7 days:\n\n${list}\n\n— SP5 Electrical App`);
  }

  // Greasing overdue
  const grease = getSheetData(SHEETS.GREASING).data || [];
  const greaseDue = grease.filter(g => { const d = parseDate(g.NextDueDate); return d && d <= in5; });
  if (greaseDue.length > 0) {
    const list = greaseDue.map(g => `• ${g.MotorID} — Due: ${g.NextDueDate}`).join('\n');
    const tgList = greaseDue.map(g => `  • <b>${g.MotorID}</b> → ${g.NextDueDate}`).join('\n');
    sendTelegram(`🟡 <b>Greasing Due — SP5</b>\n\n${tgList}\n\n📋 Total: ${greaseDue.length} motors`);
    sendEmail(`🔴 ${greaseDue.length} Motors: Greasing Due`,
      `The following motors need greasing within 5 days:\n\n${list}\n\n— SP5 Electrical App`);
  }

  // Repair overdue
  const repairs = getSheetData(SHEETS.REPAIR).data || [];
  const repairOverdue = repairs.filter(r => {
    if (r.RepairStatus !== 'Dispatched') return false;
    const d = parseDate(r.ExpectedReturnDate);
    return d && d < today;
  });
  if (repairOverdue.length > 0) {
    const list = repairOverdue.map(r => `• ${r.MotorID} → ${r.VendorName} (Expected: ${r.ExpectedReturnDate})`).join('\n');
    const tgList = repairOverdue.map(r => `  • <b>${r.MotorID}</b> → ${r.VendorName} | Due: ${r.ExpectedReturnDate}`).join('\n');
    sendTelegram(`🚨 <b>Repair Overdue — Follow Up! — SP5</b>\n\n${tgList}\n\n📋 ${repairOverdue.length} motors pending`);
    sendEmail(`🚨 ${repairOverdue.length} Motors: Repair Overdue`,
      `These motors have not returned:\n\n${list}\n\n— SP5 Electrical App`);
  }
}

// ════════════════════════════════════════════════════════════════
// AUDIT LOG
// ════════════════════════════════════════════════════════════════
function logAction(action, detail, user) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let logSheet = ss.getSheetByName('AuditLog');
    if (!logSheet) {
      logSheet = ss.insertSheet('AuditLog');
      logSheet.appendRow(['Timestamp', 'Action', 'Detail', 'User']);
    }
    logSheet.appendRow([new Date().toISOString(), action, detail, user || 'System']);
  } catch(e) { Logger.log('Log error: ' + e.message); }
}

// ════════════════════════════════════════════════════════════════
// SETUP HELPER — Run once to add Password column to Users sheet
// ════════════════════════════════════════════════════════════════
function setupUsersSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!headers.includes('Password')) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Password');
    // Set default passwords
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      sheet.getRange(i + 1, sheet.getLastColumn()).setValue('sp5@2024');
    }
    Logger.log('Password column added. Default: sp5@2024');
  }
}
