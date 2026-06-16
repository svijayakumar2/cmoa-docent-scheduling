// === CONFIGURATION ===
const KATHY_EMAIL = 'saranyav196@gmail.com';
const CLAIM_WINDOW_DAYS = 5;
const AUTO_ASSIGN_DAYS_OUT = 5;

const SHEET_SCHEDULE   = 'Schedule';
const SHEET_DOCENTS    = 'Docents';
const SHEET_SIGNUPS    = 'Signups';
const SHEET_CANCELLATIONS = 'Cancellations';
const BACKUP_FOLDER_NAME  = 'CMOA Backups';
const BACKUP_KEEP_DAYS    = 30;
const SITE_URL = 'https://svijayakumar2.github.io/cmoa-docent-scheduling/';

// Certification tags: PC = Permanent Collection, CI = Carnegie International, SCH = School Tour
// Mindful Museum has NO certification requirement (anyone can sign up)
// CI Activation uses the same CI certification
var TOUR_TAG_MAP = {
  'permanent collection': 'pc',
  'permanent collection evening': 'pc',
  'permanent collection (30 min tour)': 'pc',
  'carnegie international': 'ci',
  'carnegie international evening': 'ci',
  'ci activation tour': 'ci',
  'school tour': 'sch'
};

// Use the spreadsheet's timezone (Eastern) so dates/times don't shift
// when the script is run by someone in a different timezone
var TIMEZONE = 'America/New_York';

// =====================
// HELPERS
// =====================
function isValidEmail(email) {
  return email && email.toString().indexOf('@') !== -1;
}

// Get midnight today in Eastern timezone (so date comparisons match sheet dates)
function getTodayET() {
  var todayStr = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  return Utilities.parseDate(todayStr, TIMEZONE, 'yyyy-MM-dd');
}

function formatDateNice(date) {
  return Utilities.formatDate(date, TIMEZONE, 'EEEE, MMMM d');
}

function formatDateISO(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
}

function dayName(dayNum) {
  return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][dayNum];
}

function isCertifiedFor(docentInfo, tourType) {
  if (!docentInfo.certifiedTours) return true;
  return isCertifiedForRaw(docentInfo.certifiedTours, tourType);
}

function isCertifiedForRaw(certList, tourType) {
  if (!certList) return true;
  var requiredTag = TOUR_TAG_MAP[tourType.toLowerCase()];
  if (!requiredTag) return true;
  return certList.indexOf(requiredTag) !== -1;
}

function formatTime(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE, 'h:mm a');
  }
  return val.toString();
}

function parseTimeRange(timeVal, date) {
  var start = new Date(date);
  start.setHours(13, 0, 0, 0);

  if (timeVal instanceof Date) {
    start.setHours(timeVal.getHours(), timeVal.getMinutes(), 0, 0);
    var end = new Date(start.getTime() + 3600000);
    return [start, end];
  }

  var timeStr = (timeVal || '').toString();
  var match = timeStr.match(/(\d+)(?::(\d+))?\s*(AM|PM)?/i);
  if (match) {
    var hour = parseInt(match[1]);
    var minute = parseInt(match[2] || '0');
    var ampm = (match[3] || '').toUpperCase();
    if (ampm === 'PM' && hour < 12) hour += 12;
    if (ampm === 'AM' && hour === 12) hour = 0;
    start.setHours(hour, minute, 0, 0);
  }
  var end = new Date(start.getTime() + 3600000);
  return [start, end];
}

// Parse comma-separated day list into lowercase array
// e.g. "monday, wednesday" -> ["monday", "wednesday"]
function parseDayList(val) {
  if (!val) return [];
  return val.toString().toLowerCase().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
}

// Check if a day name matches a day number
// dayNum: 0=Sunday..6=Saturday; dayList: ["monday", "wednesday", ...]
function dayListContains(dayList, dayNum) {
  if (!dayList || dayList.length === 0) return false;
  var name = dayName(dayNum).toLowerCase();
  return dayList.indexOf(name) !== -1;
}

// =====================
// MENU
// =====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tour Scheduler')
    .addItem('Run Auto-Assignment Now', 'runAutoAssignment')
    .addItem('Send Daily Digest Now', 'sendDailyDigest')
    .addItem('Check Expired Claims', 'checkExpiredClaims')
    .addItem('Backup Now', 'dailyBackup')
    .addToUi();
}

// =====================
// WEB APP ENDPOINTS
// =====================
function doGet(e) {
  var action = (e.parameter.action || '').toString();

  if (action === 'claim') {
    return handleClaim(e.parameter.slot, e.parameter.docent);
  }

  var result = buildSchedulePayload();
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var action = payload.action;

  if (action === 'signup') {
    return handleSignup(payload.docent, payload.slots, payload.roles);
  }

  if (action === 'claim') {
    return handleClaimJSON(payload.slot, payload.docent);
  }

  if (action === 'savePreferences') {
    return handleSavePreferences(payload.docent, payload.preferences);
  }

  if (action === 'dropout') {
    return handleDropOut(payload.slot, payload.docent);
  }

  if (action === 'withdraw') {
    return handleWithdraw(payload.slot, payload.docent);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================
// GET DATA
// =====================
// Docents tab columns (indices):
//   0: Name, 1: Email, 2: Tours this Quarter, 3: Tours YTD,
//   4: Unavailable Until, 5: Certified Tours,
//   6: Preferred Days, 7: Avoid Days, 8: Lead Eligible, 9: Last Minute
function buildSchedulePayload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);
  var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);

  var schedData = schedSheet.getDataRange().getValues();
  var docentData = docentSheet.getDataRange().getValues();
  var signupData = signupSheet.getDataRange().getValues();

  // Build docent list with certification and preference info
  var docents = [];
  for (var i = 1; i < docentData.length; i++) {
    if (docentData[i][0]) {
      var certRaw = (docentData[i][5] || '').toString().trim();
      docents.push({
        name: docentData[i][0].toString(),
        certifiedTours: certRaw ? certRaw.split(',').map(function(s) { return s.trim().toLowerCase(); }) : null,
        leadEligible: (docentData[i][8] || '').toString().toLowerCase() === 'yes',
        preferredDays: (docentData[i][6] || '').toString(),
        avoidDays: (docentData[i][7] || '').toString(),
        lastMinute: (docentData[i][9] || '').toString().toLowerCase() === 'yes'
      });
    }
  }

  // Build signup maps
  var signupMap = {};
  var roleSignupMap = {};
  for (var i = 1; i < signupData.length; i++) {
    var docent = (signupData[i][1] || '').toString();
    var slotId = (signupData[i][2] || '').toString();
    var role = (signupData[i][3] || '').toString();
    if (!slotId) continue;
    if (!signupMap[slotId]) signupMap[slotId] = [];
    if (signupMap[slotId].indexOf(docent) === -1) {
      signupMap[slotId].push(docent);
    }
    if (role) {
      if (!roleSignupMap[slotId]) roleSignupMap[slotId] = {};
      if (!roleSignupMap[slotId][role]) roleSignupMap[slotId][role] = [];
      if (roleSignupMap[slotId][role].indexOf(docent) === -1) {
        roleSignupMap[slotId][role].push(docent);
      }
    }
  }

  // Build slot list
  var today = getTodayET();
  var slots = [];
  for (var i = 1; i < schedData.length; i++) {
    var row = schedData[i];
    var slotId = (row[0] || '').toString();
    var date = new Date(row[1]);
    if (isNaN(date.getTime())) continue;
    if (date < today) continue;

    var status = (row[5] || '').toString();
    var assigned = (row[6] || '').toString();
    var tourType = (row[3] || '').toString();
    var docentsNeeded = row[4] || 1;
    // Ensure docentsNeeded is numeric
    if (typeof docentsNeeded !== 'number') docentsNeeded = 1;

    var isSchoolTour = tourType.toLowerCase() === 'school tour';

    slots.push({
      slotId: slotId,
      date: formatDateISO(date),
      dateDisplay: formatDateNice(date),
      time: formatTime(row[2]),
      tourType: tourType,
      docentsNeeded: docentsNeeded,
      status: status,
      assigned: assigned,
      signups: signupMap[slotId] || [],
      details: (row[7] || '').toString(),
      tourLeadSchool: (row[8] || '').toString(),
      participantSchool: (row[9] || '').toString(),
      mindfulWelcomeDesk: (row[10] || '').toString(),
      mindfulTourLead: (row[11] || '').toString(),
      docentsNeeded_Desk: row[12] || 0,
      docentsNeeded_MindfulTour: row[13] || 0,
      docentsNeeded_Lead: isSchoolTour ? 1 : 0,
      docentsNeeded_Participant: isSchoolTour ? Math.max(docentsNeeded - 1, 0) : 0,
      roleSignups: roleSignupMap[slotId] || {}
    });
  }

  slots.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  return { docents: docents, slots: slots, tourTagMap: TOUR_TAG_MAP };
}

// =====================
// SIGNUP
// =====================
function handleSignup(docentName, slotIds, roles) {
  roles = roles || {};
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Server busy, try again.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);
    var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
    var signupData = signupSheet.getDataRange().getValues();
    var schedData = schedSheet.getDataRange().getValues();

    var openSlots = {};
    for (var i = 1; i < schedData.length; i++) {
      var status = (schedData[i][5] || '').toString();
      if (status === '' || status === 'Open') {
        openSlots[(schedData[i][0] || '').toString()] = true;
      }
    }

    var existing = {};
    for (var i = 1; i < signupData.length; i++) {
      var d = (signupData[i][1] || '').toString();
      var s = (signupData[i][2] || '').toString();
      var r = (signupData[i][3] || '').toString();
      if (d === docentName) {
        existing[s + '|' + r] = i + 1;
      }
    }

    var now = new Date();

    for (var j = 0; j < slotIds.length; j++) {
      var role = roles[slotIds[j]] || '';
      var key = slotIds[j] + '|' + role;
      if (!existing[key]) {
        signupSheet.appendRow([now, docentName, slotIds[j], role]);
      }
    }

    var slotSet = {};
    for (var j = 0; j < slotIds.length; j++) {
      var role = roles[slotIds[j]] || '';
      slotSet[slotIds[j] + '|' + role] = true;
    }

    var rowsToDelete = [];
    for (var key in existing) {
      var slotId = key.split('|')[0];
      if (!slotSet[key] && openSlots[slotId]) {
        rowsToDelete.push(existing[key]);
      }
    }
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var j = 0; j < rowsToDelete.length; j++) {
      signupSheet.deleteRow(rowsToDelete[j]);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =====================
// SAVE PREFERENCES (from website)
// =====================
function handleSavePreferences(docentName, prefs) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Server busy, try again.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var docentSheet = ss.getSheetByName(SHEET_DOCENTS);
    var docentData = docentSheet.getDataRange().getValues();

    for (var i = 1; i < docentData.length; i++) {
      if ((docentData[i][0] || '').toString() === docentName) {
        // Column G (7): Preferred Days
        docentSheet.getRange(i + 1, 7).setValue(prefs.preferredDays || '');
        // Column H (8): Avoid Days
        docentSheet.getRange(i + 1, 8).setValue(prefs.avoidDays || '');
        // Column J (10): Last Minute
        docentSheet.getRange(i + 1, 10).setValue(prefs.lastMinute ? 'Yes' : '');
        break;
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =====================
// CLAIM (cancelled slot - HTML response for email links)
// =====================
function handleClaim(slotId, docentName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return HtmlService.createHtmlOutput('<h2>Server busy, try again in a moment.</h2>');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
    var cancelSheet = ss.getSheetByName(SHEET_CANCELLATIONS);
    var docentSheet = ss.getSheetByName(SHEET_DOCENTS);

    var schedData = schedSheet.getDataRange().getValues();
    var cancelData = cancelSheet.getDataRange().getValues();
    var docentData = docentSheet.getDataRange().getValues();

    var slotRow = -1;
    for (var i = 1; i < schedData.length; i++) {
      if (schedData[i][0].toString() === slotId) { slotRow = i; break; }
    }
    if (slotRow === -1) {
      return HtmlService.createHtmlOutput('<h2>Slot not found.</h2>');
    }

    if (schedData[slotRow][5].toString() !== 'Needs Sub') {
      return HtmlService.createHtmlOutput(
        '<h2>Already claimed</h2><p>Someone else has already claimed this shift. Thanks for trying!</p>');
    }

    schedSheet.getRange(slotRow + 1, 6).setValue('Assigned');
    schedSheet.getRange(slotRow + 1, 7).setValue(docentName);

    for (var i = 1; i < cancelData.length; i++) {
      if (cancelData[i][0].toString() === slotId && cancelData[i][3] === 'Broadcast') {
        cancelSheet.getRange(i + 1, 4).setValue('Claimed');
        cancelSheet.getRange(i + 1, 5).setValue(docentName);
        break;
      }
    }

    for (var i = 1; i < docentData.length; i++) {
      if (docentData[i][0] === docentName) {
        if (isValidEmail(docentData[i][1])) {
          sendCalendarInvite(
            docentData[i][1], docentName, slotId,
            new Date(schedData[slotRow][1]),
            schedData[slotRow][2],
            schedData[slotRow][3].toString()
          );
        }
        docentSheet.getRange(i + 1, 3).setValue((docentData[i][2] || 0) + 1);
        break;
      }
    }

    var tourType = schedData[slotRow][3];
    var dateStr = formatDateNice(new Date(schedData[slotRow][1]));
    var time = schedData[slotRow][2];

    return HtmlService.createHtmlOutput(
      '<h2>Claimed!</h2>' +
      '<p>You have claimed the <strong>' + tourType + '</strong> tour on ' +
      dateStr + ' at ' + time + '.</p>' +
      '<p>A calendar invite is on its way.</p>'
    );
  } finally {
    lock.releaseLock();
  }
}

// Claim from the website (JSON response)
function handleClaimJSON(slotId, docentName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Server busy, try again.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
    var cancelSheet = ss.getSheetByName(SHEET_CANCELLATIONS);
    var docentSheet = ss.getSheetByName(SHEET_DOCENTS);

    var schedData = schedSheet.getDataRange().getValues();
    var cancelData = cancelSheet.getDataRange().getValues();
    var docentData = docentSheet.getDataRange().getValues();

    var slotRow = -1;
    for (var i = 1; i < schedData.length; i++) {
      if (schedData[i][0].toString() === slotId) { slotRow = i; break; }
    }
    if (slotRow === -1) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Slot not found.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (schedData[slotRow][5].toString() !== 'Needs Sub') {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Someone else has already claimed this shift.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    schedSheet.getRange(slotRow + 1, 6).setValue('Assigned');
    schedSheet.getRange(slotRow + 1, 7).setValue(docentName);

    for (var i = 1; i < cancelData.length; i++) {
      if (cancelData[i][0].toString() === slotId && cancelData[i][3] === 'Broadcast') {
        cancelSheet.getRange(i + 1, 4).setValue('Claimed');
        cancelSheet.getRange(i + 1, 5).setValue(docentName);
        break;
      }
    }

    for (var i = 1; i < docentData.length; i++) {
      if (docentData[i][0] === docentName) {
        if (isValidEmail(docentData[i][1])) {
          sendCalendarInvite(
            docentData[i][1], docentName, slotId,
            new Date(schedData[slotRow][1]),
            schedData[slotRow][2],
            schedData[slotRow][3].toString()
          );
        }
        docentSheet.getRange(i + 1, 3).setValue((docentData[i][2] || 0) + 1);
        break;
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =====================
// DROP OUT (docent self-service cancellation)
// Removes docent from assigned list, sets status to Needs Sub, triggers outreach
// =====================
function handleDropOut(slotId, docentName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Server busy, try again.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
    var docentSheet = ss.getSheetByName(SHEET_DOCENTS);

    var schedData = schedSheet.getDataRange().getValues();

    var slotRow = -1;
    for (var i = 1; i < schedData.length; i++) {
      if (schedData[i][0].toString() === slotId) { slotRow = i; break; }
    }
    if (slotRow === -1) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Slot not found.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var status = schedData[slotRow][5].toString();
    if (status !== 'Assigned') {
      return ContentService.createTextOutput(JSON.stringify({ error: 'This tour is not in Assigned status.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var assignedRaw = schedData[slotRow][6].toString();
    var assignedList = assignedRaw.split(',').map(function(s) { return s.trim(); }).filter(Boolean);

    if (assignedList.indexOf(docentName) === -1) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'You are not assigned to this tour.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Remove this docent from the assigned list
    var remaining = assignedList.filter(function(n) { return n !== docentName; });
    schedSheet.getRange(slotRow + 1, 7).setValue(remaining.join(', '));

    // Set status to Needs Sub
    schedSheet.getRange(slotRow + 1, 6).setValue('Needs Sub');

    // Decrement tour count
    var docentData = docentSheet.getDataRange().getValues();
    for (var i = 1; i < docentData.length; i++) {
      if (docentData[i][0].toString() === docentName) {
        var currentCount = docentData[i][2] || 0;
        if (currentCount > 0) {
          docentSheet.getRange(i + 1, 3).setValue(currentCount - 1);
        }
        break;
      }
    }

    // Notify Kathy
    MailApp.sendEmail(
      KATHY_EMAIL,
      'Docent dropped out: ' + docentName + ' - ' + schedData[slotRow][3] + ' on ' + formatDateNice(new Date(schedData[slotRow][1])),
      docentName + ' has dropped out of the ' + schedData[slotRow][3] + ' tour on ' +
      formatDateNice(new Date(schedData[slotRow][1])) + ' at ' + formatTime(schedData[slotRow][2]) + '.\n\n' +
      'The system has set this slot to "Needs Sub" and is emailing available docents automatically.\n\n' +
      (remaining.length > 0 ? 'Still assigned: ' + remaining.join(', ') : 'No docents remaining on this slot.') +
      '\n\n-- Tour Scheduler'
    );

    // Trigger the sub outreach directly (onEdit won't fire from programmatic changes)
    handleNeedsSub(ss, schedSheet, slotRow);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =====================
// WITHDRAW (remove signup/offer before assignment)
// =====================
function handleWithdraw(slotId, docentName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Server busy, try again.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);
    var signupData = signupSheet.getDataRange().getValues();

    var rowsToDelete = [];
    for (var i = 1; i < signupData.length; i++) {
      if ((signupData[i][1] || '').toString() === docentName &&
          (signupData[i][2] || '').toString() === slotId) {
        rowsToDelete.push(i + 1);
      }
    }

    if (rowsToDelete.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Signup not found.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Delete bottom-up so row indices stay valid
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var i = 0; i < rowsToDelete.length; i++) {
      signupSheet.deleteRow(rowsToDelete[i]);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =====================
// AUTO-ASSIGNMENT (daily timer)
// =====================
function runAutoAssignment() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);
  var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);

  var schedData = schedSheet.getDataRange().getValues();
  var docentData = docentSheet.getDataRange().getValues();
  var signupData = signupSheet.getDataRange().getValues();

  var today = getTodayET();
  var tally = {};
  for (var i = 1; i < docentData.length; i++) {
    var name = (docentData[i][0] || '').toString();
    if (!name) continue;
    var unavailUntil = docentData[i][4] ? new Date(docentData[i][4]) : null;
    var certRaw = (docentData[i][5] || '').toString().trim();
    tally[name] = {
      email: docentData[i][1],
      count: docentData[i][2] || 0,
      row: i + 1,
      unavailable: unavailUntil && unavailUntil >= today,
      certifiedTours: certRaw ? certRaw.split(',').map(function(s) { return s.trim().toLowerCase(); }) : null,
      leadEligible: (docentData[i][8] || '').toString().toLowerCase() === 'yes'
    };
  }

  var signupMap = {};
  var roleMap = {};
  for (var i = 1; i < signupData.length; i++) {
    var docent = (signupData[i][1] || '').toString();
    var slotId = (signupData[i][2] || '').toString();
    var role = (signupData[i][3] || '').toString();
    if (!slotId) continue;
    if (!signupMap[slotId]) signupMap[slotId] = [];
    if (signupMap[slotId].indexOf(docent) === -1) {
      signupMap[slotId].push(docent);
    }
    if (role) {
      if (!roleMap[slotId]) roleMap[slotId] = {};
      if (!roleMap[slotId][role]) roleMap[slotId][role] = [];
      if (roleMap[slotId][role].indexOf(docent) === -1) {
        roleMap[slotId][role].push(docent);
      }
    }
  }

  var cutoff = new Date(today.getTime() + AUTO_ASSIGN_DAYS_OUT * 86400000);

  var assignedTimes = {};
  for (var i = 1; i < schedData.length; i++) {
    if ((schedData[i][5] || '').toString() !== 'Assigned') continue;
    var aDate = new Date(schedData[i][1]);
    var aDateKey = formatDateISO(aDate);
    var aTimes = parseTimeRange(schedData[i][2], aDate);
    var aNames = (schedData[i][6] || '').toString().split(',').map(function(s) { return s.trim(); });
    for (var j = 0; j < aNames.length; j++) {
      var key = aNames[j] + '|' + aDateKey;
      if (!assignedTimes[key]) assignedTimes[key] = [];
      assignedTimes[key].push([aTimes[0].getTime(), aTimes[1].getTime()]);
    }
  }

  for (var i = 1; i < schedData.length; i++) {
    var row = schedData[i];
    var slotId = (row[0] || '').toString();
    var date = new Date(row[1]);
    var timeRaw = row[2];
    var tourType = (row[3] || '').toString();
    var docentsNeeded = row[4] || 1;
    if (typeof docentsNeeded !== 'number') docentsNeeded = 1;
    var status = (row[5] || '').toString();

    if (status !== '' && status !== 'Open') continue;
    if (date > cutoff || date < today) continue;

    var slotTimes = parseTimeRange(timeRaw, date);
    var slotStart = slotTimes[0].getTime();
    var slotEnd = slotTimes[1].getTime();
    var slotDateKey = formatDateISO(date);

    var signups = signupMap[slotId] || [];
    var neededDesk = row[12] || 0;
    var neededMindfulTour = row[13] || 0;
    var isSchoolTour = tourType.toLowerCase() === 'school tour';
    var neededLead = isSchoolTour ? 1 : 0;
    var neededParticipant = isSchoolTour ? Math.max(docentsNeeded - 1, 0) : 0;
    var hasMindfulRoles = neededDesk > 0 || neededMindfulTour > 0;
    var hasSchoolRoles = neededLead > 0 || neededParticipant > 0;

    function filterEligible(names) {
      return names.filter(function(n) {
        if (!tally[n]) return false;
        if (tally[n].unavailable) return false;
        if (!isCertifiedFor(tally[n], tourType)) return false;
        var key = n + '|' + slotDateKey;
        var existing = assignedTimes[key] || [];
        for (var k = 0; k < existing.length; k++) {
          if (slotStart < existing[k][1] && slotEnd > existing[k][0]) return false;
        }
        return true;
      }).sort(function(a, b) { return tally[a].count - tally[b].count; });
    }

    var assigned = [];
    var deskAssigned = [];
    var tourAssigned = [];
    var leadAssigned = [];
    var participantAssigned = [];

    if (hasMindfulRoles) {
      var deskSignups = (roleMap[slotId] && roleMap[slotId]['desk']) || [];
      var tourSignups = (roleMap[slotId] && roleMap[slotId]['tour']) || [];
      deskAssigned = filterEligible(deskSignups).slice(0, neededDesk);
      var deskSet = {};
      for (var j = 0; j < deskAssigned.length; j++) deskSet[deskAssigned[j]] = true;
      tourAssigned = filterEligible(tourSignups.filter(function(n) { return !deskSet[n]; })).slice(0, neededMindfulTour);
      assigned = deskAssigned.concat(tourAssigned);
    } else if (hasSchoolRoles) {
      var leadSignups = (roleMap[slotId] && roleMap[slotId]['lead']) || [];
      // Only lead-eligible docents can be assigned as lead
      leadSignups = leadSignups.filter(function(n) { return tally[n] && tally[n].leadEligible; });
      var participantSignups = (roleMap[slotId] && roleMap[slotId]['participant']) || [];
      leadAssigned = filterEligible(leadSignups).slice(0, neededLead);
      var leadSet = {};
      for (var j = 0; j < leadAssigned.length; j++) leadSet[leadAssigned[j]] = true;
      participantAssigned = filterEligible(participantSignups.filter(function(n) { return !leadSet[n]; })).slice(0, neededParticipant);
      assigned = leadAssigned.concat(participantAssigned);
    } else {
      assigned = filterEligible(signups).slice(0, docentsNeeded);
    }

    if (assigned.length > 0) {
      schedSheet.getRange(i + 1, 6).setValue('Assigned');
      schedSheet.getRange(i + 1, 7).setValue(assigned.join(', '));

      if (hasMindfulRoles) {
        if (deskAssigned.length > 0) schedSheet.getRange(i + 1, 11).setValue(deskAssigned.join(', '));
        if (tourAssigned.length > 0) schedSheet.getRange(i + 1, 12).setValue(tourAssigned.join(', '));
      }
      if (hasSchoolRoles) {
        if (leadAssigned.length > 0) schedSheet.getRange(i + 1, 9).setValue(leadAssigned.join(', '));
        if (participantAssigned.length > 0) schedSheet.getRange(i + 1, 10).setValue(participantAssigned.join(', '));
      }

      for (var j = 0; j < assigned.length; j++) {
        var name = assigned[j];
        tally[name].count += 1;
        docentSheet.getRange(tally[name].row, 3).setValue(tally[name].count);
        if (isValidEmail(tally[name].email)) {
          sendCalendarInvite(tally[name].email, name, slotId, date, timeRaw, tourType);
        }

        var key = name + '|' + slotDateKey;
        if (!assignedTimes[key]) assignedTimes[key] = [];
        assignedTimes[key].push([slotStart, slotEnd]);
      }
    }
  }
}

// =====================
// INSTANT STATUS CHANGE (installable onEdit trigger)
// =====================
function onEditInstallable(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_SCHEDULE) return;

  var range = e.range;
  var col = range.getColumn();
  var row = range.getRow();

  if (col !== 6 || row <= 1) return;

  var newValue = (e.value || '').toString();

  if (newValue === 'Tour Cancelled') {
    handleTourCancelled(e.source, sheet, row);
  } else if (newValue === 'Needs Sub') {
    handleNeedsSub(e.source, sheet, row);
  }
}

function handleTourCancelled(ss, schedSheet, row) {
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);

  var schedData = schedSheet.getRange(row, 1, 1, 7).getValues()[0];
  var tourType = schedData[3].toString();
  var date = new Date(schedData[1]);
  var time = formatTime(schedData[2]);
  var originallyAssigned = schedData[6].toString();

  var docentData = docentSheet.getDataRange().getValues();
  var emailMap = {};
  for (var i = 1; i < docentData.length; i++) {
    if (docentData[i][0]) emailMap[docentData[i][0].toString()] = docentData[i][1].toString();
  }

  var assignedList = originallyAssigned.split(',').map(function(s) { return s.trim(); }).filter(Boolean);

  for (var i = 0; i < assignedList.length; i++) {
    var name = assignedList[i];
    var email = emailMap[name];
    if (!isValidEmail(email)) continue;
    MailApp.sendEmail(
      email,
      'Tour cancelled: ' + tourType + ' on ' + formatDateNice(date),
      'Hi ' + name + ',\n\n' +
      'The ' + tourType + ' tour on ' + formatDateNice(date) + ' at ' + time +
      ' has been cancelled. You do not need to come.\n\n' +
      '-- Tour Scheduler'
    );

    for (var j = 1; j < docentData.length; j++) {
      if (docentData[j][0].toString() === name) {
        var currentCount = docentData[j][2] || 0;
        if (currentCount > 0) {
          docentSheet.getRange(j + 1, 3).setValue(currentCount - 1);
        }
        break;
      }
    }
  }
}

// A docent dropped out. Tour still happening. Email backups + preference-matched docents.
function handleNeedsSub(ss, schedSheet, row) {
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);
  var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);
  var cancelSheet = ss.getSheetByName(SHEET_CANCELLATIONS);

  var schedData = schedSheet.getRange(row, 1, 1, 7).getValues()[0];
  var slotId = schedData[0].toString();
  var tourType = schedData[3].toString();
  var date = new Date(schedData[1]);
  var time = formatTime(schedData[2]);
  var originallyAssigned = schedData[6].toString();

  var docentData = docentSheet.getDataRange().getValues();
  var today = getTodayET();
  var emailMap = {};
  var unavailMap = {};
  var certMap = {};
  var preferredDayMap = {};
  var avoidDayMap = {};
  var lastMinuteMap = {};
  for (var i = 1; i < docentData.length; i++) {
    if (docentData[i][0]) {
      var dName = docentData[i][0].toString();
      emailMap[dName] = docentData[i][1].toString();
      var unavailUntil = docentData[i][4] ? new Date(docentData[i][4]) : null;
      if (unavailUntil && unavailUntil >= today) unavailMap[dName] = true;
      var certRaw = (docentData[i][5] || '').toString().trim();
      certMap[dName] = certRaw ? certRaw.split(',').map(function(s) { return s.trim().toLowerCase(); }) : null;
      preferredDayMap[dName] = parseDayList(docentData[i][6]);
      avoidDayMap[dName] = parseDayList(docentData[i][7]);
      lastMinuteMap[dName] = (docentData[i][9] || '').toString().toLowerCase() === 'yes';
    }
  }

  var signupData = signupSheet.getDataRange().getValues();
  var slotSignups = [];
  for (var i = 1; i < signupData.length; i++) {
    if ((signupData[i][2] || '').toString() === slotId) {
      slotSignups.push((signupData[i][1] || '').toString());
    }
  }

  var assignedList = originallyAssigned.split(',').map(function(s) { return s.trim(); });
  var backups = slotSignups.filter(function(n) {
    if (assignedList.indexOf(n) !== -1) return false;
    if (unavailMap[n]) return false;
    if (certMap[n] && !isCertifiedForRaw(certMap[n], tourType)) return false;
    return true;
  });

  var webAppUrl = ScriptApp.getService().getUrl();
  var cancelledDay = date.getDay();
  var cancelledDayName = dayName(cancelledDay).toLowerCase();

  // Tier 1: Email direct backups (signed up but not assigned)
  for (var i = 0; i < backups.length; i++) {
    var name = backups[i];
    var email = emailMap[name];
    if (!isValidEmail(email)) continue;
    var claimUrl = webAppUrl + '?action=claim&slot=' + encodeURIComponent(slotId) +
                   '&docent=' + encodeURIComponent(name);
    MailApp.sendEmail(
      email,
      'Tour opening: ' + tourType + ' on ' + formatDateNice(date),
      'Hi ' + name + ',\n\n' +
      'A tour you signed up as available for has opened up:\n\n' +
      tourType + ' on ' + formatDateNice(date) + ' at ' + time + '\n\n' +
      'First to claim wins. Click here to claim:\n' + claimUrl + '\n\n' +
      '-- Tour Scheduler'
    );
  }

  // Tier 2: Email docents who prefer this day OR are last-minute available
  var alreadyContacted = {};
  for (var i = 0; i < backups.length; i++) alreadyContacted[backups[i]] = true;
  for (var i = 0; i < assignedList.length; i++) alreadyContacted[assignedList[i]] = true;

  var tier2Contacted = false;
  for (var name in emailMap) {
    if (alreadyContacted[name]) continue;
    if (unavailMap[name]) continue;
    if (certMap[name] && !isCertifiedForRaw(certMap[name], tourType)) continue;
    if (!isValidEmail(emailMap[name])) continue;

    // Skip if they avoid this day
    if (dayListContains(avoidDayMap[name], cancelledDay)) continue;

    // Include if they prefer this day OR are last-minute available
    var prefersDay = dayListContains(preferredDayMap[name], cancelledDay);
    var isLastMinute = lastMinuteMap[name];
    if (!prefersDay && !isLastMinute) continue;

    var email = emailMap[name];
    var claimUrl = webAppUrl + '?action=claim&slot=' + encodeURIComponent(slotId) +
                   '&docent=' + encodeURIComponent(name);
    var reason = prefersDay
      ? 'You usually tour on ' + dayName(cancelledDay) + 's'
      : 'You\'re signed up for last-minute availability';
    MailApp.sendEmail(
      email,
      'Tour opening: ' + tourType + ' on ' + formatDateNice(date),
      'Hi ' + name + ',\n\n' +
      'A ' + tourType + ' tour on ' + formatDateNice(date) + ' at ' + time +
      ' has just opened up. ' + reason + ', so we thought you might be interested.\n\n' +
      'First to claim wins. Click here to claim:\n' + claimUrl + '\n\n' +
      '-- Tour Scheduler'
    );
    tier2Contacted = true;
  }

  if (backups.length === 0 && !tier2Contacted) {
    MailApp.sendEmail(
      KATHY_EMAIL,
      'No backups available: ' + tourType + ' on ' + formatDateNice(date),
      'A docent dropped out but nobody else signed up for this slot and no ' +
      dayName(cancelledDay) + ' regulars were found.\n\n' +
      'Slot: ' + slotId + '\nTour: ' + tourType + '\nDate: ' + formatDateNice(date) +
      '\nTime: ' + time + '\n\nPlease assign manually.'
    );
    cancelSheet.appendRow([slotId, originallyAssigned, new Date(), 'Escalated', '']);
  } else {
    cancelSheet.appendRow([slotId, originallyAssigned, new Date(), 'Broadcast', '']);
  }
}

// =====================
// DAILY DIGEST
// One email per docent: upcoming tour reminders + unfilled tour requests
// =====================
function sendDailyDigest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);
  var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);

  var schedData = schedSheet.getDataRange().getValues();
  var docentData = docentSheet.getDataRange().getValues();
  var signupData = signupSheet.getDataRange().getValues();

  var today = getTodayET();
  var cutoff = new Date(today.getTime() + AUTO_ASSIGN_DAYS_OUT * 86400000);
  var day7 = new Date(today.getTime() + 7 * 86400000).getTime();
  var day2 = new Date(today.getTime() + 2 * 86400000).getTime();

  // Build docent info with preferences
  var docents = {};
  for (var i = 1; i < docentData.length; i++) {
    var name = (docentData[i][0] || '').toString();
    if (!name) continue;
    var unavailUntil = docentData[i][4] ? new Date(docentData[i][4]) : null;
    var certRaw = (docentData[i][5] || '').toString().trim();
    docents[name] = {
      email: docentData[i][1],
      unavailable: unavailUntil && unavailUntil >= today,
      certifiedTours: certRaw ? certRaw.split(',').map(function(s) { return s.trim().toLowerCase(); }) : null,
      preferredDays: parseDayList(docentData[i][6]),
      avoidDays: parseDayList(docentData[i][7]),
      lastMinute: (docentData[i][9] || '').toString().toLowerCase() === 'yes',
      reminders: [],
      needsFilling: []
    };
  }

  // Build signup map
  var signupMap = {};
  for (var i = 1; i < signupData.length; i++) {
    var docent = (signupData[i][1] || '').toString();
    var slotId = (signupData[i][2] || '').toString();
    if (!slotId) continue;
    if (!signupMap[slotId]) signupMap[slotId] = [];
    if (signupMap[slotId].indexOf(docent) === -1) signupMap[slotId].push(docent);
  }

  var kathyUnfilled = [];

  for (var i = 1; i < schedData.length; i++) {
    var row = schedData[i];
    var slotId = (row[0] || '').toString();
    var date = new Date(row[1]);
    var time = formatTime(row[2]);
    var tourType = (row[3] || '').toString();
    var docentsNeeded = row[4] || 1;
    if (typeof docentsNeeded !== 'number') docentsNeeded = 1;
    var status = (row[5] || '').toString();
    var assignedStr = (row[6] || '').toString();

    // Normalize date to midnight Eastern for comparison with day7/day2
    var dateStr = formatDateISO(date);
    var dateNorm = Utilities.parseDate(dateStr, TIMEZONE, 'yyyy-MM-dd');
    var t = dateNorm.getTime();

    // --- REMINDERS: assigned tours coming up in 2 or 7 days ---
    if (status === 'Assigned' && (t === day7 || t === day2)) {
      var window = (t === day7) ? '1 week' : '2 days';
      var assignedNames = assignedStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
      for (var j = 0; j < assignedNames.length; j++) {
        if (docents[assignedNames[j]]) {
          docents[assignedNames[j]].reminders.push(
            tourType + ' on ' + formatDateNice(date) + ' at ' + time + ' (in ' + window + ')'
          );
        }
      }
    }

    // --- UNFILLED TOURS: Open slots within the assignment window ---
    if (status !== '' && status !== 'Open') continue;
    if (date > cutoff || date < today) continue;

    var signups = signupMap[slotId] || [];
    var assignedCount = assignedStr ? assignedStr.split(',').filter(Boolean).length : 0;
    var spotsRemaining = docentsNeeded - Math.max(signups.length, assignedCount);
    if (spotsRemaining <= 0) continue;

    var slotDay = date.getDay();
    var tourLine = tourType + ' on ' + formatDateNice(date) + ' at ' + time +
                   ' (' + spotsRemaining + ' spot' + (spotsRemaining > 1 ? 's' : '') + ')';

    var anyoneNotified = false;
    for (var name in docents) {
      var d = docents[name];
      if (d.unavailable) continue;
      if (!isCertifiedFor(d, tourType)) continue;
      if (signups.indexOf(name) !== -1) continue;

      // Respect day preferences
      if (dayListContains(d.avoidDays, slotDay)) continue;

      // Only include docents who prefer this day OR are last-minute available
      var prefersDay = dayListContains(d.preferredDays, slotDay);
      var noPref = d.preferredDays.length === 0;
      if (!prefersDay && !noPref && !d.lastMinute) continue;

      d.needsFilling.push(tourLine);
      anyoneNotified = true;
    }

    if (!anyoneNotified && signups.length === 0) {
      kathyUnfilled.push(tourLine);
    }
  }

  // Send one email per docent
  for (var name in docents) {
    var d = docents[name];
    if (d.reminders.length === 0 && d.needsFilling.length === 0) continue;
    if (!isValidEmail(d.email)) continue;

    var body = 'Hi ' + name + ',\n\n';

    if (d.reminders.length > 0) {
      body += 'YOUR UPCOMING TOURS:\n';
      for (var j = 0; j < d.reminders.length; j++) {
        body += '  - ' + d.reminders[j] + '\n';
      }
      body += '\nIf you have a conflict, contact Kathy as soon as possible.\n\n';
    }

    if (d.needsFilling.length > 0) {
      body += 'TOURS THAT NEED DOCENTS:\n';
      for (var j = 0; j < d.needsFilling.length; j++) {
        body += '  - ' + d.needsFilling[j] + '\n';
      }
      body += '\nIf you can help, sign up here: ' + SITE_URL + '\n\n';
    }

    body += '-- Tour Scheduler';

    MailApp.sendEmail(d.email, 'CMOA Docent Update for ' + name, body);
  }

  if (kathyUnfilled.length > 0) {
    var kathyBody = 'The following tours have no signups and no matching docents were found:\n\n';
    for (var j = 0; j < kathyUnfilled.length; j++) {
      kathyBody += '  - ' + kathyUnfilled[j] + '\n';
    }
    kathyBody += '\nPlease assign manually.\n-- Tour Scheduler';
    MailApp.sendEmail(KATHY_EMAIL, 'Unfilled tours need attention', kathyBody);
  }
}

// =====================
// EXPIRED CLAIMS CHECK (daily timer)
// =====================
function checkExpiredClaims() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
  var cancelSheet = ss.getSheetByName(SHEET_CANCELLATIONS);

  var schedData = schedSheet.getDataRange().getValues();
  var cancelData = cancelSheet.getDataRange().getValues();

  var now = new Date();

  for (var i = 1; i < cancelData.length; i++) {
    if ((cancelData[i][3] || '').toString() !== 'Broadcast') continue;

    var cancelledAt = new Date(cancelData[i][2]);
    var ageDays = (now - cancelledAt) / 86400000;

    if (ageDays < CLAIM_WINDOW_DAYS) continue;

    var slotId = cancelData[i][0].toString();

    for (var j = 1; j < schedData.length; j++) {
      if (schedData[j][0].toString() === slotId) {
        MailApp.sendEmail(
          KATHY_EMAIL,
          'Unfilled cancellation: ' + schedData[j][3] + ' on ' + formatDateNice(new Date(schedData[j][1])),
          'Nobody claimed this slot within ' + CLAIM_WINDOW_DAYS + ' days.\n\n' +
          'Slot: ' + slotId + '\n' +
          'Tour: ' + schedData[j][3] + '\n' +
          'Date: ' + formatDateNice(new Date(schedData[j][1])) + '\n' +
          'Time: ' + formatTime(schedData[j][2]) + '\n\n' +
          'Please assign manually.'
        );
        cancelSheet.getRange(i + 1, 4).setValue('Escalated');
        break;
      }
    }
  }
}

// =====================
// CALENDAR INVITE
// =====================
function sendCalendarInvite(email, name, slotId, date, timeStr, tourType) {
  if (!isValidEmail(email)) return;
  var times = parseTimeRange(timeStr, date);
  CalendarApp.getDefaultCalendar().createEvent(
    'Tour: ' + tourType,
    times[0],
    times[1],
    {
      description: 'You are assigned to lead the ' + tourType + ' tour.\n' +
                   'Slot ID: ' + slotId + '\n\n' +
                   'If you cannot make this tour, contact Kathy immediately.',
      guests: email,
      sendInvites: true
    }
  );
}

// =====================
// QUARTERLY RESET (optional timer: Jan 1, Apr 1, Jul 1, Oct 1)
// =====================
function resetQuarterlyTally() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);
  var docentData = docentSheet.getDataRange().getValues();
  for (var i = 1; i < docentData.length; i++) {
    docentSheet.getRange(i + 1, 3).setValue(0);
  }
}

// =====================
// DAILY BACKUP
// =====================
function dailyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);

  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  var copyName = ss.getName() + ' - Backup ' + today;
  DriveApp.getFileById(ss.getId()).makeCopy(copyName, folder);

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - BACKUP_KEEP_DAYS);
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (file.getDateCreated() < cutoff) {
      file.setTrashed(true);
    }
  }
}
