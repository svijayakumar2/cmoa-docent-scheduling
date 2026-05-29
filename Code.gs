// === CONFIGURATION ===
const KATHY_EMAIL = 'CHANGE_THIS@example.com';
const CLAIM_WINDOW_DAYS = 5;
const AUTO_ASSIGN_DAYS_OUT = 21;

const SHEET_SCHEDULE   = 'Schedule';
const SHEET_DOCENTS    = 'Docents';
const SHEET_SIGNUPS    = 'Signups';
const SHEET_CANCELLATIONS = 'Cancellations';

// =====================
// MENU
// =====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tour Scheduler')
    .addItem('Run Auto-Assignment Now', 'runAutoAssignment')
    .addItem('Send Reminders Now', 'sendReminders')
    .addItem('Check Expired Claims', 'checkExpiredClaims')
    .addToUi();
}

// =====================
// WEB APP ENDPOINTS
// =====================
function doGet(e) {
  var action = (e.parameter.action || '').toString();

  // Claim a cancelled slot (link from email)
  if (action === 'claim') {
    return handleClaim(e.parameter.slot, e.parameter.docent);
  }

  // Return schedule + docent data as JSON for the website
  var result = buildSchedulePayload();
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var action = payload.action;

  if (action === 'signup') {
    return handleSignup(payload.docent, payload.slots);
  }

  if (action === 'claim') {
    return handleClaim(payload.slot, payload.docent);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================
// GET DATA
// =====================
function buildSchedulePayload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);
  var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);

  var schedData = schedSheet.getDataRange().getValues();
  var docentData = docentSheet.getDataRange().getValues();
  var signupData = signupSheet.getDataRange().getValues();

  // Build docent list (just names for the dropdown)
  var docents = [];
  for (var i = 1; i < docentData.length; i++) {
    if (docentData[i][0]) docents.push(docentData[i][0].toString());
  }

  // Build a map of existing signups: slotId -> [names]
  var signupMap = {};
  for (var i = 1; i < signupData.length; i++) {
    var docent = (signupData[i][1] || '').toString();
    var slotId = (signupData[i][2] || '').toString();
    if (!slotId) continue;
    if (!signupMap[slotId]) signupMap[slotId] = [];
    if (signupMap[slotId].indexOf(docent) === -1) {
      signupMap[slotId].push(docent);
    }
  }

  // Build slot list
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var slots = [];
  for (var i = 1; i < schedData.length; i++) {
    var row = schedData[i];
    var slotId = (row[0] || '').toString();
    var date = new Date(row[1]);
    if (isNaN(date.getTime())) continue;
    if (date < today) continue; // skip past slots

    var status = (row[5] || '').toString();
    var assigned = (row[6] || '').toString();

    slots.push({
      slotId: slotId,
      date: formatDateISO(date),
      dateDisplay: formatDateNice(date),
      time: (row[2] || '').toString(),
      tourType: (row[3] || '').toString(),
      docentsNeeded: row[4] || 1,
      status: status,
      assigned: assigned,
      signups: signupMap[slotId] || []
    });
  }

  // Sort by date
  slots.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  return { docents: docents, slots: slots };
}

// =====================
// SIGNUP
// =====================
function handleSignup(docentName, slotIds) {
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

    // Find existing signups for this docent
    var existing = {};
    for (var i = 1; i < signupData.length; i++) {
      var d = (signupData[i][1] || '').toString();
      var s = (signupData[i][2] || '').toString();
      if (d === docentName) existing[s] = i + 1; // row number (1-indexed)
    }

    var now = new Date();

    // Add new signups
    for (var j = 0; j < slotIds.length; j++) {
      if (!existing[slotIds[j]]) {
        signupSheet.appendRow([now, docentName, slotIds[j]]);
      }
    }

    // Remove unchecked signups (docent deselected a slot they previously signed up for)
    var slotSet = {};
    for (var j = 0; j < slotIds.length; j++) slotSet[slotIds[j]] = true;

    var rowsToDelete = [];
    for (var s in existing) {
      if (!slotSet[s]) {
        rowsToDelete.push(existing[s]);
      }
    }
    // Delete from bottom to top so row numbers stay valid
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
// CLAIM (cancelled slot)
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

    // Find the slot
    var slotRow = -1;
    for (var i = 1; i < schedData.length; i++) {
      if (schedData[i][0].toString() === slotId) { slotRow = i; break; }
    }
    if (slotRow === -1) {
      return HtmlService.createHtmlOutput('<h2>Slot not found.</h2>');
    }

    if (schedData[slotRow][5].toString() !== 'Cancelled') {
      return HtmlService.createHtmlOutput(
        '<h2>Already claimed</h2><p>Someone else got there first. Thanks for trying!</p>');
    }

    // Assign the slot
    schedSheet.getRange(slotRow + 1, 6).setValue('Assigned');  // Status
    schedSheet.getRange(slotRow + 1, 7).setValue(docentName);  // Assigned Docents

    // Update cancellation log
    for (var i = 1; i < cancelData.length; i++) {
      if (cancelData[i][0].toString() === slotId && cancelData[i][3] === 'Broadcast') {
        cancelSheet.getRange(i + 1, 4).setValue('Claimed');
        cancelSheet.getRange(i + 1, 5).setValue(docentName);
        break;
      }
    }

    // Send calendar invite + update tally
    for (var i = 1; i < docentData.length; i++) {
      if (docentData[i][0] === docentName) {
        sendCalendarInvite(
          docentData[i][1], docentName, slotId,
          new Date(schedData[slotRow][1]),
          schedData[slotRow][2].toString(),
          schedData[slotRow][3].toString()
        );
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

  // Build docent tally
  var tally = {};
  for (var i = 1; i < docentData.length; i++) {
    var name = (docentData[i][0] || '').toString();
    if (!name) continue;
    tally[name] = {
      email: docentData[i][1],
      count: docentData[i][2] || 0,
      row: i + 1
    };
  }

  // Build signup map: slotId -> [names]
  var signupMap = {};
  for (var i = 1; i < signupData.length; i++) {
    var docent = (signupData[i][1] || '').toString();
    var slotId = (signupData[i][2] || '').toString();
    if (!slotId) continue;
    if (!signupMap[slotId]) signupMap[slotId] = [];
    if (signupMap[slotId].indexOf(docent) === -1) {
      signupMap[slotId].push(docent);
    }
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var cutoff = new Date(today.getTime() + AUTO_ASSIGN_DAYS_OUT * 86400000);

  for (var i = 1; i < schedData.length; i++) {
    var row = schedData[i];
    var slotId = (row[0] || '').toString();
    var date = new Date(row[1]);
    var time = (row[2] || '').toString();
    var tourType = (row[3] || '').toString();
    var docentsNeeded = row[4] || 1;
    var status = (row[5] || '').toString();

    // Only assign Open slots within the assignment window
    if (status !== '' && status !== 'Open') continue;
    if (date > cutoff || date < today) continue;

    var signups = signupMap[slotId] || [];
    if (signups.length === 0) continue;

    // Sort by lowest quarterly count (fair rotation)
    var eligible = signups.filter(function(n) { return tally[n]; });
    eligible.sort(function(a, b) { return tally[a].count - tally[b].count; });

    var assigned = eligible.slice(0, docentsNeeded);
    if (assigned.length === 0) continue;

    // Update the schedule
    schedSheet.getRange(i + 1, 6).setValue('Assigned');  // Status
    schedSheet.getRange(i + 1, 7).setValue(assigned.join(', '));  // Assigned Docents

    // Update tallies and send calendar invites
    for (var j = 0; j < assigned.length; j++) {
      var name = assigned[j];
      tally[name].count += 1;
      docentSheet.getRange(tally[name].row, 3).setValue(tally[name].count);
      sendCalendarInvite(tally[name].email, name, slotId, date, time, tourType);
    }

    // Notify docents who signed up but weren't assigned
    var notAssigned = signups.filter(function(n) { return assigned.indexOf(n) === -1; });
    for (var j = 0; j < notAssigned.length; j++) {
      var name = notAssigned[j];
      if (!tally[name]) continue;
      MailApp.sendEmail(
        tally[name].email,
        'Tour assignment update: ' + tourType + ' on ' + formatDateNice(date),
        'Hi ' + name + ',\n\n' +
        'The ' + tourType + ' tour on ' + formatDateNice(date) + ' at ' + time +
        ' has been assigned to ' + assigned.join(' and ') + '.\n\n' +
        'Thanks for signing up as available. You remain on the backup list if the assigned docent cancels.\n\n' +
        '-- Tour Scheduler'
      );
    }
  }
}

// =====================
// INSTANT CANCELLATION (installable onEdit trigger)
// =====================
function onEditInstallable(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_SCHEDULE) return;

  var range = e.range;
  var col = range.getColumn();
  var row = range.getRow();

  // Column 6 = Status
  if (col !== 6 || row <= 1) return;

  var newValue = (e.value || '').toString();
  if (newValue !== 'Cancelled') return;

  var ss = e.source;
  var schedSheet = sheet;
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);
  var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);
  var cancelSheet = ss.getSheetByName(SHEET_CANCELLATIONS);

  var schedData = schedSheet.getRange(row, 1, 1, 7).getValues()[0];
  var slotId = schedData[0].toString();
  var tourType = schedData[3].toString();
  var date = new Date(schedData[1]);
  var time = schedData[2].toString();
  var originallyAssigned = schedData[6].toString();

  // Get all docent emails
  var docentData = docentSheet.getDataRange().getValues();
  var emailMap = {};
  for (var i = 1; i < docentData.length; i++) {
    if (docentData[i][0]) emailMap[docentData[i][0].toString()] = docentData[i][1].toString();
  }

  // Get signups for this slot
  var signupData = signupSheet.getDataRange().getValues();
  var slotSignups = [];
  for (var i = 1; i < signupData.length; i++) {
    if ((signupData[i][2] || '').toString() === slotId) {
      slotSignups.push((signupData[i][1] || '').toString());
    }
  }

  // Eligible backups = signed up but not originally assigned
  var assignedList = originallyAssigned.split(',').map(function(s) { return s.trim(); });
  var backups = slotSignups.filter(function(n) { return assignedList.indexOf(n) === -1; });

  var webAppUrl = ScriptApp.getService().getUrl();

  // Email backups immediately
  for (var i = 0; i < backups.length; i++) {
    var name = backups[i];
    var email = emailMap[name];
    if (!email) continue;
    var claimUrl = webAppUrl + '?action=claim&slot=' + encodeURIComponent(slotId) +
                   '&docent=' + encodeURIComponent(name);
    MailApp.sendEmail(
      email,
      'Tour opening: ' + tourType + ' on ' + formatDateNice(date),
      'Hi ' + name + ',\n\n' +
      'A tour you signed up as available for has opened up:\n\n' +
      tourType + ' on ' + formatDateNice(date) + ' at ' + time + '\n\n' +
      'First to claim wins. Click here to claim:\n' + claimUrl + '\n\n' +
      'If nobody claims within ' + CLAIM_WINDOW_DAYS + ' days, this goes to Kathy.\n\n' +
      '-- Tour Scheduler'
    );
  }

  // If no backups exist, email Kathy immediately
  if (backups.length === 0) {
    MailApp.sendEmail(
      KATHY_EMAIL,
      'No backups available: ' + tourType + ' on ' + formatDateNice(date),
      'A tour was cancelled but nobody else signed up for this slot.\n\n' +
      'Slot: ' + slotId + '\nTour: ' + tourType + '\nDate: ' + formatDateNice(date) +
      '\nTime: ' + time + '\n\nPlease assign manually.'
    );
    cancelSheet.appendRow([slotId, originallyAssigned, new Date(), 'Escalated', '']);
  } else {
    cancelSheet.appendRow([slotId, originallyAssigned, new Date(), 'Broadcast', '']);
  }
}

// =====================
// REMINDERS (daily timer)
// =====================
function sendReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);

  var schedData = schedSheet.getDataRange().getValues();
  var docentData = docentSheet.getDataRange().getValues();

  var emailMap = {};
  for (var i = 1; i < docentData.length; i++) {
    if (docentData[i][0]) emailMap[docentData[i][0].toString()] = docentData[i][1].toString();
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var day7 = new Date(today.getTime() + 7 * 86400000).getTime();
  var day2 = new Date(today.getTime() + 2 * 86400000).getTime();

  for (var i = 1; i < schedData.length; i++) {
    var row = schedData[i];
    if ((row[5] || '').toString() !== 'Assigned') continue;

    var date = new Date(row[1]);
    date.setHours(0, 0, 0, 0);
    var t = date.getTime();

    if (t !== day7 && t !== day2) continue;

    var window = (t === day7) ? '1 week' : '2 days';
    var assigned = (row[6] || '').toString().split(',').map(function(s) { return s.trim(); }).filter(Boolean);

    for (var j = 0; j < assigned.length; j++) {
      var name = assigned[j];
      var email = emailMap[name];
      if (!email) continue;
      MailApp.sendEmail(
        email,
        'Reminder: ' + row[3] + ' tour in ' + window,
        'Hi ' + name + ',\n\n' +
        'Reminder that you are leading the ' + row[3] + ' tour on ' +
        formatDateNice(date) + ' at ' + row[2] + '. This is in ' + window + '.\n\n' +
        'If you have a conflict, contact Kathy as soon as possible.\n\n' +
        '-- Tour Scheduler'
      );
    }
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

    // Find the slot in schedule
    for (var j = 1; j < schedData.length; j++) {
      if (schedData[j][0].toString() === slotId) {
        MailApp.sendEmail(
          KATHY_EMAIL,
          'Unfilled cancellation: ' + schedData[j][3] + ' on ' + formatDateNice(new Date(schedData[j][1])),
          'Nobody claimed this slot within ' + CLAIM_WINDOW_DAYS + ' days.\n\n' +
          'Slot: ' + slotId + '\n' +
          'Tour: ' + schedData[j][3] + '\n' +
          'Date: ' + formatDateNice(new Date(schedData[j][1])) + '\n' +
          'Time: ' + schedData[j][2] + '\n\n' +
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

function parseTimeRange(timeStr, date) {
  var start = new Date(date);
  start.setHours(13, 0, 0, 0); // default 1pm
  var match = timeStr.toString().match(/(\d+)(?::(\d+))?\s*(AM|PM)?/i);
  if (match) {
    var hour = parseInt(match[1]);
    var minute = parseInt(match[2] || '0');
    var ampm = (match[3] || '').toUpperCase();
    if (ampm === 'PM' && hour < 12) hour += 12;
    if (ampm === 'AM' && hour === 12) hour = 0;
    start.setHours(hour, minute, 0, 0);
  }
  var end = new Date(start.getTime() + 3600000); // 1 hour
  return [start, end];
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
// HELPERS
// =====================
function formatDateNice(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEEE, MMMM d');
}

function formatDateISO(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
