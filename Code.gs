// === CONFIGURATION ===
const KATHY_EMAIL = 'CHANGE_THIS@example.com';
const CLAIM_WINDOW_DAYS = 5;
const AUTO_ASSIGN_DAYS_OUT = 5;
const MIN_WEEKDAY_SIGNUPS = 2;

const SHEET_SCHEDULE   = 'Schedule';
const SHEET_DOCENTS    = 'Docents';
const SHEET_SIGNUPS    = 'Signups';
const SHEET_CANCELLATIONS = 'Cancellations';
const BACKUP_FOLDER_NAME  = 'CMOA Backups';
const BACKUP_KEEP_DAYS    = 30;
const SITE_URL = 'https://svijayakumar2.github.io/cmoa-docent-scheduling/';

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
    return handleClaimJSON(payload.slot, payload.docent);
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

  // Build docent list with certification info
  var docents = [];
  for (var i = 1; i < docentData.length; i++) {
    if (docentData[i][0]) {
      var certRaw = (docentData[i][5] || '').toString().trim();
      docents.push({
        name: docentData[i][0].toString(),
        certifiedTours: certRaw ? certRaw.split(',').map(function(s) { return s.trim().toLowerCase(); }) : null
      });
    }
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
      time: formatTime(row[2]),
      tourType: (row[3] || '').toString(),
      docentsNeeded: row[4] || 1,
      status: status,
      assigned: assigned,
      signups: signupMap[slotId] || [],
      tourAgesGrades: (row[7] || '').toString(),
      focusArea: (row[8] || '').toString(),
      tourLeadSchool: (row[9] || '').toString(),
      participantSchool: (row[10] || '').toString(),
      mindfulWelcomeDesk: (row[11] || '').toString(),
      mindfulTourLead: (row[12] || '').toString()
    });
  }

  // Sort by date
  slots.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  return { docents: docents, slots: slots, tourTagMap: TOUR_TAG_MAP };
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
    var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
    var signupData = signupSheet.getDataRange().getValues();
    var schedData = schedSheet.getDataRange().getValues();

    // Build a set of slot IDs that are still Open (the ones docents can see/toggle)
    var openSlots = {};
    for (var i = 1; i < schedData.length; i++) {
      var status = (schedData[i][5] || '').toString();
      if (status === '' || status === 'Open') {
        openSlots[(schedData[i][0] || '').toString()] = true;
      }
    }

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

    // Only remove signups for Open slots that the docent deselected.
    // Never touch signups for Assigned/Needs Sub slots.
    var slotSet = {};
    for (var j = 0; j < slotIds.length; j++) slotSet[slotIds[j]] = true;

    var rowsToDelete = [];
    for (var s in existing) {
      if (!slotSet[s] && openSlots[s]) {
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

    if (schedData[slotRow][5].toString() !== 'Needs Sub') {
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
          schedData[slotRow][2],
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

// Claim from the website (returns JSON instead of HTML)
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
      return ContentService.createTextOutput(JSON.stringify({ error: 'Already claimed. Someone else got there first.' }))
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
        sendCalendarInvite(
          docentData[i][1], docentName, slotId,
          new Date(schedData[slotRow][1]),
          schedData[slotRow][2],
          schedData[slotRow][3].toString()
        );
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
  var today = new Date();
  today.setHours(0, 0, 0, 0);
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
      certifiedTours: certRaw ? certRaw.split(',').map(function(s) { return s.trim().toLowerCase(); }) : null
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

  // Track assigned time ranges per docent per day: { "name|2026-06-01": [[start,end], ...] }
  var assignedTimes = {};

  // Pre-load existing assignments so we don't double-book with already-assigned slots
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
    var time = formatTime(timeRaw);
    var tourType = (row[3] || '').toString();
    var docentsNeeded = row[4] || 1;
    var status = (row[5] || '').toString();

    // Only assign Open slots within the assignment window
    if (status !== '' && status !== 'Open') continue;
    if (date > cutoff || date < today) continue;

    var slotTimes = parseTimeRange(timeRaw, date);
    var slotStart = slotTimes[0].getTime();
    var slotEnd = slotTimes[1].getTime();
    var slotDateKey = formatDateISO(date);

    var signups = signupMap[slotId] || [];

    // Sort by lowest quarterly count (fair rotation)
    // Filter out anyone who already has a conflicting assignment on this day
    var eligible = signups.filter(function(n) {
      if (!tally[n]) return false;
      if (tally[n].unavailable) return false;
      if (!isCertifiedFor(tally[n], tourType)) return false;
      var key = n + '|' + slotDateKey;
      var existing = assignedTimes[key] || [];
      for (var k = 0; k < existing.length; k++) {
        if (slotStart < existing[k][1] && slotEnd > existing[k][0]) return false;
      }
      return true;
    });
    eligible.sort(function(a, b) { return tally[a].count - tally[b].count; });

    var assigned = eligible.slice(0, docentsNeeded);
    var spotsRemaining = docentsNeeded - assigned.length;

    // If we have at least some people, assign them
    if (assigned.length > 0) {
      schedSheet.getRange(i + 1, 6).setValue('Assigned');
      schedSheet.getRange(i + 1, 7).setValue(assigned.join(', '));

      for (var j = 0; j < assigned.length; j++) {
        var name = assigned[j];
        tally[name].count += 1;
        docentSheet.getRange(tally[name].row, 3).setValue(tally[name].count);
        sendCalendarInvite(tally[name].email, name, slotId, date, timeRaw, tourType);

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

  // Column 6 = Status
  if (col !== 6 || row <= 1) return;

  var newValue = (e.value || '').toString();

  if (newValue === 'Tour Cancelled') {
    handleTourCancelled(e.source, sheet, row);
  } else if (newValue === 'Needs Sub') {
    handleNeedsSub(e.source, sheet, row);
  }
}

// The whole tour is cancelled. Notify assigned docents they don't need to come.
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
    if (!email) continue;
    MailApp.sendEmail(
      email,
      'Tour cancelled: ' + tourType + ' on ' + formatDateNice(date),
      'Hi ' + name + ',\n\n' +
      'The ' + tourType + ' tour on ' + formatDateNice(date) + ' at ' + time +
      ' has been cancelled. You do not need to come.\n\n' +
      '-- Tour Scheduler'
    );

    // Decrement their quarterly count since they won't be doing this tour
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

// A docent dropped out. Tour still happening. Blast backups + weekday regulars.
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

  // Get all docent emails and unavailability
  var docentData = docentSheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var emailMap = {};
  var unavailMap = {};
  var certMap = {};
  for (var i = 1; i < docentData.length; i++) {
    if (docentData[i][0]) {
      var dName = docentData[i][0].toString();
      emailMap[dName] = docentData[i][1].toString();
      var unavailUntil = docentData[i][4] ? new Date(docentData[i][4]) : null;
      if (unavailUntil && unavailUntil >= today) unavailMap[dName] = true;
      var certRaw = (docentData[i][5] || '').toString().trim();
      certMap[dName] = certRaw ? certRaw.split(',').map(function(s) { return s.trim().toLowerCase(); }) : null;
    }
  }

  // Get signups for this slot
  var signupData = signupSheet.getDataRange().getValues();
  var slotSignups = [];
  for (var i = 1; i < signupData.length; i++) {
    if ((signupData[i][2] || '').toString() === slotId) {
      slotSignups.push((signupData[i][1] || '').toString());
    }
  }

  // Eligible backups = signed up but not originally assigned, not unavailable, and certified
  var assignedList = originallyAssigned.split(',').map(function(s) { return s.trim(); });
  var backups = slotSignups.filter(function(n) {
    if (assignedList.indexOf(n) !== -1) return false;
    if (unavailMap[n]) return false;
    if (certMap[n] && !isCertifiedForRaw(certMap[n], tourType)) return false;
    return true;
  });

  var webAppUrl = ScriptApp.getService().getUrl();

  // Tier 1: Email direct backups immediately
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
      '-- Tour Scheduler'
    );
  }

  // Tier 2: Email docents who tend to be free on this day of the week
  // but didn't sign up for this specific slot
  var cancelledDay = date.getDay();
  var schedAllData = schedSheet.getDataRange().getValues();

  var slotDayMap = {};
  for (var i = 1; i < schedAllData.length; i++) {
    var sid = (schedAllData[i][0] || '').toString();
    var sdate = new Date(schedAllData[i][1]);
    if (!isNaN(sdate.getTime())) slotDayMap[sid] = sdate.getDay();
  }

  var weekdayCounts = {};
  for (var i = 1; i < signupData.length; i++) {
    var dName = (signupData[i][1] || '').toString();
    var sId = (signupData[i][2] || '').toString();
    if (slotDayMap[sId] === cancelledDay) {
      weekdayCounts[dName] = (weekdayCounts[dName] || 0) + 1;
    }
  }

  var alreadyContacted = {};
  for (var i = 0; i < backups.length; i++) alreadyContacted[backups[i]] = true;
  for (var i = 0; i < assignedList.length; i++) alreadyContacted[assignedList[i]] = true;

  for (var name in weekdayCounts) {
    if (alreadyContacted[name]) continue;
    if (weekdayCounts[name] < MIN_WEEKDAY_SIGNUPS) continue;
    if (unavailMap[name]) continue;
    if (certMap[name] && !isCertifiedForRaw(certMap[name], tourType)) continue;
    var email = emailMap[name];
    if (!email) continue;
    var claimUrl = webAppUrl + '?action=claim&slot=' + encodeURIComponent(slotId) +
                   '&docent=' + encodeURIComponent(name);
    MailApp.sendEmail(
      email,
      'Tour opening: ' + tourType + ' on ' + formatDateNice(date),
      'Hi ' + name + ',\n\n' +
      'A ' + tourType + ' tour on ' + formatDateNice(date) + ' at ' + time +
      ' has just opened up. You often sign up for ' + dayName(cancelledDay) +
      ' tours, so we thought you might be interested.\n\n' +
      'First to claim wins. Click here to claim:\n' + claimUrl + '\n\n' +
      '-- Tour Scheduler'
    );
  }

  // If nobody at all was contacted, email Kathy
  var weekdayRegularsContacted = false;
  for (var name in weekdayCounts) {
    if (!alreadyContacted[name] && weekdayCounts[name] >= MIN_WEEKDAY_SIGNUPS && emailMap[name]) {
      weekdayRegularsContacted = true;
      break;
    }
  }

  if (backups.length === 0 && !weekdayRegularsContacted) {
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
// DAILY DIGEST (replaces sendReminders + outreach emails)
// One email per docent combining upcoming tour reminders + unfilled tour requests
// =====================
function sendDailyDigest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
  var docentSheet = ss.getSheetByName(SHEET_DOCENTS);
  var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);

  var schedData = schedSheet.getDataRange().getValues();
  var docentData = docentSheet.getDataRange().getValues();
  var signupData = signupSheet.getDataRange().getValues();

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var cutoff = new Date(today.getTime() + AUTO_ASSIGN_DAYS_OUT * 86400000);
  var day7 = new Date(today.getTime() + 7 * 86400000).getTime();
  var day2 = new Date(today.getTime() + 2 * 86400000).getTime();

  // Build docent info
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
      reminders: [],
      needsFilling: []
    };
  }

  // Build signup map: slotId -> [names]
  var signupMap = {};
  for (var i = 1; i < signupData.length; i++) {
    var docent = (signupData[i][1] || '').toString();
    var slotId = (signupData[i][2] || '').toString();
    if (!slotId) continue;
    if (!signupMap[slotId]) signupMap[slotId] = [];
    if (signupMap[slotId].indexOf(docent) === -1) signupMap[slotId].push(docent);
  }

  // Build slotId -> day-of-week map for weekday regular detection
  var slotDayMap = {};
  for (var i = 1; i < schedData.length; i++) {
    var sd = new Date(schedData[i][1]);
    if (!isNaN(sd.getTime())) slotDayMap[(schedData[i][0] || '').toString()] = sd.getDay();
  }

  // Count each docent's signups per weekday
  var weekdaySignups = {};
  for (var i = 1; i < signupData.length; i++) {
    var dName = (signupData[i][1] || '').toString();
    var sId = (signupData[i][2] || '').toString();
    var dayOfWeek = slotDayMap[sId];
    if (dayOfWeek === undefined) continue;
    var key = dName + '|' + dayOfWeek;
    weekdaySignups[key] = (weekdaySignups[key] || 0) + 1;
  }

  var kathyUnfilled = [];

  for (var i = 1; i < schedData.length; i++) {
    var row = schedData[i];
    var slotId = (row[0] || '').toString();
    var date = new Date(row[1]);
    var time = formatTime(row[2]);
    var tourType = (row[3] || '').toString();
    var docentsNeeded = row[4] || 1;
    var status = (row[5] || '').toString();
    var assignedStr = (row[6] || '').toString();

    date.setHours(0, 0, 0, 0);
    var t = date.getTime();

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

    // --- UNFILLED TOURS: Open slots within the assignment window that need more docents ---
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
      // Skip if already signed up for this slot
      if (signups.indexOf(name) !== -1) continue;
      // Only email weekday regulars (signed up for this day of week at least MIN_WEEKDAY_SIGNUPS times)
      var wKey = name + '|' + slotDay;
      if ((weekdaySignups[wKey] || 0) < MIN_WEEKDAY_SIGNUPS) continue;

      d.needsFilling.push(tourLine);
      anyoneNotified = true;
    }

    if (!anyoneNotified && signups.length === 0) {
      kathyUnfilled.push(tourLine);
    }
  }

  // Send one email per docent (only if they have something to tell them)
  for (var name in docents) {
    var d = docents[name];
    if (d.reminders.length === 0 && d.needsFilling.length === 0) continue;
    if (!d.email) continue;

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

  // Email Kathy about any completely unfilled slots with no regulars
  if (kathyUnfilled.length > 0) {
    var kathyBody = 'The following tours have no signups and no weekday regulars were found:\n\n';
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

function parseTimeRange(timeVal, date) {
  var start = new Date(date);
  start.setHours(13, 0, 0, 0); // default 1pm

  // If Sheets gave us a Date object, just pull hours/minutes directly
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

function dayName(dayNum) {
  return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][dayNum];
}

// Check if a docent (from tally) is certified for a tour type.
// If certifiedTours is null (column blank), they're eligible for everything.
// Kathy enters comma-separated tags in the Certified Tours column, e.g. "PC, CI, MM"
// Tags: PC = Permanent Collection, CI = Carnegie International, MM = Mindful Museum, CIA = CI Activation
// The system maps each tour type to its required tag and checks if the docent has it.
var TOUR_TAG_MAP = {
  'permanent collection': 'pc',
  'permanent collection evening': 'pc',
  'carnegie international': 'ci',
  'carnegie international evening': 'ci',
  'ci activation tour': 'cia',
  'mindful museum': 'mm'
};

function isCertifiedFor(docentInfo, tourType) {
  if (!docentInfo.certifiedTours) return true;
  return isCertifiedForRaw(docentInfo.certifiedTours, tourType);
}

function isCertifiedForRaw(certList, tourType) {
  if (!certList) return true;
  var requiredTag = TOUR_TAG_MAP[tourType.toLowerCase()];
  if (!requiredTag) return true; // unknown tour type = no restriction
  return certList.indexOf(requiredTag) !== -1;
}

function formatTime(val) {
  if (!val) return '';
  // If Sheets stored it as a Date object, format it nicely
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'h:mm a');
  }
  return val.toString();
}

// =====================
// DAILY BACKUP
// =====================
function dailyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);

  // Create a dated copy
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var copyName = ss.getName() + ' - Backup ' + today;
  DriveApp.getFileById(ss.getId()).makeCopy(copyName, folder);

  // Delete backups older than BACKUP_KEEP_DAYS
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
