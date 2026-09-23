// @ts-nocheck
/***********************
 * MAIN SYNC FUNCTION
 ***********************/
function syncGoogleCalendar2026() {
  try {
    const calendarId = 'ID of your Demo GC Sessions calendar (Calendar settings → Integrate calendar → Calendar ID)';
    const sheetName  = 'The Master Sheet Name';

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet not found");

    /* ---------- HEADER ---------- */
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'S. No', 'Event ID', 'Event Title', 'Patient Name',
        'Date', 'Time', 'Language', 'Pre/Post test GC',
        'ID', 'Test Name', 'Rescheduled', 'GC assigned', 'GC Status'
      ]);
    }

    /* ---------- EXISTING SIGNATURES ---------- */
    const lastRow      = sheet.getLastRow();
    const lastCol      = sheet.getLastColumn();
    const existingData = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
      : [];

    const existingMap = {};
    existingData.forEach(r => {
      if (r[1] && r[4] instanceof Date && r[5] instanceof Date) {
        existingMap[buildEventSignature(r[1], r[4], r[5])] = true;
      }
    });

    /* ---------- FETCH CALENDAR EVENTS ---------- */
    const events = CalendarApp.getCalendarById(calendarId).getEvents(
      new Date('2026-01-01T00:00:00'),
      new Date('2026-12-31T23:59:59')
    );

    const rowsToAdd = [];
    const noteMap   = {};   // eventId -> cleaned description (applied as notes after sort)

    events
      .filter(e => e && e.getTitle() && /^genetic counsel(l)?ing/i.test(e.getTitle().trim()))
      .forEach(event => {
        const eventId = event.getId();
        const start   = event.getStartTime();
        const desc    = event.getDescription() || "";

        const eventDate = new Date(start);
        eventDate.setHours(0, 0, 0, 0);

        // Capture the note for EVERY matching event (new OR already in the sheet)
    noteMap[eventId] = cleanDescription(desc);

    const eventTime  = new Date(start);
    const signature  = buildEventSignature(eventId, eventDate, eventTime);
    if (existingMap[signature]) return;

    const testDetails = extractTestDetails(desc);

        rowsToAdd.push([
          '', // S.No
          eventId,
          event.getTitle(),
          extractLabeledField(desc, 'Patient Name'),
          eventDate,
          eventTime,
          extractLanguageClean(desc),
          extractPrePostFromTitle(event.getTitle()),
          testDetails.ids,
          testDetails.names,
          '',
          '',
          ''
        ]);
      });

    /* ---------- APPEND NEW ROWS ---------- */
    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length)
        .setValues(rowsToAdd);
    }

    /* ---------- SORT, SERIAL, NOTES & HIGHLIGHT ---------- */
    fixAndSortGC2026();
    updateSerialNumbers2026();
    applyDescriptionNotes(sheet, noteMap);
    highlightTodaySessions();
    enforceDateTimeFormat();

    Logger.log("✅ Calendar sync completed successfully");

  } catch (e) {
    Logger.log("❌ ERROR: " + e.message);
  }
}

/***********************
 * TEST DETAILS EXTRACTION
 ***********************/
function extractTestDetails(text) {
  if (!text) return { ids: "", names: "" };

  const lines = text.split(/\n|<br>/i);
  const idPattern = /(TRID-\d{4}-\d+|STRAN-\d{4}-\d+)/i;

  let ids   = [];
  let names = [];

  lines.forEach(line => {
    const match = line.match(idPattern);
    if (match) {
      const id = match[0].toUpperCase();
      ids.push(id);
      const namePart = line.replace(idPattern, '').trim();
      if (namePart) names.push(namePart);
    }
  });

  return { ids: ids.join(', '), names: names.join(', ') };
}

/***********************
 * DESCRIPTION → HOVER NOTE
 ***********************/

/** Strip HTML from a calendar description so it reads cleanly as a note */
function cleanDescription(html) {
  if (!html) return "";
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li)>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** Attach each event's description as a NOTE on its Event Title cell (col C) */
function applyDescriptionNotes(sheet, noteMap) {
  if (!noteMap || !Object.keys(noteMap).length) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const ids        = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // col B – Event ID
  const notesRange = sheet.getRange(2, 3, lastRow - 1, 1);             // col C – Event Title
  const notes      = notesRange.getNotes();
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i][0];
    if (id && noteMap[id]) notes[i][0] = noteMap[id];
  }
  notesRange.setNotes(notes);
}

/** One-time: add hover notes to rows already in the sheet. Run from the editor. */
function backfillDescriptionNotes2026() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master Sheet');
  if (!sheet || sheet.getLastRow() < 2) return;
  const events = CalendarApp.getCalendarById('Calendar synced email ID').getEvents(
    new Date('2026-01-01T00:00:00'),
    new Date('2026-12-31T23:59:59')
  );
  const noteMap = {};
  events.forEach(e => {
    if (e && e.getTitle() && /^genetic counsel(l)?ing/i.test(e.getTitle().trim())) {
      noteMap[e.getId()] = cleanDescription(e.getDescription() || "");
    }
  });
  applyDescriptionNotes(sheet, noteMap);
  Logger.log("✅ Backfilled notes for existing rows");
}

/** Returns { eventId -> note text } from column C notes of the master sheet */
function getDescriptionNotes2026() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2026_GC_Master');
  if (!sheet || sheet.getLastRow() < 2) return {};
  const n     = sheet.getLastRow() - 1;
  const ids   = sheet.getRange(2, 2, n, 1).getValues();  // col B – Event ID
  const notes = sheet.getRange(2, 3, n, 1).getNotes();   // col C – hover note
  const map = {};
  for (let i = 0; i < n; i++) {
    if (ids[i][0]) map[ids[i][0]] = notes[i][0] || "";
  }
  return map;
}

/***********************
 * SORT BY DATE & TIME
 ***********************/
function fixAndSortGC2026() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2026_GC_Master');
  if (!sheet || sheet.getLastRow() < 2) return;
  const n = sheet.getLastRow() - 1;

  const dateRange = sheet.getRange(2, 5, n, 1); // col E – Date
  const timeRange = sheet.getRange(2, 6, n, 1); // col F – Time
  const dates = dateRange.getValues();
  const times = timeRange.getValues();

  for (let i = 0; i < n; i++) {
    let d = dates[i][0];
    let t = times[i][0];

    // Coerce Date → real Date at midnight
    if (d && !(d instanceof Date)) { const p = new Date(d); if (!isNaN(p)) d = p; }
    if (d instanceof Date) { const dd = new Date(d); dd.setHours(0,0,0,0); dates[i][0] = dd; }

    // Coerce Time → full date-time on that day (so ALL times are comparable)
    if (t && !(t instanceof Date)) { const p = new Date(t); if (!isNaN(p)) t = p; }
    if (t instanceof Date && d instanceof Date) {
      const full = new Date(d);
      full.setHours(t.getHours(), t.getMinutes(), t.getSeconds(), 0);
      times[i][0] = full;
    }
  }

  dateRange.setValues(dates);
  timeRange.setValues(times);

  // Now sort + re-stamp serials + re-apply format
  sheet.getRange(2, 1, n, sheet.getLastColumn())
       .sort([{ column: 5, ascending: true }, { column: 6, ascending: true }]);
  for (let r = 2; r <= sheet.getLastRow(); r++) sheet.getRange(r, 1).setValue(r - 1);
  sheet.getRange('E:E').setNumberFormat('ddd dd mmm yyyy');
  sheet.getRange('F:F').setNumberFormat('hh:mm AM/PM');

  Logger.log("✅ Date/Time normalized and sorted");
}

/***********************
 * SERIAL NUMBERS
 ***********************/
function updateSerialNumbers2026() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2026_GC_Master');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  for (let i = 2; i <= lastRow; i++) {
    sheet.getRange(i, 1).setValue(i - 1);
  }
}

/***********************
 * DATE / TIME FORMAT
 ***********************/
function enforceDateTimeFormat() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2026_GC_Master');
  if (!sheet) return;
  sheet.getRange('E:E').setNumberFormat('ddd dd mmm yyyy');
  sheet.getRange('F:F').setNumberFormat('hh:mm AM/PM');
}

/***********************
 * HIGHLIGHT TODAY + MONTH BANDS
 ***********************/
function highlightTodaySessions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2026_GC_Master');
  if (!sheet || sheet.getLastRow() < 2) return;

  const range  = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = range.getValues();
  const today  = new Date();
  today.setHours(0, 0, 0, 0);

  // Alternating month colors (12 months, soft tones)
  const monthColors = [
    '#FDECEC', // Jan – soft red
    '#FFF3E0', // Feb – soft peach
    '#FFF8E1', // Mar – soft yellow
    '#F1F8E9', // Apr – soft lime
    '#E8F5E9', // May – soft green
    '#E0F2F1', // Jun – soft teal
    '#E1F5FE', // Jul – soft sky
    '#E3F2FD', // Aug – soft blue
    '#EDE7F6', // Sep – soft lavender
    '#F3E5F5', // Oct – soft purple
    '#FCE4EC', // Nov – soft pink
    '#EFEBE9'  // Dec – soft brown
  ];

  const bg = values.map(r => {
    // Priority 1: Rescheduled = Orange  (col K = index 10)
    if (r[10] && r[10].toString().toUpperCase() === 'RS') {
      return Array(r.length).fill('#FCE8D6');
    }
    // Priority 2: Today = Green
    if (r[4] instanceof Date) {
      const d = new Date(r[4]);
      d.setHours(0, 0, 0, 0);
      if (d.getTime() === today.getTime()) {
        return Array(r.length).fill('#B7E1A1');
      }
      // Priority 3: Month band
      return Array(r.length).fill(monthColors[d.getMonth()]);
    }
    return Array(r.length).fill(null);
  });
  range.setBackgrounds(bg);

  const weights = values.map(r => {
    if (r[4] instanceof Date) {
      const d = new Date(r[4]);
      d.setHours(0, 0, 0, 0);
      if (d.getTime() === today.getTime()) return Array(r.length).fill('bold');
    }
    return Array(r.length).fill('normal');
  });
  range.setFontWeights(weights);
}

/***********************
 * AUTO SYNC TRIGGER
 ***********************/
function createAutoSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncGoogleCalendar2026') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncGoogleCalendar2026').timeBased().everyMinutes(30).create();
}

/***********************
 * DAILY CHAT TRIGGER
 ***********************/
function createDailyGCChatTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyGCScheduleToChat') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyGCScheduleToChat')
    .timeBased().everyDays(1).atHour(8).create();
}

/***********************
 * DAILY SCHEDULE — CHAT
 ***********************/
function sendDailyGCScheduleToChat() {
  const sessions = getTodaysValidGCSessions();
  if (!sessions.length) return;
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
  sendGCChatMessage(buildTodayGCChatTable(todayStr, sessions));
}

function getTodaysValidGCSessions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2026_GC_Master');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = {
    date:     headers.indexOf('Date'),
    time:     headers.indexOf('Time'),
    patient:  headers.indexOf('Patient Name'),
    language: headers.indexOf('Language'),
    type:     headers.indexOf('Pre/Post test GC'),
    gc:       headers.indexOf('GC assigned')
  };

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const sessions = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[idx.date] || !r[idx.time]) continue;
    const d = new Date(r[idx.date]);
    d.setHours(0, 0, 0, 0);
    if (d.getTime() !== today.getTime()) continue;
    if (![idx.patient, idx.language, idx.type, idx.gc].every(c => r[c])) continue;
    sessions.push({
      time:     Utilities.formatDate(new Date(r[idx.time]), Session.getScriptTimeZone(), 'hh:mm a'),
      patient:  r[idx.patient],
      language: r[idx.language],
      type:     r[idx.type],
      gc:       r[idx.gc]
    });
  }
  return sessions;
}

function buildTodayGCChatTable(dateStr, sessions) {
  let table =
    pad('Time', 8)          + ' | ' +
    pad('Patient Name', 16) + ' | ' +
    pad('Language', 8)      + ' | ' +
    pad('Type', 9)          + ' | ' +
    pad('GC Assigned', 12)  + '\n' +
    '─'.repeat(60) + '\n';

  sessions.forEach(s => {
    table +=
      pad(s.time, 8)     + ' | ' +
      pad(s.patient, 16) + ' | ' +
      pad(s.language, 8) + ' | ' +
      pad(s.type, 9)     + ' | ' +
      pad(s.gc, 12)      + '\n';
  });

  return `📅 ${dateStr} – Today's Genetic Counseling Schedule\n\n\`\`\`\n${table}\`\`\``;
}

function sendGCChatMessage(message) {
  const webhookUrl = 'WebhookUrl;
  UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: message })
  });
}

/***********************
 * HELPERS
 ***********************/
function buildEventSignature(eventId, date, time) {
  return eventId + '|' + date.getTime() + '|' + time.getTime();
}

function extractLabeledField(text, label) {
  if (!text) return "";
  const match = text.match(new RegExp(label + "\\s*[:\\-]?\\s*(.*)", "i"));
  if (!match) return "";
  return match[1].split("<br>")[0].replace(/<[^>]*>/g, '').trim();
}

function extractPrePostFromTitle(title) {
  const t = title.toLowerCase();
  if (t.includes('pre'))  return 'Pre-test';
  if (t.includes('post')) return 'Post-test';
  return '';
}

function extractCaseIdFromText(text) {
  if (!text) return "";
  const match = text.match(/(TRID-\d{4}-\d+|ID-\d{4}-\d+)/i);
  return match ? match[0].toUpperCase() : '';
}

function extractLanguageClean(text) {
  if (!text) return "";
  const raw = extractLabeledField(text, 'Preferred Language');
  if (!raw) return "";
  const allowed = ['English','Hindi','Gujarati','Tamil','Telugu','Kannada','Malayalam','Marathi','Bengali'];
  return allowed.filter(lang => new RegExp(lang, 'i').test(raw)).join(' / ');
}

function pad(text, length) {
  text = text ? text.toString() : '';
  return text.length > length ? text.substring(0, length) : text + ' '.repeat(length - text.length);
}
