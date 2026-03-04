/**
 * Pay Period Report Generator v3
 * Bound to: 2026 check-ins spreadsheet (1rHQ1YNUUkTcmUWwy1tXksM6tWG0GUoOOG3pRrsv9DKs)
 *
 * Features:
 *  - Date picker dialog — pick start/end, auto-creates a new tab named with the range
 *  - Dynamic column detection across varying weekly sheet layouts
 *  - Fuzzy leader name matching (merges typos/partial names into one row)
 *  - Duration lookup from Ops Hub with school+workshop fuzzy matching
 *  - Color-coded rows: red=absences, yellow=incomplete check-ins, green=all verified
 *  - Action Needed column for quick HR triage
 *  - Summary stats bar, frozen headers, auto-filter
 *
 * SETUP: Paste into Extensions > Apps Script, save, reload sheet.
 */

var OPS_HUB_SS_ID = '17hnG_MZs81GFoz_lyJNVyMYocrZLweGkZ6fTxn5UCTs';
var OPS_HUB_SHEET_NAME = 'Winter/Spring 26';

var ACTIVE_STATUSES = ['leader', 'co-lead', 'onboard', 'sub'];
var ABSENT_STATUSES = ['absent'];

// Column count in output (A-J)
var NUM_COLS = 10;

/* ═══════════════════════ Menu & Triggers ═══════════════════════ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pay Period Tools')
    .addItem('Generate Report...', 'showDatePicker')
    .addItem('Refresh Current Report', 'refreshCurrentReport')
    .addItem('Set Up Daily Auto-Update (9 PM)', 'createDailyTrigger')
    .addToUi();
}

function createDailyTrigger() {
  // Auto-update finds the most recent Pay Period Report tab and refreshes it
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshCurrentReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('refreshCurrentReport')
    .timeBased().atHour(21).everyDays(1).create();
  SpreadsheetApp.getUi().alert('Daily trigger set! The most recent pay period report will auto-update at 9 PM.');
}

/* ═══════════════════════ Date Picker Dialog ═══════════════════════ */

function showDatePicker() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    '  body { font-family: Google Sans, Arial, sans-serif; padding: 16px; }' +
    '  h3 { margin: 0 0 16px 0; color: #1a73e8; }' +
    '  label { display: block; font-weight: 500; margin: 12px 0 4px 0; color: #333; }' +
    '  input[type=date] { width: 100%; padding: 8px; font-size: 14px; border: 1px solid #ddd; border-radius: 4px; }' +
    '  .buttons { margin-top: 20px; text-align: right; }' +
    '  .btn { padding: 8px 24px; font-size: 14px; border: none; border-radius: 4px; cursor: pointer; margin-left: 8px; }' +
    '  .btn-primary { background: #1a73e8; color: white; }' +
    '  .btn-primary:hover { background: #1557b0; }' +
    '  .btn-cancel { background: #f1f3f4; color: #333; }' +
    '  .btn-cancel:hover { background: #e0e0e0; }' +
    '  .hint { font-size: 12px; color: #666; margin-top: 4px; }' +
    '</style>' +
    '<h3>Generate Pay Period Report</h3>' +
    '<label for="start">Pay Period Start</label>' +
    '<input type="date" id="start" value="' + getDefaultStart_() + '">' +
    '<label for="end">Pay Period End</label>' +
    '<input type="date" id="end" value="' + getDefaultEnd_() + '">' +
    '<p class="hint">A new tab will be created named with this date range (e.g. "2/19-3/4 Pay Period Report").</p>' +
    '<div class="buttons">' +
    '  <button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '  <button class="btn btn-primary" onclick="submit()">Generate</button>' +
    '</div>' +
    '<script>' +
    '  function submit() {' +
    '    var s = document.getElementById("start").value;' +
    '    var e = document.getElementById("end").value;' +
    '    if (!s || !e) { alert("Please select both dates."); return; }' +
    '    if (new Date(s) >= new Date(e)) { alert("Start date must be before end date."); return; }' +
    '    google.script.run.withSuccessHandler(function() { google.script.host.close(); }).runReport(s, e);' +
    '  }' +
    '</script>'
  )
  .setWidth(360)
  .setHeight(320);

  SpreadsheetApp.getUi().showModalDialog(html, 'Pay Period Report');
}

/** Default start = most recent Thursday (biweekly pay period start) */
function getDefaultStart_() {
  var d = new Date();
  // Find most recent Thursday (day 4)
  var day = d.getDay();
  var diff = (day >= 4) ? (day - 4) : (day + 3);
  d.setDate(d.getDate() - diff);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Default end = 13 days after start (2-week period) */
function getDefaultEnd_() {
  var d = new Date();
  var day = d.getDay();
  var diff = (day >= 4) ? (day - 4) : (day + 3);
  d.setDate(d.getDate() - diff + 13);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* ═══════════════════════ Main Entry Points ═══════════════════════ */

/**
 * Called from the date picker dialog. Creates a new tab and generates the report.
 */
function runReport(startStr, endStr) {
  var payStart = new Date(startStr + 'T00:00:00');
  var payEnd = new Date(endStr + 'T23:59:59');
  payStart.setHours(0, 0, 0, 0);
  payEnd.setHours(23, 59, 59, 999);

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Build tab name like "2/19-3/4 Pay Period Report"
  var tabName = (payStart.getMonth() + 1) + '/' + payStart.getDate() + '-' +
                (payEnd.getMonth() + 1) + '/' + payEnd.getDate() + ' Pay Period Report';

  // Check if tab already exists
  var reportSheet = ss.getSheetByName(tabName);
  if (reportSheet) {
    // Clear and reuse
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(tabName);
  }

  // Set up config rows
  reportSheet.getRange('A2').setValue('Pay Period Start:');
  reportSheet.getRange('B2').setValue(payStart);
  reportSheet.getRange('B2').setNumberFormat('m/d/yyyy');
  reportSheet.getRange('C2').setValue('Pay Period End:');
  reportSheet.getRange('D2').setValue(payEnd);
  reportSheet.getRange('D2').setNumberFormat('m/d/yyyy');

  // Run the report
  generateReportOnSheet_(ss, reportSheet, payStart, payEnd);

  // Navigate to the new tab
  ss.setActiveSheet(reportSheet);
}

/**
 * "Refresh Current Report" — re-runs the report on whichever Pay Period Report tab is active.
 * Also used by the daily auto-update trigger.
 */
function refreshCurrentReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var active = ss.getActiveSheet();
  var reportSheet = null;

  // If the active sheet is a pay period report, refresh it
  if (active.getName().indexOf('Pay Period Report') !== -1) {
    reportSheet = active;
  } else {
    // Find the most recently created Pay Period Report tab (rightmost)
    var sheets = ss.getSheets();
    for (var i = sheets.length - 1; i >= 0; i--) {
      if (sheets[i].getName().indexOf('Pay Period Report') !== -1) {
        reportSheet = sheets[i];
        break;
      }
    }
  }

  if (!reportSheet) {
    SpreadsheetApp.getUi().alert('No Pay Period Report tab found. Use "Generate Report..." to create one.');
    return;
  }

  var startDate = reportSheet.getRange('B2').getValue();
  var endDate = reportSheet.getRange('D2').getValue();

  if (!startDate || !endDate) {
    SpreadsheetApp.getUi().alert('Could not read dates from row 2 of "' + reportSheet.getName() + '".');
    return;
  }

  var payStart = new Date(startDate);
  var payEnd = new Date(endDate);
  payStart.setHours(0, 0, 0, 0);
  payEnd.setHours(23, 59, 59, 999);

  generateReportOnSheet_(ss, reportSheet, payStart, payEnd);
  ss.setActiveSheet(reportSheet);
}

/**
 * Core report generation — runs on a given sheet with given dates.
 */
function generateReportOnSheet_(ss, reportSheet, payStart, payEnd) {
  // 1. Build duration lookup from Ops Hub
  var durationMap = buildDurationMap_();

  // 2. Parse all relevant weekly sheets
  var allSessions = [];
  var allAbsences = [];
  var sheets = ss.getSheets();

  for (var s = 0; s < sheets.length; s++) {
    var name = sheets[s].getName();
    if (!name.match(/^Week of \d+\/\d+$/i)) continue;

    var weekDate = parseWeekDate_(name);
    if (!weekDate) continue;

    var weekEnd = new Date(weekDate);
    weekEnd.setDate(weekEnd.getDate() + 6);
    if (weekEnd < payStart || weekDate > payEnd) continue;

    var parsed = parseWeeklySheet_(sheets[s], payStart, payEnd);
    allSessions = allSessions.concat(parsed.sessions);
    allAbsences = allAbsences.concat(parsed.absences);
  }

  // 3. Aggregate by leader
  var leaderData = aggregateByLeader_(allSessions, allAbsences, durationMap);

  // 4. Fuzzy-merge similar leader names
  leaderData = mergeSimiLarLeaders_(leaderData);

  // 5. Write output with formatting
  writeReport_(reportSheet, leaderData, payStart, payEnd);

  // 6. Generate criticality tab
  generateCriticalityTab_(ss, leaderData, payStart, payEnd);

  var leaderCount = Object.keys(leaderData).length;
  SpreadsheetApp.getUi().alert(
    'Report generated!\n' +
    allSessions.length + ' active sessions  |  ' +
    leaderCount + ' unique leaders (after name merge)'
  );
}

/* ═══════════════════════ Criticality Tab ═══════════════════════ */

/**
 * Generates (or refreshes) a "Leader Criticality" tab alongside the pay period report.
 * Ranks leaders by sessions descending. Tiers: CRITICAL (10+), HIGH (6-9), MEDIUM (3-5), LOW (1-2).
 * Unique school count and back-to-back flag show replaceability risk.
 */
function generateCriticalityTab_(ss, leaderData, payStart, payEnd) {
  var tabName = 'Leader Criticality';
  var tab = ss.getSheetByName(tabName);
  if (tab) {
    tab.clear();
  } else {
    tab = ss.insertSheet(tabName);
  }

  var dateRange = formatDate_(payStart) + ' – ' + formatDate_(payEnd);

  // Title
  tab.getRange('A1').setValue('Leader Criticality Report');
  tab.getRange('A1').setFontSize(14).setFontWeight('bold');
  tab.getRange('B1').setValue(dateRange).setFontSize(11).setFontColor('#666666');

  // Legend
  tab.getRange('A2').setValue('CRITICAL = 10+ sessions  |  HIGH = 6-9  |  MEDIUM = 3-5  |  LOW = 1-2');
  tab.getRange('A2').setFontSize(10).setFontColor('#555555').setFontStyle('italic');

  // Headers
  var headers = ['Rank', 'Leader Name', 'Sessions', 'Hours', 'Unique Schools', 'Back-to-Back', 'Criticality', 'Schools Covered'];
  tab.getRange(3, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  // Build sorted list
  var lKeys = Object.keys(leaderData);
  var list = [];
  for (var i = 0; i < lKeys.length; i++) {
    var ld = leaderData[lKeys[i]];
    var sessionCount = ld.sessions.length;
    if (sessionCount === 0) continue; // skip pure-absence leaders

    // Unique schools
    var schools = {};
    for (var s = 0; s < ld.sessions.length; s++) {
      if (ld.sessions[s].school) schools[ld.sessions[s].school.toLowerCase()] = ld.sessions[s].school;
    }
    var uniqueSchools = Object.keys(schools).length;
    var schoolList = Object.values(schools).join('; ');

    // Back-to-back flag
    var b2b = detectBackToBack_(ld.sessions);

    // Tier
    var tier, tierColor;
    if (sessionCount >= 10) {
      tier = 'CRITICAL'; tierColor = '#cc0000';
    } else if (sessionCount >= 6) {
      tier = 'HIGH'; tierColor = '#e67c00';
    } else if (sessionCount >= 3) {
      tier = 'MEDIUM'; tierColor = '#f4b400';
    } else {
      tier = 'LOW'; tierColor = '#0f9d58';
    }

    var hours = Math.round((ld.totalMinutes / 60) * 100) / 100;
    var displayName = ld.name;
    if (ld.mergedNames && ld.mergedNames.length > 0) {
      displayName += '  [also: ' + ld.mergedNames.join(', ') + ']';
    }

    list.push({
      name: displayName, sessions: sessionCount, hours: hours,
      uniqueSchools: uniqueSchools, b2b: b2b ? '✓ ' + b2b : '',
      tier: tier, tierColor: tierColor, schoolList: schoolList
    });
  }

  // Sort by sessions descending, then hours descending
  list.sort(function(a, b) {
    if (b.sessions !== a.sessions) return b.sessions - a.sessions;
    return b.hours - a.hours;
  });

  if (list.length === 0) return;

  var rows = [];
  var tierColors = [];
  for (var r = 0; r < list.length; r++) {
    var item = list[r];
    rows.push([r + 1, item.name, item.sessions, item.hours, item.uniqueSchools, item.b2b, item.tier, item.schoolList]);
    tierColors.push(item.tierColor);
  }

  tab.getRange(4, 1, rows.length, headers.length).setValues(rows);

  // Color the Criticality column (col 7) and Rank col (col 1)
  for (var rc = 0; rc < rows.length; rc++) {
    tab.getRange(4 + rc, 7).setFontColor(tierColors[rc]).setFontWeight('bold');
    // Light row tint based on tier
    var rowBg = null;
    if (tierColors[rc] === '#cc0000') rowBg = '#fce8e6';
    else if (tierColors[rc] === '#e67c00') rowBg = '#fef3e2';
    else if (tierColors[rc] === '#f4b400') rowBg = '#fef9e3';
    if (rowBg) tab.getRange(4 + rc, 1, 1, headers.length).setBackground(rowBg);
  }

  // Summary counts
  var critCount = list.filter(function(x) { return x.tier === 'CRITICAL'; }).length;
  var highCount = list.filter(function(x) { return x.tier === 'HIGH'; }).length;
  var medCount  = list.filter(function(x) { return x.tier === 'MEDIUM'; }).length;
  var lowCount  = list.filter(function(x) { return x.tier === 'LOW'; }).length;
  tab.getRange('C1').setValue(
    'CRITICAL: ' + critCount + '  |  HIGH: ' + highCount +
    '  |  MEDIUM: ' + medCount + '  |  LOW: ' + lowCount
  ).setFontSize(10).setFontColor('#333333');

  // Freeze header rows, auto-resize
  tab.setFrozenRows(3);
  for (var c = 1; c <= headers.length; c++) tab.autoResizeColumn(c);
  tab.setColumnWidth(6, 300); // back-to-back col can be wide
  tab.setColumnWidth(8, 350); // schools col
}

/* ═══════════════════════ Fuzzy Name Matching ═══════════════════════ */

/**
 * Merge leader entries whose names are likely the same person.
 * Handles typos (Eila/Elia), partial names, accented chars.
 */
function mergeSimiLarLeaders_(leaderData) {
  var keys = Object.keys(leaderData);
  var assigned = {};
  var groups = [];

  for (var i = 0; i < keys.length; i++) {
    if (assigned[keys[i]]) continue;
    var group = [keys[i]];
    assigned[keys[i]] = true;

    for (var j = i + 1; j < keys.length; j++) {
      if (assigned[keys[j]]) continue;
      if (areNamesSimilar_(leaderData[keys[i]].name, leaderData[keys[j]].name)) {
        group.push(keys[j]);
        assigned[keys[j]] = true;
      }
    }

    // Also check transitively: if B matched A, does C match B?
    var expanded = true;
    while (expanded) {
      expanded = false;
      for (var t = 0; t < keys.length; t++) {
        if (assigned[keys[t]]) continue;
        for (var g = 0; g < group.length; g++) {
          if (areNamesSimilar_(leaderData[group[g]].name, leaderData[keys[t]].name)) {
            group.push(keys[t]);
            assigned[keys[t]] = true;
            expanded = true;
            break;
          }
        }
      }
    }

    groups.push(group);
  }

  // Merge each group into one entry using longest name as canonical
  var merged = {};
  for (var gi = 0; gi < groups.length; gi++) {
    var grp = groups[gi];

    // Canonical = longest name (most complete version)
    var canonIdx = 0;
    for (var ci = 1; ci < grp.length; ci++) {
      if (leaderData[grp[ci]].name.length > leaderData[grp[canonIdx]].name.length) {
        canonIdx = ci;
      }
    }
    var canonKey = grp[canonIdx];

    var canon = {
      name: leaderData[canonKey].name,
      sessions: [],
      totalMinutes: 0,
      dates: {},
      workshopsSchools: [],
      checkIns: [],
      absences: [],
      mergedNames: []
    };

    for (var mi = 0; mi < grp.length; mi++) {
      var src = leaderData[grp[mi]];
      canon.sessions = canon.sessions.concat(src.sessions);
      canon.totalMinutes += src.totalMinutes;
      canon.checkIns = canon.checkIns.concat(src.checkIns);
      canon.absences = canon.absences.concat(src.absences);

      var dk = Object.keys(src.dates);
      for (var d = 0; d < dk.length; d++) canon.dates[dk[d]] = true;

      for (var w = 0; w < src.workshopsSchools.length; w++) {
        if (canon.workshopsSchools.indexOf(src.workshopsSchools[w]) === -1) {
          canon.workshopsSchools.push(src.workshopsSchools[w]);
        }
      }

      if (grp[mi] !== canonKey) {
        canon.mergedNames.push(src.name);
      }
    }

    merged[canonKey] = canon;
  }

  return merged;
}

/**
 * Check if two names likely refer to the same person.
 */
function areNamesSimilar_(name1, name2) {
  var n1 = stripAccents_(name1.toLowerCase().trim());
  var n2 = stripAccents_(name2.toLowerCase().trim());
  if (n1 === n2) return true;

  var p1 = n1.split(/\s+/);
  var p2 = n2.split(/\s+/);
  var first1 = p1[0], last1 = p1[p1.length - 1];
  var first2 = p2[0], last2 = p2[p2.length - 1];

  // Same first name + similar last name
  // Allow distance 2 only for last names with 4+ chars (catches transpositions like Eila/Elia)
  // Short last names (Ban/Bui, Lee/Leo) require distance <= 1 to avoid false positives
  if (first1 === first2) {
    var lastThreshold = (last1.length >= 4 && last2.length >= 4) ? 2 : 1;
    if (levenshtein_(last1, last2) <= lastThreshold) return true;
  }

  // One name's words are all found (with tolerance) in the other
  if (isWordSubset_(p1, p2) || isWordSubset_(p2, p1)) return true;

  // Very similar overall (short names with small edit distance)
  if (n1.length < 25 && n2.length < 25 && levenshtein_(n1, n2) <= 2) return true;

  return false;
}

/**
 * Check if every word in `shorter` appears in `longer` (edit distance <= 1 per word).
 */
function isWordSubset_(shorter, longer) {
  if (shorter.length > longer.length) return false;
  if (shorter.length === 0) return false;

  // Single-word names: only match if the word matches the FIRST word (first name)
  // of the longer name. Prevents "Lee" from bridging "Jubilyn Lee" and "Leo Lee".
  if (shorter.length === 1) {
    return levenshtein_(shorter[0], longer[0]) <= 1;
  }

  for (var i = 0; i < shorter.length; i++) {
    var found = false;
    for (var j = 0; j < longer.length; j++) {
      if (levenshtein_(shorter[i], longer[j]) <= 1) {
        found = true;
        break;
      }
    }
    if (!found) return false;
  }
  return true;
}

function levenshtein_(a, b) {
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  var matrix = [];
  for (var i = 0; i <= a.length; i++) {
    matrix[i] = [i];
  }
  for (var j = 0; j <= b.length; j++) {
    matrix[0][j] = j;
  }
  for (var i2 = 1; i2 <= a.length; i2++) {
    for (var j2 = 1; j2 <= b.length; j2++) {
      var cost = a[i2 - 1] === b[j2 - 1] ? 0 : 1;
      matrix[i2][j2] = Math.min(
        matrix[i2 - 1][j2] + 1,
        matrix[i2][j2 - 1] + 1,
        matrix[i2 - 1][j2 - 1] + cost
      );
    }
  }
  return matrix[a.length][b.length];
}

function stripAccents_(str) {
  var from = '\u00e0\u00e1\u00e2\u00e3\u00e4\u00e5\u00e8\u00e9\u00ea\u00eb\u00ec\u00ed\u00ee\u00ef\u00f2\u00f3\u00f4\u00f5\u00f6\u00f9\u00fa\u00fb\u00fc\u00fd\u00f1\u00e7';
  var to   = 'aaaaaaeeeeiiiioooooeuuuuync';
  var result = '';
  for (var i = 0; i < str.length; i++) {
    var idx = from.indexOf(str[i]);
    result += idx !== -1 ? to[idx] : str[i];
  }
  // Strip remaining non-ASCII chars (handles mojibake like "Cu√É¬©llar" → "Cullar")
  result = result.replace(/[^\x20-\x7E]/g, '');
  return result;
}

/* ═══════════════════════ Ops Hub Duration Map ═══════════════════════ */

function buildDurationMap_() {
  var opsHub = SpreadsheetApp.openById(OPS_HUB_SS_ID);
  var sheet = opsHub.getSheetByName(OPS_HUB_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var site = String(data[i][2]).trim();
    var startTime = data[i][5];
    var endTime = data[i][6];
    var lesson = String(data[i][8]).trim();
    if (!site || !lesson) continue;

    var duration = calculateDuration_(startTime, endTime);
    if (duration <= 0 || duration > 720) continue;

    var key = normalizeSchool_(site) + '|||' + normalizeWorkshop_(lesson);
    if (!map[key]) map[key] = duration;

    var schoolKey = 'SCHOOL_ONLY|||' + normalizeSchool_(site);
    if (!map[schoolKey]) map[schoolKey] = [];
    if (Array.isArray(map[schoolKey])) map[schoolKey].push(duration);
  }

  // Convert school-only arrays to median
  var keys = Object.keys(map);
  for (var k = 0; k < keys.length; k++) {
    if (keys[k].indexOf('SCHOOL_ONLY|||') === 0 && Array.isArray(map[keys[k]])) {
      var arr = map[keys[k]];
      arr.sort(function(a, b) { return a - b; });
      map[keys[k]] = arr[Math.floor(arr.length / 2)];
    }
  }
  return map;
}

function calculateDuration_(startTime, endTime) {
  if (startTime instanceof Date && endTime instanceof Date) {
    return (endTime.getTime() - startTime.getTime()) / 60000;
  }
  if (typeof startTime === 'number' && typeof endTime === 'number') {
    return (endTime - startTime) * 24 * 60;
  }
  var s = parseTimeToMinutes_(String(startTime));
  var e = parseTimeToMinutes_(String(endTime));
  if (s === null || e === null) return 0;
  return e - s;
}

function parseTimeToMinutes_(timeStr) {
  var m = timeStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
  if (!m) return null;
  var h = parseInt(m[1], 10);
  var min = parseInt(m[2], 10);
  var p = m[3].toUpperCase();
  if (p === 'PM' && h !== 12) h += 12;
  if (p === 'AM' && h === 12) h = 0;
  return h * 60 + min;
}

/* ═══════════════════════ Normalization ═══════════════════════ */

function normalizeSchool_(name) {
  return String(name).toLowerCase().trim()
    .replace(/\s+/g, ' ')
    .replace(/\b(elementary|school|academy|the|campus|extended care|of)\b/g, '')
    .replace(/\s+/g, ' ').trim();
}

function normalizeWorkshop_(name) {
  return String(name).toLowerCase().trim()
    .replace(/\(wk\s*\d+\)/gi, '')
    .replace(/\(week\s*\d+\)/gi, '')
    .replace(/part\s*\d+:?\s*/gi, '')
    .replace(/\(grades?\s*[\d\-\u2013]+\)/gi, '')
    .replace(/[\/&]/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

/* ═══════════════════════ Weekly Sheet Parser ═══════════════════════ */

function parseWeekDate_(sheetName) {
  var m = sheetName.match(/Week of (\d+)\/(\d+)/i);
  if (!m) return null;
  return new Date(2026, parseInt(m[1], 10) - 1, parseInt(m[2], 10));
}

function parseWeeklySheet_(sheet, payStart, payEnd) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { sessions: [], absences: [] };

  var headers = [];
  for (var h = 0; h < data[0].length; h++) {
    headers.push(String(data[0][h]).toLowerCase().trim().replace(/\n/g, ' '));
  }

  var cols = {
    leader:   findColumn_(headers, ['leader name', 'leader']),
    workshop: findColumn_(headers, ['workshop']),
    school:   findColumn_(headers, ['school']),
    date:     findColumn_(headers, ['date']),
    time:     findColumn_(headers, ['est time', 'time']),
    status:   findColumn_(headers, ['status']),
    checkIn:  findColumn_(headers, ['check-in status', 'check in status'])
  };

  if (cols.leader === -1 || cols.status === -1) {
    return { sessions: [], absences: [] };
  }

  var sessions = [];
  var absences = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (isSeparatorRow_(row, cols)) continue;

    var leaderRaw = String(row[cols.leader] || '').trim();
    if (!leaderRaw) continue;

    var leader = leaderRaw
      .replace(/\n[\s\S]*/g, '')
      .replace(/\s*\(ABSENT\)\s*/gi, '')
      .trim();
    if (!leader) continue;

    var status = String(row[cols.status] || '').trim().toLowerCase();
    var workshop = cols.workshop !== -1 ? String(row[cols.workshop] || '').trim() : '';
    var school   = cols.school !== -1   ? String(row[cols.school] || '').trim()   : '';
    var dateVal  = cols.date !== -1     ? row[cols.date]                          : '';
    var timeVal  = cols.time !== -1     ? String(row[cols.time] || '').trim()     : '';
    var checkIn  = cols.checkIn !== -1  ? String(row[cols.checkIn] || '').trim()  : '';

    var sessionDate = parseSessionDate_(dateVal);
    if (!sessionDate) continue;
    if (sessionDate < payStart || sessionDate > payEnd) continue;

    var isActive = false;
    for (var a = 0; a < ACTIVE_STATUSES.length; a++) {
      if (status.indexOf(ACTIVE_STATUSES[a]) !== -1) { isActive = true; break; }
    }
    var isAbsent = false;
    for (var b = 0; b < ABSENT_STATUSES.length; b++) {
      if (status.indexOf(ABSENT_STATUSES[b]) !== -1) { isAbsent = true; break; }
    }

    var dateStr = formatDate_(sessionDate);

    if (isActive) {
      sessions.push({
        leader: leader, workshop: workshop, school: school,
        date: sessionDate, dateStr: dateStr, time: timeVal,
        checkIn: checkIn, sheetName: sheet.getName()
      });
    } else if (isAbsent) {
      absences.push({
        leader: leader, date: sessionDate, dateStr: dateStr,
        school: school, workshop: workshop
      });
    }
  }
  // Second pass: scan all rows for "Co-lead: Name" references in any column.
  // This picks up co-leads who don't have their own row (e.g. Alex Chen in Week of 2/9).
  for (var ci = 1; ci < data.length; ci++) {
    var crow = data[ci];
    if (isSeparatorRow_(crow, cols)) continue;

    var cDateVal = cols.date !== -1 ? crow[cols.date] : '';
    var cSessionDate = parseSessionDate_(cDateVal);
    if (!cSessionDate) continue;
    if (cSessionDate < payStart || cSessionDate > payEnd) continue;

    var cWorkshop = cols.workshop !== -1 ? String(crow[cols.workshop] || '').trim() : '';
    var cSchool   = cols.school !== -1   ? String(crow[cols.school] || '').trim()   : '';
    var cTimeVal  = cols.time !== -1     ? String(crow[cols.time] || '').trim()     : '';
    var cDateStr  = formatDate_(cSessionDate);

    // Scan every cell in this row for "Co-lead: ..." pattern
    for (var cc = 0; cc < crow.length; cc++) {
      var cellText = String(crow[cc] || '');
      var coMatches = cellText.match(/[Cc]o-?lead:?\s*([^|\n]+)/g);
      if (!coMatches) continue;

      for (var cm = 0; cm < coMatches.length; cm++) {
        var namesPart = coMatches[cm].replace(/[Cc]o-?lead:?\s*/g, '').trim();
        var coNames = namesPart.split(/[&,]/);

        for (var cn = 0; cn < coNames.length; cn++) {
          var coName = coNames[cn].trim()
            .replace(/\n[\s\S]*/g, '')
            .replace(/\s*\(.*?\)\s*/g, '')
            .trim();
          if (!coName || coName.length < 2) continue;

          // Skip if this co-lead is already the primary leader of this row
          var primaryLeader = String(crow[cols.leader] || '').trim()
            .replace(/\n[\s\S]*/g, '').replace(/\s*\(ABSENT\)\s*/gi, '').trim();
          if (coName.toLowerCase() === primaryLeader.toLowerCase()) continue;

          // Skip if we already have this co-lead for this exact school+date+workshop
          var dupeKey = coName.toLowerCase() + '|' + cSchool.toLowerCase() + '|' + cDateStr;
          var isDupe = false;
          for (var ds = 0; ds < sessions.length; ds++) {
            var sk = sessions[ds].leader.toLowerCase() + '|' + sessions[ds].school.toLowerCase() + '|' + sessions[ds].dateStr;
            if (sk === dupeKey) { isDupe = true; break; }
          }
          if (isDupe) continue;

          sessions.push({
            leader: coName, workshop: cWorkshop, school: cSchool,
            date: cSessionDate, dateStr: cDateStr, time: cTimeVal,
            checkIn: '', sheetName: sheet.getName()
          });
        }
      }
    }
  }

  return { sessions: sessions, absences: absences };
}

function findColumn_(headers, keywords) {
  var i, k;
  for (k = 0; k < keywords.length; k++) {
    for (i = 0; i < headers.length; i++) {
      if (headers[i] === keywords[k]) return i;
    }
  }
  for (k = 0; k < keywords.length; k++) {
    for (i = 0; i < headers.length; i++) {
      if (headers[i].indexOf(keywords[k]) === 0) return i;
    }
  }
  for (k = 0; k < keywords.length; k++) {
    if (keywords[k].indexOf(' ') === -1) continue;
    for (i = 0; i < headers.length; i++) {
      if (headers[i].indexOf(keywords[k]) !== -1) return i;
    }
  }
  return -1;
}

function isSeparatorRow_(row, cols) {
  var dayNames = ['monday','tuesday','wednesday','thursday','friday','saturday','sunday'];
  var leaderVal = String(row[cols.leader] || '').trim();

  if (!leaderVal) {
    for (var c = 0; c < row.length; c++) {
      var val = String(row[c] || '').trim().toLowerCase();
      for (var d = 0; d < dayNames.length; d++) {
        if (val.indexOf(dayNames[d]) === 0) return true;
      }
    }
    return true;
  }

  var ll = leaderVal.toLowerCase();
  for (var d2 = 0; d2 < dayNames.length; d2++) {
    if (ll.indexOf(dayNames[d2]) === 0) return true;
  }

  if (cols.workshop !== -1 && cols.school !== -1) {
    var wVal = String(row[cols.workshop] || '').trim().toLowerCase();
    var sVal = String(row[cols.school] || '').trim();
    if (!sVal) {
      for (var d3 = 0; d3 < dayNames.length; d3++) {
        if (wVal.indexOf(dayNames[d3]) !== -1) return true;
      }
    }
  }
  return false;
}

/* ═══════════════════════ Date Parsing ═══════════════════════ */

function parseSessionDate_(dateVal) {
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    return new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate());
  }
  var str = String(dateVal).trim();
  if (!str) return null;

  var m1 = str.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (m1) {
    var mo = parseInt(m1[1], 10) - 1;
    var dy = parseInt(m1[2], 10);
    var yr = m1[3] ? parseInt(m1[3], 10) : 2026;
    if (yr < 100) yr += 2000;
    return new Date(yr, mo, dy);
  }

  var m2 = str.match(/^([A-Za-z]+)-(\d{1,2})$/);
  if (m2) {
    var months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
    var idx = months.indexOf(m2[1].toLowerCase().substring(0, 3));
    if (idx !== -1) return new Date(2026, idx, parseInt(m2[2], 10));
  }
  return null;
}

function formatDate_(date) {
  return (date.getMonth() + 1) + '/' + date.getDate();
}

/* ═══════════════════════ Duration Lookup ═══════════════════════ */

function lookupDuration_(school, workshop, durationMap) {
  var ns = normalizeSchool_(school);
  var nw = normalizeWorkshop_(workshop);

  var exact = ns + '|||' + nw;
  if (durationMap[exact]) return { minutes: durationMap[exact], matched: true };

  var keys = Object.keys(durationMap);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf('SCHOOL_ONLY|||') === 0) continue;
    var parts = keys[i].split('|||');
    if ((ns.indexOf(parts[0]) !== -1 || parts[0].indexOf(ns) !== -1) &&
        (nw.indexOf(parts[1]) !== -1 || parts[1].indexOf(nw) !== -1)) {
      return { minutes: durationMap[keys[i]], matched: true };
    }
  }

  var sk = 'SCHOOL_ONLY|||' + ns;
  if (durationMap[sk]) return { minutes: durationMap[sk], matched: true };

  for (var j = 0; j < keys.length; j++) {
    if (keys[j].indexOf('SCHOOL_ONLY|||') !== 0) continue;
    var ms = keys[j].replace('SCHOOL_ONLY|||', '');
    if (ns.indexOf(ms) !== -1 || ms.indexOf(ns) !== -1) {
      return { minutes: durationMap[keys[j]], matched: true };
    }
  }

  return { minutes: 60, matched: false };
}

/* ═══════════════════════ Aggregation ═══════════════════════ */

function aggregateByLeader_(sessions, absences, durationMap) {
  var leaders = {};

  for (var i = 0; i < sessions.length; i++) {
    var s = sessions[i];
    var key = s.leader.toLowerCase();
    if (!leaders[key]) {
      leaders[key] = {
        name: s.leader, sessions: [], totalMinutes: 0,
        dates: {}, workshopsSchools: [], checkIns: [], absences: [], mergedNames: []
      };
    }

    var dur = lookupDuration_(s.school, s.workshop, durationMap);
    leaders[key].sessions.push(s);
    leaders[key].totalMinutes += dur.minutes;
    leaders[key].dates[s.dateStr] = true;

    var cleanWS = s.workshop
      .replace(/\(Wk\s*\d+\)/gi, '').replace(/\(Week\s*\d+\)/gi, '').trim();
    var wsEntry = cleanWS + ' @ ' + s.school;
    if (leaders[key].workshopsSchools.indexOf(wsEntry) === -1) {
      leaders[key].workshopsSchools.push(wsEntry);
    }
    leaders[key].checkIns.push({ dateStr: s.dateStr, checkIn: s.checkIn, sheetName: s.sheetName });
  }

  for (var j = 0; j < absences.length; j++) {
    var ab = absences[j];
    var aKey = ab.leader.toLowerCase();
    if (!leaders[aKey]) {
      leaders[aKey] = {
        name: ab.leader, sessions: [], totalMinutes: 0,
        dates: {}, workshopsSchools: [], checkIns: [], absences: [], mergedNames: []
      };
    }
    leaders[aKey].absences.push(ab.dateStr + ' @ ' + ab.school);
  }
  return leaders;
}

/* ═══════════════════════ Back-to-Back Detection ═══════════════════════ */

function detectBackToBack_(sessions) {
  if (sessions.length < 2) return '';

  var groups = {};
  for (var i = 0; i < sessions.length; i++) {
    var gk = sessions[i].school.toLowerCase() + '|' + sessions[i].dateStr;
    if (!groups[gk]) groups[gk] = { school: sessions[i].school, dateStr: sessions[i].dateStr, count: 0 };
    groups[gk].count++;
  }

  var bySchool = {};
  var gKeys = Object.keys(groups);
  for (var j = 0; j < gKeys.length; j++) {
    var g = groups[gKeys[j]];
    if (g.count < 2) continue;
    var sk = g.school.toLowerCase();
    if (!bySchool[sk]) bySchool[sk] = { school: g.school, dates: [], maxCount: 0 };
    bySchool[sk].dates.push(g.dateStr);
    if (g.count > bySchool[sk].maxCount) bySchool[sk].maxCount = g.count;
  }

  var entries = [];
  var sKeys = Object.keys(bySchool);
  for (var k = 0; k < sKeys.length; k++) {
    var info = bySchool[sKeys[k]];
    info.dates.sort(function(a, b) { return new Date('2026/' + a) - new Date('2026/' + b); });
    entries.push(info.maxCount + 'x back-to-back @ ' + info.school + ' ' + info.dates.join(' AND '));
  }
  return entries.join('; ');
}

/* ═══════════════════════ Check-In Summary ═══════════════════════ */

function buildCheckInSummary_(checkIns) {
  var byDate = {};
  for (var i = 0; i < checkIns.length; i++) {
    if (!byDate[checkIns[i].dateStr]) byDate[checkIns[i].dateStr] = [];
    byDate[checkIns[i].dateStr].push(checkIns[i]);
  }

  var sorted = Object.keys(byDate).sort(function(a, b) {
    return new Date('2026/' + a) - new Date('2026/' + b);
  });

  var totalChecked = 0, totalAll = 0;
  var parts = [];
  for (var d = 0; d < sorted.length; d++) {
    var items = byDate[sorted[d]];
    var total = items.length;
    var checked = 0;
    var types = {};

    for (var c = 0; c < items.length; c++) {
      var val = items[c].checkIn;
      if (val && val.toLowerCase() !== 'false' && val.trim() !== '') {
        checked++;
        types[val] = true;
      }
    }

    totalChecked += checked;
    totalAll += total;

    if (checked === 0) {
      parts.push(sorted[d] + ': empty');
    } else if (checked === total) {
      var tl = Object.keys(types);
      var suffix = total > 1 ? ' (all ' + total + ')' : '';
      parts.push(sorted[d] + ': \u2713 ' + tl.join(', ') + suffix);
    } else {
      parts.push(sorted[d] + ': ' + checked + '/' + total + ' verified');
    }
  }

  return { text: parts.join(' | '), checked: totalChecked, total: totalAll };
}

/* ═══════════════════════ Report Writer ═══════════════════════ */

function writeReport_(sheet, leaderData, payStart, payEnd) {
  // Clear everything from row 4 down
  var lastRow = Math.max(sheet.getLastRow(), 4);
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, NUM_COLS).clearContent();
    sheet.getRange(4, 1, lastRow - 3, NUM_COLS).setBackground(null);
  }

  // ── Row 1: Title + summary stats ──
  sheet.getRange('A1').setValue('Leader Pay Period Report');
  sheet.getRange('A1').setFontWeight('bold').setFontSize(14);

  // ── Row 2: Config ──
  sheet.getRange('A2').setValue('Pay Period Start:');
  sheet.getRange('C2').setValue('Pay Period End:');

  // ── Row 3: Headers ──
  var headers = [
    'Leader Name', 'Total Sessions', 'Total Hours', 'Total Minutes',
    'Dates Taught', 'Workshops & Schools', 'Back-to-Back Details',
    'Check-In Summary', 'Absences in Period', 'Action Needed'
  ];
  sheet.getRange(3, 1, 1, NUM_COLS).setValues([headers]);
  sheet.getRange(3, 1, 1, NUM_COLS).setFontWeight('bold').setBackground('#4a86c8').setFontColor('#ffffff');

  // Sort leaders alphabetically
  var lKeys = Object.keys(leaderData);
  var sorted = [];
  for (var i = 0; i < lKeys.length; i++) sorted.push(leaderData[lKeys[i]]);
  sorted.sort(function(a, b) {
    return a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
  });

  if (sorted.length === 0) return;

  var rows = [];
  var colors = []; // background color per row
  var grandTotalSessions = 0, grandTotalMinutes = 0;

  for (var j = 0; j < sorted.length; j++) {
    var leader = sorted[j];
    var totalSessions = leader.sessions.length;
    var totalMinutes = leader.totalMinutes;
    var totalHours = Math.round((totalMinutes / 60) * 100) / 100;

    grandTotalSessions += totalSessions;
    grandTotalMinutes += totalMinutes;

    var dateKeys = Object.keys(leader.dates);
    dateKeys.sort(function(a, b) { return new Date('2026/' + a) - new Date('2026/' + b); });
    var dates = dateKeys.join(', ');

    var workshopsSchools = leader.workshopsSchools.join('; ');
    var backToBack = detectBackToBack_(leader.sessions);
    var ciResult = buildCheckInSummary_(leader.checkIns);
    var absencesStr = leader.absences.length > 0 ? leader.absences.join('; ') : 'None';

    // Display name: append merged aliases if any
    var displayName = leader.name;
    if (leader.mergedNames && leader.mergedNames.length > 0) {
      displayName += '  [also: ' + leader.mergedNames.join(', ') + ']';
    }

    // Build action flags
    var actions = [];
    if (leader.absences.length > 0) actions.push('\u26a0 Absent');
    if (ciResult.total > 0 && ciResult.checked < ciResult.total) {
      actions.push('\u26a0 Check-ins: ' + ciResult.checked + '/' + ciResult.total);
    }
    if (totalSessions === 0 && leader.absences.length > 0) actions.push('\u26a0 No active sessions');
    var actionStr = actions.length > 0 ? actions.join(' | ') : '\u2713 OK';

    // Row color
    var rowColor = null;
    if (leader.absences.length > 0) {
      rowColor = '#f4cccc'; // light red
    } else if (ciResult.total > 0 && ciResult.checked < ciResult.total) {
      rowColor = '#fff2cc'; // light yellow
    } else if (ciResult.total > 0 && ciResult.checked === ciResult.total) {
      rowColor = '#d9ead3'; // light green
    }

    rows.push([
      displayName, totalSessions, totalHours, totalMinutes,
      dates, workshopsSchools, backToBack,
      ciResult.text, absencesStr, actionStr
    ]);
    colors.push(rowColor);
  }

  // Write data
  sheet.getRange(4, 1, rows.length, NUM_COLS).setValues(rows);

  // Apply row colors
  for (var r = 0; r < colors.length; r++) {
    if (colors[r]) {
      sheet.getRange(4 + r, 1, 1, NUM_COLS).setBackground(colors[r]);
    }
  }

  // Summary stats in row 1
  var grandHours = Math.round((grandTotalMinutes / 60) * 100) / 100;
  sheet.getRange('F1').setValue(
    'Leaders: ' + sorted.length +
    '  |  Sessions: ' + grandTotalSessions +
    '  |  Total Hours: ' + grandHours +
    '  |  Period: ' + formatDate_(payStart) + ' \u2013 ' + formatDate_(payEnd)
  );
  sheet.getRange('F1').setFontSize(10).setFontColor('#666666');

  // Freeze top 3 rows
  sheet.setFrozenRows(3);

  // Auto-resize columns
  for (var c = 1; c <= NUM_COLS; c++) {
    sheet.autoResizeColumn(c);
  }

  // Set filter on data range
  var dataRange = sheet.getRange(3, 1, rows.length + 1, NUM_COLS);
  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  dataRange.createFilter();
}
