// ============================================================
// LIANA'S LOG — Google Apps Script v16
// Dashboard is now on GitHub Pages — this file is the webhook/API only
// Paste this entire file into: Extensions > Apps Script
// Then deploy as Web App (see setup guide)
// ============================================================

const LOG_SHEET_NAME = 'Log';
const SESSION_SHEET_NAME = 'LiveSession';
const LIANA_DOB = '2026-05-16';

// --- ENTRY POINT ---
function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();
  const ago = parseInt(e.parameter.ago) || 0;
  const timeParam = e.parameter.time || '';
  
  var result;
  if (action === 'status') {
    result = getStatusData();
  } else if (action === 'transcript') {
    const transcript = e.parameter.text || '';
    result = parseVoiceCommand(transcript);
  } else if (action === 'reports') {
    result = getReportsData();
  } else if (action === 'sleep_range') {
    const sleepType = (e.parameter.sleepType || '').toLowerCase();
    result = handleSleepRange(sleepType, e.parameter.startTime || '', e.parameter.endTime || '');
  } else if (action === 'note') {
    result = handleNote(e.parameter.note || '', ago, timeParam);
  } else if (action === 'trends') {
    result = getTrendsData();
  } else if (action) {
    result = processAction(action, ago, timeParam);
  } else {
    result = { status: 'ok', message: 'Liana Log API running' };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) {}
  const action = (body.action || e.parameter.action || '').toLowerCase();
  const ago = parseInt(body.ago || e.parameter.ago) || 0;
  const timeParam = body.time || e.parameter.time || '';
  
  var result;
  if (action === 'status') {
    result = getStatusData();
  } else if (action === 'transcript') {
    const transcript = body.text || e.parameter.text || '';
    result = parseVoiceCommand(transcript);
  } else if (action === 'reports') {
    result = getReportsData();
  } else if (action === 'sleep_range') {
    const sleepType = (body.sleepType || e.parameter.sleepType || '').toLowerCase();
    result = handleSleepRange(sleepType, body.startTime || e.parameter.startTime || '', body.endTime || e.parameter.endTime || '');
  } else if (action === 'note') {
    result = handleNote(body.note || e.parameter.note || '', ago, timeParam);
  } else if (action === 'trends') {
    result = getTrendsData();
  } else if (action) {
    result = processAction(action, ago, timeParam);
  } else {
    result = { status: 'error', message: 'No action specified' };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- ACTION ROUTER ---
function processAction(action, agoMinutes, timeParam) {
  var now;
  if (timeParam) {
    now = new Date(timeParam);
    if (isNaN(now.getTime())) now = new Date(); // fallback if bad date
  } else {
    now = new Date();
    if (agoMinutes && agoMinutes > 0) {
      now.setMinutes(now.getMinutes() - agoMinutes);
    }
  }
  if (action === 'left_breast') return handleBreast('left', now);
  if (action === 'right_breast') return handleBreast('right', now);
  if (action === 'pause_feeding') return handlePause(now);
  if (action === 'resume_feeding') return handleResume(now);
  if (action === 'done_feeding') return handleDone(now);
  if (action === 'pee_diaper') return logSimple(now, 'diaper', 'pee');
  if (action === 'poop_diaper') return logSimple(now, 'diaper', 'poop');
  if (action === 'mixed_diaper') return logSimple(now, 'diaper', 'mixed (pee + poop)');
  const bottleMatch = action.match(/^bottle_(\d+)$/);
  if (bottleMatch) return logSimple(now, 'bottle', parseInt(bottleMatch[1]) + ' mL');
  if (action === 'nap_start') return handleSleepStart(now, 'nap');
  if (action === 'nap_end') return handleSleepEnd(now, 'nap');
  if (action === 'bedtime') return handleSleepStart(now, 'bedtime');
  if (action === 'wake_up') return handleSleepEnd(now, 'bedtime');
  return { status: 'error', message: 'Unknown action: ' + action };
}

// Called from dashboard via google.script.run
function manualLog(action, minutesAgo, timeParam) {
  return processAction(action, parseInt(minutesAgo) || 0, timeParam || '');
}

function logSimple(now, category, detail) {
  const sheet = getOrCreateLogSheet();
  sheet.appendRow([now, category, detail, '', '', '', '', '']);
  return { status: 'ok', message: category + ': ' + detail };
}

// --- NURSING SESSION LOGIC ---

function getSessionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SESSION_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SESSION_SHEET_NAME);
    sheet.getRange('A1:B1').setValues([['Key', 'Value']]);
    const keys = [
      'active', 'current_side', 'status', 'session_start',
      'left_seconds', 'right_seconds', 'current_side_start',
      'pause_count', 'last_updated'
    ];
    keys.forEach((key, i) => {
      sheet.getRange(i + 2, 1).setValue(key);
      sheet.getRange(i + 2, 2).setValue('');
    });
  }
  return sheet;
}

function getSessionData() {
  const sheet = getSessionSheet();
  const data = {};
  const range = sheet.getRange('A2:B10').getValues();
  range.forEach(row => { if (row[0]) data[row[0]] = String(row[1]); });
  return data;
}

function setSessionData(data) {
  const sheet = getSessionSheet();
  const keys = [
    'active', 'current_side', 'status', 'session_start',
    'left_seconds', 'right_seconds', 'current_side_start',
    'pause_count', 'last_updated'
  ];
  keys.forEach((key, i) => {
    sheet.getRange(i + 2, 2).setValue(data[key] !== undefined ? data[key] : '');
  });
}

function clearSession() {
  const sheet = getSessionSheet();
  for (let i = 2; i <= 10; i++) { sheet.getRange(i, 2).setValue(''); }
}

// --- SLEEP SESSION (LiveSession rows 11-13) ---
function ensureSleepRows() {
  const sheet = getSessionSheet();
  const sleepKeys = ['sleep_active', 'sleep_type', 'sleep_start'];
  sleepKeys.forEach((key, i) => {
    const row = 11 + i;
    if (!sheet.getRange(row, 1).getValue()) {
      sheet.getRange(row, 1).setValue(key);
      sheet.getRange(row, 2).setValue('');
    }
  });
}

function getSleepData() {
  ensureSleepRows();
  const sheet = getSessionSheet();
  const data = {};
  sheet.getRange('A11:B13').getValues().forEach(row => { if (row[0]) data[String(row[0])] = String(row[1]); });
  return data;
}

function setSleepData(active, type, start) {
  ensureSleepRows();
  const sheet = getSessionSheet();
  sheet.getRange(11, 2).setValue(active);
  sheet.getRange(12, 2).setValue(type);
  sheet.getRange(13, 2).setValue(start);
}

function clearSleepData() {
  ensureSleepRows();
  const sheet = getSessionSheet();
  sheet.getRange(11, 2).setValue('');
  sheet.getRange(12, 2).setValue('');
  sheet.getRange(13, 2).setValue('');
}

function handleBreast(side, now) {
  const session = getSessionData();
  if (session.active !== 'true') {
    setSessionData({
      active: 'true', current_side: side, status: 'nursing',
      session_start: now.toISOString(), left_seconds: 0, right_seconds: 0,
      current_side_start: now.toISOString(), pause_count: 0,
      last_updated: now.toISOString()
    });
    return { status: 'ok', message: 'Started nursing on ' + side + ' breast' };
  }
  if (session.status === 'paused') {
    setSessionData({
      active: 'true', current_side: side, status: 'nursing',
      session_start: session.session_start,
      left_seconds: parseFloat(session.left_seconds) || 0,
      right_seconds: parseFloat(session.right_seconds) || 0,
      current_side_start: now.toISOString(),
      pause_count: parseFloat(session.pause_count) || 0,
      last_updated: now.toISOString()
    });
    return { status: 'ok', message: 'Resumed nursing on ' + side + ' breast' };
  }
  const elapsed = (now.getTime() - new Date(session.current_side_start).getTime()) / 1000;
  const prevSide = session.current_side;
  let leftSec = parseFloat(session.left_seconds) || 0;
  let rightSec = parseFloat(session.right_seconds) || 0;
  if (prevSide === 'left') leftSec += elapsed;
  else if (prevSide === 'right') rightSec += elapsed;
  setSessionData({
    active: 'true', current_side: side, status: 'nursing',
    session_start: session.session_start, left_seconds: leftSec, right_seconds: rightSec,
    current_side_start: now.toISOString(),
    pause_count: parseFloat(session.pause_count) || 0,
    last_updated: now.toISOString()
  });
  return { status: 'ok', message: prevSide !== side ? 'Switched to ' + side + ' breast' : 'Continuing on ' + side + ' breast' };
}

function handlePause(now) {
  const session = getSessionData();
  if (session.active !== 'true') return { status: 'ok', message: 'No active feeding session to pause' };
  if (session.status === 'paused') return { status: 'ok', message: 'Already paused' };
  const elapsed = (now.getTime() - new Date(session.current_side_start).getTime()) / 1000;
  let leftSec = parseFloat(session.left_seconds) || 0;
  let rightSec = parseFloat(session.right_seconds) || 0;
  if (session.current_side === 'left') leftSec += elapsed; else rightSec += elapsed;
  setSessionData({
    active: 'true', current_side: session.current_side, status: 'paused',
    session_start: session.session_start, left_seconds: leftSec, right_seconds: rightSec,
    current_side_start: '',
    pause_count: (parseFloat(session.pause_count) || 0) + 1,
    last_updated: now.toISOString()
  });
  return { status: 'ok', message: 'Feeding paused' };
}

function handleResume(now) {
  const session = getSessionData();
  if (session.active !== 'true') return { status: 'ok', message: 'No active feeding session to resume' };
  if (session.status !== 'paused') return { status: 'ok', message: 'Not paused — already nursing' };
  setSessionData({
    active: 'true', current_side: session.current_side, status: 'nursing',
    session_start: session.session_start,
    left_seconds: parseFloat(session.left_seconds) || 0,
    right_seconds: parseFloat(session.right_seconds) || 0,
    current_side_start: now.toISOString(),
    pause_count: parseFloat(session.pause_count) || 0,
    last_updated: now.toISOString()
  });
  return { status: 'ok', message: 'Resumed on ' + session.current_side + ' breast' };
}

function handleDone(now) {
  const session = getSessionData();
  if (session.active !== 'true') return { status: 'ok', message: 'No active feeding session' };
  let leftSec = parseFloat(session.left_seconds) || 0;
  let rightSec = parseFloat(session.right_seconds) || 0;
  let endTime = now;
  
  if (session.status === 'nursing' && session.current_side_start) {
    // Actively nursing: end time is right now, accumulate current side
    const elapsed = (now.getTime() - new Date(session.current_side_start).getTime()) / 1000;
    if (session.current_side === 'left') leftSec += elapsed; else rightSec += elapsed;
    endTime = now;
  } else if (session.status === 'paused' && session.last_updated) {
    // Paused: end time is when we last paused, not now
    endTime = new Date(session.last_updated);
  }
  
  const totalSec = leftSec + rightSec;
  const leftMin = Math.round(leftSec / 60 * 10) / 10;
  const rightMin = Math.round(rightSec / 60 * 10) / 10;
  const totalMin = Math.round(totalSec / 60 * 10) / 10;
  const sheet = getOrCreateLogSheet();
  const startTime = new Date(session.session_start);
  sheet.appendRow([
    endTime, 'nursing', 'T: ' + totalMin + ', L: ' + leftMin + ', R: ' + rightMin + ' min',
    startTime, totalMin, leftMin, rightMin, parseFloat(session.pause_count) || 0
  ]);
  clearSession();
  return { status: 'ok', message: 'Nursing done. Total: ' + totalMin + ' min (L: ' + leftMin + ', R: ' + rightMin + ')' };
}

// --- NOTES ---
function handleNote(text, agoMinutes, timeParam) {
  if (!text.trim()) return { status: 'error', message: 'Note text is required' };
  var now;
  if (timeParam) {
    now = new Date(timeParam);
    if (isNaN(now.getTime())) now = new Date();
  } else {
    now = new Date();
    if (agoMinutes && agoMinutes > 0) now.setMinutes(now.getMinutes() - agoMinutes);
  }
  const sheet = getOrCreateLogSheet();
  sheet.appendRow([now, 'note', text.trim(), '', '', '', '', '']);
  return { status: 'ok', message: 'Note saved' };
}

// --- SLEEP LOGIC ---
function handleSleepStart(now, type) {
  const sheet = getOrCreateLogSheet();
  sheet.appendRow([now, type, 'start', '', '', '', '', '']);
  setSleepData('true', type, now.toISOString());
  return { status: 'ok', message: type + ' started' };
}

function handleSleepEnd(now, type) {
  const sheet = getOrCreateLogSheet();
  const sleepData = getSleepData();
  let startTime = null;
  if (sleepData.sleep_active === 'true' && sleepData.sleep_type === type && sleepData.sleep_start) {
    startTime = new Date(sleepData.sleep_start);
    clearSleepData();
  } else {
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === type && data[i][2] === 'start') { startTime = new Date(data[i][0]); break; }
    }
  }
  let duration = '';
  if (startTime) {
    const diffMin = Math.round((now.getTime() - startTime.getTime()) / 60000);
    const hours = Math.floor(diffMin / 60);
    const mins = diffMin % 60;
    duration = hours > 0 ? hours + 'h ' + mins + 'm' : mins + 'm';
  }
  sheet.appendRow([now, type, 'end', startTime || '', duration, '', '', '']);
  return { status: 'ok', message: type + ' ended' + (duration ? ' (' + duration + ')' : '') };
}

function handleSleepRange(type, startTimeStr, endTimeStr) {
  if (!type || !startTimeStr || !endTimeStr) return { status: 'error', message: 'Missing type, startTime, or endTime' };
  const startTime = new Date(startTimeStr);
  const endTime = new Date(endTimeStr);
  if (isNaN(startTime.getTime()) || isNaN(endTime.getTime())) return { status: 'error', message: 'Invalid time values' };
  const diffMin = Math.round((endTime.getTime() - startTime.getTime()) / 60000);
  if (diffMin < 0) return { status: 'error', message: 'End time must be after start time' };
  const hours = Math.floor(diffMin / 60);
  const mins = diffMin % 60;
  const duration = hours > 0 ? hours + 'h ' + mins + 'm' : mins + 'm';
  const sheet = getOrCreateLogSheet();
  sheet.appendRow([startTime, type, 'start', '', '', '', '', '']);
  sheet.appendRow([endTime, type, 'end', startTime, duration, '', '', '']);
  return { status: 'ok', message: type + ' logged: ' + duration };
}

// --- LOG SHEET SETUP ---
function getOrCreateLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.getRange('A1:H1').setValues([['Timestamp', 'Category', 'Detail', 'Start Time', 'Total Min', 'Left Min', 'Right Min', 'Pauses']]);
    sheet.getRange('A1:H1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(4, 160);
  }
  return sheet;
}

// --- STATUS DATA ---
function getStatusData() {
  const session = getSessionData();
  const now = new Date();
  let leftSec = parseFloat(session.left_seconds) || 0;
  let rightSec = parseFloat(session.right_seconds) || 0;
  if (session.active === 'true' && session.status === 'nursing' && session.current_side_start) {
    const elapsed = (now.getTime() - new Date(session.current_side_start).getTime()) / 1000;
    if (session.current_side === 'left') leftSec += elapsed; else rightSec += elapsed;
  }
  const logSheet = getOrCreateLogSheet();
  const allData = logSheet.getDataRange().getValues();
  const recentRows = allData.slice(Math.max(1, allData.length - 20));
  const recent = recentRows.map(row => ({
    timestamp: row[0] ? new Date(row[0]).toISOString() : '',
    category: row[1] || '', detail: row[2] || '',
    startTime: row[3] ? new Date(row[3]).toISOString() : '',
    totalMin: row[4] || '', leftMin: row[5] || '', rightMin: row[6] || '', pauses: row[7] || ''
  })).reverse();
  const sleepData = getSleepData();
  const sleepElapsed = sleepData.sleep_active === 'true' && sleepData.sleep_start
    ? Math.round((now.getTime() - new Date(sleepData.sleep_start).getTime()) / 1000)
    : 0;

  return {
    nursing: {
      active: session.active === 'true', status: session.status || 'inactive',
      currentSide: session.current_side || '', leftSeconds: Math.round(leftSec),
      rightSeconds: Math.round(rightSec), totalSeconds: Math.round(leftSec + rightSec),
      pauseCount: parseFloat(session.pause_count) || 0, sessionStart: session.session_start || ''
    },
    sleep: {
      active: sleepData.sleep_active === 'true',
      type: sleepData.sleep_type || '',
      sessionStart: sleepData.sleep_start || '',
      elapsedSeconds: sleepElapsed
    },
    recentLog: recent, serverTime: now.toISOString()
  };
}

// Dashboard is hosted on GitHub Pages (index.html)

// --- REPORTS DATA ---
function getReportsData() {
  const sheet = getOrCreateLogSheet();
  const allData = sheet.getDataRange().getValues();
  
  // Get last 7 unique days that have data
  const now = new Date();
  const rows = allData.slice(1); // skip header
  
  // Find the last 7 days that have at least one log entry
  const daySet = new Set();
  rows.forEach(row => {
    if (row[0]) {
      const d = new Date(row[0]);
      daySet.add(d.toDateString());
    }
  });
  
  // Sort days descending, take last 7
  const sortedDays = Array.from(daySet)
    .sort((a, b) => new Date(b) - new Date(a))
    .slice(0, 7)
    .reverse(); // ascending for chart display
  
  // Build diaper counts per day
  const diapersByDay = {};
  const intakeByDay = {};
  sortedDays.forEach(day => {
    diapersByDay[day] = { pee: 0, poop: 0, mixed: 0, total: 0 };
    intakeByDay[day] = { bottleMl: 0, leftMin: 0, rightMin: 0 };
  });
  
  rows.forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[0]);
    const day = d.toDateString();
    const category = (row[1] || '').toLowerCase();
    const detail = (row[2] || '').toLowerCase();
    
    if (!diapersByDay[day]) return; // not in our 7-day window
    
    if (category === 'diaper') {
      if (detail === 'pee') diapersByDay[day].pee++;
      else if (detail === 'poop') diapersByDay[day].poop++;
      else if (detail.includes('mixed') || detail.includes('both')) diapersByDay[day].mixed++;
      diapersByDay[day].total++;
    }
  });
  
  // Format labels as "Jun 15"
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const labels = sortedDays.map(day => {
    const d = new Date(day);
    return months[d.getMonth()] + ' ' + d.getDate();
  });
  
  // Build bottle feeds — all bottles in the 7-day window
  // Each entry: { day (label), timeMinutes (minutes since midnight), ml, timeStr, dateStr }
  const bottleFeeds = [];
  
  rows.forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[0]);
    const day = d.toDateString();
    const category = (row[1] || '').toLowerCase();
    const detail = (row[2] || '');
    
    if (!diapersByDay[day]) return; // not in our 7-day window
    
    if (category === 'bottle') {
      const mlMatch = detail.match(/(\d+)/);
      if (!mlMatch) return;
      const ml = parseInt(mlMatch[1]);
      const hours = d.getHours();
      const minutes = d.getMinutes();
      const timeMinutes = hours * 60 + minutes;
      
      // Format time as "2:34 AM"
      const ampm = hours >= 12 ? 'PM' : 'AM';
      const h12 = hours % 12 || 12;
      const timeStr = h12 + ':' + String(minutes).padStart(2,'0') + ' ' + ampm;
      
      // Date label "Jun 15"
      const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      const dateStr = monthNames[d.getMonth()] + ' ' + d.getDate();
      
      // Day index (0 = oldest)
      const dayIndex = sortedDays.indexOf(day);
      
      bottleFeeds.push({ dayIndex, dateStr, timeMinutes, timeStr, ml });
    }
  });
  
  // Build nursing feeds — use start time (col D) and end time (col A), total min (col E)
  const nursingFeeds = [];
  const monthNamesN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  rows.forEach(row => {
    if (!row[0]) return;
    const endTime = new Date(row[0]);
    const category = (row[1] || '').toLowerCase();
    const detail = (row[2] || '');
    const startTimeRaw = row[3];
    const totalMin = parseFloat(row[4]) || 0;
    const leftMin = parseFloat(row[5]) || 0;
    const rightMin = parseFloat(row[6]) || 0;

    if (category !== 'nursing') return;
    if (!startTimeRaw || totalMin <= 0) return;

    const startTime = new Date(startTimeRaw);
    const day = startTime.toDateString();
    if (!diapersByDay[day]) return; // not in our 7-day window

    const hours = startTime.getHours();
    const minutes = startTime.getMinutes();
    const timeMinutes = hours * 60 + minutes;

    const ampm = hours >= 12 ? 'PM' : 'AM';
    const h12 = hours % 12 || 12;
    const timeStr = h12 + ':' + String(minutes).padStart(2,'0') + ' ' + ampm;
    const dateStr = monthNamesN[startTime.getMonth()] + ' ' + startTime.getDate();
    const dayIndex = sortedDays.indexOf(day);
    if (dayIndex === -1) return;

    nursingFeeds.push({ dayIndex, dateStr, timeMinutes, timeStr, totalMin, leftMin, rightMin });
  });

  // Build sleep starts — naps from start rows; bedtimes from end rows (day = evening start date)
  const napStarts = [];
  const bedtimeStarts = [];
  const durationByStartKey = {};
  const bedtimeCompletedStartMs = new Set();

  rows.forEach(row => {
    if (!row[0]) return;
    const category = (row[1] || '').toLowerCase();
    const detail = (row[2] || '').toLowerCase();
    if (category !== 'nap' || detail !== 'end' || !row[3]) return;
    const startKey = 'nap|' + new Date(row[3]).getTime();
    const dur = parseDurationMin(row[4]);
    if (dur) durationByStartKey[startKey] = dur;
  });

  rows.forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[0]);
    const day = d.toDateString();
    const category = (row[1] || '').toLowerCase();
    const detail = (row[2] || '').toLowerCase();

    if (!diapersByDay[day]) return;
    if (detail !== 'start') return;
    if (category !== 'nap') return;

    const hours = d.getHours();
    const minutes = d.getMinutes();
    const timeMinutes = hours * 60 + minutes;
    const ampm = hours >= 12 ? 'PM' : 'AM';
    const h12 = hours % 12 || 12;
    const timeStr = h12 + ':' + String(minutes).padStart(2,'0') + ' ' + ampm;
    const dateStr = months[d.getMonth()] + ' ' + d.getDate();
    const dayIndex = sortedDays.indexOf(day);
    if (dayIndex === -1) return;

    const startKey = 'nap|' + d.getTime();
    napStarts.push({
      dayIndex, dateStr, timeMinutes, timeStr,
      durationMin: durationByStartKey[startKey] || null
    });
  });

  // Bedtime: attribute to the evening start day (e.g. 9:30 PM Jun 21 → Jun 21, not wake day)
  rows.forEach(row => {
    if (!row[0] || !row[3]) return;
    const category = (row[1] || '').toLowerCase();
    const detail = (row[2] || '').toLowerCase();
    if (category !== 'bedtime' || detail !== 'end') return;

    const startTime = new Date(row[3]);
    const day = startTime.toDateString();
    if (!diapersByDay[day]) return;

    const startMs = startTime.getTime();
    bedtimeCompletedStartMs.add(startMs);

    const hours = startTime.getHours();
    const minutes = startTime.getMinutes();
    const timeMinutes = hours * 60 + minutes;
    const ampm = hours >= 12 ? 'PM' : 'AM';
    const h12 = hours % 12 || 12;
    const timeStr = h12 + ':' + String(minutes).padStart(2,'0') + ' ' + ampm;
    const dateStr = months[startTime.getMonth()] + ' ' + startTime.getDate();
    const dayIndex = sortedDays.indexOf(day);
    if (dayIndex === -1) return;

    bedtimeStarts.push({
      dayIndex, dateStr, timeMinutes, timeStr,
      durationMin: parseDurationMin(row[4])
    });
  });

  // In-progress bedtime (started but not yet ended)
  rows.forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[0]);
    const day = d.toDateString();
    const category = (row[1] || '').toLowerCase();
    const detail = (row[2] || '').toLowerCase();

    if (!diapersByDay[day]) return;
    if (category !== 'bedtime' || detail !== 'start') return;
    if (bedtimeCompletedStartMs.has(d.getTime())) return;

    const hours = d.getHours();
    const minutes = d.getMinutes();
    const timeMinutes = hours * 60 + minutes;
    const ampm = hours >= 12 ? 'PM' : 'AM';
    const h12 = hours % 12 || 12;
    const timeStr = h12 + ':' + String(minutes).padStart(2,'0') + ' ' + ampm;
    const dateStr = months[d.getMonth()] + ' ' + d.getDate();
    const dayIndex = sortedDays.indexOf(day);
    if (dayIndex === -1) return;

    bedtimeStarts.push({
      dayIndex, dateStr, timeMinutes, timeStr,
      durationMin: null
    });
  });

  // Daily intake totals — attributed to log timestamp day (col A)
  rows.forEach(row => {
    if (!row[0]) return;
    const day = new Date(row[0]).toDateString();
    if (!intakeByDay[day]) return;
    const category = (row[1] || '').toLowerCase();

    if (category === 'bottle') {
      const mlMatch = String(row[2] || '').match(/(\d+)/);
      if (mlMatch) intakeByDay[day].bottleMl += parseInt(mlMatch[1]);
    } else if (category === 'nursing') {
      const leftMin = parseFloat(row[5]) || 0;
      const rightMin = parseFloat(row[6]) || 0;
      if (leftMin > 0 || rightMin > 0) {
        intakeByDay[day].leftMin += leftMin;
        intakeByDay[day].rightMin += rightMin;
      }
    }
  });

  const round1 = v => Math.round(v * 10) / 10;

  return {
    diapers: {
      labels: labels,
      pee: sortedDays.map(d => diapersByDay[d].pee),
      poop: sortedDays.map(d => diapersByDay[d].poop),
      mixed: sortedDays.map(d => diapersByDay[d].mixed),
      total: sortedDays.map(d => diapersByDay[d].total)
    },
    bottles: {
      labels: labels,
      feeds: bottleFeeds
    },
    nursing: {
      labels: labels,
      feeds: nursingFeeds
    },
    sleep: {
      labels: labels,
      naps: napStarts,
      bedtimes: bedtimeStarts
    },
    intake: {
      labels: labels,
      bottleMl: sortedDays.map(d => intakeByDay[d].bottleMl),
      leftMin: sortedDays.map(d => round1(intakeByDay[d].leftMin)),
      rightMin: sortedDays.map(d => round1(intakeByDay[d].rightMin))
    }
  };
}

function parseDurationMin(str) {
  if (!str) return null;
  const s = String(str).trim();
  const hMatch = s.match(/(\d+)\s*h/);
  const mMatch = s.match(/(\d+)\s*m/);
  let total = 0;
  if (hMatch) total += parseInt(hMatch[1]) * 60;
  if (mMatch) total += parseInt(mMatch[1]);
  return total > 0 ? total : null;
}


// --- CLAUDE VOICE PARSING ---
function parseVoiceCommand(transcript) {
  var systemPrompt = "You are a baby logging assistant for a newborn named Liana.\n" +
    "Your job is to parse a parent's natural speech into a structured log action.\n\n" +
    "Available actions:\n" +
    "- left_breast: start/switch to left breast nursing\n" +
    "- right_breast: start/switch to right breast nursing\n" +
    "- pause_feeding: pause nursing session\n" +
    "- resume_feeding: resume nursing session\n" +
    "- done_feeding: end nursing session\n" +
    "- pee_diaper: wet/pee diaper\n" +
    "- poop_diaper: dirty/poop diaper\n" +
    "- mixed_diaper: both pee and poop diaper\n" +
    "- bottle_NNN: bottle feed where NNN is the exact mL (e.g. bottle_85 for 85 mL)\n" +
    "- nap_start: nap began\n" +
    "- nap_end: nap ended\n" +
    "- bedtime: night sleep started\n" +
    "- wake_up: woke up from night sleep\n\n" +
    "Also detect if the event happened in the past. If the user says 'X minutes ago' or 'X hours ago', return agoMinutes as a number. Otherwise return 0.\n\n" +
    "Respond ONLY with a JSON object, no markdown, no explanation:\n" +
    "{\"action\": \"action_name_here\", \"agoMinutes\": 0, \"confirmation\": \"A short friendly spoken confirmation\", \"error\": null}\n\n" +
    "If you cannot understand the command, return:\n" +
    "{\"action\": null, \"agoMinutes\": 0, \"confirmation\": null, \"error\": \"I didn't catch that. Try saying something like: left breast, pee diaper, or bottle 90.\"}";

  var payload = {
    model: "claude-sonnet-4-6",
    max_tokens: 200,
    system: systemPrompt,
    messages: [{ role: "user", content: transcript }]
  };

  var options = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY"),
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", options);
    var json = JSON.parse(response.getContentText());
    var text = json.content[0].text.replace(/```json|```/g, '').trim();
    var parsed = JSON.parse(text);
    if (parsed.action) {
      var result = processAction(parsed.action, parsed.agoMinutes || 0, '');
      return {
        action: parsed.action,
        agoMinutes: parsed.agoMinutes || 0,
        confirmation: parsed.confirmation || result.message,
        error: null
      };
    }
    return parsed;
  } catch(err) {
    return { action: null, agoMinutes: 0, confirmation: null, error: "Something went wrong. Please try again." };
  }
}


// --- TRENDS DATA (for Advice tab) ---
function getTrendsData() {
  const sheet = getOrCreateLogSheet();
  const allData = sheet.getDataRange().getValues();
  const rows = allData.slice(1);
  const now = new Date();
  const fiveDaysAgo = new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000);
  const oneDayAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  function fmtTime(d) {
    const h = d.getHours(), m = d.getMinutes();
    return (h % 12 || 12) + ':' + String(m).padStart(2,'0') + ' ' + (h >= 12 ? 'PM' : 'AM');
  }
  function fmtDate(d) { return months[d.getMonth()] + ' ' + d.getDate(); }

  const recentRows = rows.filter(r => r[0] && new Date(r[0]) >= fiveDaysAgo);
  if (recentRows.length === 0) {
    return { status: 'ok', summary: 'No data found in the last 5 days.', generatedAt: now.toISOString() };
  }

  const daySet = new Set();
  recentRows.forEach(r => { if (r[0]) daySet.add(new Date(r[0]).toDateString()); });
  const sortedDays = Array.from(daySet).sort((a, b) => new Date(a) - new Date(b));

  const dayData = {};
  sortedDays.forEach(day => {
    dayData[day] = { diapers: {pee:0, poop:0, mixed:0}, bottles: [], nursing: [], naps: [], bedtimes: [], notes: [] };
  });

  recentRows.forEach(r => {
    if (!r[0]) return;
    const d = new Date(r[0]);
    const category = (r[1]||'').toLowerCase();
    const detail = String(r[2]||'');

    if (category === 'diaper') {
      const det = detail.toLowerCase();
      const day = d.toDateString();
      if (!dayData[day]) return;
      if (det === 'pee') dayData[day].diapers.pee++;
      else if (det === 'poop') dayData[day].diapers.poop++;
      else if (det.includes('mixed')) dayData[day].diapers.mixed++;
    } else if (category === 'bottle') {
      const day = d.toDateString();
      if (!dayData[day]) return;
      const ml = parseInt((detail.match(/(\d+)/)||[])[1]) || 0;
      dayData[day].bottles.push(fmtTime(d) + ': ' + ml + 'mL');
    } else if (category === 'nursing' && r[3]) {
      const startTime = new Date(r[3]);
      const day = startTime.toDateString();
      if (!dayData[day]) return;
      const totalMin = parseFloat(r[4]) || 0;
      const leftMin = parseFloat(r[5]) || 0;
      const rightMin = parseFloat(r[6]) || 0;
      dayData[day].nursing.push(fmtTime(startTime) + '–' + fmtTime(d) + ' (' + totalMin + 'min, L:' + leftMin + ' R:' + rightMin + ')');
    } else if (category === 'nap' && detail.toLowerCase() === 'end' && r[3]) {
      const startTime = new Date(r[3]);
      const day = startTime.toDateString();
      if (!dayData[day]) return;
      dayData[day].naps.push(fmtTime(startTime) + '–' + fmtTime(d) + ' (' + (r[4]||'?') + ')');
    } else if (category === 'bedtime' && detail.toLowerCase() === 'end' && r[3]) {
      const startTime = new Date(r[3]);
      const day = startTime.toDateString();
      if (!dayData[day]) return;
      dayData[day].bedtimes.push(fmtTime(startTime) + '–' + fmtTime(d) + ' (' + (r[4]||'?') + ')');
    } else if (category === 'note') {
      const day = d.toDateString();
      if (!dayData[day]) return;
      dayData[day].notes.push(fmtTime(d) + ': ' + detail);
    }
  });

  let text = '=== LAST 5 DAYS ===\n\n';
  sortedDays.forEach(day => {
    const dd = dayData[day];
    text += '--- ' + fmtDate(new Date(day)) + ' ---\n';
    const totalD = dd.diapers.pee + dd.diapers.poop + dd.diapers.mixed;
    if (totalD > 0) text += 'Diapers: ' + totalD + ' (pee:' + dd.diapers.pee + ' poop:' + dd.diapers.poop + ' mixed:' + dd.diapers.mixed + ')\n';
    if (dd.bottles.length > 0) {
      const totalMl = dd.bottles.reduce((s,b) => s + (parseInt((b.match(/(\d+)mL/)||[])[1])||0), 0);
      text += 'Bottles: ' + dd.bottles.length + ' feeds, ' + totalMl + 'mL total (' + dd.bottles.join('; ') + ')\n';
    }
    if (dd.nursing.length > 0) text += 'Nursing: ' + dd.nursing.length + ' sessions (' + dd.nursing.join('; ') + ')\n';
    if (dd.naps.length > 0) text += 'Naps: ' + dd.naps.join('; ') + '\n';
    if (dd.bedtimes.length > 0) text += 'Bedtime: ' + dd.bedtimes.join('; ') + '\n';
    if (dd.notes.length > 0) text += 'Notes:\n' + dd.notes.map(n => '  ' + n).join('\n') + '\n';
    text += '\n';
  });

  text += '=== LAST 24 HOURS DETAILED TIMELINE ===\n';
  const last24 = rows
    .filter(r => r[0] && new Date(r[0]) >= oneDayAgo)
    .sort((a, b) => new Date(a[0]) - new Date(b[0]));

  if (last24.length === 0) {
    text += 'No events in the last 24 hours.\n';
  } else {
    last24.forEach(r => {
      const d = new Date(r[0]);
      const category = (r[1]||'').toLowerCase();
      const detail = String(r[2]||'');
      let line = fmtDate(d) + ' ' + fmtTime(d) + ' — ';
      if (category === 'nursing' && r[3]) {
        const st = new Date(r[3]);
        line += 'Nursing: ' + fmtTime(st) + '–' + fmtTime(d) + ', ' + (parseFloat(r[4])||0) + 'min (L:' + (parseFloat(r[5])||0) + ' R:' + (parseFloat(r[6])||0) + ')';
      } else if ((category === 'nap' || category === 'bedtime') && detail.toLowerCase() === 'end' && r[3]) {
        line += category + ' ended: started ' + fmtTime(new Date(r[3])) + ', duration ' + (r[4]||'?');
      } else {
        line += category + ': ' + detail;
      }
      text += line + '\n';
    });
  }

  const nowH = now.getHours(), nowM = now.getMinutes();
  const nowStr = months[now.getMonth()] + ' ' + now.getDate() + ', ' + now.getFullYear() +
    ' at ' + (nowH%12||12) + ':' + String(nowM).padStart(2,'0') + ' ' + (nowH>=12?'PM':'AM');

  const dob = new Date(LIANA_DOB);
  const ageMs = now.getTime() - dob.getTime();
  const ageWeeks = Math.floor(ageMs / (7 * 24 * 60 * 60 * 1000));
  const ageDays = Math.floor(ageMs / (24 * 60 * 60 * 1000));

  const systemPrompt =
    "You are analyzing baby tracking data for Liana (DOB May 16, 2026) and providing sleep guidance.\n" +
    "Current date/time: " + nowStr + ". Liana's age: " + ageWeeks + " weeks (" + ageDays + " days).\n\n" +
    "Write your response in EXACTLY two sections, using these exact markdown headers:\n\n" +
    "## Trends\n\n" +
    "Report only what happened. No developmental norms, no language like 'should' or 'typically'.\n" +
    "Just what the data shows, what the notes say, what co-occurred.\n" +
    "Use **5-Day Patterns** and **Last 24 Hours** as sub-headers with bullet points.\n\n" +
    "## Sleep Advice\n\n" +
    "Apply Marc Weissbluth's method (Healthy Sleep Habits, Happy Child) calibrated to " + ageWeeks + " weeks:\n" +
    "- Watch for drowsy cues (eye rubbing, yawning, glazed stare) and put down before overtired\n" +
    "- Overtiredness causes cortisol and makes sleep harder to initiate\n" +
    "- 0–6 weeks: wake windows ~45–60 min; 5–7 short naps; bedtime often late (9–11 PM)\n" +
    "- 6–12 weeks: wake windows ~1–1.5 hours; first morning nap becomes predictable; fussiness peaks ~6 wks then eases; bedtime begins moving earlier\n" +
    "- 3–4 months: wake windows ~1.5–2 hours; bedtime consolidates toward 6–8 PM\n" +
    "- Morning nap is most restorative — protect it first\n" +
    "- No formal sleep training before 4–6 months; focus on drowsy-but-awake\n" +
    "Reference specific patterns from today's data. Be concrete about timing.";

  const result = claudeCall(systemPrompt, 'Current time: ' + nowStr + '\n\n' + text, 1500);
  if (result.status === 'error') {
    const errMsg = 'API Error: ' + result.message;
    return { status: 'ok', summary: errMsg, sleepAdvice: errMsg, generatedAt: new Date().toISOString() };
  }

  const fullText = result.text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  // Find the Sleep Advice section with progressively looser patterns
  let sepIdx = -1;
  let sepLen = 0;
  const pats = [
    /\n#{1,3}[^\n]*sleep advice[^\n]*\n/i,
    /\n\*{1,2}[^\n]*sleep advice[^\n]*\n/i,
    /\nsleep advice[^\n]*\n/i
  ];
  for (let pi = 0; pi < pats.length; pi++) {
    const m = fullText.match(pats[pi]);
    if (m) { sepIdx = m.index; sepLen = m[0].length; break; }
  }

  const trendsRaw = sepIdx >= 0 ? fullText.substring(0, sepIdx).trim() : fullText;
  const summary = trendsRaw.replace(/^#{1,3}[^\n]*trends[^\n]*\n?/i, '').trim();
  const sleepAdvice = sepIdx >= 0 ? fullText.substring(sepIdx + sepLen).trim() : '';

  // If parsing found no Sleep Advice section, surface full text for diagnosis
  const sleepContent = sleepAdvice || '[Sleep Advice section not found in response]\n\nFull Claude output:\n\n' + fullText;

  return {
    status: 'ok',
    summary: summary,
    sleepAdvice: sleepContent,
    generatedAt: new Date().toISOString()
  };
}

function claudeCall(systemPrompt, userContent, maxTokens) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) return { status: 'error', message: 'ANTHROPIC_API_KEY not set in Script Properties' };

  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': key,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: maxTokens,
      system: systemPrompt,
      messages: [{ role: 'user', content: userContent }]
    }),
    muteHttpExceptions: true
  };

  let rawText = '';
  try {
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    rawText = response.getContentText();
    const json = JSON.parse(rawText);
    if (json.error) return { status: 'error', message: json.error.type + ': ' + json.error.message };
    if (!json.content || !json.content.length) return { status: 'error', message: 'API returned empty content array. Raw: ' + rawText.substring(0, 300) };
    const text = json.content[0].text;
    if (text === undefined || text === null) return { status: 'error', message: 'content[0].text missing. Raw: ' + rawText.substring(0, 300) };
    return { status: 'ok', text: String(text) };
  } catch(err) {
    return { status: 'error', message: err.toString() + ' | Raw response: ' + rawText.substring(0, 300) };
  }
}


function initialSetup() {
  getOrCreateLogSheet();
  getSessionSheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('Setup complete!');
}