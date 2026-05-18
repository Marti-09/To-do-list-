// ═══════════════════════════════════════════════════════════
// AttendanceSync.gs — CrossChex Cloud → Google Sheets
// Paste into: Google Sheets → Extensions → Apps Script
// Trigger: Time-based, daily, 5:00–6:00 PM
// ═══════════════════════════════════════════════════════════

const CROSSCHEX_API_URL    = 'https://api.us.crosschexcloud.com/';
const CROSSCHEX_API_KEY    = '91b049b867bd282c196a4f0e22724df1';
const CROSSCHEX_API_SECRET = 'bb88467757b488677cb17bd9b052dc97';
const SPREADSHEET_ID_GS    = '10BtYmx-wtCw95YFTRaRkBYiGZRQf7VQoprN4r7tmIa4';
const SHEET_TAB            = 'Attendance';

// ── ENTRY POINT ───────────────────────────────────────────
function syncAttendance() {
  const token              = getToken();
  const { monday, saturday } = getWeekRange();
  const records            = getAttendanceRecords(token, monday, saturday);
  const byEmployee         = groupByEmployee(records);
  writeToSheet(byEmployee, monday);
  Logger.log('Done. Employees synced: ' + Object.keys(byEmployee).length);
}

// ── AUTH ──────────────────────────────────────────────────
function getToken() {
  const res = callApi({
    header: {
      nameSpace:   'authorize.token',
      nameAction:  'token',
      version:     '1.0',
      requestId:   Utilities.getUuid(),
      timestamp:   new Date().toISOString()
    },
    payload: {
      api_key:    CROSSCHEX_API_KEY,
      api_secret: CROSSCHEX_API_SECRET
    }
  });
  if (!res.payload?.token) throw new Error('CrossChex: no token returned. Check API key/secret.');
  Logger.log('Token acquired. Expires: ' + res.payload.expires);
  return res.payload.token;
}

// ── FETCH ATTENDANCE ──────────────────────────────────────
function getAttendanceRecords(token, startDate, endDate) {
  let allRecords = [];
  let page = 1;
  while (true) {
    const res = callApi({
      header: {
        nameSpace:  'attendance.record',
        nameAction: 'getrecord',
        version:    '1.0',
        requestId:  Utilities.getUuid(),
        timestamp:  new Date().toISOString()
      },
      authorize: { type: 'token', token: token },
      payload: {
        begin_time: toApiTime(startDate),
        end_time:   toApiTime(endDate),
        order:      'asc',
        page:       String(page),
        per_page:   '100'
      }
    });
    const list = res.payload?.list || [];
    Logger.log('Page ' + page + ': ' + list.length + ' records');
    if (page === 1 && list.length > 0) Logger.log('Sample record: ' + JSON.stringify(list[0]));
    allRecords = allRecords.concat(list);
    if (list.length < 100) break;
    page++;
  }
  return allRecords;
}

// ── GROUP BY EMPLOYEE ─────────────────────────────────────
function groupByEmployee(records) {
  const map = {};
  records.forEach(r => {
    const name = r.employee
      ? (r.employee.first_name + ' ' + r.employee.last_name).trim()
      : ('ID-' + (r.emp_code || r.id || '?'));
    if (!map[name]) map[name] = [];
    map[name].push(r);
  });
  return map;
}

// ── WRITE TO SHEET ────────────────────────────────────────
function writeToSheet(byEmployee, monday) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID_GS);
  let sheet   = ss.getSheetByName(SHEET_TAB);
  if (!sheet) sheet = ss.insertSheet(SHEET_TAB);

  sheet.getRange(1, 1, 1, 10).setValues([
    ['Employee','Mon','Tue','Wed','Thu','Fri','Sat','Days','Hours','LastSync']
  ]);

  const syncTime = new Date().toISOString();
  // Employees who stay as 'IN' when tap-out is missing (no assumed hours)
  const NO_DEFAULT_HOURS = ['mb', 'simran'];

  const rows = Object.keys(byEmployee).sort().map(name => {
    const recs        = byEmployee[name];
    let totalMins     = 0;
    let anyHours      = false;
    const useDefault  = !NO_DEFAULT_HOURS.some(ex => name.toLowerCase().includes(ex));

    const dayCells = [0,1,2,3,4,5].map(offset => {
      const dayDate = new Date(monday);
      dayDate.setDate(monday.getDate() + offset);
      const dayStr = Utilities.formatDate(dayDate, 'UTC', 'yyyy-MM-dd');

      const dayRecs = recs
        .filter(r => {
          const ct = r.checktime || r.check_time || r.time || '';
          return ct.startsWith(dayStr);
        })
        .sort((a, b) => (a.checktime || a.check_time || '').localeCompare(b.checktime || b.check_time || ''));

      if (!dayRecs.length) return '';

      const firstTap = dayRecs[0].checktime || dayRecs[0].check_time || '';
      if (dayRecs.length < 2) {
        if (useDefault) {
          totalMins += 480;
          anyHours   = true;
          return '08:00-16:00';
        }
        return 'IN';
      }

      const lastTap  = dayRecs[dayRecs.length - 1].checktime || dayRecs[dayRecs.length - 1].check_time || '';
      const inDate   = new Date(firstTap);
      const outDate  = new Date(lastTap);
      const mins     = Math.round((outDate - inDate) / 60000);
      totalMins     += mins;
      anyHours       = true;
      return formatTime(firstTap) + '-' + formatTime(lastTap);
    });

    const daysPresent  = dayCells.filter(c => c !== '').length;
    const hoursDecimal = anyHours ? parseFloat((totalMins / 60).toFixed(2)) : '';
    return [name, ...dayCells, daysPresent, hoursDecimal, syncTime];
  });

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 10).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 10).setValues(rows);
}

// ── HELPERS ───────────────────────────────────────────────
function callApi(payload) {
  const response = UrlFetchApp.fetch(CROSSCHEX_API_URL, {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const json = JSON.parse(response.getContentText());
  if (json.payload?.type === 'AUTH_ERROR') throw new Error('CrossChex AUTH_ERROR — check API key/secret');
  return json;
}

function getWeekRange() {
  const now  = new Date();
  const day  = now.getDay();
  const diff = (day === 0) ? -6 : 1 - day;
  const monday = new Date(now);
  monday.setDate(now.getDate() + diff);
  monday.setHours(0, 0, 0, 0);
  const saturday = new Date(monday);
  saturday.setDate(monday.getDate() + 5);
  saturday.setHours(23, 59, 59, 0);
  return { monday, saturday };
}

function toApiTime(date) {
  return date.toISOString().replace('Z', '+00:00');
}

function formatTime(isoString) {
  if (!isoString) return '';
  const d = new Date(isoString);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
}
