/**
 * לוח חופשות חטיבת דביר אבו - צד השרת (Google Apps Script)
 * הנתונים נשמרים בגוגל שיט: לשונית העובדים הקיימת + לשונית "חופשות".
 */
const CONFIG = {
  SHEET_ID: '1AGX6sL9M_HcDvxPi4BQ6aUEVuamlVIg46iouRZGc-uY',
  ADMINS: ['dvir@amsalem.com'],          // מי מאשר חופשות ועורך הכול
  LEAVES_SHEET: 'חופשות',
  REPORT_DATE: '2026-09-23',              // התאריך שהיתרות בגיליון נכונות אליו
  TZ: 'Asia/Jerusalem',
  TITLE: 'לוח חופשות חטיבת דביר אבו'
};

const EMP_COLS = {
  id: 'מספר עובד', first: 'שם פרטי', last: 'שם משפחה', dept: 'מחלקה',
  quota: 'מכסת חופש שנתית', used: 'ניצול חופשה שנתית', credit: 'זיכוי חופשה שנתי',
  balance: 'יתרת חופשה', toUse2026: 'יתרה למימוש 2026'
};
const EMP_NUMS = ['quota', 'used', 'credit', 'balance', 'toUse2026'];
const LEAVE_HEAD = ['מזהה', 'מספר עובד', 'מתאריך', 'עד תאריך', 'חצי יום', 'כבר ירד מהיתרה', 'סטטוס', 'הערה', 'נרשם ע"י', 'עודכן'];
const STATUS_OK = 'מאושר', STATUS_PENDING = 'ממתין לאישור';

/* ---------- web app ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- helpers ---------- */
function ss_() { return SpreadsheetApp.openById(CONFIG.SHEET_ID); }

function me_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  return { email: email, isAdmin: CONFIG.ADMINS.map(function (a) { return a.toLowerCase(); }).indexOf(email) >= 0 };
}

function empSheet_() {
  const sheets = ss_().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    if (sh.getName() === CONFIG.LEAVES_SHEET || sh.getLastRow() < 1) continue;
    const top = sh.getRange(1, 1, Math.min(5, sh.getLastRow()), sh.getLastColumn()).getValues();
    if (top.some(function (r) { return r.indexOf(EMP_COLS.id) >= 0; })) return sh;
  }
  throw new Error('לא נמצאה לשונית עם עמודת "מספר עובד"');
}

function leavesSheet_() {
  const ss = ss_();
  var sh = ss.getSheetByName(CONFIG.LEAVES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.LEAVES_SHEET);
    sh.getRange(1, 1, 1, LEAVE_HEAD.length).setValues([LEAVE_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange('A:D').setNumberFormat('@');   // מזהה, מספר עובד ותאריכים נשמרים כטקסט
    sh.getRange('E2:F1000').insertCheckboxes();
    sh.getRange('G2:G1000').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList([STATUS_OK, STATUS_PENDING], true).build());
  }
  return sh;
}

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);   // 27.09.2026 או 27/09/2026
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function num_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v || '').replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

/* ---------- read ---------- */
function readEmployees_() {
  const values = empSheet_().getDataRange().getValues();
  const hi = values.findIndex(function (r) { return r.indexOf(EMP_COLS.id) >= 0; });
  const head = values[hi], idx = {};
  Object.keys(EMP_COLS).forEach(function (k) {
    idx[k] = head.indexOf(EMP_COLS[k]);
    if (idx[k] < 0) throw new Error('חסרה עמודה בגיליון העובדים: ' + EMP_COLS[k]);
  });
  const out = [];
  values.slice(hi + 1).forEach(function (r) {
    const id = String(r[idx.id] || '').replace(/[^\d]/g, '');
    if (!id) return;
    const e = { id: id };
    Object.keys(EMP_COLS).forEach(function (k) {
      if (k === 'id') return;
      e[k] = EMP_NUMS.indexOf(k) >= 0 ? num_(r[idx[k]]) : String(r[idx[k]] || '').trim();
    });
    out.push(e);
  });
  return out;
}

function readLeaves_() {
  const sh = leavesSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, LEAVE_HEAD.length).getValues()
    .map(function (r) {
      return {
        id: String(r[0] || ''), emp: String(r[1] || '').replace(/[^\d]/g, ''),
        from: normDate_(r[2]), to: normDate_(r[3]),
        half: r[4] === true, taken: r[5] === true,
        status: String(r[6]).trim() === STATUS_OK ? 'approved' : 'pending',
        note: String(r[7] || ''), by: String(r[8] || ''),
        at: r[9] instanceof Date ? r[9].toISOString() : String(r[9] || '')
      };
    })
    .filter(function (l) { return l.id && l.emp && l.from && l.to; });
}

function getData() {
  return { employees: readEmployees_(), leaves: readLeaves_(), me: me_(), reportDate: CONFIG.REPORT_DATE };
}

/* ---------- write ---------- */
function findLeaveRow_(sh, id) {
  if (sh.getLastRow() < 2) return -1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function saveLeave(l) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const user = me_();
    if (!user.email) throw new Error('לא זוהה משתמש. צריך להיכנס עם חשבון הארגון.');
    const sh = leavesSheet_();
    const emps = readEmployees_();
    const emp = String(l.emp || '');
    if (!emps.some(function (e) { return e.id === emp; })) throw new Error('עובד לא נמצא');
    const from = normDate_(l.from), to = normDate_(l.to);
    if (!from || !to || to < from) throw new Error('תאריכים לא תקינים');

    const leaves = readLeaves_();
    const clash = leaves.find(function (x) { return x.emp === emp && x.id !== l.id && x.from <= to && x.to >= from; });
    if (clash) throw new Error('כבר רשומה לעובד חופשה בין ' + clash.from + ' ל-' + clash.to);

    var row = -1, existing = null;
    if (l.id) {
      row = findLeaveRow_(sh, l.id);
      existing = leaves.find(function (x) { return x.id === l.id; }) || null;
      if (row < 0) throw new Error('החופשה כבר לא קיימת. רענן את הדף.');
      if (!user.isAdmin && existing && existing.by && existing.by.toLowerCase() !== user.email)
        throw new Error('אפשר לערוך רק חופשות שהוספת בעצמך');
    }
    const status = user.isAdmin ? (l.status === 'approved' ? STATUS_OK : STATUS_PENDING)
      : (existing && existing.status === 'approved' ? STATUS_OK : STATUS_PENDING);
    const id = l.id || Utilities.getUuid().slice(0, 12);
    const values = [[id, emp, from, to, !!(l.half && from === to), !!l.taken, status,
      String(l.note || '').slice(0, 200), existing && existing.by ? existing.by : user.email, new Date()]];
    if (row < 0) row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, LEAVE_HEAD.length).setValues(values);
    sh.getRange(row, 5, 1, 2).insertCheckboxes().setValues([[values[0][4], values[0][5]]]);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function deleteLeave(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const user = me_();
    const sh = leavesSheet_();
    const row = findLeaveRow_(sh, id);
    if (row < 0) return getData();
    const by = String(sh.getRange(row, 9).getValue() || '').toLowerCase();
    if (!user.isAdmin && by && by !== user.email) throw new Error('אפשר למחוק רק חופשות שהוספת בעצמך');
    sh.deleteRow(row);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

/* ---------- הרצה חד-פעמית: יוצר את לשונית החופשות ומעביר את החופשות שכבר נרשמו ---------- */
function setup() {
  const sh = leavesSheet_();
  if (sh.getLastRow() > 1) return;
  const by = CONFIG.ADMINS[0], now = new Date();
  const rows = [
    ['seed-1718', '1718', '2026-09-27', '2026-09-27', false, false, STATUS_OK, 'חופשה רגילה', by, now],
    ['seed-1931', '1931', '2026-09-02', '2026-09-30', false, true, STATUS_OK, 'חופשה ארה"ב', by, now],
    ['seed-1489', '1489', '2026-09-02', '2026-09-30', false, true, STATUS_OK, 'חופשה ארה"ב', by, now]
  ];
  sh.getRange(2, 1, rows.length, LEAVE_HEAD.length).setValues(rows);
  sh.getRange(2, 5, rows.length, 2).insertCheckboxes().setValues(rows.map(function (r) { return [r[4], r[5]]; }));
}
