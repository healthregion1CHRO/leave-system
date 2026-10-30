/**
 * ระบบวันลาบุคลากร — Google Apps Script (ฝั่งเซิร์ฟเวอร์)
 * ผู้ใช้งาน = ชื่อ-นามสกุล, รหัสผ่านเริ่มต้น = เลขบัตรประชาชน 13 หลัก
 * เลขบัตรและรหัสผ่านเก็บแบบเข้ารหัสทางเดียว (salted SHA-256) ไม่เก็บตัวเลขจริง
 */

var SH = {
  STAFF: 'บุคลากร',
  LEAVES: 'การลา',
  CARRY: 'ยอดยกมา',
  TYPES: 'สิทธิ์การลา',
  HOLIDAYS: 'วันหยุด',
  SETTINGS: 'ตั้งค่า'
};
var STAFF_HEAD = ['id', 'ชื่อ-สกุล', 'ตำแหน่ง', 'ระดับ', 'กลุ่มงาน', 'งาน', 'ประเภทบุคลากร', 'บทบาท',
  'เลขบัตรประชาชน (กรอกชั่วคราว)', 'salt', 'hash', 'ควรเปลี่ยนรหัสผ่าน', 'เข้าใช้ล่าสุด'];
var LEAVE_HEAD = ['id', 'staffId', 'ชื่อ-สกุล', 'ประเภทการลา', 'ฉุกเฉิน', 'ตั้งแต่', 'ถึง', 'ช่วงเวลา',
  'วันทำการ', 'ปีงบประมาณ', 'หมายเหตุ', 'เลขที่ใบลา', 'บันทึกโดย', 'เวลาบันทึก'];
var CARRY_HEAD = ['staffId', 'ชื่อ-สกุล', 'ปีงบประมาณ', 'ลาป่วยยกมา', 'ลากิจยกมา', 'ลาพักผ่อนยกมา'];
var TYPES_HEAD = ['รหัส', 'ประเภทบุคลากร', 'ลาป่วย/ปี', 'ลากิจ/ปี', 'ลาพักผ่อน/ปี', 'ลาพักผ่อนสะสมรวมไม่เกิน'];
var HOL_HEAD = ['วันที่ (ปปปป-ดด-วว ค.ศ.)', 'ชื่อวันหยุด'];

var DEFAULT_TYPES = [
  ['gov', 'ข้าราชการ', 60, 45, 10, 20],
  ['emp', 'พนักงานราชการ', 30, 10, 10, 15],
  ['moph', 'พนักงานกระทรวงสาธารณสุข', 30, 10, 10, 15],
  ['temp', 'ลูกจ้างชั่วคราว', 30, 10, 10, 15]
];
var DEFAULT_HOLIDAYS = [
  ['2026-10-13', 'วันนวมินทรมหาราช'], ['2026-10-23', 'วันปิยมหาราช'],
  ['2026-12-07', 'ชดเชยวันคล้ายวันพระบรมราชสมภพ ร.9 / วันพ่อแห่งชาติ'],
  ['2026-12-10', 'วันรัฐธรรมนูญ'], ['2026-12-31', 'วันสิ้นปี'], ['2027-01-01', 'วันขึ้นปีใหม่'],
  ['2027-02-22', 'ชดเชยวันมาฆบูชา'],
  ['2027-04-06', 'วันจักรี'], ['2027-04-13', 'วันสงกรานต์'], ['2027-04-14', 'วันสงกรานต์'],
  ['2027-04-15', 'วันสงกรานต์'], ['2027-05-04', 'วันฉัตรมงคล'], ['2027-05-20', 'วันวิสาขบูชา'],
  ['2027-06-03', 'วันเฉลิมพระชนมพรรษาสมเด็จพระราชินี'], ['2027-07-19', 'วันเข้าพรรษา'],
  ['2027-07-28', 'วันเฉลิมพระชนมพรรษา ร.10'], ['2027-08-12', 'วันแม่แห่งชาติ']
];
var KIND_CODES = { sick: 1, personal: 1, vacation: 1 };
var KIND_ALIASES = { 'ลาป่วย': 'sick', 'ลากิจ': 'personal', 'ลากิจส่วนตัว': 'personal', 'ลาพักผ่อน': 'vacation' };
var SESSION_SECONDS = 6 * 60 * 60;
var MAX_FAILS = 5;

// ---------------------------------------------------------------- เมนูในชีต
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบวันลา')
    .addItem('1) ตั้งค่าเริ่มต้นระบบ (ทำครั้งแรกครั้งเดียว)', 'setupSystem')
    .addItem('2) เข้ารหัสเลขบัตรที่กรอกในชีตบุคลากร', 'menuHashPending')
    .addSeparator()
    .addItem('ดูลิงก์เข้าใช้งานระบบ', 'menuShowUrl')
    .addToUi();
}

function setupSystem() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SH.STAFF, STAFF_HEAD, [6, 7, 9]);
  ensureSheet_(ss, SH.LEAVES, LEAVE_HEAD, [6, 7, 11, 12]);
  ensureSheet_(ss, SH.CARRY, CARRY_HEAD, []);
  var t = ensureSheet_(ss, SH.TYPES, TYPES_HEAD, []);
  if (t.getLastRow() < 2) t.getRange(2, 1, DEFAULT_TYPES.length, 6).setValues(DEFAULT_TYPES);
  var h = ensureSheet_(ss, SH.HOLIDAYS, HOL_HEAD, [1]);
  if (h.getLastRow() < 2) h.getRange(2, 1, DEFAULT_HOLIDAYS.length, 2).setValues(DEFAULT_HOLIDAYS);
  var s = ensureSheet_(ss, SH.SETTINGS, ['ค่า', 'ข้อมูล'], []);
  if (s.getLastRow() < 2) s.getRange(2, 1, 1, 2).setValues([['startFY', 2570]]);
  pepper_();
  var sheet1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (sheet1 && ss.getSheets().length > 1 && sheet1.getLastRow() === 0) ss.deleteSheet(sheet1);

  if (staffRows_().filter(function (r) { return r.role === 'HR'; }).length === 0) {
    var n = ui.prompt('สร้างบัญชีผู้ดูแล (HR) คนแรก', 'กรอกชื่อ-นามสกุลของคุณ (เช่น นางสาวสุดา ใจงาม)', ui.ButtonSet.OK_CANCEL);
    if (n.getSelectedButton() !== ui.Button.OK) return;
    var c = ui.prompt('สร้างบัญชีผู้ดูแล (HR) คนแรก', 'กรอกเลขบัตรประชาชน 13 หลัก (ใช้เป็นรหัสผ่านเริ่มต้น)', ui.ButtonSet.OK_CANCEL);
    if (c.getSelectedButton() !== ui.Button.OK) return;
    var cid = digits_(c.getResponseText());
    if (!validCid_(cid)) { ui.alert('เลขบัตรประชาชนไม่ถูกต้อง กรุณาเรียกเมนูนี้ใหม่อีกครั้ง'); return; }
    var salt = Utilities.getUuid();
    appendStaff_({ name: n.getResponseText().trim(), position: '', level: '', group: 'กลุ่มงานบริหารทรัพยากรบุคคล', section: '', ptype: 'gov', role: 'HR' }, salt, hash_(salt, cid));
  }
  ui.alert('ตั้งค่าเสร็จแล้ว\n\nขั้นต่อไป: เมนู ส่วนขยาย → Apps Script → ปุ่ม "ทำให้ใช้งานได้" (Deploy) → การทำให้ใช้งานได้รายการใหม่ → เว็บแอป\n• เรียกใช้ในฐานะ: ฉัน\n• ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน');
}

function menuHashPending() {
  var n = withLock_(hashPending_);
  SpreadsheetApp.getUi().alert(n ? ('เข้ารหัสเลขบัตรแล้ว ' + n + ' คน และลบเลขบัตรออกจากชีตแล้ว') : 'ไม่มีเลขบัตรที่รอเข้ารหัส');
}

function menuShowUrl() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? ('ลิงก์ระบบ:\n' + url) : 'ยังไม่ได้ Deploy เป็นเว็บแอป (ดูคู่มือขั้นตอนที่ 4)');
}

// ---------------------------------------------------------------- เว็บแอป
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบวันลาบุคลากร')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------- API: เข้าสู่ระบบ
function login(username, password) {
  var key = 'f_' + normName_(username).slice(0, 200);
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) throw new Error('ใส่รหัสผิดเกิน ' + MAX_FAILS + ' ครั้ง กรุณารอ 15 นาทีแล้วลองใหม่');
  if (staffRows_().some(function (r) { return r.pendingCid || !r.id; })) withLock_(hashPending_);
  var want = normName_(username);
  var pw = String(password || '');
  var cands = staffRows_().filter(function (r) { return normName_(r.name) === want && r.hash; });
  var hit = null;
  cands.forEach(function (r) { if (!hit && pwMatches_(r, pw)) hit = r; });
  if (!hit) {
    cache.put(key, String(fails + 1), 900);
    throw new Error('ชื่อ-นามสกุลหรือรหัสผ่านไม่ถูกต้อง');
  }
  cache.remove(key);
  var token = Utilities.getUuid() + Utilities.getUuid();
  var sess = { id: hit.id, role: hit.role, v: hit.salt };
  cache.put('s_' + token, JSON.stringify(sess), SESSION_SECONDS);
  withLock_(function () { setStaffCell_(findStaff_(hit.id).row, 13, new Date()); });
  return { token: token, mustChange: hit.mustChange, data: buildData_(sess) };
}

// รหัสเริ่มต้น (เลขบัตร) พิมพ์มีขีด/เว้นวรรคได้ แต่หลังเปลี่ยนรหัสแล้วต้องตรงทุกตัวอักษร
function pwMatches_(r, pw) {
  pw = String(pw || '');
  return hash_(r.salt, pw) === r.hash || (r.mustChange && hash_(r.salt, digits_(pw)) === r.hash);
}

function logout(token) { CacheService.getScriptCache().remove('s_' + token); return true; }

function bootstrap(token) { return { data: buildData_(session_(token)) }; }

function changePassword(token, oldPw, newPw) {
  var s = session_(token);
  newPw = String(newPw || '');
  if (newPw.length < 8) throw new Error('รหัสผ่านใหม่ต้องยาวอย่างน้อย 8 ตัวอักษร');
  if (/^\d{13}$/.test(newPw)) throw new Error('รหัสผ่านใหม่ห้ามเป็นตัวเลข 13 หลัก (เลขบัตรประชาชน)');
  var salt = Utilities.getUuid();
  withLock_(function () {
    var r = findStaff_(s.id); // หาแถวภายใน lock เพราะแถวอาจเลื่อนเมื่อมีการลบบุคลากร
    if (r.salt !== s.v) throw new Error('SESSION: รหัสผ่านถูกเปลี่ยนแล้ว กรุณาเข้าสู่ระบบใหม่');
    if (!pwMatches_(r, oldPw)) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    sheet_(SH.STAFF).getRange(r.row, 10, 1, 3).setValues([[salt, hash_(salt, newPw), false]]);
  });
  rebind_(token, s, salt);
  return true;
}

// session ปัจจุบันใช้ต่อได้หลังเปลี่ยนรหัสของตัวเอง ส่วน session อื่นของบัญชีนี้หมดอายุ
function rebind_(token, s, salt) {
  s.v = salt;
  CacheService.getScriptCache().put('s_' + token, JSON.stringify(s), SESSION_SECONDS);
}

// ---------------------------------------------------------------- API: HR
function saveStaff(token, st) {
  var s = requireHR_(token);
  st = st || {};
  var name = String(st.name || '').trim();
  if (!name) throw new Error('กรอกชื่อ-สกุล');
  var ptype = typeCode_(st.ptype);
  var role = st.role === 'HR' ? 'HR' : 'บุคลากร';
  var cid = digits_(st.cid);
  if (cid && !validCid_(cid)) throw new Error('เลขบัตรประชาชนไม่ถูกต้อง ตรวจสอบอีกครั้ง');
  withLock_(function () {
    var rows = staffRows_();
    var dup = rows.filter(function (r) { return normName_(r.name) === normName_(name) && r.id !== st.id; });
    if (dup.length) throw new Error('มีชื่อ "' + name + '" อยู่ในระบบแล้ว ชื่อผู้ใช้ต้องไม่ซ้ำกัน');
    if (st.id) {
      var r = findStaff_(st.id, rows);
      if (r.id === s.id && role !== 'HR') throw new Error('เปลี่ยนบทบาทของตัวเองออกจาก HR ไม่ได้');
      var sh = sheet_(SH.STAFF);
      sh.getRange(r.row, 2, 1, 7).setValues([[name, st.position || '', st.level || '', st.group || '', st.section || '', ptype, role]]);
      if (cid) {
        var salt = Utilities.getUuid();
        sh.getRange(r.row, 10, 1, 3).setValues([[salt, hash_(salt, cid), true]]);
        if (r.id === s.id) rebind_(token, s, salt);
      }
      if (r.name !== name) renameInSheet_(SH.LEAVES, 2, 3, r.id, name), renameInSheet_(SH.CARRY, 1, 2, r.id, name);
    } else {
      if (!cid) throw new Error('กรอกเลขบัตรประชาชนเพื่อใช้เป็นรหัสผ่านเริ่มต้น');
      var salt2 = Utilities.getUuid();
      var nid = appendStaff_({ name: name, position: st.position, level: st.level, group: st.group, section: st.section, ptype: ptype, role: role }, salt2, hash_(salt2, cid));
      // บุคลากรใหม่ไม่มีวันลาพักผ่อนสะสมจากปีก่อน: บันทึกยอดยกมา 0 ไม่ให้ระบบคำนวณย้อนไปถึงปีเริ่มต้น
      var fyNow = fyOf_(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'));
      if (fyNow > (Number(getSetting_('startFY')) || 2570)) sheet_(SH.CARRY).appendRow([nid, name, fyNow, '', '', 0]);
    }
  });
  return { data: buildData_(s) };
}

function deleteStaff(token, id) {
  var s = requireHR_(token);
  if (id === s.id) throw new Error('ลบบัญชีของตัวเองไม่ได้');
  withLock_(function () {
    deleteWhere_(SH.LEAVES, 2, id);
    deleteWhere_(SH.CARRY, 1, id);
    deleteWhere_(SH.STAFF, 1, id);
  });
  return { data: buildData_(s) };
}

function resetPassword(token, id, cid) {
  var s = requireHR_(token);
  cid = digits_(cid);
  if (!validCid_(cid)) throw new Error('เลขบัตรประชาชนไม่ถูกต้อง');
  withLock_(function () {
    var r = findStaff_(id);
    var salt = Utilities.getUuid();
    sheet_(SH.STAFF).getRange(r.row, 10, 1, 3).setValues([[salt, hash_(salt, cid), true]]);
    if (r.id === s.id) rebind_(token, s, salt);
  });
  return { data: buildData_(s) };
}

function addLeave(token, l) {
  var s = requireHR_(token);
  if (!KIND_CODES[l.kind]) throw new Error('ประเภทการลาไม่ถูกต้อง');
  if (!isIso_(l.start) || !isIso_(l.end) || l.end < l.start) throw new Error('วันที่ลาไม่ถูกต้อง');
  var days = Number(l.days);
  if (!(days > 0) || days > 366) throw new Error('จำนวนวันทำการไม่ถูกต้อง');
  withLock_(function () {
    var r = findStaff_(l.staffId);
    var me = findStaff_(s.id);
    var sh = sheet_(SH.LEAVES), row = sh.getLastRow() + 1;
    if (row > sh.getMaxRows()) sh.insertRowAfter(sh.getMaxRows());
    // วันที่ เลขที่ใบลา และหมายเหตุ เก็บเป็นข้อความ (กัน 12/2569 กลายเป็นวันที่ หรือ = กลายเป็นสูตร)
    ['F', 'G', 'K', 'L'].forEach(function (c) { sh.getRange(c + row).setNumberFormat('@'); });
    sh.getRange(row, 1, 1, 14).setValues([[Utilities.getUuid(), r.id, r.name, l.kind, !!l.emergency, l.start, l.end,
      l.start === l.end ? (l.half || 'full') : 'full', days, fyOf_(l.start), String(l.note || ''), String(l.docNo || ''), me.name, new Date()]]);
  });
  return { data: buildData_(s) };
}

function deleteLeave(token, leaveId) {
  var s = requireHR_(token);
  withLock_(function () { deleteWhere_(SH.LEAVES, 1, leaveId); });
  return { data: buildData_(s) };
}

function saveCarry(token, staffId, fy, vals) {
  var s = requireHR_(token);
  fy = Number(fy);
  withLock_(function () {
    var r = findStaff_(staffId);
    var sh = sheet_(SH.CARRY);
    var data = sh.getDataRange().getValues();
    var row = -1;
    for (var i = 1; i < data.length; i++) if (String(data[i][0]) === staffId && Number(data[i][2]) === fy) { row = i + 1; break; }
    var v = function (x) { return (x === '' || x == null) ? '' : Number(x); };
    var out = [staffId, r.name, fy, v(vals.sick), v(vals.personal), v(vals.vacation)];
    var empty = out[3] === '' && out[4] === '' && out[5] === '';
    if (row > 0) { if (empty) sh.deleteRow(row); else sh.getRange(row, 1, 1, 6).setValues([out]); }
    else if (!empty) sh.appendRow(out);
  });
  return { data: buildData_(s) };
}

function saveTypes(token, types, startFY) {
  var s = requireHR_(token);
  withLock_(function () {
    var sh = sheet_(SH.TYPES);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var t = types[String(data[i][0])];
      if (!t) continue;
      sh.getRange(i + 1, 3, 1, 4).setValues([[Number(t.sick) || 0, Number(t.personal) || 0, Number(t.vacation) || 0, Number(t.carryCap) || 0]]);
    }
    if (startFY) setSetting_('startFY', Number(startFY));
  });
  return { data: buildData_(s) };
}

function addHoliday(token, d, n) {
  var s = requireHR_(token);
  if (!isIso_(d) || !String(n || '').trim()) throw new Error('ระบุวันที่และชื่อวันหยุด');
  withLock_(function () { deleteWhere_(SH.HOLIDAYS, 1, d); sheet_(SH.HOLIDAYS).appendRow([d, String(n).trim()]); });
  return { data: buildData_(s) };
}

function deleteHoliday(token, d) {
  var s = requireHR_(token);
  withLock_(function () { deleteWhere_(SH.HOLIDAYS, 1, d); });
  return { data: buildData_(s) };
}

function exportSheet(token, title, rows) {
  var s = requireHR_(token);
  var name = String(title || 'สรุปวันลา').replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 90);
  if (Object.keys(SH).some(function (k) { return SH[k] === name.trim(); })) throw new Error('ชื่อชีตซ้ำกับชีตของระบบ');
  if (!rows || !rows.length) throw new Error('ไม่มีข้อมูลให้ส่งออก');
  var width = rows[0].length;
  rows = rows.map(function (r) { var x = r.slice(0, width); while (x.length < width) x.push(''); return x; });
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, rows.length, width).setValues(rows);
  sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#d3f2e8');
  sh.setFrozenRows(1);
  return { sheet: name, data: buildData_(s) };
}

// ---------------------------------------------------------------- สร้างข้อมูลส่งให้หน้าเว็บ
function buildData_(s) {
  var staff = staffRows_();
  var me = null;
  staff.forEach(function (r) { if (r.id === s.id) me = r; });
  if (!me) throw new Error('SESSION: ไม่พบบัญชีของคุณ กรุณาเข้าสู่ระบบใหม่');
  if (me.salt !== s.v) throw new Error('SESSION: รหัสผ่านถูกเปลี่ยน กรุณาเข้าสู่ระบบใหม่');
  var isHR = me.role === 'HR';
  var pub = function (r) {
    var o = { id: r.id, name: r.name, position: r.position, level: r.level, group: r.group, section: r.section, ptype: r.ptype, role: r.role };
    if (isHR) { o.hasPassword = !!r.hash; o.mustChange = r.mustChange; o.lastLogin = r.lastLogin; }
    return o;
  };
  var leaves = rowsOf_(SH.LEAVES).map(function (x) {
    var k = String(x[3]).trim();
    return { id: String(x[0]), staffId: String(x[1]), kind: KIND_ALIASES[k] || k, emergency: x[4] === true || String(x[4]).toUpperCase() === 'TRUE',
      start: isoCell_(x[5]), end: isoCell_(x[6]), half: String(x[7] || 'full'), days: Number(x[8]) || 0, fy: Number(x[9]) || fyOf_(isoCell_(x[5])),
      note: String(x[10] || ''), docNo: String(x[11] || ''), recordedBy: String(x[12] || '') };
  });
  var carry = rowsOf_(SH.CARRY).map(function (x) {
    return { staffId: String(x[0]), fy: Number(x[2]), sick: x[3] === '' ? null : Number(x[3]), personal: x[4] === '' ? null : Number(x[4]), vacation: x[5] === '' ? null : Number(x[5]) };
  });
  var types = {};
  rowsOf_(SH.TYPES).forEach(function (x) {
    types[String(x[0])] = { label: String(x[1]), sick: Number(x[2]) || 0, personal: Number(x[3]) || 0, vacation: Number(x[4]) || 0, carryCap: Number(x[5]) || 0 };
  });
  var holidays = rowsOf_(SH.HOLIDAYS).map(function (x) { return { d: isoCell_(x[0]), n: String(x[1] || '') }; })
    .filter(function (h) { return h.d; });
  if (!isHR) {
    staff = [me];
    leaves = leaves.filter(function (l) { return l.staffId === me.id; });
    carry = carry.filter(function (c) { return c.staffId === me.id; });
  }
  return { me: pub(me), isHR: isHR, staff: staff.map(pub), leaves: leaves, carry: carry,
    config: { types: types, holidays: holidays, startFY: Number(getSetting_('startFY')) || 2570 } };
}

// ---------------------------------------------------------------- ตัวช่วย
function session_(token) {
  var raw = token && CacheService.getScriptCache().get('s_' + token);
  if (!raw) throw new Error('SESSION: หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่');
  CacheService.getScriptCache().put('s_' + token, raw, SESSION_SECONDS);
  return JSON.parse(raw);
}
function requireHR_(token) {
  var s = session_(token);
  var r = findStaff_(s.id);
  if (r.salt !== s.v) throw new Error('SESSION: รหัสผ่านถูกเปลี่ยน กรุณาเข้าสู่ระบบใหม่');
  if (r.role !== 'HR') throw new Error('เฉพาะเจ้าหน้าที่ HR เท่านั้น');
  return s;
}
function sheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('ยังไม่ได้ตั้งค่าระบบ: เปิดชีตแล้วเลือกเมนู ระบบวันลา → ตั้งค่าเริ่มต้นระบบ');
  return sh;
}
function rowsOf_(name) {
  var v = sheet_(name).getDataRange().getValues();
  return v.slice(1).filter(function (r) { return String(r[0]).trim() !== ''; });
}
function staffRows_() {
  var v = sheet_(SH.STAFF).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var x = v[i];
    if (!String(x[1]).trim()) continue;
    out.push({ row: i + 1, id: String(x[0]), name: String(x[1]).trim(), position: String(x[2] || ''), level: String(x[3] || ''),
      group: String(x[4] || ''), section: String(x[5] || ''), ptype: typeCode_(x[6]), role: String(x[7]).trim().toUpperCase() === 'HR' ? 'HR' : 'บุคลากร',
      pendingCid: digits_(x[8]), salt: String(x[9] || ''), hash: String(x[10] || ''),
      mustChange: x[11] === true || String(x[11]).toUpperCase() === 'TRUE', lastLogin: x[12] instanceof Date ? x[12].toISOString() : String(x[12] || '') });
  }
  return out;
}
function findStaff_(id, rows) {
  rows = rows || staffRows_();
  for (var i = 0; i < rows.length; i++) if (rows[i].id === id) return rows[i];
  throw new Error('ไม่พบบุคลากรในระบบ');
}
function appendStaff_(st, salt, hash) {
  var sh = sheet_(SH.STAFF), id = 's' + Utilities.getUuid().slice(0, 8);
  sh.appendRow([id, st.name, st.position || '', st.level || '', st.group || '', st.section || '',
    typeCode_(st.ptype), st.role === 'HR' ? 'HR' : 'บุคลากร', '', salt, hash, true, '']);
  return id;
}
function setStaffCell_(row, col, val) { sheet_(SH.STAFF).getRange(row, col).setValue(val); }
function hashPending_() {
  var sh = sheet_(SH.STAFF);
  var v = sh.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < v.length; i++) {
    if (!String(v[i][1]).trim()) continue;
    if (!String(v[i][0]).trim()) { sh.getRange(i + 1, 1).setValue('s' + Utilities.getUuid().slice(0, 8)); }
    if (!String(v[i][7]).trim()) sh.getRange(i + 1, 8).setValue('บุคลากร');
    var cid = digits_(v[i][8]);
    if (!cid) continue;
    if (!validCid_(cid)) { sh.getRange(i + 1, 9).setNote('เลขบัตรไม่ถูกต้อง ตรวจสอบอีกครั้ง'); continue; }
    var salt = Utilities.getUuid();
    sh.getRange(i + 1, 9, 1, 4).setValues([['', salt, hash_(salt, cid), true]]);
    sh.getRange(i + 1, 9).clearNote();
    n++;
  }
  return n;
}
function deleteWhere_(name, col, val) {
  var sh = sheet_(name);
  var v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(isoCellOrRaw_(v[i][col - 1])) === String(val)) sh.deleteRow(i + 1);
}
function renameInSheet_(name, idCol, nameCol, id, newName) {
  var sh = sheet_(name);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][idCol - 1]) === id) sh.getRange(i + 1, nameCol).setValue(newName);
}
function ensureSheet_(ss, name, head, textCols) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#d3f2e8');
    sh.setFrozenRows(1);
  }
  (textCols || []).forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  return sh;
}
function getSetting_(k) {
  var v = rowsOf_(SH.SETTINGS);
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === k) return v[i][1];
  return null;
}
function setSetting_(k, val) {
  var sh = sheet_(SH.SETTINGS);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === k) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([k, val]);
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('ระบบกำลังบันทึกข้อมูลอื่นอยู่ กรุณาลองใหม่');
  try { return fn(); } finally { lock.releaseLock(); }
}
function pepper_() {
  var p = PropertiesService.getScriptProperties();
  var v = p.getProperty('PEPPER');
  if (!v) { v = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('PEPPER', v); }
  return v;
}
function hash_(salt, secret) {
  var h = salt + '|' + String(secret) + '|' + pepper_();
  for (var i = 0; i < 200; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8)
      .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  }
  return h;
}
function digits_(s) { return String(s == null ? '' : s).replace(/\D/g, ''); }
function validCid_(id) {
  if (!/^\d{13}$/.test(id)) return false;
  var sum = 0;
  for (var i = 0; i < 12; i++) sum += Number(id.charAt(i)) * (13 - i);
  return (11 - (sum % 11)) % 10 === Number(id.charAt(12));
}
function normName_(n) {
  // NFKC ให้ "ำ" กับ "ํา" ตรงกัน และตัดอักขระล่องหน (zero-width) ที่ติดมากับการคัดลอก
  return String(n || '').normalize('NFKC').replace(/[​-‍﻿]/g, '').replace(/\s+/g, ' ').trim()
    .replace(/^(นางสาว|นาย|นาง|น\.ส\.|ด\.ช\.|ด\.ญ\.|ดร\.|ว่าที่\s?ร\.?ต\.?(หญิง)?)\s*/, '')
    .replace(/\s+/g, '').toLowerCase();
}
function typeCode_(v) {
  var s = String(v || '').replace(/ํา/g, 'ำ').trim(); // "ํา" สองตัว → "ำ"
  var map = { 'ข้าราชการ': 'gov', 'พนักงานราชการ': 'emp', 'พนักงานกระทรวงสาธารณสุข': 'moph', 'พกส.': 'moph', 'ลูกจ้างชั่วคราว': 'temp' };
  if (map[s]) return map[s];
  if (['gov', 'emp', 'moph', 'temp'].indexOf(s) >= 0) return s;
  if (/^พนักงานกระทรวง|^พกส/.test(s)) return 'moph';
  if (/^พนักงานราชการ|^พนง\.?\s?ราชการ/.test(s)) return 'emp';
  if (/^ลูกจ้างประจำ/.test(s)) return 'gov'; // สิทธิ์ลาตามระเบียบกระทรวงการคลังใกล้เคียงข้าราชการ
  if (/^ลูกจ้าง/.test(s)) return 'temp';
  return 'gov';
}
function isIso_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }
function isoCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  var s = String(v || '').trim();
  return isIso_(s) ? s : '';
}
var TZ_;
function tz_() { return TZ_ || (TZ_ = SpreadsheetApp.getActive().getSpreadsheetTimeZone()); }
function isoCellOrRaw_(v) { return v instanceof Date ? isoCell_(v) : v; }
function fyOf_(iso) { var y = Number(iso.slice(0, 4)), m = Number(iso.slice(5, 7)); return (m >= 10 ? y + 1 : y) + 543; }
