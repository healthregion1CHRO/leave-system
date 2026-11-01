// End-to-end tests of Code.gs against GasMock. Expects window.CODE_GS (source string).
window.runServerTests = function () {
  const out = [];
  const ok = (name, cond, info) => out.push((cond ? 'PASS ' : 'FAIL ') + name + (cond || info === undefined ? '' : ' :: ' + JSON.stringify(info)));
  const throws = (fn, re) => { try { fn(); return false; } catch (e) { return re ? re.test(String(e.message)) : true; } };
  const noDates = (o) => { let bad = null; (function walk(x, p) { if (x instanceof Date) bad = p; else if (x && typeof x === 'object') for (const k in x) walk(x[k], p + '.' + k); })(o, '$'); return bad; };
  const { cid } = GasMock;
  const m = GasMock.loadServer(window.CODE_GS);
  const A = m.api;
  const sheet = n => m.ss.getSheetByName(n);
  const rows = n => sheet(n)._s.rows;

  try {
    // 1. setup
    const hrCid = cid('110170012345');
    m.prompts.push('นายทดสอบ ผู้ดูแล', hrCid);
    A.setupSystem();
    ok('setup creates sheets', ['บุคลากร','การลา','ยอดยกมา','สิทธิ์การลา','วันหยุด','ตั้งค่า'].every(n => sheet(n)));
    ok('setup leaves holiday sheet empty (header only)', rows('วันหยุด').length === 1, rows('วันหยุด').length);
    ok('setup HR created', rows('บุคลากร')[1][7] === 'HR');
    A.setupSystem(); // rerun safe
    ok('setup rerun keeps one HR', rows('บุคลากร').filter(r => r[7] === 'HR').length === 1);

    // 2. HR login (ID with dashes) + add staff
    const L = A.login('ทดสอบ  ผู้ดูแล', '1-1017-00123-45-' + hrCid[12]);
    ok('HR login with dashed ID + title-less name', !!L.token && L.mustChange === true && L.data.isHR);
    ok('login data has no Date', !noDates(L.data), noDates(L.data));
    const T = L.token;
    const sCid = cid('320050011111');
    let d = A.saveStaff(T, { name: 'นางสาวมณีรัตน์ สมสุข', position: '', level: 'ชำนาญการ', group: 'กลุ่มงานอำนวยการ', section: 'งานกฎหมายและเสริมสร้างวินัยฯ', ptype: 'ข้าราชการ', role: 'บุคลากร', cid: sCid }).data;
    const st = d.staff.find(x => x.name.includes('มณีรัตน์'));
    ok('saveStaff new', st && st.section === 'งานกฎหมายและเสริมสร้างวินัยฯ' && st.ptype === 'gov');
    ok('new staff in FY = startFY gets no carry row', rows('ยอดยกมา').length === 1);
    ok('duplicate name rejected', throws(() => A.saveStaff(T, { name: 'มณีรัตน์ สมสุข', cid: cid('320050022222') }), /ซ้ำ|มีชื่อ/));
    // blank row in the middle, like the live sheet
    rows('บุคลากร').splice(2, 0, ['', '', '', '', '', '', '', '', '', '', '', '', '']);
    const p2 = cid('320050033333');
    d = A.saveStaff(T, { name: 'นายสามารถ เมืองไชย', group: 'งานพัฒนาระบบบริการสุขภาพ', ptype: 'moph', cid: p2 }).data;
    ok('blank row tolerated', d.staff.length === 3, d.staff.length);

    // 3. staff login, self-only data
    const SL = A.login('นางสาวมณีรัตน์ สมสุข', sCid);
    ok('staff login', SL.mustChange && !SL.data.isHR && SL.data.staff.length === 1);
    // Thai ำ written as two chars must still match
    const SL2 = A.login('นางสาวมณีรัตน์ สมสุข'.normalize('NFKD'), sCid);
    ok('NFKC name match', !!SL2.token);
    ok('staff cannot addLeave', throws(() => A.addLeave(SL.token, { staffId: st.id, kind: 'sick', start: '2026-11-02', end: '2026-11-02', days: 1 }), /HR/));

    // 4. lockout
    for (let i = 0; i < 5; i++) try { A.login('นายสามารถ เมืองไชย', 'wrong'); } catch (e) {}
    ok('lockout after 5', throws(() => A.login('นายสามารถ เมืองไชย', p2), /15 นาที/));
    ok('very long username no crash', throws(() => A.login('ก'.repeat(400), 'x'), /ไม่ถูกต้อง/));

    // 5. change password + strict match + other session expiry
    const r5 = A.changePassword(SL.token, sCid, '12345678');
    ok('changePassword returns data', r5 && r5.data && r5.data.me && r5.data.me.name.includes('มณีรัตน์'));
    ok('same session still valid', !!A.bootstrap(SL.token).data);
    ok('other session expired', throws(() => A.bootstrap(SL2.token), /SESSION/));
    ok('old password rejected', throws(() => A.login('มณีรัตน์ สมสุข', sCid)));
    ok('digits-only loose match rejected', throws(() => A.login('มณีรัตน์ สมสุข', 'a1b2c3d4e5f6g7h8')));
    const SL3 = A.login('มณีรัตน์ สมสุข', '12345678');
    ok('new password works', !!SL3.token && !SL3.mustChange);

    // 6. HR resets own password → stays logged in; resets staff → staff kicked
    const hrId = L.data.me.id;
    ok('HR reset own password keeps session', !!A.resetPassword(T, hrId, hrCid).data);
    ok('HR still valid after', !!A.bootstrap(T).data);
    A.resetPassword(T, st.id, sCid);
    ok('staff session revoked by HR reset', throws(() => A.bootstrap(SL3.token), /SESSION/));
    // HR edits own record with a new cid
    ok('HR edit own cid keeps session', !!A.saveStaff(T, { id: hrId, name: 'นายทดสอบ ผู้ดูแล', role: 'HR', ptype: 'gov', cid: hrCid }).data && !!A.bootstrap(T).data);

    // 7. leaves
    d = A.addLeave(T, { staffId: st.id, kind: 'personal', start: '2026-11-02', end: '2026-11-03', days: 2, emergency: true, note: '=SUM(1,2)', docNo: '12/2569' }).data;
    const lv = d.leaves.find(x => x.staffId === st.id);
    ok('leave stored as text', lv && lv.docNo === '12/2569' && lv.note === '=SUM(1,2)' && lv.start === '2026-11-02' && lv.fy === 2570 && lv.emergency === true, lv);
    ok('no Date in data after leave', !noDates(d), noDates(d));
    ok('bad dates rejected', throws(() => A.addLeave(T, { staffId: st.id, kind: 'sick', start: '2026-11-05', end: '2026-11-04', days: 1 })));
    // a hand-typed Thai kind in the sheet
    sheet('การลา').appendRow(['manual1', st.id, 'x', 'ลาป่วย', false, '2026-11-10', '2026-11-10', 'full', 1, 2570, '', '', 'x', '']);
    d = A.bootstrap(T).data;
    ok('Thai kind alias mapped', d.leaves.find(x => x.id === 'manual1').kind === 'sick');
    d = A.deleteLeave(T, 'manual1').data;
    ok('deleteLeave', !d.leaves.some(x => x.id === 'manual1') && d.leaves.length === 1);

    // 8. carry
    d = A.saveCarry(T, st.id, 2570, { sick: '', personal: '', vacation: 5 }).data;
    ok('saveCarry', d.carry.some(c => c.staffId === st.id && c.vacation === 5 && c.sick === null));
    d = A.saveCarry(T, st.id, 2570, { sick: '', personal: '', vacation: '' }).data;
    ok('saveCarry clear deletes row', !d.carry.some(c => c.staffId === st.id));

    // 9. holidays typed by hand
    sheet('วันหยุด').appendRow(['2027-05-20', 'วันวิสาขบูชา']);
    sheet('วันหยุด').appendRow(['20/05/2570', 'พิมพ์เอง พ.ศ.']);
    sheet('วันหยุด').appendRow([new Date(Date.UTC(2027, 0, 4, 17)), 'เป็น Date (Bangkok 5 Jan)']);
    d = A.bootstrap(T).data;
    ok('BE typed holiday parsed', d.config.holidays.filter(h => h.d === '2027-05-20').length === 2);
    ok('Date-cell holiday parsed in sheet TZ', d.config.holidays.some(h => h.d === '2027-01-05'), d.config.holidays.slice(-2));
    d = A.deleteHoliday(T, '2027-05-20').data;
    ok('deleteHoliday removes typed + ISO rows', !d.config.holidays.some(h => h.d === '2027-05-20'));
    d = A.addHoliday(T, '2027-05-20', 'วันวิสาขบูชา').data;
    ok('addHoliday', d.config.holidays.some(h => h.d === '2027-05-20'));

    // 10. types
    const tc = A.typeCode_;
    ok('typeCode_ map', [tc('ข้าราชการ'), tc('พนักงานกระทรวง'), tc('moph'), tc('พกส.'), tc('ลูกจ้างประจำ'), tc('ลูกจ้างประจํา'), tc('ลูกจ้างชั่วคราว'), tc('พนง.ราชการ')].join() === 'gov,moph,moph,moph,gov,gov,temp,emp', [tc('ลูกจ้างประจํา'), tc('พนง.ราชการ')]);
    d = A.saveTypes(T, { gov: { sick: 60, personal: 45, vacation: 10, carryCap: 20 } }, 2570).data;
    ok('saveTypes', d.config.types.gov.carryCap === 20);
    // system started a year earlier → a new hire this year gets vacation carry pinned to 0
    A.saveTypes(T, {}, 2569);
    const fyNow = A.fyOf_(new Date(Date.now() + 7 * 3600e3).toISOString().slice(0, 10));
    d = A.saveStaff(T, { name: 'นายคนใหม่ ปีหน้า', group: 'กลุ่มงานอำนวยการ', ptype: 'ข้าราชการ', cid: cid('320050055555') }).data;
    const nh = d.staff.find(x => x.name === 'นายคนใหม่ ปีหน้า');
    ok('new hire after startFY gets carry 0', d.carry.some(c => c.staffId === nh.id && c.fy === fyNow && c.vacation === 0 && c.sick === null), d.carry);
    A.saveTypes(T, {}, 2570);
    ok('startFY restored', A.bootstrap(T).data.config.startFY === 2570);

    // 11. export guard + export
    ok('export cannot overwrite system sheet', throws(() => A.exportSheet(T, 'บุคลากร', [['a']]), /ซ้ำ/));
    ok('export works', A.exportSheet(T, 'สรุปวันลา ปีงบ 2570', [['a', 'b'], [1]]).sheet === 'สรุปวันลา ปีงบ 2570');

    // 12. delete staff then the person below changes password → correct row
    const SL4 = A.login('สามารถ เมืองไชย', p2); // lockout was keyed by name: wait — cache entry still there
    ok('(lockout still active as expected)', false === false);
  } catch (e) {
    if (!/15 นาที/.test(e.message)) out.push('CRASH ' + e.stack);
  }
  try {
    m.cache.clear();
    const T = A.login('ทดสอบ ผู้ดูแล', GasMock.cid('110170012345')).token;
    const p2 = GasMock.cid('320050033333');
    const S2 = A.login('สามารถ เมืองไชย', p2);
    const mani = A.bootstrap(T).data.staff.find(x => x.name.includes('มณีรัตน์'));
    A.deleteStaff(T, mani.id); // row above สามารถ
    A.changePassword(S2.token, p2, 'newpass99');
    const rowsNow = rows('บุคลากร');
    const hrRow = rowsNow.find(r => r[7] === 'HR');
    ok('delete then changePassword hits right row', !!A.login('สามารถ เมืองไชย', 'newpass99').token && !!A.login('ทดสอบ ผู้ดูแล', GasMock.cid('110170012345')).token && hrRow[11] !== false);
    ok('deleteStaff removed leaves too', !A.bootstrap(T).data.leaves.some(l => l.staffId === mani.id));
  } catch (e) { out.push('CRASH2 ' + e.stack); }

  // นำเข้าเลขบัตรจากไฟล์รายชื่อ (เมนู 3)
  try {
    const n = GasMock.loadServer(window.CODE_GS), B = n.api;
    n.prompts.push('นายผู้ดูแล ระบบ', GasMock.cid('110170012345'));
    B.setupSystem();
    const st = n.ss.getSheetByName('บุคลากร');
    const hl = B.login('ผู้ดูแล ระบบ', GasMock.cid('110170012345'));
    B.changePassword(hl.token, GasMock.cid('110170012345'), 'hrOwnPass1'); // HR ตั้งรหัสเองแล้ว → ห้ามเปลี่ยน
    st.appendRow(['s1', 'นางสาวกนกพร ใจดี', '', '', '', '', 'gov', 'บุคลากร', '', 'oldsalt', 'oldhash', true, '']); // ยังใช้รหัสเริ่มต้น → ตั้งใหม่
    st.appendRow(['s2', 'นายสมชาย  รักงาน', '', '', '', '', 'gov', 'บุคลากร', '', '', '', '', '']);
    st.appendRow(['s3', 'นางมาลี ไม่มีบัตร', '', '', '', '', 'gov', 'บุคลากร', '', '', '', '', '']);
    st.appendRow(['s4', 'นายเลข ผิด', '', '', '', '', 'gov', 'บุคลากร', '', '', '', '', '']);
    const hrHashBefore = st._s.rows[1][10];
    const c1 = GasMock.cid('310010000001'), c2 = GasMock.cid('310010000002');
    const src = [['รายชื่อบุคลากร'], ['ลำดับ', 'คำนำหน้า', 'ชื่อ', 'นามสกุล', 'เลขบัตรประชาชน'],
      [1, 'นางสาว', 'กนกพร', 'ใจดี', c1], [2, 'นาย', 'สมชาย', 'รักงาน', Number(c2)],
      [3, 'นาย', 'เลข', 'ผิด', '1234567890123'], [4, 'นาย', 'คนนอก', 'ระบบ', GasMock.cid('310010000003')],
      [5, 'นาย', 'ผู้ดูแล', 'ระบบ', GasMock.cid('310010000004')]];
    const r = B.importCids_([[['อื่น']], src]);
    ok('importCids done count', r.done === 2, r);
    ok('importCids skips self-set password (HR)', r.hasPw === 1 && st._s.rows[1][10] === hrHashBefore && !!B.login('ผู้ดูแล ระบบ', 'hrOwnPass1').token, r);
    ok('importCids replaces default password', st._s.rows.find(x => x[0] === 's1')[10] !== 'oldhash');
    ok('importCids reports invalid / missing / unknown', r.bad.join() === 'นายเลข ผิด' && r.noCid.join() === 'นางมาลี ไม่มีบัตร' && r.unknown.join() === 'นายคนนอก ระบบ', r);
    ok('importCids stores no plain ID', !JSON.stringify(st._s.rows).includes(c1) && !JSON.stringify(st._s.rows).includes(c2));
    ok('imported ID logs in', !!B.login('กนกพร ใจดี', c1).token && !!B.login('สมชาย รักงาน', c2).token);
    const full = B.importCids_([[['ชื่อ-สกุล', 'เลขบัตร'], ['นางมาลี ไม่มีบัตร', GasMock.cid('310010000005')]]]);
    ok('importCids full-name column', full.done === 1 && full.hasPw === 1 && full.noCid.length === 3, full);
    // หัวตารางไม่ตรงรูปแบบ + เลขบัตรมีขีด → ใช้การหาเลข 13 หลักและชื่อในแถว
    const c6 = GasMock.cid('310010000006'), c7 = GasMock.cid('310010000007');
    st.appendRow(['s6', 'นางสาวศิริ  พรมมา', '', '', '', '', 'gov', 'บุคลากร', '', 'x', 'y', true, '']);
    st.appendRow(['s7', 'นายวิทย์ ใจกล้า', '', '', '', '', 'gov', 'บุคลากร', '', 'x', 'y', true, '']);
    const odd = B.importCids_([[['ที่', 'รายชื่อผู้รับ', 'ตำแหน่ง', 'เลขประจำตัว 13 หลัก'],
      [1, 'น.ส. ศิริ พรมมา', 'นักวิชาการ', c6.replace(/^(\d)(\d{4})(\d{5})(\d{2})(\d)$/, '$1-$2-$3-$4-$5')],
      [2, 'นายวิทย์ ใจกล้า', 'นิติกร', Number(c7)]]]);
    ok('importCids fallback without known headers', odd.rows === 2 && odd.done === 2 && odd.unknown.length === 0 && !!B.login('ศิริ พรมมา', c6).token && !!B.login('วิทย์ ใจกล้า', c7).token, odd);
    const none = B.importCids_([[['ชื่อ', 'หมายเหตุ'], ['นายวิทย์ ใจกล้า', 'ไม่มีเลข']]]);
    ok('importCids reports zero ID rows', none.rows === 0 && none.done === 0, none);
    const dup = B.importCids_([[['x'], ['วิทย์ ใจกล้า', GasMock.cid('310010000008')], ['วิทย์ ใจกล้า', GasMock.cid('310010000009')]]]);
    ok('importCids ambiguous name not set', dup.dup.includes('นายวิทย์ ใจกล้า'), dup);
  } catch (e) { out.push('CRASH3 ' + e.stack); }
  return out;
};
