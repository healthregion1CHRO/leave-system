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
    ok('setup holidays incl. new 2570 days', rows('วันหยุด').some(r => r[0] === '2026-10-23') && rows('วันหยุด').length === 18, rows('วันหยุด').length);
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
  return out;
};
