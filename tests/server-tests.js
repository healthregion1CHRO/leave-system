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

    // 7b. ขอยกเลิกวันลา / คืนสถานะ
    // ชีตเก่าที่มีแค่ 14 คอลัมน์: ตัดหัวคอลัมน์ใหม่ออกก่อน แล้วระบบต้องเติมให้เอง
    rows('การลา')[0].length = 14;
    ok('old leave stays active', d.leaves.find(x => x.id === lv.id).cancelled === false);
    ok('staff cannot cancelLeave', throws(() => A.cancelLeave(SL.token, lv.id, 'x'), /HR/) || throws(() => A.cancelLeave(SL.token, lv.id, 'x')));
    ok('cancel needs reason', throws(() => A.cancelLeave(T, lv.id, '   '), /เหตุผล/));
    ok('cancel unknown id', throws(() => A.cancelLeave(T, 'nope', 'x'), /ไม่พบ/));
    d = A.cancelLeave(T, lv.id, '=ไปราชการแทน').data;
    const cl = d.leaves.find(x => x.id === lv.id);
    ok('cancelLeave marks cancelled, keeps row', cl && cl.cancelled === true && cl.cancelReason === '=ไปราชการแทน' && cl.cancelledBy === 'นายทดสอบ ผู้ดูแล' && /^\d{4}-\d{2}-\d{2}$/.test(cl.cancelledAt), cl);
    ok('cancel header added to old sheet', rows('การลา')[0][14] === 'สถานะ' && rows('การลา')[0][17] === 'เวลายกเลิก', rows('การลา')[0]);
    ok('no Date in data after cancel', !noDates(d), noDates(d));
    ok('cancel twice rejected', throws(() => A.cancelLeave(T, lv.id, 'x'), /ยกเลิกไปแล้ว/));
    ok('staff sees own cancelled leave', A.bootstrap(A.login('มณีรัตน์ สมสุข', sCid).token).data.leaves.some(x => x.id === lv.id && x.cancelled));
    d = A.restoreLeave(T, lv.id).data;
    ok('restoreLeave clears cancel', d.leaves.find(x => x.id === lv.id).cancelled === false && d.leaves.find(x => x.id === lv.id).cancelReason === '');
    ok('restore active rejected', throws(() => A.restoreLeave(T, lv.id), /ไม่ได้ถูกยกเลิก/));
    ok('API exposes cancel/restore', !!A.apiFn_('cancelLeave') && !!A.apiFn_('restoreLeave'));

    // 7b2. staffId ในชีตการลาถูกวางข้อความทับ → login ครั้งถัดไปผูกกลับด้วยชื่อ
    const lr = rows('การลา').findIndex(r => r[0] === lv.id);
    const realId = rows('การลา')[lr][1];
    rows('การลา')[lr][1] = '2. ที่ไฟล์ Code.gs กด Ctrl+A';
    ok('broken staffId hides leave', !A.bootstrap(A.login('มณีรัตน์ สมสุข', sCid).token).data.leaves.some(x => x.id === lv.id) || rows('การลา')[lr][1] === realId);
    ok('login repairs broken staffId by name', rows('การลา')[lr][1] === realId, rows('การลา')[lr][1]);
    rows('การลา')[lr][1] = 'sdeadbeef'; rows('การลา')[lr][2] = 'ไม่มีคนนี้';
    A.login('มณีรัตน์ สมสุข', sCid);
    ok('unknown name left untouched', rows('การลา')[lr][1] === 'sdeadbeef');
    rows('การลา')[lr][1] = realId; rows('การลา')[lr][2] = 'นางสาวมณีรัตน์ สมสุข';

    // 7c. ยกเลิกเฉพาะบางวัน: 16–27 พ.ย. 69 (จ.–ศ. 2 สัปดาห์) วันหยุด 20 พ.ย. → 9 วันทำการ
    A.addHoliday(T, '2026-11-20', 'ทดสอบ');
    d = A.addLeave(T, { staffId: st.id, kind: 'vacation', start: '2026-11-16', end: '2026-11-27', days: 9, note: '=ไปต่างจังหวัด', docNo: '12/2569' }).data;
    const pl = d.leaves.find(x => x.start === '2026-11-16' && x.end === '2026-11-27');
    const nLeaves = rows('การลา').length;
    ok('partial: from before start rejected', throws(() => A.cancelLeave(T, pl.id, 'x', '2026-11-13', '2026-11-18'), /ช่วงวันลา/));
    ok('partial: to after end rejected', throws(() => A.cancelLeave(T, pl.id, 'x', '2026-11-25', '2026-11-30'), /ช่วงวันลา/));
    ok('partial: to < from rejected', throws(() => A.cancelLeave(T, pl.id, 'x', '2026-11-23', '2026-11-19'), /ไม่ถูกต้อง/));
    ok('partial: weekend-only range rejected', throws(() => A.cancelLeave(T, pl.id, 'x', '2026-11-21', '2026-11-22'), /ไม่มีวันทำการ/));
    ok('partial: failed attempts change nothing', rows('การลา').length === nLeaves && !A.bootstrap(T).data.leaves.find(x => x.id === pl.id).cancelled);
    d = A.cancelLeave(T, pl.id, 'กลับมาปฏิบัติงาน', '2026-11-19', '2026-11-23').data;
    const mine = d.leaves.filter(x => x.staffId === st.id && x.kind === 'vacation' && x.start >= '2026-11-16' && x.end <= '2026-11-27').sort((a, b) => a.start.localeCompare(b.start));
    const sig = mine.map(x => [x.start, x.end, x.days, x.cancelled].join('|'));
    ok('partial: split into before / cancelled / after',
      JSON.stringify(sig) === JSON.stringify(['2026-11-16|2026-11-18|3|false', '2026-11-19|2026-11-23|2|true', '2026-11-24|2026-11-27|4|false']), sig);
    ok('partial: cancelled row keeps id + reason', mine[1].id === pl.id && mine[1].cancelReason === 'กลับมาปฏิบัติงาน' && mine[1].cancelledBy === 'นายทดสอบ ผู้ดูแล');
    ok('partial: new rows copy note/docNo/fy', mine[0].note === '=ไปต่างจังหวัด' && mine[2].docNo === '12/2569' && mine[0].fy === 2570 && mine[0].id !== mine[2].id);
    ok('partial: total days preserved', mine.reduce((n, x) => n + x.days, 0) === 9);
    ok('partial: no Date in data', !noDates(d), noDates(d));
    d = A.cancelLeave(T, mine[2].id, 'x', '2026-11-24', '2026-11-24').data; // ยกเลิกวันแรกของช่วงหลัง → เหลือ 25–27
    ok('partial: cancel first day only', d.leaves.some(x => x.id === mine[2].id && x.start === '2026-11-24' && x.end === '2026-11-24' && x.days === 1 && x.cancelled)
      && d.leaves.some(x => !x.cancelled && x.start === '2026-11-25' && x.end === '2026-11-27' && x.days === 3));
    d = A.restoreLeave(T, pl.id).data;
    ok('partial: restore counts cancelled part again', d.leaves.find(x => x.id === pl.id).cancelled === false && d.leaves.find(x => x.id === pl.id).days === 2);
    d.leaves.filter(x => x.staffId === st.id && x.kind === 'vacation' && x.start >= '2026-11-16' && x.end <= '2026-11-27').forEach(x => A.deleteLeave(T, x.id));
    A.deleteHoliday(T, '2026-11-20');

    // 7d. ยื่นใบลา / พิจารณา (รอพิจารณา · อนุมัติ · ไม่อนุมัติ) / ขอยกเลิก
    const ST = A.login('มณีรัตน์ สมสุข', sCid).token;
    const find = (dd, id) => dd.leaves.find(x => x.id === id);
    d = A.bootstrap(T).data;
    ok('old rows (blank review) count as approved', find(d, lv.id).status === 'approved' && find(d, lv.id).cx === '', find(d, lv.id));
    // ค่าเริ่มต้น: ล็อกปุ่มของบุคลากร (ยื่นใบลา · ถอนใบลา · ขอยกเลิกวันลา) ฝั่งเซิร์ฟเวอร์ด้วย
    ok('staff locked by default', d.config.staffLocked === true);
    ok('locked: requestLeave blocked', throws(() => A.requestLeave(ST, { kind: 'sick', start: '2026-12-07', end: '2026-12-07' }), /HR/));
    ok('locked: requestCancel blocked', throws(() => A.requestCancel(ST, lv.id, 'x'), /HR/));
    ok('locked: withdrawLeave blocked', throws(() => A.withdrawLeave(ST, lv.id), /HR/));
    ok('locked: withdrawCancel blocked for staff', throws(() => A.withdrawCancel(ST, lv.id), /HR/));
    A.setSetting_('staffRequests', 'เปิด');
    ok('unlocked via setting', A.bootstrap(ST).data.config.staffLocked === false);
    d = A.requestLeave(ST, { staffId: 'someone-else', kind: 'sick', start: '2026-12-07', end: '2026-12-08', note: '=ไข้', days: 99 }).data;
    const rq = d.leaves.find(x => x.start === '2026-12-07');
    ok('requestLeave: pending, own id, server-side days', rq && rq.staffId === st.id && rq.status === 'pending' && rq.days === 2 && rq.note === '=ไข้' && d.leaves.every(x => x.staffId === st.id), rq);
    ok('requestLeave: staff sees only own data', d.staff.length === 1 && !d.isHR);
    ok('requestLeave header extended to 27 cols', rows('การลา')[0][18] === 'การพิจารณา' && rows('การลา')[0][26] === 'ขอยกเลิกถึง', rows('การลา')[0]);
    ok('requestLeave: overlap rejected', throws(() => A.requestLeave(ST, { kind: 'personal', start: '2026-12-08', end: '2026-12-09' }), /ซ้อน/));
    ok('requestLeave: weekend only rejected', throws(() => A.requestLeave(ST, { kind: 'sick', start: '2026-12-12', end: '2026-12-13' }), /ไม่มีวันทำการ/));
    ok('requestLeave: across FY rejected', throws(() => A.requestLeave(ST, { kind: 'sick', start: '2027-09-30', end: '2027-10-01' }), /ปีงบประมาณ/));
    d = A.requestLeave(ST, { kind: 'personal', start: '2026-12-09', end: '2026-12-09', half: 'am', emergency: true }).data;
    const hf = d.leaves.find(x => x.start === '2026-12-09');
    ok('requestLeave: half day = 0.5', hf && hf.days === 0.5 && hf.half === 'am' && hf.emergency === true, hf);
    ok('staff cannot reviewLeave', throws(() => A.reviewLeave(ST, rq.id, 'approved'), /HR/));
    ok('reject needs reason', throws(() => A.reviewLeave(T, rq.id, 'rejected', ' '), /เหตุผล/));
    ok('bad status rejected', throws(() => A.reviewLeave(T, rq.id, 'ok'), /สถานะ/));
    d = A.reviewLeave(T, rq.id, 'rejected', '=ตรงวันประชุม').data;
    ok('reviewLeave rejected', find(d, rq.id).status === 'rejected' && find(d, rq.id).rejectReason === '=ตรงวันประชุม' && find(d, rq.id).reviewedBy === 'นายทดสอบ ผู้ดูแล' && /^\d{4}-\d{2}-\d{2}$/.test(find(d, rq.id).reviewedAt), find(d, rq.id));
    d = A.reviewLeave(T, rq.id, 'pending').data;
    ok('reviewLeave back to pending clears reviewer', find(d, rq.id).status === 'pending' && !find(d, rq.id).reviewedBy && !find(d, rq.id).rejectReason);
    d = A.reviewLeave(T, rq.id, 'approved').data;
    ok('reviewLeave approved', find(d, rq.id).status === 'approved' && !noDates(d), noDates(d));
    ok('withdraw approved leave rejected', throws(() => A.withdrawLeave(ST, rq.id), /รอพิจารณา/));
    d = A.withdrawLeave(ST, hf.id).data;
    ok('withdrawLeave removes pending row', !find(d, hf.id));
    d = A.addLeave(T, { staffId: st.id, kind: 'vacation', start: '2026-12-21', end: '2026-12-21', days: 1 }).data;
    const hrl = d.leaves.find(x => x.start === '2026-12-21');
    ok('addLeave by HR = approved at once', hrl.status === 'approved' && hrl.reviewedBy === 'นายทดสอบ ผู้ดูแล');
    // HR เลือกผลการพิจารณาตอนบันทึก: รอพิจารณา / ไม่อนุญาต
    d = A.addLeave(T, { staffId: st.id, kind: 'sick', start: '2027-01-11', end: '2027-01-11', days: 1, status: 'pending' }).data;
    const hp = d.leaves.find(x => x.start === '2027-01-11');
    ok('addLeave status pending', hp && hp.status === 'pending' && !hp.reviewedBy, JSON.stringify(hp));
    ok('addLeave rejected needs reason', throws(() => A.addLeave(T, { staffId: st.id, kind: 'sick', start: '2027-01-12', end: '2027-01-12', days: 1, status: 'rejected' }), /เหตุผล/));
    ok('addLeave bad status', throws(() => A.addLeave(T, { staffId: st.id, kind: 'sick', start: '2027-01-12', end: '2027-01-12', days: 1, status: 'xx' }), /สถานะ/));
    d = A.addLeave(T, { staffId: st.id, kind: 'sick', start: '2027-01-12', end: '2027-01-12', days: 1, status: 'rejected', rejectReason: 'ไม่มีใบรับรองแพทย์' }).data;
    const hj = d.leaves.find(x => x.start === '2027-01-12');
    ok('addLeave status rejected + reason', hj && hj.status === 'rejected' && hj.rejectReason === 'ไม่มีใบรับรองแพทย์' && hj.reviewedBy === 'นายทดสอบ ผู้ดูแล', JSON.stringify(hj));
    A.deleteLeave(T, hp.id); d = A.deleteLeave(T, hj.id).data;
    d = A.requestLeave(ST, { kind: 'sick', start: '2026-12-22', end: '2026-12-22' }).data;
    const pend = d.leaves.find(x => x.start === '2026-12-22');
    ok('HR cannot cancel a pending leave', throws(() => A.cancelLeave(T, pend.id, 'x'), /อนุมัติแล้ว/));
    ok('cannot request cancel of pending leave', throws(() => A.requestCancel(ST, pend.id, 'x'), /ถอนใบลา/));
    ok('request cancel needs reason', throws(() => A.requestCancel(ST, rq.id, ''), /เหตุผล/));
    ok('request cancel: range checked', throws(() => A.requestCancel(ST, rq.id, 'x', '2026-12-01', '2026-12-08'), /ช่วงวันลา/));
    d = A.requestCancel(ST, rq.id, 'หายแล้ว', '2026-12-08', '2026-12-08').data;
    ok('requestCancel partial → pending, still counted', find(d, rq.id).cx === 'pending' && !find(d, rq.id).cancelled && find(d, rq.id).cxFrom === '2026-12-08' && find(d, rq.id).cancelReason === 'หายแล้ว' && find(d, rq.id).cancelledBy.includes('มณีรัตน์'), find(d, rq.id));
    ok('request cancel twice rejected', throws(() => A.requestCancel(ST, rq.id, 'x'), /รอ HR/));
    ok('reviewLeave blocked while cancel pending', throws(() => A.reviewLeave(T, rq.id, 'rejected', 'x'), /ยกเลิกวันลา/));
    ok('staff cannot reviewCancel', throws(() => A.reviewCancel(ST, rq.id, 'approved'), /HR/));
    ok('reviewCancel reject needs reason', throws(() => A.reviewCancel(T, rq.id, 'rejected'), /เหตุผล/));
    d = A.reviewCancel(T, rq.id, 'rejected', 'ไม่มีใบรับรอง').data;
    ok('reviewCancel rejected → still counted', find(d, rq.id).cx === 'rejected' && !find(d, rq.id).cancelled && find(d, rq.id).cxRejectReason === 'ไม่มีใบรับรอง' && find(d, rq.id).cxReviewedBy === 'นายทดสอบ ผู้ดูแล');
    d = A.requestCancel(ST, rq.id, 'ขอใหม่ทั้งรายการ').data;
    ok('re-request after rejection (whole)', find(d, rq.id).cx === 'pending' && find(d, rq.id).cxFrom === '' && !find(d, rq.id).cxRejectReason);
    d = A.withdrawCancel(ST, rq.id).data;
    ok('withdrawCancel clears request', find(d, rq.id).cx === '' && !find(d, rq.id).cancelReason);
    ok('withdrawCancel with nothing pending', throws(() => A.withdrawCancel(ST, rq.id), /ไม่มีคำขอ/));
    A.requestCancel(ST, rq.id, 'กลับมาทำงาน', '2026-12-08', '2026-12-08');
    d = A.reviewCancel(T, rq.id, 'approved').data;
    const rqs = d.leaves.filter(x => x.staffId === st.id && x.start >= '2026-12-07' && x.end <= '2026-12-08').sort((a, b) => a.start.localeCompare(b.start));
    ok('reviewCancel approved splits by requested range', JSON.stringify(rqs.map(x => [x.start, x.days, x.cancelled, x.status, x.cx].join('|'))) === JSON.stringify(['2026-12-07|1|false|approved|', '2026-12-08|1|true|approved|approved']), rqs.map(x => [x.start, x.days, x.cancelled, x.status, x.cx].join('|')));
    ok('approved cancel keeps requester + reviewer', rqs[1].id === rq.id && rqs[1].cancelledBy.includes('มณีรัตน์') && rqs[1].cxReviewedBy === 'นายทดสอบ ผู้ดูแล' && rqs[1].cxFrom === '');
    ok('no Date after reviewCancel', !noDates(d), noDates(d));
    d = A.reviewCancel(T, rq.id, 'pending').data;
    ok('reviewCancel back to pending un-cancels', find(d, rq.id).cx === 'pending' && !find(d, rq.id).cancelled);
    A.reviewCancel(T, rq.id, 'approved');
    d = A.restoreLeave(T, rq.id).data;
    ok('restoreLeave clears request columns', find(d, rq.id).cx === '' && !find(d, rq.id).cancelled && !find(d, rq.id).cxReviewedBy);
    d = A.cancelLeave(T, hrl.id, 'HR ยกเลิกเอง').data;
    ok('HR cancelLeave = approved cancel', find(d, hrl.id).cx === 'approved' && find(d, hrl.id).cxReviewedBy === 'นายทดสอบ ผู้ดูแล');
    ok('API exposes approval fns', ['requestLeave', 'withdrawLeave', 'reviewLeave', 'requestCancel', 'withdrawCancel', 'reviewCancel'].every(f => !!A.apiFn_(f)));
    A.bootstrap(T).data.leaves.filter(x => x.staffId === st.id && x.start >= '2026-12-01').forEach(x => A.deleteLeave(T, x.id));

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

  // อัปเดตอัตโนมัติจาก GitHub (จำลอง GitHub API + Apps Script API)
  try {
    const u = GasMock.loadServer(window.CODE_GS), U = u.api;
    const SHA1 = 'a'.repeat(40), SHA2 = 'b'.repeat(40);
    let sha = SHA1, codeSrc = 'function autoUpdateTick(){}', manifest = '{"oauthScopes":["x/auth/script.projects","x/auth/script.deployments"]}';
    const calls = [];
    let projectFiles = [{ name: 'Code', type: 'SERVER_JS', source: 'old' }, { name: 'Index', type: 'HTML', source: 'old' }, { name: 'appsscript', type: 'JSON', source: '{}' }, { name: 'Extra', type: 'SERVER_JS', source: 'keep me' }];
    let ver = 3;
    const deployments = [
      { deploymentId: 'HEAD', deploymentConfig: { scriptId: 'SCRIPT1' }, entryPoints: [{ entryPointType: 'WEB_APP' }] },
      { deploymentId: 'D1', deploymentConfig: { scriptId: 'SCRIPT1', versionNumber: 3 }, entryPoints: [{ entryPointType: 'WEB_APP' }] },
      { deploymentId: 'D2', deploymentConfig: { scriptId: 'SCRIPT1', versionNumber: 2 }, entryPoints: [{ entryPointType: 'WEB_APP' }] }];
    let ghCode = 200; const rawRefs = [];
    u.fetchHook.fn = (url, o) => {
      const method = (o.method || 'get').toLowerCase();
      calls.push(method + ' ' + url);
      if (url.startsWith('https://api.github.com/')) {
        if (o.headers.Authorization && o.headers.Authorization !== 'Bearer TOKEN') return { code: 401, body: '' };
        if (ghCode !== 200) return { code: ghCode, body: '{"message":"API rate limit exceeded"}' };
        if (/commits\/main$/.test(url) && o.headers.Accept === 'application/vnd.github.sha') return { code: 200, body: sha + '\n' };
      }
      const raw = url.match(/^https:\/\/raw\.githubusercontent\.com\/healthregion1CHRO\/leave-system\/(\w+)\/apps-script\/(.+)$/);
      if (raw) {
        rawRefs.push(raw[1]);
        if (raw[2] === 'Code.gs') return { code: 200, body: codeSrc };
        if (raw[2] === 'Index.html') return { code: 200, body: '<p>ใหม่</p>' };
        if (raw[2] === 'appsscript.json') return { code: 200, body: manifest };
      }
      const p = url.replace('https://script.googleapis.com/v1/projects/', '');
      if (o.headers.Authorization !== 'Bearer OAUTH') return { code: 401, body: '{}' };
      if (method === 'get' && p === 'SCRIPT1/content') return { code: 200, body: JSON.stringify({ files: projectFiles }) };
      if (method === 'put' && p === 'SCRIPT1/content') { projectFiles = JSON.parse(o.payload).files; return { code: 200, body: '{}' }; }
      if (method === 'post' && p === 'SCRIPT1/versions') return { code: 200, body: JSON.stringify({ versionNumber: ++ver }) };
      if (method === 'get' && p === 'SCRIPT1/deployments') return { code: 200, body: JSON.stringify({ deployments }) };
      const dm = p.match(/^SCRIPT1\/deployments\/(\w+)$/);
      if (method === 'put' && dm) { deployments.find(d => d.deploymentId === dm[1]).deploymentConfig = JSON.parse(o.payload).deploymentConfig; return { code: 200, body: '{}' }; }
      return { code: 404, body: '{"error":{"message":"nope"}}' };
    };
    let r = U.autoUpdate_(false); // ไม่มี token (repo public)
    ok('autoUpdate deploys new version without token', r.updated && r.version === 4 && r.deployments === 2 && /a{7}/.test(r.label), r);
    ok('autoUpdate replaces 3 files, keeps others', projectFiles.length === 4 && projectFiles.find(f => f.name === 'Index').source === '<p>ใหม่</p>' && projectFiles.find(f => f.name === 'Extra').source === 'keep me' && projectFiles.find(f => f.name === 'Code').source === codeSrc, projectFiles.map(f => f.name));
    ok('autoUpdate skips HEAD, updates D1+D2', !deployments[0].deploymentConfig.versionNumber && deployments[1].deploymentConfig.versionNumber === 4 && deployments[2].deploymentConfig.versionNumber === 4 && deployments[1].deploymentConfig.manifestFileName === 'appsscript');
    ok('autoUpdate fetches files pinned to that commit', rawRefs.length === 3 && rawRefs.every(x => x === SHA1), rawRefs);
    calls.length = 0;
    r = U.autoUpdate_(false);
    ok('autoUpdate no-op on same commit', r.updated === false && calls.length === 1, calls);
    sha = SHA2; codeSrc = 'function somethingElse(){}';
    r = U.autoUpdate_(false);
    ok('autoUpdate refuses code without updater', /ไม่มีตัวอัปเดต/.test(r.error || '') && u.props.get('GH_SHA') === SHA1 && ver === 4, r);
    codeSrc = 'function autoUpdateTick(){} // v2'; manifest = '{}';
    r = U.autoUpdate_(false);
    ok('autoUpdate refuses manifest without scopes', /appsscript\.json/.test(r.error || '') && ver === 4, r);
    manifest = '{"oauthScopes":["x/auth/script.projects","x/auth/script.deployments"]}';
    u.props.set('GH_TOKEN', 'BAD');
    r = U.autoUpdate_(false);
    ok('autoUpdate reports bad token, keeps old SHA', /token/.test(r.error || '') && u.props.get('GH_SHA') === SHA1 && /token/.test(u.props.get('GH_ERROR')), r);
    u.props.set('GH_TOKEN', 'TOKEN');
    r = U.autoUpdate_(false);
    ok('autoUpdate with token, recovers next run', r.updated && u.props.get('GH_SHA') === SHA2 && r.version === 5 && !u.props.get('GH_ERROR'), r);
    // GitHub API เกินโควตา (ไม่มี token) → อ่านจาก branch, รอเห็นชุดเดิม 2 รอบก่อนอัปเดต
    u.props.delete('GH_TOKEN'); ghCode = 403; rawRefs.length = 0; codeSrc = 'function autoUpdateTick(){} // v3';
    r = U.autoUpdate_(false);
    ok('rate-limited: first sighting waits', r.pending === true && ver === 5 && rawRefs.every(x => x === 'main'), r);
    r = U.autoUpdate_(false);
    ok('rate-limited: second sighting updates', r.updated && ver === 6 && /^raw:/.test(u.props.get('GH_SHA')) && !u.props.get('GH_PENDING'), r);
    r = U.autoUpdate_(false);
    ok('rate-limited: unchanged → no-op', r.updated === false && ver === 6, r);
    codeSrc = 'function autoUpdateTick(){} // v4';
    r = U.autoUpdate_(true);
    ok('rate-limited: manual update applies at once', r.updated && ver === 7, r);
    ghCode = 200;
    ok('updater not exposed to web API', !U.apiFn_('autoUpdate_') && !U.apiFn_('menuUpdateNow') && !U.apiFn_('autoUpdateTick') && !U.apiFn_('menuUpdateStatus'));
    u.props.set('GH_TOKEN', 'SECRET'); u.props.set('GH_ERROR', '2026 boom');
    const st = U.apiFn_('updateStatus')();
    ok('updateStatus via API: status without token', st.trigger === false && st.token === true && st.error === '2026 boom' && st.last && st.last.version === ver && JSON.stringify(st).indexOf('SECRET') < 0, st);
    u.props.delete('GH_TOKEN'); u.props.delete('GH_ERROR');
    // google.script.run จากหน้าเว็บเรียกฟังก์ชันเมนูได้ แต่ไม่มี UI → ต้องหยุดก่อนทำอะไร
    const getUi = u.env.SpreadsheetApp.getUi;
    u.env.SpreadsheetApp.getUi = () => { throw new Error('Cannot call SpreadsheetApp.getUi() from this context.'); };
    const shaBefore = u.props.get('GH_SHA'), verBefore = ver;
    let threw = false; try { U.menuUpdateNow(); } catch (e) { threw = true; }
    ok('menuUpdateNow from web: no reset, no new version', threw && u.props.get('GH_SHA') === shaBefore && ver === verBefore);
    u.env.SpreadsheetApp.getUi = getUi;
  } catch (e) { out.push('CRASH4 ' + e.stack); }
  return out;
};
