// Drives Index.html through the DOM, with google.script.run backed by Code.gs in GasMock.
(function () {
  const log = [];
  window.__errors = [];
  window.addEventListener('error', e => window.__errors.push(e.message));
  window.addEventListener('unhandledrejection', e => window.__errors.push('unhandled: ' + (e.reason && e.reason.message || e.reason)));
  const m = GasMock.loadServer(window.CODE_GS); const A = m.api; const { cid } = GasMock;
  const HR_CID = cid('110170012345'), ST_CID = cid('320050011111');
  m.prompts.push('นายทดสอบ ผู้ดูแล', HR_CID); A.setupSystem();
  const T0 = A.login('ทดสอบ ผู้ดูแล', HR_CID).token;
  A.saveStaff(T0, { name: 'นางสาวบานเย็น แสงแก้ว', level: 'ชำนาญการ', group: 'กลุ่มงานยุทธศาสตร์และแผนงาน', section: 'งาน Digital Health & Data', ptype: 'ข้าราชการ', cid: ST_CID });
  A.saveStaff(T0, { name: 'นายกิตติ ก้อนแก้ว', group: 'กลุ่มงานอำนวยการ', section: 'งานที่ไม่อยู่ในรายการ', ptype: 'พนักงานกระทรวง', cid: cid('320050044444') });
  m.cache.clear();
  const hasDate = o => { let b = false; (function w(x) { if (x instanceof Date) b = true; else if (x && typeof x === 'object') for (const k in x) w(x[k]); })(o); return b; };
  window.google = { script: { run: (function mk(ok, fail) {
    return new Proxy({}, { get: (_, fn) => {
      if (fn === 'withSuccessHandler') return h => mk(h, fail);
      if (fn === 'withFailureHandler') return h => mk(ok, h);
      return (...args) => setTimeout(() => {
        try { const r = A[fn](...args); if (hasDate(r)) { log.push('FAIL Date returned by ' + fn); ok && ok(null); } else ok && ok(JSON.parse(JSON.stringify(r === undefined ? null : r))); }
        catch (e) { fail && fail(new Error(e.message)); }
      }, 5);
    } });
  })(null, null) } };

  const $ = s => document.querySelector(s);
  const wait = ms => new Promise(r => setTimeout(r, ms || 60));
  const ok = (n, c, info) => log.push((c ? 'PASS ' : 'FAIL ') + n + (c || info === undefined ? '' : ' :: ' + info));
  const setv = async (sel, v, ev = 'change') => { const el = $(sel); if (!el) throw new Error('no ' + sel); el.value = v; el.dispatchEvent(new Event('input', { bubbles: true })); if (ev === 'change') el.dispatchEvent(new Event('change', { bubbles: true })); await wait(); };
  const click = async sel => { const el = typeof sel === 'string' ? $(sel) : sel; if (!el) throw new Error('no ' + sel); el.click(); await wait(120); };
  const text = () => document.body.innerText;
  const leaves = () => m.ss.getSheetByName('การลา')._s.rows.slice(1);
  const staffRow = name => m.ss.getSheetByName('บุคลากร')._s.rows.find(r => r[1] === name);
  const login = async (u, p) => { await setv('#lg_user', u); await setv('#lg_pw', p, 'input'); $('#loginForm').requestSubmit(); await wait(200); };

  window.runClientTests = async function () {
    try {
      await wait(200);
      ok('login page shows', !!$('#loginForm'));
      await login('สมมติ ไม่มีชื่อ', '123');
      ok('wrong login shows error', /ไม่ถูกต้อง/.test(text()));
      await login('นายทดสอบ ผู้ดูแล', HR_CID);
      ok('HR login → first-login password modal', !!$('#pwForm'));
      await click('[data-act="pw-close"]');
      ok('summary visible', /สรุป/.test(text()) && /บานเย็น/.test(text()));

      // record leave, change dates after manual days override
      await click('[data-act="tab"][data-v="record"]');
      const staffSel = $('#f_staff'); const bid = [...staffSel.options].find(o => /บานเย็น/.test(o.text)).value;
      await setv('#f_staff', bid);
      await setv('#f_start_txt', '02/11/2569'); await setv('#f_end_txt', '02/11/2569');
      await setv('#f_days', '1.', 'input');
      ok('typing "1." keeps field', $('#f_days').value === '' || $('#f_days').value === '1.' , JSON.stringify($('#f_days').value));
      await setv('#f_days', '2', 'input');
      await setv('#f_end_txt', '04/11/2569');
      ok('date change clears manual days', $('#f_days').value === '', $('#f_days').value);
      ok('half-day select disabled for multi-day', $('#f_half').disabled === true);
      await setv('#f_note', 'ไข้หวัด', 'input'); $('#f_note').dispatchEvent(new Event('change', { bubbles: true })); await wait();
      ok('save button still in DOM after note change', !!document.querySelector('#leaveForm button[type=submit]'));
      await click('#leaveForm button[type=submit]'); await wait(200);
      let L = leaves();
      ok('leave saved with auto days 3', L.length === 1 && Number(L[0][8]) === 3 && L[0][10] === 'ไข้หวัด', JSON.stringify(L));

      // half day + holiday handling
      await setv('#f_start_txt', '23/10/2569'); await setv('#f_end_txt', '23/10/2569');
      ok('holiday 23 Oct counts 0 days', /\(คำนวณอัตโนมัติ 0 วัน\)/.test(text()), (text().match(/คำนวณอัตโนมัติ[^)]*\)/) || [])[0]);
      await setv('#f_start_txt', '05/11/2569'); await setv('#f_end_txt', '05/11/2569');
      ok('half-day enabled for single day', $('#f_half').disabled === false);
      await setv('#f_half', 'am');
      await click('#leaveForm button[type=submit]'); await wait(200);
      L = leaves();
      ok('half-day saved 0.5', L.length === 2 && Number(L[1][8]) === 0.5 && L[1][7] === 'am', JSON.stringify(L[1]));

      // calendar open then click save → should submit on first click
      await setv('#f_start_txt', '06/11/2569'); await setv('#f_end_txt', '06/11/2569');
      await click('[data-act="tdp-open"]');
      ok('calendar opened', !!$('.cal'));
      await click('#leaveForm button[type=submit]'); await wait(200);
      ok('save works while calendar open', leaves().length === 3, leaves().length);

      // edit staff with section not in ORG → unchanged
      await click('[data-act="tab"][data-v="staff"]');
      const kid = staffRow('นายกิตติ ก้อนแก้ว')[0];
      await click(`[data-act="edit-staff"][data-id="${kid}"]`);
      ok('staff form open', !!$('#staffForm'));
      ok('legacy section shown', /งานที่ไม่อยู่ในรายการ/.test($('#s_section').innerHTML));
      $('#staffForm').requestSubmit(); await wait(250);
      ok('section preserved after edit', staffRow('นายกิตติ ก้อนแก้ว')[5] === 'งานที่ไม่อยู่ในรายการ', staffRow('นายกิตติ ก้อนแก้ว')[5]);
      const bRow = staffRow('นางสาวบานเย็น แสงแก้ว');
      await click(`[data-act="edit-staff"][data-id="${bRow[0]}"]`);
      ok('ORG group/section preselected', $('#s_group').value === 'กลุ่มงานยุทธศาสตร์และแผนงาน' && $('#s_section').value === 'งาน Digital Health & Data', $('#s_group').value + '|' + $('#s_section').value);
      await click('[data-act="cancel-staff"]');

      // person view + delete a leave
      await click('[data-act="tab"][data-v="summary"]');
      await click(`[data-act="person"][data-id="${bRow[0]}"]`);
      ok('person view shows balances', /คงเหลือ/.test(text()));
      const askBtn = document.querySelector('[data-act="ask"]');
      await click(askBtn); await click('[data-act="del-leave"]'); await wait(200);
      ok('delete leave', leaves().length === 2, leaves().length);
      // vacation card: 10/yr, no carry yet → quota 10; enter carry 5 → quota 15
      const card = re => [...document.querySelectorAll('.bal')].find(b => re.test(b.innerText) && /รวมสิทธิ์ปีนี้/.test(b.innerText));
      const vacCard = () => card(/ลาพักผ่อน/);
      const quotaOf = card => { const m = card.innerText.match(/รวมสิทธิ์ปีนี้\s*([\d.]+)/); return m && Number(m[1]); };
      ok('vacation quota 10 without carry', quotaOf(vacCard()) === 10, vacCard() && vacCard().innerText.slice(0, 200));
      await setv('#carry_vacation', '5', 'input'); await click('[data-act="save-carry"]'); await wait(200);
      const cr = m.ss.getSheetByName('ยอดยกมา')._s.rows.find(r => r[0] === bRow[0]);
      ok('carry saved to sheet', cr && Number(cr[5]) === 5 && cr[3] === '' && cr[4] === '', JSON.stringify(cr));
      ok('vacation quota 15 with carry 5', quotaOf(vacCard()) === 15, quotaOf(vacCard()));
      const sickCard = card(/ลาป่วย/);
      // remaining sick leaves after the delete: 2–4 Nov (3) + half day 5 Nov (0.5) → used 3.5, left 56.5
      ok('sick: quota 60, used 3.5, left 56.5', quotaOf(sickCard) === 60 && /56\.5/.test(sickCard.innerText) && /3\.5 วัน/.test(sickCard.innerText), sickCard && sickCard.innerText.replace(/\s+/g, ' ').slice(0, 300));

      // settings: add holiday via UI
      await click('[data-act="tab"][data-v="settings"]');
      if ($('#h_d_txt')) { await setv('#h_d_txt', '05/12/2569'); await setv('#h_n', 'ทดสอบ', 'input'); $('#holForm').requestSubmit(); await wait(200); }
      ok('add holiday via UI', m.ss.getSheetByName('วันหยุด')._s.rows.some(r => r[0] === '2026-12-05'));

      // change own password via modal
      await click('[data-act="pw-open"]');
      await setv('#p_old', HR_CID, 'input'); await setv('#p_new', 'hrpass123', 'input'); await setv('#p_new2', 'hrpass123', 'input');
      $('#pwForm').requestSubmit(); await wait(250);
      ok('HR change password stays logged in', !$('#pwForm') && !$('#loginForm') && /เปลี่ยนรหัสผ่านแล้ว/.test(text()));

      // logout, staff login, self view
      await click('[data-act="logout"]'); await wait(150);
      ok('logout → login page', !!$('#loginForm'));
      await login('บานเย็น แสงแก้ว', ST_CID);
      if ($('#pwForm')) await click('[data-act="pw-close"]');
      ok('staff sees own leave only', /บานเย็น/.test(text()) && !/กิตติ/.test(text()) && /ลาป่วย/.test(text()));
    } catch (e) { log.push('CRASH ' + e.stack); }
    ok('no JS errors', window.__errors.length === 0, JSON.stringify(window.__errors));
    return log;
  };
})();
