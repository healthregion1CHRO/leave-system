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
  A.addHoliday(T0, '2026-10-23', 'วันปิยมหาราช'); // no default holidays any more — seed one for the holiday test
  A.saveStaff(T0, { name: 'นางสาวบานเย็น แสงแก้ว', level: 'ชำนาญการ', group: 'กลุ่มงานยุทธศาสตร์และสารสนเทศ+สุขภาพดิจิทัล', section: 'งานสุขภาพดิจิทัล', ptype: 'ข้าราชการ', cid: ST_CID });
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
      // ค้นหาบุคลากร: พิมพ์บางส่วนของชื่อ → เลือกด้วยแป้นลูกศร/Enter
      $('#f_staff_q').focus(); await wait();
      ok('staff search opens full list on focus', !$('#f_staff_list').hidden && document.querySelectorAll('#f_staff_list li[data-sp]').length === 3, document.querySelectorAll('#f_staff_list li[data-sp]').length);
      await setv('#f_staff_q', 'ไม่มีชื่อนี้', 'input');
      ok('staff search: no match message', /ไม่พบบุคลากร/.test($('#f_staff_list').innerText));
      await setv('#f_staff_q', 'สุขภาพดิจิทัล', 'input');
      ok('staff search by section', [...document.querySelectorAll('#f_staff_list li[data-sp]')].map(li => li.innerText).join('|').includes('บานเย็น') && document.querySelectorAll('#f_staff_list li[data-sp]').length === 1);
      await setv('#f_staff_q', 'แก้ว', 'input');
      ok('staff search by partial name', document.querySelectorAll('#f_staff_list li[data-sp]').length === 2 && !!$('#f_staff_list mark'));
      const kd = k => $('#f_staff_q').dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true }));
      const target = [...document.querySelectorAll('#f_staff_list li[data-sp]')].findIndex(li => /บานเย็น/.test(li.innerText));
      for (let i = 0; i < target; i++) { kd('ArrowDown'); await wait(20); }
      kd('Enter'); await wait();
      const bid = $('#f_staff').value;
      ok('staff picked with keyboard', /บานเย็น/.test($('#f_staff_q').value) && bid && $('#f_staff_list').hidden && /สุขภาพดิจิทัล/.test($('.sp-sel').innerText), $('.sp-sel') && $('.sp-sel').innerText);
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

      // ปฏิทินการลา: เลื่อนไป พ.ย. 2569 (2–4 พ.ย. 3 วัน + 5 พ.ย. ครึ่งเช้า + 6 พ.ย.)
      await click('[data-act="tab"][data-v="calendar"]');
      ok('calendar tab shows month grid', !!$('.lc-grid') && document.querySelectorAll('.lc-wd').length === 7);
      const now = new Date(), diff = (2026 * 12 + 10) - (now.getFullYear() * 12 + now.getMonth());
      for (let i = 0; i < Math.abs(diff); i++) await click(`[data-act="cal-nav"][data-d="${diff > 0 ? 1 : -1}"]`);
      ok('calendar shows พฤศจิกายน 2569', /พฤศจิกายน 2569/.test($('.lc-title').innerText), $('.lc-title').innerText);
      const cell = d => $(`[data-act="cal-day"][data-d="${d}"]`);
      ok('leave days marked on grid', /\b1\b/.test(cell('2026-11-03').querySelector('.lc-count')?.innerText || '') && !!cell('2026-11-05').querySelector('.lc-k'), cell('2026-11-03').innerText);
      ok('day cell shows type + count, no names', /ลาป่วย\s*1/.test(cell('2026-11-03').querySelector('.lc-kinds')?.innerText || '') && !cell('2026-11-03').querySelector('.lc-chip') && !/บานเย็น/.test(cell('2026-11-03').innerText) && !cell('2026-11-07').querySelector('.lc-kinds'), cell('2026-11-03').innerText.replace(/\s+/g, ' '));
      ok('weekend not marked', !cell('2026-11-07').querySelector('.lc-count'));
      await click(cell('2026-11-05'));
      ok('day detail lists person with half-day', /ผู้ลาวันที่ 5 พ.ย. 2569/.test(text()) && /ครึ่งเช้า/.test(text()));
      ok('month list has 3 leaves', document.querySelectorAll('.tbl-wrap')[1]?.querySelectorAll('tbody tr').length === 3);
      const tileDays = k => Number((($(`.cx-tile[data-act="cal-kind"][data-k="${k}"] .val`) || {}).innerText || '').replace(/[^\d.]/g, ''));
      const sickDays = tileDays('sick');
      ok('summary tiles per type', !!$('.lc-stats') && sickDays > 0 && tileDays('all') === sickDays + tileDays('personal') + tileDays('vacation') && /คน · \d+ ครั้ง/.test($('.cx-tile[data-k="sick"]').innerText), [tileDays('all'), sickDays, tileDays('personal'), tileDays('vacation')].join(','));
      await click('.lc-per summary');
      const perRows = [...document.querySelectorAll('.lc-per-tbl tbody tr')];
      ok('per-person summary table', perRows.length >= 2 && /บานเย็น/.test(perRows[0].innerText) && /รวม/.test(perRows[perRows.length - 1].innerText) && perRows[perRows.length - 1].innerText.includes(String(sickDays)), perRows.map(r => r.innerText.replace(/\s+/g, ' ')).join(' | '));
      await click('.lc-per summary');
      ok('per-person summary collapses', !$('.lc-per-tbl'));
      await click('.cx-tile[data-act="cal-kind"][data-k="vacation"]');
      ok('tile click filters calendar', $('#cal_kind').value === 'vacation' && !cell('2026-11-03').querySelector('.lc-count') && tileDays('sick') === sickDays);
      await click('.cx-tile[data-act="cal-kind"][data-k="all"]');
      await setv('#cal_kind', 'vacation');
      ok('kind filter hides sick leave', !cell('2026-11-03').querySelector('.lc-count'));
      await setv('#cal_kind', 'all');

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
      ok('ORG group/section preselected', $('#s_group').value === 'กลุ่มงานยุทธศาสตร์และสารสนเทศ+สุขภาพดิจิทัล' && $('#s_section').value === 'งานสุขภาพดิจิทัล', $('#s_group').value + '|' + $('#s_section').value);
      await click('[data-act="cancel-staff"]');

      // person view + delete a leave
      await click('[data-act="tab"][data-v="summary"]');
      await click(`[data-act="person"][data-id="${bRow[0]}"]`);
      ok('person view shows balances', /คงเหลือ/.test(text()));
      const askBtn = document.querySelector('[data-act="ask"][data-k^="del:"]');
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

      // ขอยกเลิกวันลา: ครึ่งวัน 5 พ.ย. (แถวแรก เรียงล่าสุดก่อน) → ใช้ไป 3, คงเหลือ 57
      const usedSick = () => { const c = card(/ลาป่วย/); const mm = c && c.innerText.match(/ประจำปี \(ปีงบ \d+\)\s*([\d.]+) วัน/); return mm && Number(mm[1]); };
      await click('[data-act="ask"][data-k^="cxl:"]');
      ok('cancel form opens with reason box', !!$('#cxl_reason') && document.activeElement === $('#cxl_reason'));
      $('#cxlForm').requestSubmit(); await wait(200);
      ok('cancel without reason blocked', !!$('#cxl_reason') && !leaves().some(r => r[14] === 'ยกเลิก'));
      await setv('#cxl_reason', 'กลับมาปฏิบัติงาน', 'input'); $('#cxlForm').requestSubmit(); await wait(250);
      const cRow = leaves().find(r => r[14] === 'ยกเลิก');
      ok('cancel saved to sheet', cRow && Number(cRow[8]) === 0.5 && cRow[15] === 'กลับมาปฏิบัติงาน' && leaves().length === 2, JSON.stringify(cRow));
      ok('cancelled leave not counted (used 3)', usedSick() === 3 && /57/.test(card(/ลาป่วย/).innerText), usedSick());
      ok('cancelled row shown with reason', !!$('tr.cancelled') && /ยกเลิกแล้ว/.test($('tr.cancelled').innerText) && /กลับมาปฏิบัติงาน/.test($('tr.cancelled').innerText));
      await click('[data-act="ask"][data-k^="rst:"]'); await click('[data-act="restore-leave"]'); await wait(200);
      ok('restore counts again (used 3.5)', usedSick() === 3.5 && !$('tr.cancelled') && !leaves().some(r => r[14] === 'ยกเลิก'), usedSick());
      // ยกเลิกไว้อีกครั้งเพื่อตรวจมุมมองของบุคลากรตอนท้าย
      await click('[data-act="ask"][data-k^="cxl:"]'); await setv('#cxl_reason', 'ทดสอบ', 'input'); $('#cxlForm').requestSubmit(); await wait(250);

      // เมนูยกเลิกวันลา: กรองตามประเภท ยกเลิก/คืนสถานะ จากเมนูนี้ได้
      await click('[data-act="tab"][data-v="cancel"]');
      const tbls = () => [...document.querySelectorAll('#app table')].map(t => t.tBodies[0].rows.length);
      const tileVal = k => ($('.cx-tile[data-k="' + k + '"] .val') || {}).innerText || '';
      ok('cancel menu opens', /ยกเลิกวันลา ปีงบประมาณ/.test($('#app').innerText) && /^1/.test(tileVal('sick')) && /^1/.test(tileVal('all')) && /^0/.test(tileVal('vacation')), tileVal('all') + ' | ' + tileVal('sick'));
      ok('cancel menu lists active + history', JSON.stringify(tbls()) === '[1,1]' && /ทดสอบ/.test($('tr.cancelled').innerText), JSON.stringify(tbls()));
      await click('.chip[data-act="cx-kind"][data-k="vacation"]');
      ok('filter by type (vacation → none)', tbls().length === 0 && /ไม่มีรายการลาที่ตรงเงื่อนไข/.test($('#app').innerText), JSON.stringify(tbls()));
      await click('.chip[data-act="cx-kind"][data-k="sick"]');
      // ค้นหาชื่อบุคคล: พิมพ์บางส่วนของชื่อ → เลือกจากรายการ
      const cxPick = async q => { $('#cx_staff_q').focus(); await wait(); await setv('#cx_staff_q', q, 'input');
        const li = $('#cx_staff_list li[data-sp]'); if (!li) throw new Error('no match for ' + q);
        li.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); await wait(150); };
      $('#cx_staff_q').focus(); await wait();
      ok('cancel menu: name search opens list', !$('#cx_staff_list').hidden && document.querySelectorAll('#cx_staff_list li[data-sp]').length === 3);
      await setv('#cx_staff_q', 'กิตติ', 'input');
      ok('cancel menu: name search filters', document.querySelectorAll('#cx_staff_list li[data-sp]').length === 1 && !!$('#cx_staff_list mark'));
      await cxPick('กิตติ');
      ok('cancel menu: pick person filters table', $('#cx_staff_q').value.includes('กิตติ') && $('#cxstaff').value !== 'all' && tbls().length === 0);
      await click('[data-act="sp-clear"][data-for="cx_staff_q"]');
      ok('cancel menu: clear person = everyone', $('#cxstaff').value === 'all' && $('#cx_staff_q').value === '' && JSON.stringify(tbls()) === '[1,1]', JSON.stringify(tbls()));
      await cxPick('บานเย็น');
      ok('filter by type + staff (sick)', JSON.stringify(tbls()) === '[1,1]', JSON.stringify(tbls()));
      await click('#app tbody tr:not(.cancelled) [data-act="ask"][data-k^="cxl:"]');
      await setv('#cxl_reason', 'ยกเลิกจากเมนู', 'input'); $('#cxlForm').requestSubmit(); await wait(250);
      ok('cancel from menu saved', leaves().filter(r => r[14] === 'ยกเลิก').length === 2 && leaves().some(r => r[15] === 'ยกเลิกจากเมนู') && /^2/.test(tileVal('sick')) && JSON.stringify(tbls()) === '[2]', JSON.stringify(tbls()));
      const rr = [...document.querySelectorAll('#app tr.cancelled')].find(r => /ยกเลิกจากเมนู/.test(r.innerText));
      await click(rr.querySelector('[data-k^="rst:"]')); await click('[data-act="restore-leave"]'); await wait(200);
      ok('restore from menu', leaves().filter(r => r[14] === 'ยกเลิก').length === 1 && !leaves().some(r => r[15] === 'ยกเลิกจากเมนู') && JSON.stringify(tbls()) === '[1,1]', JSON.stringify(tbls()));
      // ยกเลิกเฉพาะบางวัน: ลาป่วย 2–4 พ.ย. (3 วัน) ยกเลิกเฉพาะ 3 พ.ย.
      await click('#app tbody tr:not(.cancelled) [data-act="ask"][data-k^="cxl:"]');
      ok('cancel form has date range (default = whole leave)', !!$('#cxl_from') && !!$('#cxl_to') && $('#cxl_from').value === '2026-11-02' && $('#cxl_to').value === '2026-11-04' && /3 วันทำการ \(ทั้งรายการ\)/.test($('#cxl_days').innerText), $('#cxl_days') && $('#cxl_days').innerText);
      await setv('#cxl_reason', 'มาปฏิบัติงาน 3 พ.ย.', 'input');
      await setv('#cxl_from', '2026-11-03'); await setv('#cxl_to', '2026-11-02');
      ok('range check: to before from warned', /ไม่ก่อน/.test($('#cxl_days').innerText) && $('#cxl_reason').value === 'มาปฏิบัติงาน 3 พ.ย.');
      $('#cxlForm').requestSubmit(); await wait(250);
      ok('to before from not saved', leaves().filter(r => r[14] === 'ยกเลิก').length === 1 && !!$('#cxlForm'));
      await setv('#cxl_to', '2026-11-03');
      ok('preview shows 1 day cancelled, 2 kept', /ยกเลิก 1 วันทำการ/.test($('#cxl_days').innerText) && /คงเหลือเป็นวันลา 2 วัน/.test($('#cxl_days').innerText), $('#cxl_days').innerText);
      $('#cxlForm').requestSubmit(); await wait(250);
      const part = leaves().filter(r => r[1] === leaves().find(x => x[15] === 'มาปฏิบัติงาน 3 พ.ย.')[1] && r[3] === 'sick' && r[5] >= '2026-11-02' && r[6] <= '2026-11-04').map(r => [r[5], r[6], r[8], r[14] || '-'].join('|')).sort();
      ok('partial cancel splits leave in sheet', JSON.stringify(part) === JSON.stringify(['2026-11-02|2026-11-02|1|-', '2026-11-03|2026-11-03|1|ยกเลิก', '2026-11-04|2026-11-04|1|-']), JSON.stringify(part));
      ok('partial cancel shown in menu', JSON.stringify(tbls()) === '[2,2]' && /^2/.test(tileVal('sick')), JSON.stringify(tbls()));
      const pr = [...document.querySelectorAll('#app tr.cancelled')].find(r => /มาปฏิบัติงาน 3 พ\.ย\./.test(r.innerText));
      await click(pr.querySelector('[data-k^="rst:"]')); await click('[data-act="restore-leave"]'); await wait(200);
      ok('restore partial cancel', leaves().filter(r => r[14] === 'ยกเลิก').length === 1 && JSON.stringify(tbls()) === '[3,1]', JSON.stringify(tbls()));
      await click('[data-act="ask"][data-k^="cxl:"]');
      const oneRow = $('#cxlForm') && $('#cxlForm').closest('tr');
      ok('single-day leave: no date range', oneRow && !/ – /.test(oneRow.cells[2].innerText) && !$('#cxl_from'), oneRow && oneRow.cells[2].innerText);
      await click('[data-act="cancel"]');

      // settings: add holiday via UI
      await click('[data-act="tab"][data-v="settings"]');
      if ($('#h_d_txt')) { await setv('#h_d_txt', '05/12/2569'); await setv('#h_n', 'ทดสอบ', 'input'); $('#holForm').requestSubmit(); await wait(200); }
      ok('add holiday via UI', m.ss.getSheetByName('วันหยุด')._s.rows.some(r => r[0] === '2026-12-05'));

      // settings: leave entitlement rows are read-only until แก้ไข
      const typeRow = code => m.ss.getSheetByName('สิทธิ์การลา')._s.rows.find(r => r[0] === code);
      ok('types read-only by default', !$('[data-cfg]') && document.querySelectorAll('[data-act="edit-type"]').length === 4);
      await click('[data-act="edit-type"][data-id="emp"]');
      ok('edit shows inputs for one row only', document.querySelectorAll('[data-cfg]').length === 4 && !!$('[data-cfg="emp.sick"]') && !!$('[data-act="save-type"]'));
      ok('other edit buttons disabled while editing', $('[data-act="edit-type"][data-id="gov"]').disabled === true);
      await setv('[data-cfg="emp.personal"]', '12', 'input');
      await click('[data-act="cancel-type"]');
      ok('cancel discards', !$('[data-cfg]') && Number(typeRow('emp')[3]) === 10, JSON.stringify(typeRow('emp')));
      await click('[data-act="edit-type"][data-id="emp"]');
      await setv('[data-cfg="emp.personal"]', '-1', 'input');
      await click('[data-act="save-type"]');
      ok('negative rejected, still editing', !!$('[data-cfg]') && Number(typeRow('emp')[3]) === 10);
      await setv('[data-cfg="emp.personal"]', '12', 'input');
      await click('[data-act="save-type"]'); await wait(200);
      ok('save-type writes sheet and closes edit', !$('[data-cfg]') && Number(typeRow('emp')[3]) === 12 && Number(typeRow('gov')[3]) === 45, JSON.stringify(typeRow('emp')));

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
      ok('staff sees cancelled leave, no HR buttons', !!$('tr.cancelled') && !$('[data-k^="rst:"]') && !$('[data-act="rv"]') && !$('[data-k^="del:"]'));
      ok('staff has no cancel menu', $('#tabs').hidden && !/ยกเลิกวันลา ปีงบประมาณ/.test($('#app').innerText));
      // ค่าเริ่มต้น: ล็อกปุ่มของบุคลากร → ปุ่มสีเทา กดไม่ได้
      const lockedLabels = [...document.querySelectorAll('#app button.locked')].map(b => b.disabled && b.textContent.trim());
      ok('locked: request button grey + disabled', !$('[data-act="req-open"]') && lockedLabels.includes('+ ยื่นใบลา') && /ติดต่อเจ้าหน้าที่ HR/.test(text()), JSON.stringify(lockedLabels));
      ok('locked: cancel request button disabled', lockedLabels.includes('ขอยกเลิกวันลา') && !$('[data-k^="cxl:"]') && !$('[data-k^="wd:"]'), JSON.stringify(lockedLabels));
      m.api.setSetting_('staffRequests', 'เปิด'); m.cache.clear();
      await click('[data-act="logout"]'); await wait(150);
      await login('บานเย็น แสงแก้ว', ST_CID);
      if ($('#pwForm')) await click('[data-act="pw-close"]');
      ok('unlocked: buttons active again', !!$('[data-act="req-open"]') && !$('#app button.locked'));

      // ยื่นใบลา → HR อนุมัติ → ขอยกเลิก → HR ไม่อนุมัติ (มีเหตุผล)
      const row = id => leaves().find(r => r[0] === id);
      await click('[data-act="req-open"]');
      ok('request form opens', !!$('#leaveForm') && !$('#f_staff_q') && !$('[data-act="req-open"]'));
      $('input[name="f_kind"][value="vacation"]').click(); await wait();
      await setv('#f_start_txt', '15/03/2570'); await setv('#f_end_txt', '16/03/2570');
      ok('request form shows days + balance', /ลา\s*2\s*วันทำการ/.test($('#calcBox').innerText), $('#calcBox').innerText);
      await setv('#f_note', 'พักผ่อนกับครอบครัว', 'input');
      $('#leaveForm').requestSubmit(); await wait(250);
      const req = leaves().find(r => r[5] === '2027-03-15');
      ok('requestLeave saved as pending', req && req[18] === 'รอพิจารณา' && req[8] === 2 && req[3] === 'vacation' && !$('#leaveForm'), JSON.stringify(req));
      ok('staff sees pending status + withdraw', /รอพิจารณา/.test($('#app').innerText) && !!$(`[data-k="wd:${req[0]}"]`) && /รอพิจารณาอีก 2 วัน/.test($('#app').innerText));
      await click('[data-act="logout"]'); await wait(150);
      await login('นายทดสอบ ผู้ดูแล', 'hrpass123');
      ok('HR pending badge', ($('#tabs .tab-badge') || {}).textContent === '1');
      await click('[data-act="tab"][data-v="pending"]');
      ok('pending tab lists request', /บานเย็น/.test($('#app').innerText) && !!$(`[data-act="rv"][data-id="${req[0]}"][data-s="approved"]`));
      await click(`[data-act="rv"][data-t="leave"][data-id="${req[0]}"][data-s="approved"]`); await wait(150);
      ok('HR approve writes sheet', row(req[0])[18] === 'อนุมัติ' && row(req[0])[19] === 'นายทดสอบ ผู้ดูแล' && !$('#tabs .tab-badge'), JSON.stringify(row(req[0])));
      await click('[data-act="logout"]'); await wait(150);
      await login('บานเย็น แสงแก้ว', ST_CID);
      if ($('#pwForm')) await click('[data-act="pw-close"]');
      await click(`[data-k="cxl:${req[0]}"]`);
      ok('staff cancel form says request', /ส่งคำขอยกเลิก/.test($('#cxlForm').innerText));
      await setv('#cxl_reason', 'งานด่วน', 'input');
      $('#cxlForm').requestSubmit(); await wait(250);
      ok('requestCancel → pending, still counted', row(req[0])[14] === 'ขอยกเลิก' && row(req[0])[15] === 'งานด่วน' && !!$(`[data-act="withdraw-cancel"][data-id="${req[0]}"]`), JSON.stringify(row(req[0])));
      await click('[data-act="logout"]'); await wait(150);
      await login('นายทดสอบ ผู้ดูแล', 'hrpass123');
      ok('cancel request badge on cancel tab', ($('#tabs [data-v="cancel"] .tab-badge') || {}).textContent === '1' && !$('#tabs [data-v="pending"] .tab-badge'));
      await click('[data-act="tab"][data-v="pending"]');
      ok('pending tab has no cancel requests', !$(`[data-act="rv"][data-t="cx"]`) && !/คำขอยกเลิก/.test($('#app').innerText));
      await click('[data-act="tab"][data-v="cancel"]');
      ok('opening cancel menu clears filters', ($('[data-act="cx-kind"][aria-pressed="true"]') || {}).dataset?.k === 'all' && $('#cxstaff').value === 'all' && (!$('#cxq') || $('#cxq').value === ''));
      await click(`[data-act="rv"][data-t="cx"][data-id="${req[0]}"][data-s="rejected"]`);
      ok('reject opens reason form', !!$('#rjForm'));
      $('#rjForm').requestSubmit(); await wait(150);
      ok('reject without reason blocked', row(req[0])[14] === 'ขอยกเลิก' && /เหตุผล/.test(text()));
      await setv('#rj_reason', 'ช่วงนั้นคนพอ', 'input');
      $('#rjForm').requestSubmit(); await wait(250);
      ok('HR reject cancel writes sheet', row(req[0])[14] === 'ไม่อนุมัติการยกเลิก' && row(req[0])[24] === 'ช่วงนั้นคนพอ' && !$('#tabs .tab-badge'), JSON.stringify(row(req[0])));
    } catch (e) { log.push('CRASH ' + e.stack); }
    ok('no JS errors', window.__errors.length === 0, JSON.stringify(window.__errors));
    return log;
  };
})();
