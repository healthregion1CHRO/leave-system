// รันชุดทดสอบด้วย Node + Playwright (ใช้ทั้งในเครื่องและบน GitHub Actions)
// ใช้:  npm i -D playwright && npx playwright install chromium && node tests/run-tests.js
const fs = require('fs'), os = require('os'), path = require('path');
const { chromium } = require('playwright');
const repo = path.join(__dirname, '..');
const out = fs.mkdtempSync(path.join(os.tmpdir(), 'leave-tests-'));
for (const f of ['gasmock.js', 'server-tests.js', 'client-tests.js']) fs.copyFileSync(path.join(__dirname, f), path.join(out, f));
const code = fs.readFileSync(path.join(repo, 'apps-script/Code.gs'), 'utf8');
const codeJs = '<script>window.CODE_GS=' + JSON.stringify(code).replace(/<\//g, '<\\/') + ';</script>';
fs.writeFileSync(path.join(out, 'server-run.html'), `<!doctype html><meta charset=utf-8><body><pre id=o>running</pre>${codeJs}<script src=gasmock.js></script><script src=server-tests.js></script><script>try{o.textContent=runServerTests().join('\\n')}catch(e){o.textContent='CRASH '+e.stack}</script>`);
let idx = fs.readFileSync(path.join(repo, 'apps-script/Index.html'), 'utf8');
const i = idx.indexOf('<script>');
idx = idx.slice(0, i) + codeJs + '<script src=gasmock.js></script><script src=client-tests.js></script>' + idx.slice(i);
idx = idx.replace('</body>', "<pre id=__out>running</pre><script>runClientTests().then(l=>{__out.textContent=l.join('\\n')})</script></body>");
fs.writeFileSync(path.join(out, 'client-run.html'), idx);
(async () => {
  const b = await chromium.launch({ args: ['--allow-file-access-from-files'] });
  let fail = 0;
  for (const [p, id] of [['server-run.html', 'o'], ['client-run.html', '__out']]) {
    const pg = await b.newPage(); const errs = [];
    pg.on('pageerror', e => errs.push(e.message));
    await pg.goto('file://' + path.join(out, p));
    await pg.waitForFunction(id => { const e = document.getElementById(id); return e && e.textContent !== 'running'; }, id, { timeout: 120000 }).catch(() => {});
    const r = await pg.$eval('#' + id, e => e.textContent).catch(() => 'NO OUTPUT');
    console.log('== ' + p + '\n' + r + (errs.length ? '\nPAGEERR ' + errs.join('\n') : ''));
    if (!/PASS/.test(r) || /^(FAIL|CRASH)/m.test(r)) fail++;
  }
  await b.close();
  console.log(fail ? '\nมีการทดสอบไม่ผ่าน' : '\nผ่านทั้งหมด');
  process.exit(fail ? 1 : 0);
})();
