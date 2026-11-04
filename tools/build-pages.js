// สร้างหน้าเว็บ GitHub Pages จาก apps-script/Index.html (ไฟล์เดียวกับที่ Apps Script ใช้)
// ผลลัพธ์อยู่ในโฟลเดอร์ _site: index.html (เติม config.js ให้) + config.js + install-guide.html
// ใช้:  node tools/build-pages.js [โฟลเดอร์ผลลัพธ์]
const fs = require('fs'), path = require('path');
const root = path.join(__dirname, '..');
const out = path.resolve(process.argv[2] || path.join(root, '_site'));
const src = fs.readFileSync(path.join(root, 'apps-script/Index.html'), 'utf8');
const i = src.indexOf('<script>');
if (i < 0) throw new Error('ไม่พบ <script> ใน apps-script/Index.html');
const html = src.slice(0, i) + '<script src="config.js"></script>\n' + src.slice(i);
fs.mkdirSync(out, { recursive: true });
fs.writeFileSync(path.join(out, 'index.html'), html);
for (const f of ['config.js', 'install-guide.html']) {
  const p = path.join(root, 'docs', f);
  if (fs.existsSync(p)) fs.copyFileSync(p, path.join(out, f));
}
fs.writeFileSync(path.join(out, '.nojekyll'), '');
const cfg = fs.readFileSync(path.join(out, 'config.js'), 'utf8');
if (!/LEAVE_API_URL\s*=\s*"https:\/\/script\.google\.com\/macros\/s\/[^"]+\/exec"/.test(cfg)) throw new Error('docs/config.js ยังไม่มีลิงก์ /exec ของ Apps Script');
console.log('สร้างหน้าเว็บแล้ว: ' + out + ' (' + Math.round(html.length / 1024) + ' KB)');
