# ทดสอบเวอร์ชัน Apps Script ก่อนขึ้น GitHub: รัน Code.gs บนชีตจำลอง และทดสอบ Index.html ใน Chrome แบบ headless
# วิธีใช้:  powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1
$tests = $PSScriptRoot
$repo = Split-Path $tests -Parent
$chrome = @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $chrome) { throw 'ไม่พบ Chrome หรือ Edge' }
$out = Join-Path $env:TEMP 'leave-system-tests'
New-Item -ItemType Directory -Force $out | Out-Null
Copy-Item "$tests\gasmock.js", "$tests\server-tests.js", "$tests\client-tests.js" $out -Force
$utf8 = [Text.UTF8Encoding]::new($false)
$code = [IO.File]::ReadAllText("$repo\apps-script\Code.gs")
$codeJs = "<script>window.CODE_GS=" + ((ConvertTo-Json $code -Compress) -replace '</', '<\/') + ";</script>"

# server
[IO.File]::WriteAllText("$out\server-run.html", "<!doctype html><meta charset=utf-8><body><pre id=o>running</pre>$codeJs<script src=gasmock.js></script><script src=server-tests.js></script><script>try{o.textContent=runServerTests().join('\n')}catch(e){o.textContent='CRASH '+e.stack}</script>", $utf8)
# client: Index.html with google.script.run backed by the simulated server
$idx = [IO.File]::ReadAllText("$repo\apps-script\Index.html")
$i = $idx.IndexOf('<script>')
$idx = $idx.Substring(0, $i) + "$codeJs<script src=gasmock.js></script><script src=client-tests.js></script>" + $idx.Substring($i)
$idx = $idx.Replace('</body>', "<pre id=__out>running</pre><script>runClientTests().then(l=>{__out.textContent=l.join('\n')})</script></body>")
[IO.File]::WriteAllText("$out\client-run.html", $idx, $utf8)

$fail = 0
foreach ($p in @(@('server-run.html', 'o'), @('client-run.html', '__out'))) {
  $url = 'file:///' + ($out -replace '\\', '/') + '/' + $p[0]
  # Chrome writes harmless warnings to stderr; discard them via cmd so PowerShell 5.1 doesn't turn them into errors
  $dom = cmd /c "`"$chrome`" --headless=new --disable-gpu --allow-file-access-from-files --virtual-time-budget=60000 --dump-dom `"$url`" 2>nul" | Out-String
  $res = [Net.WebUtility]::HtmlDecode([regex]::Match($dom, "(?s)<pre id=`"$($p[1])`">(.*?)</pre>").Groups[1].Value)
  Write-Output "== $($p[0])"; Write-Output $res
  if ($res -notmatch 'PASS' -or $res -match '(?m)^(FAIL|CRASH)') { $fail++ }
}
if ($fail) { Write-Output "`nมีการทดสอบไม่ผ่าน"; exit 1 } else { Write-Output "`nผ่านทั้งหมด" }
