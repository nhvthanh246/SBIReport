// Real Chromium smoke test. Uses an isolated temporary profile, never the user's browser.
// Run: node tool/tests/browser-warnings.cjs [path-to-chrome-or-edge] [--dataquytac]
const fs = require('fs'), path = require('path'), os = require('os'), assert = require('assert/strict');
const { spawn } = require('child_process'), { pathToFileURL } = require('url');
const root = path.resolve(__dirname, '../..');
const dataQuyTac = process.argv.includes('--dataquytac');
const browser = process.argv.slice(2).find(a=>!a.startsWith('--')) || [
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'
].find(p => fs.existsSync(p));
assert(browser, 'Pass a local Chromium executable as the first argument');
const tempRoot = fs.realpathSync(os.tmpdir());
const profile = fs.mkdtempSync(path.join(tempRoot, 'excel-warning-ui-'));
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
let child, socket, sequence = 0, stderr = '';
const pending = new Map(), exceptions = [];
function call(method, params = {}) {
  return new Promise((resolve, reject) => {
    const id = ++sequence;
    const timer = setTimeout(() => { pending.delete(id); reject(new Error('CDP timeout: ' + method)); }, 20000);
    pending.set(id, { resolve, reject, timer });
    socket.send(JSON.stringify({ id, method, params }));
  });
}
async function evaluate(expression) {
  const result = await call('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
  assert(!result.exceptionDetails, JSON.stringify(result.exceptionDetails));
  return result.result.value;
}
async function until(expression) {
  const limit = Date.now() + 25000;
  while (Date.now() < limit) { if (await evaluate(expression)) return; await delay(100); }
  throw new Error('UI timeout: ' + expression);
}
(async () => {
  try {
    child = spawn(browser, ['--headless=new', '--remote-debugging-port=0', '--remote-debugging-address=127.0.0.1',
      '--user-data-dir=' + profile, '--no-first-run', '--no-default-browser-check', '--disable-extensions', 'about:blank'],
      { windowsHide: true, stdio: ['ignore', 'ignore', 'pipe'] });
    child.stderr.on('data', chunk => { stderr += chunk; });
    child.on('error', error => { stderr += error.message; });
    const limit = Date.now() + 12000;
    while (!stderr.includes('DevTools listening on ') && Date.now() < limit && child.exitCode === null) await delay(100);
    const match = stderr.match(/DevTools listening on (ws:\/\/[^\s]+)/);
    assert(match, 'Browser failed to start: ' + stderr.slice(-1800));
    const endpoint = new URL(match[1]);
    const pages = await (await fetch('http://' + endpoint.host + '/json/list')).json();
    socket = new WebSocket(pages.find(p => p.type === 'page').webSocketDebuggerUrl);
    await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
    socket.onmessage = event => {
      const message = JSON.parse(event.data);
      if (message.method === 'Runtime.exceptionThrown') exceptions.push(message.params);
      const waiter = pending.get(message.id);
      if (!waiter) return;
      pending.delete(message.id); clearTimeout(waiter.timer);
      if (message.error) waiter.reject(new Error(JSON.stringify(message.error))); else waiter.resolve(message.result);
    };
    await call('Runtime.enable'); await call('Page.enable');
    await call('Emulation.setDeviceMetricsOverride', { width: 1280, height: 1000, deviceScaleFactor: 1, mobile: false });
    await call('Page.navigate', { url: pathToFileURL(path.join(root, 'tool/index.html')).href });
    await until("document.readyState === 'complete' && !!document.getElementById('cfg-text').value");
    assert.equal(await evaluate("document.getElementById('start-time-source').value"),'Requested time');
    const directory = path.join(root, dataQuyTac ? 'DataQuyTac/File thô' : 'ExtraData/Daily Report 12.10.2026');
    const files = fs.readdirSync(directory).filter(n => n.endsWith('.xlsx') && !n.includes('20261012 -')).map(n => path.join(directory, n));
    if(dataQuyTac) files.push(path.join(root,'Data/test',fs.readdirSync(path.join(root,'Data/test')).find(n=>n.endsWith('20260912.xlsx'))));
    assert.equal(files.length, 4);
    const doc = await call('DOM.getDocument');
    const input = await call('DOM.querySelector', { nodeId: doc.root.nodeId, selector: '#dz-input' });
    await call('DOM.setFileInputFiles', { nodeId: input.nodeId, files });
    await until("!document.getElementById('btn-run').disabled && document.getElementById('busy').classList.contains('hidden')");
    await evaluate("document.getElementById('btn-run').click()");
    await until("!document.getElementById('sec-export').classList.contains('hidden') && document.getElementById('busy').classList.contains('hidden')");
    if(!dataQuyTac){
      for(const text of ['Đánh giá khách hàng','O1070','G2001','C2017','CC-VN'])
        assert.equal(await evaluate(`document.getElementById('issues').textContent.includes(${JSON.stringify(text)})`),true,text);
      assert.equal(await evaluate("document.getElementById('issues').textContent.includes('có thể đang gán sai case')"),false);
    }
    assert.equal(await evaluate("document.getElementById('warning-panel').nextElementSibling.id"), 'marked-downloads');
    assert.equal(await evaluate("document.querySelector('#issues .warning-list').open"), true);
    const missingEmail = dataQuyTac ? 'anh.np@altius-link.com.vn' : 'dao.pt@altius-link.com.vn';
    assert.equal(await evaluate(`document.getElementById('issues').textContent.includes(${JSON.stringify(missingEmail)})`), true);
    assert.equal(await evaluate(`document.getElementById('issues').innerText.includes(${JSON.stringify(missingEmail)})`), false,
      'Long staff lists must start collapsed, without discarding the evidence');
    assert.equal(await evaluate("document.getElementById('issues').innerText.includes('KPI đang bằng 0 dù có công việc')"), true,
      'The consequence and next action must remain visible');
    assert.equal(await evaluate("Array.from(document.querySelectorAll('#issues .issue-detail')).every(d => !d.open)"), true);
    await evaluate("window.__hoursIssue = Array.from(document.querySelectorAll('#issues .issue')).find(n=>n.querySelector('.it').textContent.includes('người chưa có giờ làm hợp lệ')); __hoursIssue.querySelector('.issue-detail summary').click()");
    assert.equal(await evaluate(`document.getElementById('issues').innerText.includes(${JSON.stringify(missingEmail)})`), true);
    const missingHoursCount = await evaluate("__hoursIssue.querySelectorAll('.detail-table tbody tr').length");
    assert.equal(missingHoursCount,dataQuyTac ? 6 : 5);
    assert.equal(await evaluate("__hoursIssue.querySelector('.detail-impact .detail-label').innerText"),'Ảnh hưởng');
    assert.equal(await evaluate("__hoursIssue.querySelector('.detail-action .detail-label').innerText"),'Cần làm');
    await evaluate("__hoursIssue.querySelector('.issue-detail summary').click()");
    await evaluate("document.querySelector('#issues .error .issue-detail summary').click()");
    assert.equal(await evaluate(`document.getElementById('issues').innerText.includes('${dataQuyTac?'C34':'G79'}')`), true);
    assert.equal(await evaluate("document.querySelectorAll('#issues .error .detail-spot-table tbody tr').length"),dataQuyTac ? 2 : 11);
    assert.equal(await evaluate("document.querySelector('#issues .error .detail-spot-table tbody td').innerText"),dataQuyTac ? 'NGUYEN THI KHANH LY' : 'CC-VN');
    if(dataQuyTac){
      assert.equal(await evaluate("document.querySelector('#issues .error').innerText.includes('ly.ntk1@altius-link.com.vn')"),true);
      assert.equal(await evaluate("document.querySelector('#issues .error').innerText.includes('Bổ sung mã case → email')"),true);
      assert.equal(await evaluate("JSON.parse(document.getElementById('cfg-text').value).rulesVersion"),6);
      assert.equal(await evaluate("JSON.parse(document.getElementById('cfg-text').value).approvedAdjustments.enabled"),false);
      assert.equal(await evaluate("Array.from(document.querySelectorAll('#processing-notes .detail-spot-table tbody tr')).filter(r=>r.textContent.includes('M2026')).some(r=>r.cells[1].textContent==='0.00')"),true);
    }
    assert.equal(await evaluate("parseInt(getComputedStyle(document.querySelector('#issues .error .detail-spot-table thead th')).fontWeight)>=700"),true);
    await evaluate("document.querySelector('#issues .error .issue-detail summary').click()");
    assert.equal(await evaluate("getComputedStyle(document.querySelector('.marked-download')).backgroundColor"), 'rgb(255, 213, 79)');
    await evaluate("window.__download = null; Writer.download = (bytes, name) => { window.__download = {bytes, name}; }");
    const sourceName = dataQuyTac ? 'File báo cáo tổng' : 'Daily Report CCVN';
    await evaluate(`Array.from(document.querySelectorAll('.marked-download')).find(b => b.getAttribute('aria-label').includes(${JSON.stringify(sourceName)})).click()`);
    await until('!!window.__download');
    const source = files.find(p => path.basename(p).startsWith(sourceName));
    const original64 = fs.readFileSync(source).toString('base64');
    const marked = await evaluate(`(() => {
      const original = XLSX.read(Uint8Array.from(atob(${JSON.stringify(original64)}), c => c.charCodeAt(0)), {type:'array', cellFormula:true, cellNF:true});
      const copy = XLSX.read(__download.bytes, {type:'array', cellFormula:true, cellNF:true});
      let compared = 0;
      for (const sheet of original.SheetNames) for (const [addr, cell] of Object.entries(original.Sheets[sheet])) {
        if (addr[0] === '!') continue;
        const other = copy.Sheets[sheet][addr];
        for (const key of ['v', 't', 'f', 'z', 'w']) if (JSON.stringify(cell[key]) !== JSON.stringify(other && other[key])) throw Error(sheet + '!' + addr + ' changed ' + key);
        compared++;
      }
      return {name:__download.name, firstSheet:copy.SheetNames[0], compared};
    })()`);
    assert.match(marked.name, /_DA-DANH-DAU\.xlsx$/); assert.equal(marked.firstSheet, 'Loi can sua');
    const artifact = path.join(__dirname, 'artifacts'); fs.mkdirSync(artifact, { recursive: true });
    for (const mode of ['light', 'dark-mobile']) {
      await call('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: mode === 'light' ? 'light' : 'dark' }] });
      if (mode === 'dark-mobile') await call('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true });
      await evaluate("document.getElementById('marked-downloads').scrollIntoView({block:'center'})");
      assert.equal(await evaluate('document.documentElement.scrollWidth <= window.innerWidth'), true, mode + ' must not overflow');
      const shot = await call('Page.captureScreenshot', { format: 'png' });
      fs.writeFileSync(path.join(artifact, 'warnings-' + mode + '.png'), Buffer.from(shot.data, 'base64'));
      // Also exercise expanded tables: keyboard scrolling, sticky headers and no page overflow.
      await evaluate("document.querySelectorAll('#issues .issue-detail').forEach(d=>d.open=true)");
      assert.equal(await evaluate('document.documentElement.scrollWidth <= window.innerWidth'),true,mode+' expanded tables must not overflow the page');
      assert.equal(await evaluate("getComputedStyle(document.querySelector('#issues .detail-table thead th')).position"),'sticky');
      await evaluate("__hoursIssue.querySelector('.detail-table-scroll').focus(); __hoursIssue.querySelector('.detail-impact').scrollIntoView({block:'start'})");
      if(mode==='dark-mobile'){
        assert.equal(await evaluate("__hoursIssue.querySelector('.detail-table-scroll').scrollWidth > __hoursIssue.querySelector('.detail-table-scroll').clientWidth"),true);
        await call('Input.dispatchKeyEvent',{type:'keyDown',key:'ArrowRight',code:'ArrowRight',windowsVirtualKeyCode:39});
        await call('Input.dispatchKeyEvent',{type:'keyUp',key:'ArrowRight',code:'ArrowRight',windowsVirtualKeyCode:39});
        await until("__hoursIssue.querySelector('.detail-table-scroll').scrollLeft > 0");
      }
      const detailShot=await call('Page.captureScreenshot',{format:'png'});
      fs.writeFileSync(path.join(artifact,'warning-details-'+mode+'.png'),Buffer.from(detailShot.data,'base64'));
      await evaluate("document.querySelectorAll('#issues .issue-detail').forEach(d=>d.open=false)");
    }
    // Replacing the same-named input must lock every old annotation download.
    await call('DOM.setFileInputFiles', { nodeId: input.nodeId, files: [source] });
    await until("!document.getElementById('stale-note').classList.contains('hidden')");
    assert.equal(await evaluate("Array.from(document.querySelectorAll('.marked-download')).every(b => b.disabled)"), true);
    assert.equal(exceptions.length, 0, JSON.stringify(exceptions));
    console.log('PASS: real Chromium / '+(dataQuyTac?'DataQuyTac with SBI v6':'October 12 with SBI v6')+'; warning placement, missing-hours notice, yellow CTA, light/390px dark layout, same-name upload lock; ' + marked.compared + ' annotated source cells unchanged.');
  } finally {
    if (socket) { socket.close(); socket = null; }
    if (child) child.kill();
    // Delete only the validated, uniquely created test profile under the temp directory.
    assert.equal(path.dirname(path.resolve(profile)), tempRoot);
    assert(path.basename(profile).startsWith('excel-warning-ui-'));
    try { fs.rmSync(profile, { recursive: true, force: true, maxRetries: 6, retryDelay: 100 }); } catch (_) { /* Chrome may still hold a profile file. */ }
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
