// UI regression without a browser: real warning aggregation and click handlers.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert/strict');
const tool = path.resolve(__dirname, '..');
function element(tag = 'div') {
  const classes = new Set(), listeners = {};
  return { tagName: tag, children: [], listeners, style: {}, disabled: false, textContent: '',
    set innerHTML(value) { this.children = []; this.textContent = ''; },
    get className() { return [...classes].join(' '); },
    set className(value) { classes.clear(); value.split(/\s+/).filter(Boolean).forEach(c => classes.add(c)); },
    classList: { add: c => classes.add(c), remove: c => classes.delete(c), contains: c => classes.has(c),
      toggle(c, on) { if (on) classes.add(c); else classes.delete(c); } },
    appendChild(child) { this.children.push(child); },
    setAttribute(k, v) { this[k] = v; }, addEventListener(type, fn) { listeners[type] = fn; }
  };
}
const html = fs.readFileSync(path.join(tool, 'index.html'), 'utf8');
assert.match(html, /id="issues"[^>]*><\/div>\s*<\/section>\s*<div id="marked-downloads"/,
  'Marked downloads must immediately follow the warning panel');
assert(html.indexOf('id="marked-downloads"') < html.indexOf('id="fanpage"'));
const ids = new Map([...html.matchAll(/id="([^"]+)"/g)].map(m => [m[1], element()]));
const timers = [], downloads = [], builds = [];
function descendants(node) { return node.children.flatMap(c => [c, ...descendants(c)]); }
function markedButtons() { return descendants(ids.get('marked-downloads')).filter(n => n.tagName === 'button'); }
function textOf(node) { return node.textContent + node.children.map(textOf).join(' '); }
const ctx = vm.createContext({ console, setTimeout: fn => timers.push(fn), clearTimeout() {},
  document: { getElementById: id => { assert(ids.has(id), id); return ids.get(id); },
    createElement: element, querySelectorAll: () => markedButtons() }, addEventListener() {},
  Writer: { download: (bytes, name) => downloads.push({ bytes, name }) },
  Markup: { numToCol: n => String.fromCharCode(65 + n), build: (bytes, entry) => { builds.push(entry); return { bytes }; } }
});
ctx.window = ctx;
for (const file of ['config.js', 'parse.js', 'pipeline.js'])
  vm.runInContext(fs.readFileSync(path.join(tool, 'src', file), 'utf8'), ctx);
vm.runInContext(fs.readFileSync(path.join(tool, 'src/app.js'), 'utf8').replace(
  "document.addEventListener('DOMContentLoaded', init);",
  'globalThis.ui = {state, renderIssues, showTopIssue, markStale, clearStale, fileSignature, renderNeeds, applyFanpage, run, renderPatchReport, issueEl};'), ctx);
const { ui } = ctx, { state } = ui;
state.cfg = ctx.AppConfig.defaults();
state.files.ccvn = { name: 'cases.xlsx', bytes: new Uint8Array([1, 2, 3]) };
state.session = { warnings: [], droppedRows: [] };
state.blocks = { warnings: [] };
state.ranSignature = ui.fileSignature();
function resetWarnings() {
  state.uiWarnings = []; state.extraWarnings = []; state.fanpageWarnings = []; state.blocks.warnings = [];
}
resetWarnings();
state.extraWarnings.push({ level: 'error', message: 'Không tra được email', detail: 'Kiểm tra tài khoản dùng chung',
  display: {title:'1 tên chưa tìm được email', summary:'Case chưa được tính cho nhân viên. Kiểm tra bảng tên và email.'},
  spots: [{ file: 'cases.xlsx', sheet: 'Cases', row: 98, col: 6, value: 'CC-VN', why: 'Xác nhận người nhận case' }] });
state.blocks.warnings.push({ level: 'warn', message: 'Thiếu giờ làm', detail: 'KPI bằng 0, cần kiểm tra OP' });
state.fanpageWarnings.push({ level: 'warn', message: 'Thiếu follow', detail: 'Dán lại số' });
state.session.warnings = [{ level: 'info', message: 'Đã chuẩn hóa', detail: 'Giữ nguyên dữ liệu gốc' }];
ui.renderIssues();
assert(ids.get('warning-panel').classList.contains('has-errors'));
assert.match(ids.get('warning-summary').textContent, /1 lỗi · 2 cảnh báo/);
const firstBody = ids.get('issues').children[0].children[1];
assert.equal(firstBody.children[0].textContent, '1 tên chưa tìm được email');
assert.match(firstBody.children[1].textContent, /Kiểm tra bảng tên và email/);
const extraDetail = firstBody.children.find(c => c.tagName === 'details');
assert.equal(extraDetail.open, false, 'Only supporting details start collapsed, not the warning itself');
assert.match(textOf(extraDetail), /Không tra được email/);
assert.match(textOf(extraDetail), /Kiểm tra tài khoản dùng chung/);
assert.match(textOf(extraDetail), /G98/);
const sourceTable = descendants(extraDetail).find(n => n.tagName === 'table');
assert(sourceTable, 'Source locations must be presented as a table');
assert.match(textOf(sourceTable), /Nội dung ô trong file gốc/);
assert.match(textOf(sourceTable), /CC-VN/);
assert.match(textOf(sourceTable), /Xác nhận người nhận case/);
const payload = '<img src=x onerror=alert(1)>';
const tableProbe = {
  level:'warn', message:'Thông tin nguyên gốc', detail:'Giữ giải thích đầy đủ', spotTotal:70,
  display: {title:'Kiểm tra dữ liệu', summary:'Tóm tắt ngắn', impact:'KPI bằng 0', action:'Kiểm tra giờ làm',
    tables:[{title:'Nhân viên cần kiểm tra', headers:['Tên','Giờ làm'], rows:[[payload,0],['Tên dài', '']]}]},
  spots:[...Array.from({length:15},(_,i)=>({file:'A.xlsx',sheet:'Hours',row:i+1,col:6,value:i===0?'':0,why:payload})),
    {file:'B.xlsx',sheet:'Cases',row:2,col:3,value:'giá trị dài '.repeat(50),why:'Xác nhận dữ liệu'}]
};
const probeNode = ui.issueEl('warn', tableProbe.message, tableProbe.detail, tableProbe);
const probeNodes = descendants(probeNode);
assert.equal(probeNodes.filter(n=>n.tagName==='table').length,3, 'One staff table and one source table per file/sheet');
assert.equal(probeNodes.filter(n=>n.tagName==='caption').length,3);
assert(probeNodes.some(n=>n.tagName==='strong'&&n.textContent==='Ảnh hưởng'));
assert(probeNodes.some(n=>n.tagName==='strong'&&n.textContent==='Cần làm'));
assert.match(textOf(probeNode),/G15/,'Source locations after the old 12-cell limit must remain available');
assert.equal((textOf(probeNode).match(/16\/70/g)||[]).length,1,'Show the global cap once, not once per file');
assert(probeNodes.some(n=>n.tagName==='td'&&n.textContent==='0'&&!n.classList.contains('detail-empty')));
assert(probeNodes.some(n=>n.tagName==='td'&&n.textContent==='(trống)'&&n.classList.contains('detail-empty')));
assert(probeNodes.some(n=>n.tagName==='th'&&n.scope==='row'&&n.textContent===payload));
assert(!probeNodes.some(n=>['img','script'].includes(n.tagName)), 'Workbook text must never become HTML');
assert.match(textOf(probeNode),/Giữ giải thích đầy đủ/);
assert(probeNodes.filter(n=>n.className==='detail-table-scroll').every(n=>n.tabindex==='0'&&n.role==='region'&&n['aria-label']));
ctx.XLSX={utils:{encode_cell:({r,c})=>String.fromCharCode(65+c)+(r+1)}};
state.files.ops=[{name:'SBI.xlsx',wb:{Sheets:{Hours:{C2:{t:'n',v:46278,w:'13/09/2026'},M2:{t:'n',v:0,w:'0.00'}}}}}];
const formatted=ui.issueEl('info','Đã xử lý','', {level:'info',spots:[
  {file:'SBI.xlsx',sheet:'Hours',row:2,col:2,value:'46278',why:'Ngày trong file'},
  {file:'SBI.xlsx',sheet:'Hours',row:2,col:12,value:'0',why:'Giờ việc khác'}]});
assert.match(textOf(formatted),/13\/09\/2026/);assert.match(textOf(formatted),/0\.00/);
assert.doesNotMatch(textOf(formatted),/46278/);assert.match(textOf(formatted),/Các ô tool đã xử lý/);
state.files.ops=[];
const fold = ids.get('issues').children.find(c => c.tagName === 'details');
assert.equal(fold.open, true, 'Data warnings must be expanded by default');
assert.equal(ids.get('processing-notes').children[0].open, false);
assert.match(textOf(ids.get('issues')), /G98/);
let button = markedButtons()[0];
assert.equal(button.className, 'marked-download'); assert.equal(button.disabled, false);
assert.match(button['aria-label'], /cases.xlsx/);
button.listeners.click(); timers.shift()();
assert.equal(builds.length, 1); assert.equal(downloads[0].name, 'cases_DA-DANH-DAU.xlsx');
ui.markStale();
assert.equal(button.disabled, true);
timers.shift()();
assert.equal(button.disabled, true, 'The completion timer must not re-enable stale downloads');
button.listeners.click(); assert.equal(timers.length, 0);
ui.clearStale(); ui.renderIssues();
button = markedButtons()[0]; button.listeners.click();
const oldSignature = ui.fileSignature();
state.files.ccvn = { name: 'cases.xlsx', bytes: new Uint8Array([9]) };
assert.notEqual(ui.fileSignature(), oldSignature, 'Re-upload with the same name is still new input');
timers.shift()();
assert.equal(downloads.length, 1, 'Pending highlight download must not mix old warnings with new input');
assert.equal(state.stale, true);
ui.clearStale(); state.ranSignature = ui.fileSignature(); resetWarnings();
state.session.warnings = [];
state.fanpageText = 'SBIR 431722'; ui.applyFanpage();
assert.match(textOf(ids.get('issues')), /Chưa đọc được số follow/);
state.fanpageText = 'SBIR 431722\nSMILES 149874\nDCOM 393729'; ui.applyFanpage();
assert.doesNotMatch(textOf(ids.get('issues')), /Chưa đọc được số follow/);
assert.equal(state.blocks.fanpage.follows.DCOM, 393729);
assert(ids.get('warning-panel').classList.contains('is-clear'));
assert.match(ids.get('warning-summary').textContent, /chưa thể xác nhận/);
assert(ids.get('marked-downloads').classList.contains('hidden'));
assert.equal(markedButtons().length, 0, 'Clear old download controls when warnings disappear');
state.blocks = state.session = null;
ui.showTopIssue('Không đọc được file', 'Thử nạp lại');
assert(ids.get('warning-panel').classList.contains('has-errors'));
assert.match(textOf(ids.get('issues')), /Không đọc được file/);
assert.equal(markedButtons().length, 0);
ui.showTopIssue('Không đọc được file', 'Thử nạp lại');
assert.equal(state.uiWarnings.length, 1, 'Repeated failures must not inflate issue counts');
ui.renderPatchReport(null, 'Mẫu KPI có phạm vi tổng không thống nhất');
assert.match(textOf(ids.get('issues')), /Mẫu KPI có phạm vi tổng không thống nhất/,
  'Export template failures must also be visible in the prominent warning panel');
state.blocks = { warnings: [] };
state.files.session = {name:'broken.xlsx'};
ids.get('cfg-text').value = JSON.stringify(state.cfg);
ctx.Pipeline.processSession = () => { throw new Error('Dữ liệu bị hỏng'); };
ui.run(); timers.shift()();
assert.equal(state.blocks, null, 'Failed processing must invalidate previous output');
assert.equal(markedButtons().length, 0);
assert.equal(ids.get('metrics').children.length, 0);
assert(ids.get('sec-export').classList.contains('hidden'));
assert(!ids.get('sec-result').classList.contains('is-stale'), 'Failure warning must remain legible');
assert.match(textOf(ids.get('issues')), /Dữ liệu bị hỏng/);
console.log('PASS: prominent warning counts, expanded data warnings, adjacent yellow downloads, source locations, fanpage updates, stale/same-name input guards and empty/error states.');
