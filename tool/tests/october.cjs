const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert/strict');
const root = path.resolve(__dirname, '../..');
const ctx = vm.createContext({ console, Uint8Array, ArrayBuffer, TextEncoder, TextDecoder, Date, setTimeout });
for (const f of ['vendor/xlsx.full.min.js', 'vendor/fflate.min.js', 'src/config.js', 'src/parse.js', 'src/pipeline.js', 'src/writer.js', 'src/xlsxpatch.js']) {
  vm.runInContext(fs.readFileSync(path.join(root, 'tool', f), 'utf8'), ctx, { filename: f });
}
const { XLSX, AppConfig, Parse, Pipeline, XlsxPatch } = ctx;
const dir = path.join(root, 'Lệch dữ liệu 04.10', 'Dữ liệu gốc đưa vào');
function load(prefix) {
  const name = fs.readdirSync(dir).find(n => n.startsWith(prefix));
  const bytes = new Uint8Array(fs.readFileSync(path.join(dir, name)));
  return { name, bytes, wb: XLSX.read(bytes, { type: 'array', cellFormula: true }) };
}
const base = load('Altius'), ses = load('Session'), cc = load('Daily'), op = load('SBI');
const cfg = AppConfig.defaults(), serial = Parse.toDateSerial('2026-10-04');
cfg.op.acceptSwappedDates = true;
const session = Pipeline.processSession(ses.wb, cfg, ses.name);
const ccvn = Pipeline.processCcvn(cc.wb, cfg, cc.name);
const ops = [Pipeline.processOpFile(op.wb, cfg, serial, op.name)];
const mail = Pipeline.inspectMailSf(base.wb, cfg, []);
const blocks = Pipeline.buildBlocks({session, ccvn, ops, emailTable: mail.table}, cfg, serial);
function row(ws, r, cols) {
  return cols.map(c => { const x = ws[c+r]; return [c+r, x && x.v, x && x.f]; });
}
if (process.argv.includes('--summary')) {
  console.log('rows', blocks.block1.rows.length, blocks.block3.rows.length, 'topics', blocks.summary.topics, 'swapped', ops[0].stats);
  console.log('channels', blocks.summary.counts);
  console.log('drop',session.droppedRows);
  console.log('JP rows',blocks.block1.rows.filter(r => r[4] === 'bao.nq@altius-link.com.vn').reduce((a,r) => (a[r[2]]=(a[r[2]]||0)+1,a),{}));
} else if (process.argv.includes('--inspect')) {
  console.log('rows', blocks.block1.rows.length, blocks.block3.rows.length, 'topics', blocks.summary.topics, 'unknown', blocks.summary.unknownTopics);
  console.log('other source', ops[0].other.rows.slice(0,32).map(x => [x.excelRow,...x.cells.slice(2,13)]));
  for (const r of [4,5,6,20,28,29,70]) console.log('Sheet5', row(base.wb.Sheets.Sheet5,r,['A','B','C','D','E','F','G','H','I','J','O','P','Q','R','S','T','U','V']));
  console.log('JP',row(base.wb.Sheets['Daily KPI Result'],54,['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W']));
  console.log('topic3',Object.entries(blocks.summary.unknownTopics));
} else {
  const out = XlsxPatch.patch(base.bytes, blocks, cfg, {masterWb:base.wb});
  const written = XLSX.read(out.bytes, {type:'array', cellFormula:true});
  assert.equal(blocks.block1.rows.length,2440);
  assert.equal(blocks.block2.rows.length,348);
  assert.equal(blocks.block3.rows.length,176);
  assert.equal(written.Sheets['Summarize_team VietNam'].I36.v,621);
  const other = written.Sheets[cfg.targets.other.sheet];
  let checked = 0;
  for (let i=0;i<ops[0].other.rows.length;i++) {
    const raw = ops[0].other.rows[i].cells[12];
    const expected = raw === 0 ? '' : raw; // DataQuyTac: zero other-hours display as blank.
    assert.equal(other['M'+(i+6)]?.v ?? '',expected ?? '', 'Other hours row '+(i+6));
    checked++;
  }
  assert.equal(other.F6.f,'SUM(G6:M13)');
  assert.equal(other.M13.v,0.5);
  assert.equal(other.M21.v,1);
  assert.equal(other.M33.v,0.16);
  for(const name of ['Summarize_team VietNam','202610 KPI Result']) {
    const ws=base.wb.Sheets[name], got=written.Sheets[name];
    const hr=name.startsWith('Summarize')?4:5;
    for(let c=0;c<40;c++) {
      const col=XLSX.utils.encode_col(c), date=ws[col+hr]?.v;
      if(typeof date!=='number'||date<45000||date>=serial) continue;
      for(let r=6;r<=70;r++) {
        assert.equal(got[col+r]?.v,ws[col+r]?.v, name+' history '+col+r);
        assert.equal(got[col+r]?.f,ws[col+r]?.f, name+' historical formula '+col+r);
        checked++;
      }
    }
  }
  const legacy=AppConfig.defaults(); legacy.op.other.copyColumns.pop();
  assert.equal(AppConfig.migrate(legacy).op.other.copyColumns.length,11);
  delete legacy.summary.topicRows;
  legacy.op.acceptSwappedDates=true;
  const migrated=AppConfig.migrate(legacy);
  assert.equal(migrated.summary.topicRows.length,13,'Old config receives topic counting');
  assert.equal(migrated.op.acceptSwappedDates,true,'Preserve user date choice');
  const legacyBlocks=Pipeline.buildBlocks({session,ccvn,ops,emailTable:mail.table},legacy,serial);
  assert.equal(legacyBlocks.block3.rows[7][10],0.5,'Saved v1 config must copy M');
  // A second run with fewer rows must clear inline strings AND old M values.
  const small=Object.assign({},blocks,{block1:Object.assign({},blocks.block1,{rows:blocks.block1.rows.slice(0,1)}),
    block3:Object.assign({},blocks.block3,{rows:blocks.block3.rows.slice(0,8)})});
  const again=XlsxPatch.patch(out.bytes,small,cfg,{masterWb:written});
  const reread=XLSX.read(again.bytes,{type:'array'});
  assert.equal(reread.Sheets[cfg.targets.other.sheet].M21?.v,undefined);
  assert.equal(reread.Sheets[cfg.targets.cskh.sheet].C12?.v,undefined);
  const empty=Object.assign({},blocks,{block3:Object.assign({},blocks.block3,{rows:[]})});
  const cleared=XLSX.read(XlsxPatch.patch(out.bytes,empty,cfg,{masterWb:written}).bytes,{type:'array'});
  assert.equal(cleared.Sheets[cfg.targets.other.sheet].M13?.v,undefined);
  assert.equal(cleared.Sheets[cfg.targets.other.sheet].E6?.v,undefined);
  fs.mkdirSync(path.join(__dirname,'artifacts'), {recursive:true});
  fs.writeFileSync(path.join(__dirname,'artifacts','AltiuslinkVN-20261004.xlsx'),out.bytes);
  console.log('PASS: October 4 data, '+checked+' source/history cells, old config migration, repeated and empty exports.');
}
module.exports = {ctx,base,blocks,cfg,ops,serial};
