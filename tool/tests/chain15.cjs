const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert/strict');
const root = path.resolve(__dirname, '../..');
const ctx = vm.createContext({console,Uint8Array,ArrayBuffer,TextEncoder,TextDecoder,Date,setTimeout});
for(const f of ['vendor/xlsx.full.min.js','vendor/fflate.min.js','src/config.js','src/parse.js','src/pipeline.js','src/writer.js','src/xlsxpatch.js','src/selftest.js']) {
  vm.runInContext(fs.readFileSync(path.join(root,'tool',f),'utf8'),ctx,{filename:f});
}
const {XLSX,AppConfig,Parse,Pipeline,XlsxPatch,SelfTest}=ctx;
function load(folder,prefix){
  const dir=path.join(root,'Data',folder),name=fs.readdirSync(dir).find(n=>n.startsWith(prefix));
  const bytes=new Uint8Array(fs.readFileSync(path.join(dir,name)));
  return {name,bytes,wb:XLSX.read(bytes,{type:'array',cellFormula:true})};
}
const base=load('input/14092026','Altius'),ref=load('input/15092026','Altius');
const se=load('input/15092026','Session'),sf=load('input/15092026','Daily'),op=load('input','SBI');
// Historical output assertions use an explicit old business profile, not current defaults.
const cfg=require('./fixtures/rules-v4.cjs')(AppConfig.defaults()),serial=Parse.toDateSerial('2026-09-15');
const session=Pipeline.processSession(se.wb,cfg,se.name),ccvn=Pipeline.processCcvn(sf.wb,cfg,sf.name);
const ops=[Pipeline.processOpFile(op.wb,cfg,serial,op.name)];
const mail=Pipeline.inspectMailSf(base.wb,cfg,[]);
const blocks=Pipeline.buildBlocks({session,ccvn,ops,emailTable:mail.table},cfg,serial);
const out=XlsxPatch.patch(base.bytes,blocks,cfg,{masterWb:base.wb});
const wb=XLSX.read(out.bytes,{type:'array',cellFormula:true});
const checks=SelfTest.compareWithReference(out.bytes,ref.wb,blocks,cfg,{baseWb:base.wb,
  oddRateEmails:out.oddRateEmails,affectedEmails:mail.affected,input:{ccvn,ops}});
assert.equal(session.keptRows.length,1997);
assert.equal(blocks.block1.rows.length,2741);
assert.equal(blocks.block2.rows.length,548);
assert.equal(blocks.block3.rows.length,176);
assert.equal(mail.corrections.length,0);
session.keptRows.forEach((row,i)=>{
  const addr='F'+(i+11);
  assert.equal(wb.Sheets['1. CSKH'][addr].v,row[Parse.findColumn(session.headers,'First response time')]);
  assert.equal(wb.Sheets['1. CSKH'][addr].v,ref.wb.Sheets['1. CSKH'][addr].v);
});
assert.match(checks[0].detail,/^1 ô lệch/,'Only source topic differs, no Start Time mismatches');
assert.match(checks[2].detail,/tool 176 dòng, bản chuẩn 192 dòng/);
assert.match(checks[4].detail,/5\/39 ô lệch/);
let history=0;
for(const name of [cfg.summary.sheet,'202609 KPI Result']){
  const hr=name===cfg.summary.sheet?4:5;
  for(let c=0;c<40;c++){
    const col=XLSX.utils.encode_col(c),d=base.wb.Sheets[name][col+hr]?.v;
    if(typeof d!=='number'||d<45000||d>=serial)continue;
    for(let r=6;r<=70;r++){
      assert.equal(wb.Sheets[name][col+r]?.v,base.wb.Sheets[name][col+r]?.v);
      assert.equal(wb.Sheets[name][col+r]?.f,base.wb.Sheets[name][col+r]?.f);
      history++;
    }
  }
}
assert.equal(history,1820);
assert.equal(checks.find(c=>c.name.startsWith('8.')).status,'pass');
// Reference inputs are used only for this diagnostic, never production output.
const fixtureBlocks={...blocks};
for(const key of ['block1','block2','block3']){
  const b=blocks[key],a=XLSX.utils.decode_cell(b.anchor);
  const rows=Parse.sheetToMatrix(ref.wb.Sheets[b.sheet]).rows.slice(a.r)
    .map(r=>Array.from({length:b.headers.length},(_,c)=>r[a.c+c]??''));
  while(rows.length&&rows.at(-1).every(v=>v===''))rows.pop();
  fixtureBlocks[key]={...b,rows};
}
fixtureBlocks.kpi=Pipeline.computeKpi(fixtureBlocks,cfg);
const fixture=XlsxPatch.patch(base.bytes,fixtureBlocks,cfg,{masterWb:base.wb});
const diagnostic=SelfTest.compareWithReference(fixture.bytes,ref.wb,fixtureBlocks,cfg,{baseWb:base.wb});
assert.equal(diagnostic[4].status,'pass',JSON.stringify(diagnostic[4]));
fs.mkdirSync(path.join(__dirname,'artifacts'),{recursive:true});
fs.writeFileSync(path.join(__dirname,'artifacts','AltiuslinkVN-20260915.xlsx'),out.bytes);
if(process.argv.includes('--compare'))console.log(JSON.stringify(checks,null,2));
console.log('PASS: day 15 - 1997 First response start times, 1820 history cells, pivot filters, all 39 reference-input KPI results. Remaining source differences are reported.');
