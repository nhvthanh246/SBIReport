const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert/strict');
const root = path.resolve(__dirname, '../..');
const ctx = vm.createContext({console,Uint8Array,ArrayBuffer,TextEncoder,TextDecoder,Date,setTimeout});
for (const f of ['vendor/xlsx.full.min.js','vendor/fflate.min.js','src/config.js','src/parse.js','src/pipeline.js','src/writer.js','src/xlsxpatch.js','src/selftest.js']) {
  vm.runInContext(fs.readFileSync(path.join(root,'tool',f),'utf8'),ctx,{filename:f});
}
const {XLSX,AppConfig,Parse,Pipeline,XlsxPatch,SelfTest} = ctx;
function load(folder,prefix) {
  const dir=path.join(root,'Data',folder), name=fs.readdirSync(dir).find(n=>n.startsWith(prefix));
  const bytes=new Uint8Array(fs.readFileSync(path.join(dir,name)));
  return {name,bytes,wb:XLSX.read(bytes,{type:'array',cellFormula:true})};
}
const base=load('input/13092026','File báo cáo tổng');
const ref=load('input/14092026','Altius');
const sf=load('input/14092026','Daily');
const se=load('input/14092026','Session');
const op=load('input','SBI');
// Historical output assertions use an explicit old business profile, not current defaults.
const cfg=require('./fixtures/rules-v4.cjs')(AppConfig.defaults()), serial=Parse.toDateSerial('2026-09-14');
const session=Pipeline.processSession(se.wb,cfg,se.name), ccvn=Pipeline.processCcvn(sf.wb,cfg,sf.name);
const ops=[Pipeline.processOpFile(op.wb,cfg,serial,op.name)];
const mail=Pipeline.inspectMailSf(base.wb,cfg,[]);
const blocks=Pipeline.buildBlocks({session,ccvn,ops,emailTable:mail.table},cfg,serial);
const out=XlsxPatch.patch(base.bytes,blocks,cfg,{masterWb:base.wb});
const wb=XLSX.read(out.bytes,{type:'array',cellFormula:true});
if(process.argv.includes('--inspect')) {
  const ws=ref.wb.Sheets[cfg.targets.sfcase.sheet];
  const diffs={};
  blocks.block2.rows.forEach((row,i)=>row.forEach((v,c)=>{
    const exp=ws[XLSX.utils.encode_cell({r:i+6,c:c+1})]?.v;
    if(String(v??'').trim().toLowerCase()===String(exp??'').trim().toLowerCase())return;
    const k=blocks.block2.headers[c]; const d=diffs[k]??={count:0,samples:[]}; d.count++;
    if(d.samples.length<3)d.samples.push({cell:XLSX.utils.encode_cell({r:i+6,c:c+1}),value:v,expected:exp});
  }));
  console.log('SF differences',JSON.stringify(diffs,null,2));
  console.log('Lookup base/ref', [base,ref].map(f=>({file:f.name,rows:Parse.sheetToMatrix(f.wb.Sheets['mail SF']).rows.map((r,i)=>({row:i+1,name:r[2],email:r[3]})).filter(r=>String(r.email).includes('ly.ntk'))})));
  const rws=ref.wb.Sheets[cfg.targets.other.sheet];
  console.log('Hours differences',ops[0].other.rows.flatMap((rec,i)=>{
    const dif=[];
    for(let c=3;c<11;c++){
      const v=rec.cells[c+2],e=rws[XLSX.utils.encode_cell({r:i+5,c:c+2})]?.v;
      if(typeof v==='number'&&typeof e==='number'&&Math.abs(v-e)<1e-9)continue;
      if(String(v??'')!==String(e??''))dif.push({sourceRow:rec.excelRow,email:rec.cells[4],col:XLSX.utils.encode_col(c+2),outRow:i+6,v,expected:e});
    }return dif;
  }));
  const refHours=Parse.sheetToMatrix(rws).rows.slice(5).filter(r=>r[4]&&Parse.toDateSerial(r[2])===serial);
  console.log('Hours reference count',refHours.length,'missing people',Array.from(new Set(refHours.map(r=>r[4]))).filter(e=>!blocks.block3.rows.some(r=>r[2]===e)));
  console.log('Reference summary', ['S11','S13','S14','S16'].map(a=>[a,ref.wb.Sheets[cfg.summary.sheet][a]]));
  console.log('KPI29',ref.wb.Sheets['202609 KPI Result'].D29, blocks.kpi.forEmail(ref.wb.Sheets['202609 KPI Result'].D29.v));
}
const checks=SelfTest.compareWithReference(out.bytes,ref.wb,blocks,cfg,{baseWb:base.wb,oddRateEmails:out.oddRateEmails,
  affectedEmails:mail.affected,input:{ccvn,ops}});
assert.equal(session.keptRows.length,1929);
const firstResponse=Parse.findColumn(session.headers,'First response time');
session.keptRows.forEach((row,i)=>{
  assert.equal(wb.Sheets['1. CSKH']['F'+(i+11)].v,row[firstResponse]);
  assert.notEqual(wb.Sheets['1. CSKH']['F'+(i+11)].v,ref.wb.Sheets['1. CSKH']['F'+(i+11)].v);
});
const requestedCfg=AppConfig.defaults();
requestedCfg.session.stagingColumns[3]={out:'Requested time',from:'Requested time'};
const requestedSession=Pipeline.processSession(se.wb,AppConfig.migrate(requestedCfg),se.name);
requestedSession.staging.rows.forEach((row,i)=>assert.equal(row[3],ref.wb.Sheets['1. CSKH']['F'+(i+11)].v));
assert.equal(mail.corrections.length,1);
assert.equal(wb.Sheets['mail SF'].C43.v,'NGUYEN THI KHANH LY (3M)');
assert.equal(base.wb.Sheets['mail SF'].C43.v,'NGUYEN THI KHANH LY','Do not mutate base workbook');
assert.equal(Pipeline.mailSfCorrections(wb,cfg).length,0,'Correction is idempotent');
const conflict=JSON.parse(JSON.stringify(base.wb));
conflict.Sheets['mail SF'].C34.v='NGUYEN THI KHANH LY (3M)';
assert.equal(Pipeline.mailSfCorrections(conflict,cfg).length,0,'Never overwrite conflicting identity');
let correctedEmails=0, statusDifferences=0, supersededEmails=0;
blocks.block2.rows.forEach((row,i)=>{
  const r=i+7;
  if ((wb.Sheets[cfg.targets.sfcase.sheet]['G'+r]?.v??'') !== (ref.wb.Sheets[cfg.targets.sfcase.sheet]['G'+r]?.v??'')) {
    // October 11-12 supersede the three blanket account assignments.
    assert(['PHAM THI THU HOAI','VU NHU QUYNH','LUU NGOC QUANG'].includes(Parse.normName(row[4])));
    supersededEmails++;
  }
  if(row[5]==='ly.ntk1@altius-link.com.vn') correctedEmails++;
  const status=ccvn.dataRows[i][Parse.findColumn(ccvn.headers,'Case Status')];
  assert.equal(row[2],status,'Status must come from actual Salesforce export');
  if(row[2]!==ref.wb.Sheets[cfg.targets.sfcase.sheet]['D'+r]?.v)statusDifferences++;
});
assert.equal(correctedEmails,18);
assert.equal(supersededEmails,51);
assert.equal(statusDifferences,127);
assert.match(checks[0].detail,/tool 2576 dòng, bản chuẩn 2577 dòng/);
assert.match(checks[1].detail,/230 ô lệch/,'128 source cells + 102 name/email cells from superseded assignment rules');
assert.match(checks[2].detail,/tool 160 dòng, bản chuẩn 192 dòng/);
assert.match(checks[4].detail,/6\/39 ô lệch/,'Do not excuse missing hours or duplicate names');
assert.equal(checks.find(c=>c.name.startsWith('6.')).status,'pass');
assert.equal(checks.find(c=>c.name.startsWith('8.')).status,'pass');
let historyCells=0;
for(const name of [cfg.summary.sheet,'202609 KPI Result']){
  const hdr=name===cfg.summary.sheet?4:5;
  for(let c=0;c<40;c++){
    const col=XLSX.utils.encode_col(c),d=base.wb.Sheets[name][col+hdr]?.v;
    if(typeof d!=='number'||d<45000||d>=serial)continue;
    for(let r=6;r<=70;r++){
      assert.equal(wb.Sheets[name][col+r]?.v,base.wb.Sheets[name][col+r]?.v);
      assert.equal(wb.Sheets[name][col+r]?.f,base.wb.Sheets[name][col+r]?.f);
      historyCells++;
    }
  }
}
const legacy=AppConfig.defaults();
legacy.rulesVersion=5;
delete legacy.session.startTimeMappingVersion;
legacy.session.stagingColumns[3]={out:'Requested time',from:'Requested time'};
delete legacy.mailSf.nameCorrections; delete legacy.kpi.setsumekaiWeight;
const migrated=AppConfig.migrate(legacy);
assert.equal(migrated.session.stagingColumns[3].from,'Requested time');
assert.equal(legacy.session.stagingColumns[3].from,'Requested time');
assert.equal(AppConfig.migrate(requestedCfg).session.stagingColumns[3].from,'Requested time','Preserve explicit v2 choice');
assert.equal(AppConfig.migrate(migrated).session.stagingColumns[3].from,'Requested time','Migration is idempotent');
assert.equal(migrated.mailSf.nameCorrections.length,0);
assert.equal(migrated.kpi.setsumekaiWeight,45);
legacy.session.stagingColumns[3]={out:'Custom start',from:'Assign time'};
assert.equal(AppConfig.migrate(legacy).session.stagingColumns[3].from,'Requested time','SBI confirmation supersedes pre-v6 Start Time choices');
// Comparison regression: strings sharing a numeric prefix must NOT compare equal.
const altered=JSON.parse(JSON.stringify(wb));
const cskh=altered.Sheets[cfg.targets.cskh.sheet];
cskh.F11={t:'s',v:'2026-09-14 00:00:00+07:00'};
cskh.L11={t:'s',v:'3. Different topic'};
cskh.M11={t:'s',v:'12345678901234567'};
const actual=JSON.parse(JSON.stringify(wb));
actual.Sheets[cfg.targets.cskh.sheet].L11={t:'s',v:'3. Original topic'};
actual.Sheets[cfg.targets.cskh.sheet].M11={t:'s',v:'12345678901234568'};
altered.Sheets[cfg.targets.sfcase.sheet].G7={t:'s',v:'ly.ntk1@altius-link.com.vn'};
actual.Sheets[cfg.targets.sfcase.sheet].G7={t:'s',v:'ly.ntk@altius-link.com.vn'};
const strict=SelfTest.compareWithReference(XLSX.write(actual,{type:'array',bookType:'xlsx'}),altered,blocks,cfg);
assert.match(strict[0].detail,/3 ô lệch/);
assert.match(strict[1].detail,/1 ô lệch/);
// Diagnostic fixture ONLY: with the reference's actual rows/hours, all 39 KPI
// results must be reproducible. Never merge reference data into production input.
const referenceBlocks={...blocks};
for(const key of ['block1','block2','block3']){
  const block=blocks[key],a=XLSX.utils.decode_cell(block.anchor);
  const rows=Parse.sheetToMatrix(ref.wb.Sheets[block.sheet]).rows.slice(a.r)
    .map(r=>Array.from({length:block.headers.length},(_,c)=>r[a.c+c]??''));
  while(rows.length&&rows.at(-1).every(v=>v===''))rows.pop();
  referenceBlocks[key]={...block,rows};
}
referenceBlocks.kpi=Pipeline.computeKpi(referenceBlocks,cfg);
const fixture=XlsxPatch.patch(base.bytes,referenceBlocks,cfg,{masterWb:base.wb});
const fixtureWb=XLSX.read(fixture.bytes,{type:'array'});
const refKpiCheck=SelfTest.compareWithReference(fixture.bytes,ref.wb,referenceBlocks,cfg,{baseWb:base.wb})[4];
assert.equal(refKpiCheck.status,'pass',JSON.stringify(refKpiCheck));
assert.equal(referenceBlocks.kpi.forEmail('anh.bn1@altius-link.com.vn').E,201);
assert.equal(referenceBlocks.kpi.forEmail('anh.bn1@altius-link.com.vn').W,2.204761904761905);
assert.equal(fixtureWb.Sheets['202609 KPI Result'].R7.v,1.3095795162459634);
if (require.main === module) {
  if(process.argv.includes('--compare'))console.log(JSON.stringify(checks,null,2));
  console.log('PASS: day 14 - 1929 start times and Requested-time override, 18 corrected emails; 51 assignments differ intentionally under October 11-12 rules; '+historyCells+' history cells, strict comparison, all 39 reference-input KPI results.');
}
fs.mkdirSync(path.join(__dirname,'artifacts'),{recursive:true});
fs.writeFileSync(path.join(__dirname,'artifacts','AltiuslinkVN-20260914.xlsx'),out.bytes);
fs.writeFileSync(path.join(__dirname,'artifacts','chain14-reference-fixture.xlsx'),fixture.bytes);
module.exports={ctx,base,ref,blocks,cfg,serial,ops,out,wb,checks,session,ccvn,mail};
