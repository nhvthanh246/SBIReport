// Regression tests for numeric hours, template-defined team membership and aliases.
// Reference data is used only to assert results and in the explicitly named fixture.
const fs=require('fs'),path=require('path'),vm=require('vm'),assert=require('assert/strict');
const root=path.resolve(__dirname,'../..');
const ctx=vm.createContext({console,Uint8Array,ArrayBuffer,TextEncoder,TextDecoder,Date,setTimeout});
for(const f of ['vendor/xlsx.full.min.js','vendor/fflate.min.js','src/config.js','src/parse.js','src/pipeline.js','src/writer.js','src/xlsxpatch.js','src/selftest.js'])
  vm.runInContext(fs.readFileSync(path.join(root,'tool',f),'utf8'),ctx,{filename:f});
const {XLSX,AppConfig,Parse,Pipeline,XlsxPatch,SelfTest}=ctx;
const dir=path.join(root,'ExtraData','Daily Report 12.10.2026');
function load(prefix){
  const names=fs.readdirSync(dir).filter(n=>n.startsWith(prefix)&&n.endsWith('.xlsx'));
  assert.equal(names.length,1,prefix);
  const name=names[0],bytes=new Uint8Array(fs.readFileSync(path.join(dir,name)));
  return {name,bytes,wb:XLSX.read(bytes,{type:'array',cellFormula:true})};
}
const base=load('AltiuslinkVN_Daily Report_(New)-20261011'),ref=load('AltiuslinkVN_Daily Report_(New)-20261012');
// Historical output assertions use an explicit old business profile, not current defaults.
const se=load('Session'),sf=load('Daily'),op=load('SBI'),cfg=require('./fixtures/rules-v4.cjs')(AppConfig.defaults()),serial=Parse.toDateSerial('2026-10-12');
const session=Pipeline.processSession(se.wb,cfg,se.name),ccvn=Pipeline.processCcvn(sf.wb,cfg,sf.name);
const ops=[Pipeline.processOpFile(op.wb,cfg,serial,op.name)],mail=Pipeline.inspectMailSf(base.wb,cfg,[]);
const input={session,ccvn,ops,emailTable:mail.table},blocks=Pipeline.buildBlocks(input,cfg,serial);
// Missing hours must be visible before export, without importing another day's hours.
const dataBeforeWarnings=JSON.stringify([blocks.block1.rows,blocks.block2.rows,blocks.block3.rows]);
const rosterWarnings=[];
Pipeline.checkRoster({emailTable:mail.table},cfg,base.wb,blocks,rosterWarnings);
const noHours=rosterWarnings.find(w=>w.message.includes('thiếu giờ hợp lệ'));
assert(noHours,'Missing KPI hours must be warned during processing');
assert.match(noHours.detail,/dao\.pt@altius-link\.com\.vn/);
assert.match(noHours.detail,/2026-10-12/);
assert.equal(noHours.level,'warn');
assert.match(noHours.display.title,/người chưa có giờ làm hợp lệ/);
assert.match(noHours.display.summary,/KPI đang bằng 0/);
assert(!noHours.display.summary.includes('@'),'Keep the long staff list in the original detail');
assert.equal(noHours.display.tables[0].rows.length,4);
const daoHours=noHours.display.tables[0].rows.find(row=>row[1]==='dao.pt@altius-link.com.vn');
assert(daoHours,'Missing-hours table must include the affected staff');
assert.equal(daoHours[3],blocks.kpi.forEmail(daoHours[1]).hours,'Show the actual KPI hours, including zero');
assert.match(noHours.display.action,/ngày khác/);
const missingEmail=blocks.warnings.find(w=>w.message.includes('tên không tra được email'));
assert(missingEmail.display.summary.includes('từng case'),'Shared-account caveat stays in the short warning');
assert.equal(missingEmail.spots.length,11,'Concise wording must preserve every source location');
assert.equal(JSON.stringify(missingEmail.display.tables[0].rows),JSON.stringify([['CC-VN',11,'Chưa tìm thấy']]));
assert.match(missingEmail.display.impact,/đang trống/);
for(const warning of [...blocks.warnings,...rosterWarnings,...session.warnings,...ops[0].warnings]){
  if(!warning.display)continue;
  assert(warning.display.title.length<100,warning.display.title);
  assert(warning.display.summary.length<200,warning.display.summary);
}
assert.equal(JSON.stringify([blocks.block1.rows,blocks.block2.rows,blocks.block3.rows]),dataBeforeWarnings,
  'Data diagnostics must not alter the calculated blocks');
const out=XlsxPatch.patch(base.bytes,blocks,cfg,{masterWb:base.wb});
const wb=XLSX.read(out.bytes,{type:'array',cellFormula:true});
const checks=SelfTest.compareWithReference(out.bytes,ref.wb,blocks,cfg,{baseWb:base.wb,oddRateEmails:out.oddRateEmails,affectedEmails:mail.affected,input});
const other=cfg.targets.other.sheet;
assert.equal(blocks.appliedAdjustments.length,0,'Do not invent October 12 exceptions');
assert.equal(blocks.block3.rows.length,144,'Do not silently import Dao hours dated October 11');
assert.equal(wb.Sheets[other].G94.t,'n');
assert.equal(wb.Sheets[other].G94.v,6.12);
assert(Math.abs(wb.Sheets[other].F94.v-ref.wb.Sheets[other].F86.v)<1e-9);
assert.equal(wb.Sheets[other].F94.f,'SUM(G94:M101)');
assert.equal(wb.Sheets['Daily KPI Result'].F22.f,'IFERROR(VLOOKUP($D22,Sheet5!$A$74:$B$132,2,0),0)');
assert.equal(base.wb.Sheets['Daily KPI Result'].F22.f,'IFERROR(VLOOKUP($D22,Sheet5!$A$74:$B$132,2,0),"")','Base formula stays unchanged');
assert.equal(op.wb.Sheets[other].G1977.v,'6.12','Original workbook unchanged');
assert.equal(ops[0].other.rows[88].cells[6],'6.12','Original parsed rows unchanged');
assert(blocks.warnings.some(w=>w.spots?.some(s=>s.row===1977&&s.col===6)));
const quang=blocks.block2.rows.filter(r=>Parse.normName(r[4])==='LUU NGOC QUANG');
assert.equal(quang.length,4);
for(const r of quang){
  const index=blocks.block2.rows.indexOf(r)+7;
  assert.equal(r[5],ref.wb.Sheets[cfg.targets.sfcase.sheet]['G'+index].v);
}
assert.deepEqual(Array.from(out.teamRows),Array.from({length:28},(_,i)=>i+15));
assert.equal(out.teamEmails.length,28);
assert(!out.teamEmails.includes('linh.dp1@altius-link.com.vn'));
assert.equal(wb.Sheets['202610 KPI Result'].P7.v,blocks.kpi.team(out.teamEmails));
const linhRow=Object.keys(wb.Sheets['202610 KPI Result']).find(a=>/^D\d+$/.test(a)&&wb.Sheets['202610 KPI Result'][a].v==='linh.dp1@altius-link.com.vn').slice(1);
assert.equal(wb.Sheets['202610 KPI Result']['P'+linhRow].v,blocks.kpi.forEmail('linh.dp1@altius-link.com.vn').W,'Excluded from team still gets individual KPI');
assert.notEqual(blocks.kpi.team(out.teamEmails),blocks.kpi.team([...out.teamEmails,'linh.dp1@altius-link.com.vn']));
const noFormulas=JSON.parse(JSON.stringify(base.wb));
for(const s of Object.values(noFormulas.Sheets))for(const cell of Object.values(s))if(cell&&typeof cell==='object')delete cell.f;
const noFormulaOut=XlsxPatch.patch(base.bytes,blocks,cfg,{masterWb:noFormulas});
assert.deepEqual(Array.from(noFormulaOut.teamRows),Array.from(out.teamRows),'Read SUM from XML even when UI does not parse formulas');
const duplicateRoster=JSON.parse(JSON.stringify(base.wb));
duplicateRoster.Sheets[cfg.kpi.daily.sheet].D16={...duplicateRoster.Sheets[cfg.kpi.daily.sheet].D15};
assert.throws(()=>XlsxPatch.patch(base.bytes,blocks,cfg,{masterWb:duplicateRoster}),/email bị trùng/);
assert.match(checks[1].detail,/^60 ô lệch/,'30 remaining assignments are not guessed');
assert.match(checks[0].detail,/^1634 ô lệch/,'Do not change Start Time or source descriptions to hide differences');
assert.equal(checks.find(c=>c.name.startsWith('6.')).status,'pass');
assert.equal(checks.find(c=>c.name.startsWith('8.')).status,'pass');
let history=0;
for(const name of [cfg.summary.sheet,'202610 KPI Result']){
  const header=name===cfg.summary.sheet?4:5;
  for(let c=0;c<40;c++){
    const col=XLSX.utils.encode_col(c),day=base.wb.Sheets[name][col+header]?.v;
    if(typeof day!=='number'||day<45000||day>=serial)continue;
    for(let r=6;r<=70;r++){
      assert.equal(wb.Sheets[name][col+r]?.v,base.wb.Sheets[name][col+r]?.v);
      assert.equal(wb.Sheets[name][col+r]?.f,base.wb.Sheets[name][col+r]?.f);history++;
    }
  }
}
assert.equal(history,1430);
// Strict conversion: only hours, full numeric strings, finite values, no guessed units.
for(const [text,want] of [['6.12',6.12],[' 0 ',0],['.5',.5],['+1.25',1.25],['1e-2',.01]])assert.equal(Parse.normalizeHours(text),want);
for(const text of ['', ' ', '`', '6 giờ','6,12','6:12','1,234.56','NaN','Infinity','1e999','0x10','=6.12'])assert.equal(Parse.normalizeHours(text),text);
const probeOps=[{...ops[0],other:{...ops[0].other,rows:ops[0].other.rows.map(r=>({...r,cells:r.cells.slice()}))}}];
probeOps[0].other.rows[88].cells[6]='6 giờ';
probeOps[0].other.rows[95].cells[12]='1.5';
probeOps[0].other.rows[88].cells[11]='6.12'; // task label must remain text
const probe=Pipeline.buildBlocks({...input,ops:probeOps},cfg,serial);
assert.equal(probe.block3.rows[88][4],'6 giờ');
assert.equal(probe.block3.rows[88][9],'6.12');
assert.equal(probe.block3.rows[95][10],1.5);
assert(Math.abs(probe.block3.rows[88][3]-(.882352941176471+1.5))<1e-9);
assert.equal(probe.kpi.p3['cuc.ht@altius-link.com.vn'].cskh,0,'Do not parse a numeric prefix from text');
assert(probe.warnings.some(w=>w.message.includes('ô giờ không phải số')));
// Derive team membership from all eleven SUM formulas, not a fixed row boundary.
const cols=['E','F','I','K','L','N','O','Q','R','T','U'];
function sums(end){return Object.fromEntries(cols.map(c=>[c,`SUM($${c}$15:$${c}$${end})`]));}
assert.equal(XlsxPatch.teamRowsFromFormulas(sums(43),cfg.kpi.daily).length,29);
assert.equal(XlsxPatch.teamRowsFromFormulas(sums(42),cfg.kpi.daily).length,28);
const noncontiguous=Object.fromEntries(cols.map(c=>[c,`SUM(${c}15:${c}20,${c}22:${c}42)`]));
assert(!XlsxPatch.teamRowsFromFormulas(noncontiguous,cfg.kpi.daily).includes(21));
assert.throws(()=>XlsxPatch.teamRowsFromFormulas({...sums(42),F:'SUM(F15:F43)'},cfg.kpi.daily),/không cùng phạm vi/);
for(const bad of ['', 'SUBTOTAL(9,E15:E42)', 'SUM(Other!E15:E42)', 'SUM(E15:E42)+1', 'SUM(E15:E42,E42)', 'SUM(E14:E42)', 'SUM(E15:E999999)'])
  assert.throws(()=>XlsxPatch.teamRowsFromFormulas({...sums(42),E:bad},cfg.kpi.daily),/KPI nhóm/);
const legacy=AppConfig.clone(cfg);legacy.rulesVersion=3;legacy.requestedByAliases['LUU NGOC QUANG']='NGUYEN THAO NGUYEN';
legacy.op.acceptSwappedDates=false;legacy.session.stagingColumns[3]={out:'Requested time',from:'Requested time'};
const migrated=AppConfig.migrate(legacy);
assert.equal(migrated.rulesVersion,6);assert.equal(migrated.requestedByAliases['LUU NGOC QUANG'],'NGUYEN THAO NGUYEN');
assert.equal(migrated.op.acceptSwappedDates,false);assert.equal(migrated.session.stagingColumns[3].from,'Requested time');
assert.equal(legacy.requestedByAliases['LUU NGOC QUANG'],'NGUYEN THAO NGUYEN');
legacy.requestedByAliases['LUU NGOC QUANG']='CUSTOM';
assert.equal(AppConfig.migrate(legacy).requestedByAliases['LUU NGOC QUANG'],'CUSTOM');
const explicit=AppConfig.clone(cfg);explicit.requestedByAliases['LUU NGOC QUANG']='NGUYEN THAO NGUYEN';
assert.equal(AppConfig.migrate(explicit).requestedByAliases['LUU NGOC QUANG'],'NGUYEN THAO NGUYEN','Preserve deliberate v4 customizations');
// With the reference's actual assignments/hours, verify all 40 KPI results.
// This diagnostic fixture is NOT the production output and never edits source inputs.
const referenceBlocks={...blocks};
for(const key of ['block1','block2','block3']){
  const b=blocks[key],a=XLSX.utils.decode_cell(b.anchor);
  const rows=Parse.sheetToMatrix(ref.wb.Sheets[b.sheet]).rows.slice(a.r)
    .map(r=>Array.from({length:b.headers.length},(_,c)=>r[a.c+c]??''));
  while(rows.length&&rows.at(-1).every(v=>v===''))rows.pop();
  referenceBlocks[key]={...b,rows};
}
referenceBlocks.kpi=Pipeline.computeKpi(referenceBlocks,cfg);
const fixture=XlsxPatch.patch(base.bytes,referenceBlocks,cfg,{masterWb:base.wb});
const diagnostic=SelfTest.compareWithReference(fixture.bytes,ref.wb,referenceBlocks,cfg,{baseWb:base.wb});
assert.equal(diagnostic[4].status,'pass',JSON.stringify(diagnostic[4]));
if(require.main===module){
  if(process.argv.includes('--compare'))console.log(JSON.stringify(checks,null,2));
  fs.mkdirSync(path.join(__dirname,'artifacts'),{recursive:true});
  fs.writeFileSync(path.join(__dirname,'artifacts','AltiuslinkVN-20261012.xlsx'),out.bytes);
  fs.writeFileSync(path.join(__dirname,'artifacts','chain12-reference-fixture.xlsx'),fixture.bytes);
  console.log('PASS: October 12 - numeric hours, 4 Quang identities, template team range, strict conversion and guards, migration; '+history+' history cells unchanged; diagnostic reference-input KPI 40/40. Remaining source differences are still reported.');
}
module.exports={ctx,base,ref,op,cfg,blocks,out,wb,checks,fixture};
