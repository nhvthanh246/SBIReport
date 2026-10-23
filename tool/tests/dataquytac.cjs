// Current default business rules: raw -> cleaned -> final, without reading reference
// values into the production pipeline. Missing source data remains a strict mismatch.
const fs=require('fs'),path=require('path'),assert=require('assert/strict');
const audit=require('./dataquytac-audit.cjs');
const {ctx,files,cfg,serial,session,ccvn,ops,mail,input,blocks,patched,report,stages,blockRows,same}=audit;
const {XLSX,AppConfig,Parse,Pipeline,Writer}=ctx;
const plain=o=>JSON.parse(JSON.stringify(o));
const wb=XLSX.read(patched.bytes,{type:'array',cellFormula:true});
function compareSheets(actual,expected,name,expectedAt){
  const a=Parse.sheetToMatrix(actual.Sheets[name]).rows,b=Parse.sheetToMatrix(expected.Sheets[name]).rows;
  let checked=0;
  for(let r=0;r<Math.max(a.length,b.length);r++)for(let c=0;c<Math.max(a[r]?.length||0,b[r]?.length||0);c++){
    const want=expectedAt?expectedAt(r,c,b[r]?.[c]):b[r]?.[c];
    assert(same(a[r]?.[c],want),name+'!'+XLSX.utils.encode_cell({r,c})+': '+a[r]?.[c]+' != '+want);checked++;
  }
  return checked;
}
assert.equal(cfg.rulesVersion,6);assert.equal(cfg.rulesSource,AppConfig.DEFAULT_CONFIG.rulesSource);
assert.equal(cfg.session.stagingColumns[3].from,'Requested time');
assert.equal(cfg.approvedAdjustments.enabled,false);assert.equal(cfg.approvedAdjustments.cases.length,0);
assert.equal(blocks.appliedAdjustments.length,0);assert.equal(mail.corrections.length,0);
assert.equal(serial,46278);assert.deepEqual(plain(session.stats),{total:1567,dropped:12,kept:1555});
const cleanSession=XLSX.read(Writer.buildSessionWorkbook(session,cfg),{type:'array'});
const cleanSf=XLSX.read(Writer.buildCcvnWorkbook(ccvn,cfg),{type:'array'});
let cleanedCells=0;
const requestedCol=Parse.findColumn(session.headers,'Requested time');
for(const name of files.cleanSession.wb.SheetNames)cleanedCells+=compareSheets(cleanSession,files.cleanSession.wb,name,
  name===session.staging.name?(r,c,old)=>c===3?(r===0?'Requested time':session.keptRows[r-1]?.[requestedCol]):old:null);
for(const name of files.cleanCcvn.wb.SheetNames)cleanedCells+=compareSheets(cleanSf,files.cleanCcvn.wb,name);
assert.equal(stages.sessionStaging.count,1555);assert.equal(stages.salesforceRawToClean.count,0);
assert(stages.sessionStaging.differences.every(d=>XLSX.utils.decode_cell(d.cell).c===3));
assert.equal(stages.block1.count,1555);assert.equal(stages.block2.count,22);assert.equal(stages.block3.count,81);
assert(stages.block1.differences.every(d=>XLSX.utils.decode_cell(d.cell).c===5),'Only intentional Start Time differences vs the old output');
session.keptRows.forEach((row,i)=>assert.equal(blocks.block1.rows[i][5],row[requestedCol]));
assert.equal(blocks.block1.rows.length,2218);assert.equal(blocks.block2.rows.length,379);assert.equal(blocks.block3.rows.length,128);
const ref2=blockRows(files.reference,blocks.block2),ref3=blockRows(files.reference,blocks.block3);
blocks.block3.rows.forEach((row,r)=>row.forEach((v,c)=>assert(same(v,ref3[r][c]),'Shared hours '+r+','+c)));
assert(stages.block3.differences.every(d=>XLSX.utils.decode_cell(d.cell).r>=133),'Only the 24 absent source rows may differ');
assert(stages.block2.differences.every(d=>d.actual==='ly.ntk@altius-link.com.vn'&&d.expected==='ly.ntk1@altius-link.com.vn'));
const creator=Parse.findColumn(ccvn.headers,'Created By');
for(const [from,to] of Object.entries(cfg.requestedByAliases)){
  const rows=ccvn.dataRows.map((r,i)=>({r,i})).filter(({r})=>Parse.normName(r[creator])===Parse.normName(from));
  assert.equal(rows.length,to?19:6);
  for(const {r,i} of rows){
    assert.equal(Parse.normName(blocks.block2.rows[i][4]),Parse.normName(to||from));
    assert.equal(blocks.block2.rows[i][5],ref2[i][5]);
    assert.equal(Parse.normName(r[creator]),Parse.normName(from),'Do not mutate raw creator');
  }
}
assert.equal(blocks.block2.rows.filter(r=>r[5]==='ly.ntk@altius-link.com.vn').length,62);
assert.equal(blocks.block2.rows.filter(r=>r[5]==='ly.ntk1@altius-link.com.vn').length,0);
assert.equal(wb.Sheets['mail SF'].C43.v,files.base.wb.Sheets['mail SF'].C43.v,'No invented (3M) identity');
const missingPeople=['ly.ntk1@altius-link.com.vn','nguyen.nt@altius-link.com.vn','anh.np@altius-link.com.vn'];
for(const email of missingPeople){
  for(const sheet of Object.values(files.op.wb.Sheets))assert(!Object.values(sheet).some(c=>c&&c.v===email),'No hours for '+email+' anywhere in OP');
  assert(!blocks.block3.rows.some(r=>r[2]===email),'No fabricated 7+1 hours for '+email);
  assert.equal(blocks.kpi.forEmail(email).hours,0);
}
const warnings=[];
Pipeline.inspectMailSf(files.base.wb,cfg,warnings,files.base.name);
const duplicate=warnings.find(w=>w.level==='error'&&w.message.includes('ghi trùng'));
assert(duplicate);assert.equal(duplicate.spots.length,2);assert.equal(duplicate.display.tables[0].rows.length,2);
assert.match(duplicate.display.action,/mã case.*email/);assert.match(duplicate.display.action,/bảng giờ/);
const before=JSON.stringify([blocks.block1.rows,blocks.block2.rows,blocks.block3.rows]);
Pipeline.checkRoster(input,cfg,files.base.wb,blocks,warnings);
const noHours=warnings.find(w=>w.message.includes('thiếu giờ hợp lệ'));
assert.equal(noHours.display.tables[0].rows.length,6);
for(const email of missingPeople)assert(noHours.display.tables[0].rows.some(r=>r[1]===email));
assert.equal(noHours.display.tables[0].rows.find(r=>r[1]===missingPeople[0])[0],'NGUYEN THI KHANH LY');
assert.equal(JSON.stringify([blocks.block1.rows,blocks.block2.rows,blocks.block3.rows]),before);
const zeros=blocks.warnings.find(w=>w.message.includes('ô có 0 giờ việc khác'));
assert.equal(zeros.spots.length,3);assert(zeros.spots.every(s=>s.col===12&&String(s.value)==='0'));
for(const s of zeros.spots)assert.equal(files.op.wb.Sheets[s.sheet]['M'+s.row].v,0,'Original zero unchanged');
// Explicit column scope: only M zeros, never G zero, task labels, or nonzero M.
const probeOps=plain(ops);
probeOps[0].other.rows[0].cells[6]=0;probeOps[0].other.rows[0].cells[11]=0;
probeOps[0].other.rows[0].cells[12]='0';probeOps[0].other.rows[1].cells[12]=0.25;
const probe=Pipeline.buildBlocks({...input,ops:probeOps},cfg,serial);
assert.equal(probe.block3.rows[0][4],0);assert.equal(probe.block3.rows[0][9],0);
assert.equal(probe.block3.rows[0][10],'');assert.equal(probe.block3.rows[1][10],0.25);
const noBlank=AppConfig.clone(cfg);noBlank.op.other.zeroAsBlankColumns=[];
assert.equal(Pipeline.buildBlocks({...input,ops:probeOps},noBlank,serial).block3.rows[0][10],0);
// Already cleaned input must be idempotent, not lose another set of rows/columns.
const fromClean=Pipeline.buildBlocks({...input,
  session:Pipeline.processSession(files.cleanSession.wb,cfg,files.cleanSession.name),ccvn:audit.cleanCcvn},cfg,serial);
for(const key of ['block1','block2','block3'])assert.deepEqual(plain(fromClean[key].rows),plain(blocks[key].rows));
const reversed={...ccvn,dataRows:ccvn.dataRows.slice().reverse(),dataRowNumbers:ccvn.dataRowNumbers.slice().reverse()};
assert.deepEqual(plain(Pipeline.buildBlocks({...input,ccvn:reversed},cfg,serial).block2.rows),plain(blocks.block2.rows).reverse());
// Every pre-v6 stored default migrates once; unrelated choices survive.
for(const version of [0,1,2,3,4]){
  const old=require('./fixtures/rules-v4.cjs')(AppConfig.defaults());old.rulesVersion=version;
  delete old.op.other.zeroAsBlankColumns;old.summary.ignoredChannels.push('CUSTOM CHANNEL');
  const migrated=AppConfig.migrate(old);
  assert.equal(migrated.rulesVersion,6);assert.equal(migrated.rulesSource,AppConfig.DEFAULT_CONFIG.rulesSource);
  assert.equal(migrated.requestedByAliases['VU NHU QUYNH'],'NGUYEN PHUONG ANH');
  assert.equal(migrated.requestedByAliases['LUU NGOC QUANG'],'NGUYEN THAO NGUYEN');
  assert.equal(migrated.approvedAdjustments.enabled,false);assert.equal(migrated.approvedAdjustments.cases.length,5,'Keep disabled records for audit');
  assert.equal(migrated.mailSf.nameCorrections.length,0);assert.deepEqual(plain(migrated.op.other.zeroAsBlankColumns),[12]);
  assert(migrated.summary.ignoredChannels.includes('CUSTOM CHANNEL'));
  assert.deepEqual(plain(AppConfig.migrate(migrated)),plain(migrated));
}
const custom=AppConfig.defaults();custom.requestedByAliases['VU NHU QUYNH']=false;
custom.requestedByAliases['LUU NGOC QUANG']='CUSTOM PERSON';custom.op.acceptSwappedDates=false;
assert.deepEqual(plain(AppConfig.migrate(custom)),plain(custom));
const legacyCustom=AppConfig.clone(custom);legacyCustom.rulesVersion=4;
assert.equal(AppConfig.migrate(legacyCustom).requestedByAliases['LUU NGOC QUANG'],'CUSTOM PERSON');
// Never relabel expected source gaps as passed output checks.
for(const n of ['4.','6.','8.'])assert.equal(report.checks.find(c=>c.name.startsWith(n)).status,'pass');
for(const n of ['1.','2.','3.','5.'])assert.equal(report.checks.find(c=>c.name.startsWith(n)).status,'fail');
assert.match(report.checks.find(c=>c.name.startsWith('5.')).detail,/5\/39/);
// Full history values AND formulas, not just the cells managed by SelfTest.
let history=0;
for(const name of [cfg.summary.sheet,'202609 KPI Result']){
  const header=name===cfg.summary.sheet?4:5;
  for(let c=0;c<40;c++){
    const col=XLSX.utils.encode_col(c),day=files.base.wb.Sheets[name][col+header]?.v;
    if(typeof day!=='number'||day<45000||day>=serial)continue;
    for(let r=6;r<=70;r++)for(const field of ['v','f']){
      assert.equal(wb.Sheets[name][col+r]?.[field],files.base.wb.Sheets[name][col+r]?.[field]);history++;
    }
  }
}
// Excel verification fixture: detail ranges vs verified source-backed output,
// summary vs reference; daily formulas vs the independently calculated monthly KPI.
const ranges=['block1','block2','block3'].map(key=>({sheet:blocks[key].sheet,anchor:blocks[key].anchor,rows:blocks[key].rows}));
const cells=cfg.summary.cells.concat(cfg.summary.topicRows).map(s=>({sheet:cfg.summary.sheet,cell:'R'+s.row,value:files.reference.wb.Sheets[cfg.summary.sheet]['R'+s.row]?.v??''}));
let individuals=0;
for(let r=6;r<=60;r++){
  const email=wb.Sheets['202609 KPI Result']['D'+r]?.v;if(typeof email!=='string'||!email.includes('@'))continue;
  const dr=Array.from({length:46},(_,i)=>i+15).find(i=>wb.Sheets['Daily KPI Result']['D'+i]?.v===email);if(!dr)continue;
  cells.push({sheet:'Daily KPI Result',cell:'W'+dr,value:wb.Sheets['202609 KPI Result']['Q'+r]?.v??''});individuals++;
}
cells.push({sheet:'Daily KPI Result',cell:'W14',value:wb.Sheets['202609 KPI Result'].Q7.v});
const artifact=path.join(__dirname,'artifacts');fs.mkdirSync(artifact,{recursive:true});
fs.writeFileSync(path.join(artifact,'DataQuyTac-20260913.xlsx'),patched.bytes);
fs.writeFileSync(path.join(artifact,'dataquytac-excel-expected.json'),JSON.stringify({ranges,cells}));
console.log('PASS: DataQuyTac with SBI v6 - '+cleanedCells+' cleaned cells checked with Requested-time expectation; 1555 intentional Start Time differences vs old output, other CSKH cells unchanged; 128 source hours rows and 22 summary cells match. '+individuals+' KPI formulas prepared for Excel verification; '+history+' history value/formula checks.');
console.log('KNOWN SOURCE GAPS (not passes): 22 SF emails, 24 missing hours rows (81 cells), 5/39 monthly KPI cells. No output backfill.');
