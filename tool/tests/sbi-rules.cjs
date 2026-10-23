// SBI-confirmed rules from docs/VẤN ĐỀ.docx. Output is comparison evidence only.
const fs=require('fs'),path=require('path'),vm=require('vm'),assert=require('assert/strict');
const root=path.resolve(__dirname,'../..');
const ctx=vm.createContext({console,Uint8Array,ArrayBuffer,TextEncoder,TextDecoder,Date,setTimeout});
for(const f of ['vendor/xlsx.full.min.js','vendor/fflate.min.js','src/config.js','src/parse.js','src/pipeline.js','src/writer.js','src/xlsxpatch.js','src/selftest.js','src/markup.js'])
  vm.runInContext(fs.readFileSync(path.join(root,'tool',f),'utf8'),ctx,{filename:f});
const {XLSX,AppConfig,Parse,Pipeline,XlsxPatch,SelfTest,Markup}=ctx;
const plain=x=>JSON.parse(JSON.stringify(x));
function load(folder,prefix){
  const dir=path.join(root,folder),names=fs.readdirSync(dir).filter(n=>n.startsWith(prefix)&&n.endsWith('.xlsx'));
  assert.equal(names.length,1,folder+'/'+prefix);
  const name=names[0],bytes=new Uint8Array(fs.readFileSync(path.join(dir,name)));
  return {name,bytes,wb:XLSX.read(bytes,{type:'array',cellFormula:true,cellNF:true})};
}
const cfg=AppConfig.defaults();
assert.equal(cfg.rulesVersion,6);assert.equal(cfg.session.stagingColumns[3].from,'Requested time');
for(const version of [0,1,2,3,4,5])for(const source of ['First response time','Assign time','Requested time']){
  const old=AppConfig.defaults();old.rulesVersion=version;
  old.session.stagingColumns[3]={out:source,from:source};old.session.startTimeMappingVersion=2;
  old.requestedByAliases['VU NHU QUYNH']='CUSTOM PERSON';
  old.approvedAdjustments={enabled:true,source:'ExtraData/Daily Report 11.10.2026/Output',cases:[],hours:[{from:0.941176470588235,to:.94}]};
  delete old.sourceReviews;
  const before=JSON.stringify(old),m=AppConfig.migrate(old);
  assert.equal(m.session.stagingColumns[3].from,'Requested time');
  assert.equal(m.session.startTimeMappingVersion,3);assert.equal(m.rulesVersion,6);
  assert.equal(m.requestedByAliases['VU NHU QUYNH'],'CUSTOM PERSON');
  assert.equal(m.approvedAdjustments.enabled,false);assert.equal(m.approvedAdjustments.hours.length,1);
  assert.equal(m.sourceReviews.customerFeedbackTopic,true);
  assert.deepEqual(plain(AppConfig.migrate(m)),plain(m));assert.equal(JSON.stringify(old),before);
}
const chosen=AppConfig.defaults();chosen.session.stagingColumns[3]={out:'Assign time',from:'Assign time'};
assert.equal(AppConfig.migrate(chosen).session.stagingColumns[3].from,'Assign time','Explicit post-v6 choice stays visible');
const folder='ExtraData/Daily Report 12.10.2026';
const se=load(folder,'Session'),sf=load(folder,'Daily'),op=load(folder,'SBI'),
  base=load(folder,'AltiuslinkVN_Daily Report_(New)-20261011'),ref=load(folder,'AltiuslinkVN_Daily Report_(New)-20261012');
const sources=[se,sf,op,base],snapshots=sources.map(f=>JSON.stringify(f.wb));
const serial=Parse.toDateSerial('2026-10-12');
const session=Pipeline.processSession(se.wb,cfg,se.name),ccvn=Pipeline.processCcvn(sf.wb,cfg,sf.name);
const input={session,ccvn,ops:[Pipeline.processOpFile(op.wb,cfg,serial,op.name)],emailTable:Pipeline.inspectMailSf(base.wb,cfg,[]).table};
const blocks=Pipeline.buildBlocks(input,cfg,serial),warnings=[];
Pipeline.checkRoster(input,cfg,base.wb,blocks,warnings);
assert.equal(session.staging.rows.length,1604);
const requested=Parse.findColumn(session.headers,'Requested time');
session.keptRows.forEach((row,i)=>{
  assert.equal(blocks.block1.rows[i][5],row[requested]);
  assert.equal(blocks.block1.rows[i][5],ref.wb.Sheets['1. CSKH']['F'+(i+11)].v);
});
assert.equal(Pipeline.detectReportDate(session,cfg).serial,serial);
const feedback=blocks.warnings.find(w=>w.message.includes('Đánh giá khách hàng'));
assert(feedback);assert.equal(feedback.level,'warn');assert.equal(feedback.spots.length,1);
assert.equal(feedback.spots[0].row,1070);assert.equal(feedback.spots[0].col,14);
assert.equal(blocks.block1.rows[1073-11][9],se.wb.Sheets[session.sheetName].O1070.v);
assert.equal(blocks.summary.topics['12'],1,'Warn but retain the source topic');
// Optional descriptions must not be deleted just to match a handwritten output.
for(let r=1927;r<=1955;r++)assert.equal(blocks.block1.rows[r-11][7],'ok');
const review=blocks.warnings.find(w=>w.message.includes('ô giờ đã được SBI'));
assert(review);assert.equal(review.spots.length,1);assert.equal(review.spots[0].row,2001);assert.equal(review.spots[0].col,6);
const nguyen=blocks.block3.rows.find(r=>r[2]==='nguyen.nt@altius-link.com.vn'&&r[4]===5);
assert(nguyen,'Keep Nguyen source 5; never backfill 5.24');
assert.equal(blocks.block3.rows.length,144);assert(!blocks.block3.rows.some(r=>r[2]==='dao.pt@altius-link.com.vn'));
const dateWarning=warnings.find(w=>w.message.includes('giờ khai ở ngày khác'));
assert(dateWarning.spots.some(s=>s.row===2017&&s.col===2));
assert(dateWarning.display.tables[0].rows.some(r=>r[3]==='C2017:C2024'));
assert.match(dateWarning.display.action,/lịch sử đúng/);
assert.equal(blocks.summary.counts.viber,1,'The detail Viber row must be counted despite old summary 0');
assert.equal(blocks.block2.rows.filter(r=>r[4]==='NGUYEN PHUONG ANH').length,36);
const shared=blocks.warnings.find(w=>w.message.includes('tên không tra được email'));
assert.equal(shared.spots.length,11);assert.match(shared.display.action,/CC-VN/);
assert.equal(blocks.block2.rows.filter(r=>r[4]==='CC-VN'&&r[5]==='').length,11);
assert(!blocks.warnings.some(w=>w.level==='error'&&/chủ tài khoản/.test(w.message)));
// Assignment is all to the selected name, independent of order and never split half.
const assigned=AppConfig.clone(cfg);assigned.requestedByAliases['VU NHU QUYNH']='PHAM THANH DAO';
assigned.requestedByAliases['CC-VN']='NGUYEN THI THANH THUY';
const assignedBlocks=Pipeline.buildBlocks(input,assigned,serial);
assert.equal(assignedBlocks.block2.rows.filter(r=>r[4]==='PHAM THANH DAO').length,36);
const rawCreator=Parse.findColumn(ccvn.headers,'Created By');
ccvn.dataRows.forEach((r,i)=>{if(r[rawCreator]==='CC-VN'){
  assert.equal(assignedBlocks.block2.rows[i][4],'NGUYEN THI THANH THUY');
  assert.equal(assignedBlocks.block2.rows[i][5],'thuy.ntt1@altius-link.com.vn');
}});
assert.deepEqual(plain(assignedBlocks.block3.rows),plain(blocks.block3.rows));
assert.deepEqual(plain(assignedBlocks.block1.rows),plain(blocks.block1.rows));
const reverse={...ccvn,dataRows:ccvn.dataRows.slice().reverse(),dataRowNumbers:ccvn.dataRowNumbers.slice().reverse()};
assert.deepEqual(plain(Pipeline.buildBlocks({...input,ccvn:reverse},assigned,serial).block2.rows),plain(assignedBlocks.block2.rows).reverse());
// Targeted numeric-input review must disappear after correction; never flag other days.
const corrected=plain(input);
corrected.ops[0].other.rows.find(r=>r.excelRow===2001).cells[6]=5.24;
assert(!Pipeline.buildBlocks(corrected,cfg,serial).warnings.some(w=>w.message.includes('ô giờ đã được SBI')));
assert(!Pipeline.buildBlocks(input,cfg,serial+1).warnings.some(w=>w.message.includes('ô giờ đã được SBI')));
// Missing Requested column stays blank with an error; no silent fallback to First response.
const missing=plain(se.wb);delete missing.Sheets[session.sheetName][XLSX.utils.encode_cell({r:0,c:requested})];
const noRequested=Pipeline.processSession(missing,cfg,se.name);
assert(noRequested.warnings.some(w=>w.level==='error'&&w.message.includes('Requested time')));
assert(noRequested.staging.rows.every(r=>r[3]===''));
// Source locations survive removed bot rows and OP filtering; OP topics also warned.
const probe=plain(input),opRec=probe.ops[0].cskh.rows[0];opRec.cells[9]='12. Đánh giá khách hàng';
const opFeedback=Pipeline.buildBlocks(probe,cfg,serial).warnings.find(w=>w.message.includes('Đánh giá khách hàng'));
assert(opFeedback.spots.some(s=>s.file===op.name&&s.row===opRec.excelRow&&s.col===9));
// October 11 case hours stay full precision under the new default, not the old .94 override.
const op11=load('ExtraData/Daily Report 11.10.2026/Input','SBI'),day11=Parse.toDateSerial('2026-10-11');
const b11=Pipeline.buildBlocks({ops:[Pipeline.processOpFile(op11.wb,cfg,day11,op11.name)]},cfg,day11);
const cuc=b11.block3.rows.find(r=>r[2]==='cuc.ht@altius-link.com.vn'&&typeof r[6]==='number');
assert.equal(cuc[6],op11.wb.Sheets[cfg.targets.other.sheet].I1833.v);assert.notEqual(cuc[6],.94);
assert.equal(b11.appliedAdjustments.length,0);
const out=XlsxPatch.patch(base.bytes,blocks,cfg,{masterWb:base.wb}),wb=XLSX.read(out.bytes,{type:'array',cellFormula:true,sheetStubs:true});
assert.equal(wb.Sheets[cfg.targets.sfcase.sheet].G73.f,'IFERROR(VLOOKUP($F73,\'mail SF\'!$C$3:$D$69,2,0),"")',
  'Unassigned CC-VN remains blank rather than #N/A after native Excel recalculation');
const checks=SelfTest.compareWithReference(out.bytes,ref.wb,blocks,cfg,{baseWb:base.wb,input,oddRateEmails:out.oddRateEmails});
assert.match(checks[0].detail,/^30 ô lệch/,'29 optional descriptions and 1 source topic, no Start Time mismatch');
assert.equal(checks.find(c=>c.name.startsWith('6.')).status,'pass');
assert.equal(checks.find(c=>c.name.startsWith('8.')).status,'pass');
assert.equal(out.teamRows.length,28);
for(const w of [...blocks.warnings,...warnings,...session.warnings,...input.ops[0].warnings])if(w.display){
  assert(w.display.title.length<100,w.display.title);assert(w.display.summary.length<200,w.display.summary);
}
// New warnings must produce a source-preserving marked copy with exact original positions.
const entries=Pipeline.collectSpots([[feedback,review,dateWarning]]);
for(const file of [se,op]){
  const entry=entries.find(e=>e.file===file.name);assert(entry);
  const marked=XLSX.read(Markup.build(file.bytes,entry).bytes,{type:'array',cellFormula:true,cellNF:true});
  assert(marked.Sheets['Loi can sua']);
  const spots=file===se?feedback.spots:[...review.spots,...dateWarning.spots];
  for(const s of spots){const cell=XLSX.utils.encode_cell({r:s.row-1,c:s.col});
    for(const key of ['v','t','f','z'])assert.equal(marked.Sheets[s.sheet][cell]?.[key],file.wb.Sheets[s.sheet][cell]?.[key],cell+' '+key);
  }
}
sources.forEach((f,i)=>assert.equal(JSON.stringify(f.wb),snapshots[i],f.name+' mutated'));
const artifacts=path.join(__dirname,'artifacts');fs.mkdirSync(artifacts,{recursive:true});
fs.writeFileSync(path.join(artifacts,'SBI-v6-20261012.xlsx'),out.bytes);
const ranges=['block1','block2','block3'].map(k=>({sheet:blocks[k].sheet,anchor:blocks[k].anchor,rows:blocks[k].rows}));
const cells=[];
for(let r=6;r<=60;r++){
  const email=wb.Sheets['202610 KPI Result']['D'+r]?.v;if(typeof email!=='string'||!email.includes('@'))continue;
  const dr=Array.from({length:46},(_,i)=>i+15).find(i=>wb.Sheets['Daily KPI Result']['D'+i]?.v===email);
  if(dr)cells.push({sheet:'Daily KPI Result',cell:'W'+dr,value:wb.Sheets['202610 KPI Result']['P'+r]?.v??''});
}
cells.push({sheet:'Daily KPI Result',cell:'W14',value:wb.Sheets['202610 KPI Result'].P7.v});
cells.push({sheet:cfg.summary.sheet,cell:'Q14',value:1});
fs.writeFileSync(path.join(artifacts,'sbi-v6-excel-expected.json'),JSON.stringify({ranges,cells}));
console.log('PASS: SBI v6 migration; 1604 Requested times; source topic/hour/date warnings; all-to-assigned case policy; optional notes; no rounding/backfill; Viber=1; original files unchanged; marked cells preserved; history and 7 Pivot fields.');
console.log('Remaining comparisons:',JSON.stringify(checks.map(c=>({name:c.name,status:c.status,detail:c.detail}))));
