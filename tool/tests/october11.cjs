const fs=require('fs'),path=require('path'),vm=require('vm'),assert=require('assert/strict');
const root=path.resolve(__dirname,'../..');
const ctx=vm.createContext({console,Uint8Array,ArrayBuffer,TextEncoder,TextDecoder,Date,setTimeout});
for(const f of ['vendor/xlsx.full.min.js','vendor/fflate.min.js','src/config.js','src/parse.js','src/pipeline.js','src/writer.js','src/xlsxpatch.js','src/selftest.js'])
  vm.runInContext(fs.readFileSync(path.join(root,'tool',f),'utf8'),ctx,{filename:f});
const {XLSX,AppConfig,Parse,Pipeline,XlsxPatch,SelfTest}=ctx;
const dir=path.join(root,'ExtraData','Daily Report 11.10.2026');
function load(folder,prefix){
  const folderPath=path.join(dir,folder),name=fs.readdirSync(folderPath).find(n=>n.startsWith(prefix)&&n.endsWith('.xlsx'));
  const bytes=new Uint8Array(fs.readFileSync(path.join(folderPath,name)));
  return {name,bytes,wb:XLSX.read(bytes,{type:'array',cellFormula:true,cellHTML:false,cellText:false})};
}
const base=load('Input','Altius'),ref=load('Output','Altius'),se=load('Input','Session'),sf=load('Input','Daily'),op=load('Input','SBI');
// Historical output assertions use an explicit old business profile, not current defaults.
const cfg=require('./fixtures/rules-v4.cjs')(AppConfig.defaults()),serial=Parse.toDateSerial('2026-10-11');
if(process.argv.includes('--swapped'))cfg.op.acceptSwappedDates=true;
const session=Pipeline.processSession(se.wb,cfg,se.name),ccvn=Pipeline.processCcvn(sf.wb,cfg,sf.name);
const ops=[Pipeline.processOpFile(op.wb,cfg,serial,op.name)];
const mail=Pipeline.inspectMailSf(base.wb,cfg,[]);
const blocks=Pipeline.buildBlocks({session,ccvn,ops,emailTable:mail.table},cfg,serial);
const out=XlsxPatch.patch(base.bytes,blocks,cfg,{masterWb:base.wb});
const wb=XLSX.read(out.bytes,{type:'array',cellFormula:true});
const checks=SelfTest.compareWithReference(out.bytes,ref.wb,blocks,cfg,{baseWb:base.wb,oddRateEmails:out.oddRateEmails,affectedEmails:mail.affected,input:{ccvn,ops}});
assert.equal(session.stats.kept,1399);
assert.equal(session.stats.dropped,14);
assert.equal(ops[0].stats.swapped,102);
assert.equal(blocks.block1.rows.length,2097);
assert.equal(blocks.block2.rows.length,377);
assert.equal(blocks.block3.rows.length,160);
assert.equal(blocks.appliedAdjustments.length,6);
for(const check of checks)assert.equal(check.status,'pass',JSON.stringify(check));
assert.equal(out.oddRateEmails&&Object.keys(out.oddRateEmails).length||0,0,'No silently skipped KPI profiles');
assert.equal(wb.Sheets[cfg.targets.other.sheet].I110.v,0.94);
assert(Math.abs(wb.Sheets[cfg.targets.other.sheet].F142.v-7.995294117647059)<1e-9);
let history=0;
for(const name of [cfg.summary.sheet,'202610 KPI Result']){
  const header=name===cfg.summary.sheet?4:5;
  for(let c=0;c<40;c++){
    const col=XLSX.utils.encode_col(c),date=base.wb.Sheets[name][col+header]?.v;
    if(typeof date!=='number'||date<45000||date>=serial)continue;
    for(let r=6;r<=70;r++){
      assert.equal(wb.Sheets[name][col+r]?.v,base.wb.Sheets[name][col+r]?.v);
      assert.equal(wb.Sheets[name][col+r]?.f,base.wb.Sheets[name][col+r]?.f);history++;
    }
  }
}
assert.equal(history,1300);
const input={session,ccvn,ops,emailTable:mail.table};
const disabled=AppConfig.clone(cfg);disabled.approvedAdjustments.enabled=false;
const raw=Pipeline.buildBlocks(input,disabled,serial);
assert.equal(raw.appliedAdjustments.length,0);
assert.equal(raw.block3.rows[104][6],0.941176470588235,'Raw hours preserved when adjustment is disabled');
assert.equal(ccvn.dataRows.find(r=>r[1]==='20261012218289')[5],'VU NHU QUYNH','Original Salesforce data stays unchanged');
assert.equal(ops[0].other.rows[104].cells[8],0.941176470588235,'Original OP data stays unchanged');
const anotherDay=Pipeline.buildBlocks(input,cfg,serial+1);
assert.equal(anotherDay.appliedAdjustments.length,0,'Never transfer approved adjustments to another day');
const reordered={...ccvn,dataRows:ccvn.dataRows.slice().reverse()};
const shuffled=Pipeline.buildBlocks({...input,ccvn:reordered},cfg,serial);
const assigned=shuffled.block2.rows.filter(r=>r[4]==='PHAM THANH DAO').map(r=>r[0]).sort();
assert.deepEqual(Array.from(assigned),Array.from(cfg.approvedAdjustments.cases,r=>r.caseNumber).sort());
const conflict=AppConfig.clone(cfg);conflict.approvedAdjustments.cases.push({...conflict.approvedAdjustments.cases[0]});
assert.throws(()=>Pipeline.buildBlocks(input,conflict,serial),/Case Number/);
const wrongHours=AppConfig.clone(cfg);wrongHours.approvedAdjustments.hours[0].from=1;
assert.throws(()=>Pipeline.buildBlocks(input,wrongHours,serial),/không khớp/);
const wrongOwner=AppConfig.clone(cfg);wrongOwner.approvedAdjustments.cases[0].from='WRONG OWNER';
assert.throws(()=>Pipeline.buildBlocks(input,wrongOwner,serial),/không khớp/);
const missingTarget=AppConfig.clone(cfg);missingTarget.approvedAdjustments.cases[0].to='UNKNOWN PERSON';
assert.throws(()=>Pipeline.buildBlocks(input,missingTarget,serial),/không có email/);
const legacy=AppConfig.clone(cfg);delete legacy.rulesVersion;delete legacy.approvedAdjustments;
legacy.op.acceptSwappedDates=false;
legacy.requestedByAliases['VU NHU QUYNH']='NGUYEN PHUONG ANH';
legacy.requestedByAliases['PHAM THI THU HOAI']='TRAN LAN HUONG';
const migrated=AppConfig.migrate(legacy);
assert.equal(migrated.requestedByAliases['VU NHU QUYNH'],'NGUYEN PHUONG ANH');
assert.equal(migrated.requestedByAliases['PHAM THI THU HOAI'],false);
assert.equal(migrated.op.acceptSwappedDates,true);
assert.equal(migrated.approvedAdjustments.cases.length,0);
assert.equal(migrated.approvedAdjustments.enabled,false);
const custom=AppConfig.clone(cfg);custom.op.acceptSwappedDates=false;custom.approvedAdjustments.enabled=false;
custom.requestedByAliases['VU NHU QUYNH']='CUSTOM';
assert.equal(AppConfig.migrate(custom).op.acceptSwappedDates,false);
assert.equal(AppConfig.migrate(custom).approvedAdjustments.enabled,false);
assert.equal(AppConfig.migrate(custom).requestedByAliases['VU NHU QUYNH'],'CUSTOM');
// An incorrect configured alias must be reported, never excused as a reference error.
const badAlias=AppConfig.clone(disabled);badAlias.requestedByAliases['PHAM THI THU HOAI']='TRAN LAN HUONG';
const badBlocks=Pipeline.buildBlocks(input,badAlias,serial);
const badOutput=XlsxPatch.patch(base.bytes,badBlocks,badAlias,{masterWb:base.wb});
const strict=SelfTest.compareWithReference(badOutput.bytes,ref.wb,badBlocks,badAlias,{baseWb:base.wb});
assert.equal(strict[1].status,'fail');assert(!strict[1].detail.includes('không phải lỗi tool'));
if(require.main===module){
  if(process.argv.includes('--compare'))console.log(JSON.stringify({session:session.stats,op:ops[0].stats,rows:[blocks.block1.rows.length,blocks.block2.rows.length,blocks.block3.rows.length],checks},null,2));
  console.log('PASS: October 11 - 33757 detail cells, 22 summary and 40 KPI cells match; '+history+' historical cells/formulas unchanged; six scoped adjustments, migration and conflict guards verified.');
  fs.mkdirSync(path.join(__dirname,'artifacts'),{recursive:true});
  fs.writeFileSync(path.join(__dirname,'artifacts','AltiuslinkVN-20261011.xlsx'),out.bytes);
  const ranges=['block1','block2','block3'].map(key=>{
    const b=blocks[key],a=XLSX.utils.decode_cell(b.anchor);
    return {sheet:b.sheet,anchor:b.anchor,rows:b.rows.map((_,i)=>b.headers.map((_,j)=>
      ref.wb.Sheets[b.sheet][XLSX.utils.encode_cell({r:a.r+i,c:a.c+j})]?.v??''))};
  });
  const cells=[];
  for(const spec of [...cfg.summary.cells,...cfg.summary.topicRows])cells.push({sheet:cfg.summary.sheet,cell:'P'+spec.row,value:ref.wb.Sheets[cfg.summary.sheet]['P'+spec.row]?.v??''});
  for(let r=6;r<=60;r++){
    const email=ref.wb.Sheets['202610 KPI Result']['D'+r]?.v;
    if(typeof email!=='string'||!email.includes('@'))continue;
    const dr=Array.from({length:46},(_,i)=>i+15).find(i=>ref.wb.Sheets['Daily KPI Result']['D'+i]?.v===email);
    if(!dr)continue;
    cells.push({sheet:'202610 KPI Result',cell:'O'+r,value:ref.wb.Sheets['202610 KPI Result']['O'+r]?.v??''});
    cells.push({sheet:'Daily KPI Result',cell:'W'+dr,value:ref.wb.Sheets['202610 KPI Result']['O'+r]?.v??''});
  }
  cells.push({sheet:'202610 KPI Result',cell:'O7',value:ref.wb.Sheets['202610 KPI Result'].O7.v});
  cells.push({sheet:'Daily KPI Result',cell:'W14',value:ref.wb.Sheets['202610 KPI Result'].O7.v});
  fs.writeFileSync(path.join(__dirname,'artifacts','october11-excel-expected.json'),JSON.stringify({ranges,cells}));
}
module.exports={ctx,base,ref,se,sf,op,cfg,serial,session,ccvn,ops,mail,blocks,out,wb,checks};
