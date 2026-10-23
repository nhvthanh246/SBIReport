// Inspect the authoritative raw -> business-cleaned -> final report evidence.
// No file in DataQuyTac or Data is modified; optional JSON goes to test artifacts.
const fs=require('fs'),path=require('path'),vm=require('vm'),crypto=require('crypto');
const root=path.resolve(__dirname,'../..');
const ctx=vm.createContext({console,Uint8Array,ArrayBuffer,TextEncoder,TextDecoder,Date,setTimeout});
for(const file of ['vendor/xlsx.full.min.js','vendor/fflate.min.js','src/config.js','src/parse.js','src/pipeline.js','src/writer.js','src/xlsxpatch.js','src/selftest.js'])
  vm.runInContext(fs.readFileSync(path.join(root,'tool',file),'utf8'),ctx,{filename:file});
const {XLSX,AppConfig,Parse,Pipeline,XlsxPatch,SelfTest}=ctx;
function read(folder,prefix){
  const dir=path.join(root,folder),names=fs.readdirSync(dir).filter(n=>n.startsWith(prefix)&&n.endsWith('.xlsx'));
  if(names.length!==1)throw Error(folder+'/'+prefix+': '+names.join(', '));
  const name=names[0],bytes=new Uint8Array(fs.readFileSync(path.join(dir,name)));
  return {name,bytes,wb:XLSX.read(bytes,{type:'array',cellFormula:true}),path:path.join(folder,name),
    hash:crypto.createHash('sha256').update(bytes).digest('hex')};
}
const files={
  rawSession:read('DataQuyTac/File thô','Download trên web về_Session'),
  rawCcvn:read('DataQuyTac/File thô','Download trên web về_Daily'),
  op:read('DataQuyTac/File thô','File báo cáo ngày'),
  cleanSession:read('DataQuyTac/File thô đã qua xử lý','Session'),
  cleanCcvn:read('DataQuyTac/File thô đã qua xử lý','Daily'),
  reference:read('DataQuyTac/Output','File báo cáo tổng'),
  base:read('Data/test','File báo cáo tổng_AltiuslinkVN_Daily Report_(New)-20260912')
};
const cfg=AppConfig.defaults(),session=Pipeline.processSession(files.rawSession.wb,cfg,files.rawSession.name),
  ccvn=Pipeline.processCcvn(files.rawCcvn.wb,cfg,files.rawCcvn.name),
  cleanCcvn=Pipeline.processCcvn(files.cleanCcvn.wb,cfg,files.cleanCcvn.name);
const day=Pipeline.detectReportDate(session,cfg),serial=day.serial;
const ops=[Pipeline.processOpFile(files.op.wb,cfg,serial,files.op.name)];
const mailWarnings=[],mail=Pipeline.inspectMailSf(files.base.wb,cfg,mailWarnings,files.base.name);
const input={session,ccvn,ops,emailTable:mail.table},blocks=Pipeline.buildBlocks(input,cfg,serial);
const rosterWarnings=[];
Pipeline.checkRoster(input,cfg,files.base.wb,blocks,rosterWarnings);
function norm(v){return String(v??'').replace(/\r\r\n|\r\n|\r/g,'\n').trim();}
function same(a,b){return norm(a)===norm(b)||(typeof a==='number'&&typeof b==='number'&&Math.abs(a-b)<1e-9);}
function compareMatrices(actual,expected,headers,offset={r:0,c:0}){
  const byColumn={},differences=[];
  for(let r=0;r<Math.max(actual.length,expected.length);r++)for(let c=0;c<headers.length;c++){
    const a=actual[r]?.[c],b=expected[r]?.[c];if(same(a,b))continue;
    const delta={cell:XLSX.utils.encode_cell({r:r+offset.r,c:c+offset.c}),column:headers[c],actual:a??'',expected:b??''};
    differences.push(delta);const slot=byColumn[c]??={header:headers[c],count:0,pairs:{},samples:[]};slot.count++;
    const pair=norm(a)+' → '+norm(b);slot.pairs[pair]=(slot.pairs[pair]||0)+1;
    if(slot.samples.length<5)slot.samples.push(delta);
  }
  return {actualRows:actual.length,expectedRows:expected.length,count:differences.length,
    byColumn:Object.values(byColumn).map(col=>({...col,pairs:Object.entries(col.pairs).sort((a,b)=>b[1]-a[1]).slice(0,15)})),differences};
}
function blockRows(file,block){
  const a=XLSX.utils.decode_cell(block.anchor),m=Parse.sheetToMatrix(file.wb.Sheets[block.sheet]);
  const rows=m.rows.slice(a.r).map(row=>Array.from({length:block.headers.length},(_,c)=>row[a.c+c]??''));
  while(rows.length&&rows.at(-1).every(v=>v===''))rows.pop();return rows;
}
function compact(result){const {differences,...summary}=result;return summary;}
const stages={};
const rawId=Parse.findColumn(ccvn.headers,'Case Number'),cleanId=Parse.findColumn(cleanCcvn.headers,'Case Number');
const cleanCases=new Map(cleanCcvn.dataRows.map(row=>[norm(row[cleanId]),row]));
const matched=ccvn.dataRows.map(row=>cleanCases.get(norm(row[rawId]))||[]);
const aligned=matched.map(row=>ccvn.headers.map(h=>row[Parse.findColumn(cleanCcvn.headers,h)]??''));
stages.salesforceRawToClean=compareMatrices(ccvn.dataRows,aligned,ccvn.headers,{r:ccvn.headerIndex+1,c:0});
const stagingSheet=files.cleanSession.wb.SheetNames.find(n=>n!==session.sheetName);
const stagingRows=Parse.sheetToMatrix(files.cleanSession.wb.Sheets[stagingSheet]).rows.slice(1).filter(r=>r.some(v=>norm(v)!==''));
stages.sessionStaging=compareMatrices(session.staging.rows,stagingRows,session.staging.headers,{r:1,c:0});
for(const key of ['block1','block2','block3'])stages[key]=compareMatrices(blocks[key].rows,blockRows(files.reference,blocks[key]),blocks[key].headers,XLSX.utils.decode_cell(blocks[key].anchor));
const patched=XlsxPatch.patch(files.base.bytes,blocks,cfg,{masterWb:files.base.wb});
const checks=SelfTest.compareWithReference(patched.bytes,files.reference.wb,blocks,cfg,{baseWb:files.base.wb,input,affectedEmails:mail.affected,oddRateEmails:patched.oddRateEmails});
const inventory=Object.fromEntries(Object.entries(files).map(([key,f])=>[key,{path:f.path,sha256:f.hash,
  sheets:f.wb.SheetNames.map(name=>({name,range:f.wb.Sheets[name]['!ref'],sample:Parse.sheetToMatrix(f.wb.Sheets[name]).rows.slice(0,key==='rawCcvn'||key==='cleanCcvn'?8:2).map(row=>row.map(v=>String(v??'').slice(0,100)))}))}]));
const oldFiles={
  rawSession:read('Data/input/13092026','Session'),rawCcvn:read('Data/input/13092026','Daily'),
  op:read('Data/input/13092026','File báo cáo ngày'),reference:read('Data/input/13092026','File báo cáo tổng'),
  cleanSession:read('Data/output','Session'),cleanCcvn:read('Data/output','Daily')
};
const report={day,inventory,identicalToOldData:Object.fromEntries(Object.keys(oldFiles).map(k=>[k,files[k].hash===oldFiles[k].hash])),
  stats:{session:session.stats,ccvn:ccvn.stats,cleanCcvn:cleanCcvn.stats,ops:ops.map(o=>o.stats)},
  stages:Object.fromEntries(Object.entries(stages).map(([k,v])=>[k,compact(v)])),checks,
  warnings:mailWarnings.concat(rosterWarnings,session.warnings,ccvn.warnings,...ops.map(o=>o.warnings),blocks.warnings).map(w=>({level:w.level,message:w.message,detail:w.detail,display:w.display,spots:w.spots})),
  patchReport:patched.report};
if(require.main===module){
  fs.mkdirSync(path.join(__dirname,'artifacts'),{recursive:true});
  fs.writeFileSync(path.join(__dirname,'artifacts','dataquytac-audit.json'),JSON.stringify({...report,fullDiffs:Object.fromEntries(Object.entries(stages).map(([k,v])=>[k,v.differences]))},null,2));
  console.log(JSON.stringify(process.argv.includes('--inventory')?{day,inventory,identicalToOldData:report.identicalToOldData}:
    {day,identicalToOldData:report.identicalToOldData,stats:report.stats,stages:report.stages,checks},null,2));
}
module.exports={ctx,files,cfg,serial,session,ccvn,cleanCcvn,ops,mail,input,blocks,patched,report,stages,blockRows,same};
