// Read-only source audit. JSON contains aggregate statistics, never customer data.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = path.resolve(__dirname, '../..');
const ctx = vm.createContext({console, Uint8Array, ArrayBuffer, TextEncoder, TextDecoder, Date});
for (const f of ['vendor/xlsx.full.min.js','src/config.js','src/parse.js','src/pipeline.js']) {
  vm.runInContext(fs.readFileSync(path.join(root,'tool',f),'utf8'),ctx,{filename:f});
}
const {XLSX, Parse, Pipeline, AppConfig} = ctx;
const folder = path.join(root,'Tổng hợp dữ liệu tháng');
const dailyFolder = path.join(folder,'Tháng 09');
function read(file) {
  return XLSX.read(new Uint8Array(fs.readFileSync(file)),{type:'array',sheets:['1. CSKH'],cellFormula:false,cellHTML:false,cellText:false});
}
function chatRows(wb) {
  return Parse.sheetToMatrix(wb.Sheets['1. CSKH']).rows.filter(r=>String(r[2]).trim()==='CHAT BOT');
}
function key(row) { return JSON.stringify([row[4],row[6],String(row[10]??'').trim()]); }
function times(rows) {
  let increasing=0, comparisons=0, invalid=0, overHour=0, negative=0;
  const minutes=[];
  rows.forEach((r,i)=>{
    const start=Date.parse(r[5]),end=Date.parse(r[6]);
    if(!Number.isFinite(start)||!Number.isFinite(end)){invalid++;return;}
    const dt=(end-start)/60000;minutes.push(dt);
    if(dt>60)overHour++;if(dt<0)negative++;
    if(i&&Number.isFinite(Date.parse(rows[i-1][5]))){comparisons++;if(start>Date.parse(rows[i-1][5]))increasing++;}
  });
  minutes.sort((a,b)=>a-b);
  return {rows:rows.length,medianMinutes:minutes[Math.floor(minutes.length/2)],overHour,negative,invalid,increasing,comparisons};
}
const fullName=fs.readdirSync(folder).find(n=>n.endsWith('.xlsx'));
const fullRows=chatRows(read(path.join(folder,fullName)));
const monthHistogram={};
const fullByDay=new Map();
for(const row of fullRows){
  const day=Parse.toDateSerial(row[1]);
  const month=String(row[5]??'').slice(0,7);monthHistogram[month]=(monthHistogram[month]||0)+1;
  if(!fullByDay.has(day))fullByDay.set(day,[]);fullByDay.get(day).push(row);
}
const reports=[];
for(const file of fs.readdirSync(dailyFolder).filter(n=>n.endsWith('.xlsx')).sort()){
  const day=Number(file.match(/202609(\d\d)/)[1]);
  const rows=chatRows(read(path.join(dailyFolder,file)));
  const serial=Parse.toDateSerial('2026-09-'+String(day).padStart(2,'0'));
  const aggregateRows=fullByDay.get(serial)||[];
  const index=new Map();
  for(const r of aggregateRows){const k=key(r);if(!index.has(k))index.set(k,[]);index.get(k).push(r);}
  let matched=0,equalStart=0,ambiguous=0;
  for(const r of rows){const candidates=index.get(key(r));if(!candidates)continue;if(candidates.length!==1){ambiguous++;continue;}matched++;if(r[5]===candidates[0][5])equalStart++;}
  const report={day,...times(rows),fullMonth:{rows:aggregateRows.length,matched,equalStart,ambiguous}};
  const sourceDir=path.join(root,'Data/input',String(day).padStart(2,'0')+'092026');
  if(fs.existsSync(sourceDir)){
    const sessionFile=fs.readdirSync(sourceDir).find(n=>n.startsWith('Session'));
    if(sessionFile){
      const input=XLSX.read(new Uint8Array(fs.readFileSync(path.join(sourceDir,sessionFile))),{type:'array'});
      const session=Pipeline.processSession(input,AppConfig.defaults(),sessionFile);
      const idx=n=>Parse.findColumn(session.headers,n);
      const map=new Map();
      for(const r of session.keptRows){const k=JSON.stringify([r[idx('Supporter')],r[idx('End time')],String(r[idx('Sender id')]??'').trim()]);if(!map.has(k))map.set(k,[]);map.get(k).push(r);}
      const fields=['Requested time','Assign time','First response time'];
      const exact={sourceRows:session.keptRows.length,matched:0,ambiguous:0,unmatched:0,fields:Object.fromEntries(fields.map(n=>[n,0]))};
      for(const r of rows){const candidates=map.get(key(r));if(!candidates){exact.unmatched++;continue;}if(candidates.length!==1){exact.ambiguous++;continue;}exact.matched++;for(const n of fields)if(r[5]===candidates[0][idx(n)])exact.fields[n]++;}
      report.rawSource=exact;
    }
  }
  reports.push(report);
  console.error('Audited day '+day);
}
const result={fullMonthFile:fullName,fullMonthStartMonths:monthHistogram,
  dailyFiles:reports.length,totalChatRows:reports.reduce((n,r)=>n+r.rows,0),
  fullMonthMatched:reports.reduce((n,r)=>n+r.fullMonth.matched,0),reports};
const artifacts=path.join(__dirname,'artifacts');
fs.mkdirSync(artifacts,{recursive:true});
fs.writeFileSync(path.join(artifacts,'month-times-audit.json'),JSON.stringify(result,null,2));
console.log(JSON.stringify(result,null,2));
