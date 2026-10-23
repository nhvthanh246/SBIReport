// Audit report rules against source tables inside all daily reports.
// Only diagnostic artifacts are written. No source workbook or app code is changed.
const fs=require('fs'),path=require('path'),vm=require('vm');
const root=path.resolve(__dirname,'../..');
const ctx=vm.createContext({console,Uint8Array,ArrayBuffer,TextEncoder,TextDecoder,Date});
for(const f of ['vendor/xlsx.full.min.js','src/config.js','src/parse.js','src/pipeline.js'])
  vm.runInContext(fs.readFileSync(path.join(root,'tool',f),'utf8'),ctx,{filename:f});
const {XLSX,Parse:P,Pipeline:Pipe,AppConfig}=ctx,cfg=AppConfig.defaults();
const dir=path.join(root,'Tổng hợp dữ liệu tháng','Tháng 09');
function load(file){return XLSX.read(new Uint8Array(fs.readFileSync(file)),{type:'array',cellFormula:true,cellHTML:false,cellText:false});}
function data(wb,sheet,anchor,width){
  const a=XLSX.utils.decode_cell(anchor),m=P.sheetToMatrix(wb.Sheets[sheet]).rows;
  const rows=m.slice(a.r).map(r=>Array.from({length:width},(_,c)=>r[a.c+c]??''));
  while(rows.length&&rows.at(-1).every(v=>v===''))rows.pop();return {rows};
}
function near(a,b){return Math.abs(Number(a??0)-Number(b??0))<1e-8;}
function val(ws,r,c){return ws[XLSX.utils.encode_cell({r:r-1,c})]?.v;}
function dayCol(wb,sheet,header,serial){return (P.sheetToMatrix(wb.Sheets[sheet]).rows[header-1]||[]).findIndex(v=>P.toDateSerial(v)===serial);}
const opName=fs.readdirSync(path.join(root,'Data/input')).find(n=>n.startsWith('SBI'));
const op=load(path.join(root,'Data/input',opName));
const days=[];
for(const file of fs.readdirSync(dir).filter(n=>n.endsWith('.xlsx')).sort()){
  const day=Number(file.match(/202609(\d\d)/)[1]),serial=P.toDateSerial('2026-09-'+String(day).padStart(2,'0'));
  const wb=load(path.join(dir,file));
  const blocks={block1:data(wb,cfg.targets.cskh.sheet,cfg.targets.cskh.anchor,14),
    block2:data(wb,cfg.targets.sfcase.sheet,cfg.targets.sfcase.anchor,7),
    block3:data(wb,cfg.targets.other.sheet,cfg.targets.other.anchor,11)};
  const summary=Pipe.computeSummary(blocks.block1,cfg,[]),kpi=Pipe.computeKpi(blocks,cfg);
  const s=wb.Sheets[cfg.summary.sheet],col=dayCol(wb,cfg.summary.sheet,4,serial);
  const sumDiff=[];
  for(const spec of cfg.summary.cells){
    let expected;
    if(spec.from==='customerFormula')expected=summary.distinct+['zalo','viber','line','tel'].reduce((n,k)=>n+(summary.counts[k]||0),0);
    else if(spec.from==='ccvnCount')expected=blocks.block2.rows.length;
    else expected=summary.counts[spec.from]||0;
    const actual=val(s,spec.row,col);
    if(!near(expected,actual))sumDiff.push({row:spec.row,source:spec.from,expected,actual});
  }
  for(const spec of cfg.summary.topicRows){
    const expected=summary.topics[spec.topic]||0,actual=val(s,spec.row,col);
    if(!near(expected,actual))sumDiff.push({row:spec.row,source:'topic '+spec.topic,expected,actual});
  }
  const daily=wb.Sheets[cfg.kpi.daily.sheet],month=wb.Sheets['202609 KPI Result'];
  const monthCol=dayCol(wb,'202609 KPI Result',5,serial);
  const monthlyByEmail=new Map();for(let r=6;r<=60;r++){const e=P.normText(val(month,r,3));if(e.includes('@'))monthlyByEmail.set(e,{r,v:val(month,r,monthCol)});}
  const kpiDiff=[],dailyMonthDiff=[],missingHours=[],emails=[];let checked=0,rateSkipped=0;
  for(let r=15;r<=60;r++){
    const e=P.normText(val(daily,r,3));if(!e.includes('@'))continue;
    if(r<=43)emails.push(e);
    const computed=kpi.forEmail(e),actual=val(daily,r,22),monthly=monthlyByEmail.get(e);
    const sameRates=cfg.kpi.daily.rateColumns.every((c,i)=>{
      const cell=daily[XLSX.utils.encode_cell({r:r-1,c})],m=cell?.f?.match(/\*\s*(\d+(?:\.\d+)?)\s*\)/);
      return m&&Number(m[1])===cfg.kpi.tracks[i].rate;
    });
    if(sameRates){checked++;if(!near(computed.W,actual))kpiDiff.push({email:e,row:r,computed:computed.W,actual,E:computed.E,actualE:val(daily,r,4),hours:computed.hours});}
    else rateSkipped++;
    if(monthly&&!near(actual,monthly.v))dailyMonthDiff.push({email:e,dailyRow:r,monthRow:monthly.r,daily:actual,month:monthly.v});
    if(computed.hasWork&&!computed.hours)missingHours.push({email:e,actual});
  }
  const helperDiff=[],sets=[];
  const sheet5=wb.Sheets.Sheet5;
  for(let r=5;r<=70;r++){
    const e=P.normText(val(sheet5,r,0));if(!e.includes('@'))continue;
    const expected=kpi.forEmail(e).E,actual=val(sheet5,r,14);
    if(!near(expected,actual))helperDiff.push({email:e,row:r,expected,actual,formula:sheet5['O'+r]?.f});
    if(kpi.p1[e]?.sets)sets.push({email:e,count:kpi.p1[e].sets,V:val(sheet5,r,21),formula:sheet5['V'+r]?.f,O:actual,
      otherHours:blocks.block3.rows.filter(row=>P.normText(row[2])===e&&P.normText(row[9]).includes('setsumekai')).map(row=>row[10])});
  }
  const names={};for(const row of blocks.block2.rows){const n=P.normName(row[4]);names[n]=(names[n]||0)+1;}
  const aliases=Object.entries(cfg.requestedByAliases).map(([from,to])=>({from,to,sourceNameCases:names[P.normName(from)]||0,targetNameCases:names[P.normName(to)]||0}));
  const input=Pipe.processOpFile(op,cfg,serial,opName);
  function hoursMap(rows){const map=new Map();for(const row of rows){const e=P.normText(row[2]);if(!e.includes('@'))continue;if(!map.has(e))map.set(e,[]);map.get(e).push(row);}return map;}
  const refHours=hoursMap(blocks.block3.rows);
  const srcHours=hoursMap(input.other.rows.map(rec=>cfg.op.other.copyColumns.map(c=>rec.cells[c]??'')));
  const missingPeople=[...refHours.keys()].filter(e=>!srcHours.has(e));
  const extraPeople=[...srcHours.keys()].filter(e=>!refHours.has(e));
  const hoursDiff=[];
  for(const [e,rows] of refHours){const source=srcHours.get(e);if(!source)continue;
    const a=rows.reduce((n,r)=>n+cfg.kpi.tracks.reduce((v,t)=>v+(Number(r[t.hoursColumn])||0),0),0);
    const b=source.reduce((n,r)=>n+cfg.kpi.tracks.reduce((v,t)=>v+(Number(r[t.hoursColumn])||0),0),0);
    if(!near(a,b))hoursDiff.push({email:e,referenceKpiHours:a,inputKpiHours:b});
  }
  const mail=Pipe.inspectMailSf(wb,cfg,[]);
  const result={day,rows:[blocks.block1.rows.length,blocks.block2.rows.length,blocks.block3.rows.length],
    summaryDiff:sumDiff,unknownChannels:summary.unknownChannels,kpiChecked:checked,rateSkipped,kpiDiff,dailyMonthDiff,
    team:{computed:kpi.team(emails),daily:val(daily,14,22),monthly:val(month,7,monthCol)},
    missingHours,helperDiff,sets,aliases,op:{rows:input.other.rows.length,missingPeople,extraPeople,hoursDiff},
    mailCorrections:mail?.corrections.length,mailDuplicates:mail?.duplicates.length};
  const sourceDir=path.join(root,'Data/input',String(day).padStart(2,'0')+'092026');
  if(fs.existsSync(sourceDir)){
    const sfFile=fs.readdirSync(sourceDir).find(n=>n.startsWith('Daily'));
    if(sfFile){
      const cc=Pipe.processCcvn(load(path.join(sourceDir,sfFile)),cfg,sfFile);
      const ix=n=>P.findColumn(cc.headers,n),refByCase=new Map(blocks.block2.rows.map(r=>[String(r[0]),r]));
      const raw={matched:0,statusChanges:{},customerDiff:0,customerCaseOnly:0,aliases:{}};
      for(const row of cc.dataRows){const ref=refByCase.get(String(row[ix('Case Number')]));if(!ref)continue;raw.matched++;
        const status=row[ix('Case Status')];if(status!==ref[2]){const k=ref[2]+' -> '+status;raw.statusChanges[k]=(raw.statusChanges[k]||0)+1;}
        const customer=String(row[ix('Customer information')]??'');if(customer!==String(ref[1])){raw.customerDiff++;if(customer.toLowerCase()===String(ref[1]).toLowerCase())raw.customerCaseOnly++;}
        const from=P.normName(row[ix('Created By')]);if(cfg.requestedByAliases[from]){const to=P.normName(ref[4]);const k=from+' -> '+to;raw.aliases[k]=(raw.aliases[k]||0)+1;}
      }result.raw=raw;
    }
  }
  days.push(result);console.error('Audited rules day '+day);
}
const result={days};
fs.mkdirSync(path.join(__dirname,'artifacts'),{recursive:true});
fs.writeFileSync(path.join(__dirname,'artifacts','month-rules-audit.json'),JSON.stringify(result,null,2));
console.log(JSON.stringify(days.map(d=>({day:d.day,summary:d.summaryDiff.length,kpi:d.kpiDiff.length,checked:d.kpiChecked,
  dailyMonth:d.dailyMonthDiff.length,helper:d.helperDiff.length,sets:d.sets.length,missingOP:d.op.missingPeople.length,hoursDiff:d.op.hoursDiff.length})),null,2));
