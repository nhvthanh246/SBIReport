// Source formatting in warnings and annotated XLSX; raw input is never changed.
const fs=require('fs'),path=require('path'),vm=require('vm'),assert=require('assert/strict');
const root=path.resolve(__dirname,'../..'),ctx=vm.createContext({console,Uint8Array,ArrayBuffer,TextEncoder,TextDecoder,Date});
for(const file of ['vendor/xlsx.full.min.js','vendor/fflate.min.js','src/parse.js','src/config.js','src/pipeline.js','src/markup.js'])
  vm.runInContext(fs.readFileSync(path.join(root,'tool',file),'utf8'),ctx);
const {XLSX,Parse,Pipeline,AppConfig,Markup}=ctx;
const read=bytes=>XLSX.read(bytes,{type:'array',cellNF:true,cellFormula:true,cellStyles:true});
function load(folder){const dir=path.join(root,folder),name=fs.readdirSync(dir).find(n=>n.includes('SBI')&&n.endsWith('.xlsx'));
  const bytes=new Uint8Array(fs.readFileSync(path.join(dir,name)));return {name,bytes,wb:read(bytes)};}
const source=load('DataQuyTac/File thô'),sheet='3. Other Tasks & Working time';
const before=JSON.stringify(source.wb);
assert.equal(Parse.sourceCellText(source.wb,sheet,2026,2,46278),'13/09/2026');
assert.equal(Parse.sourceCellText(source.wb,sheet,2026,12,0),'0.00');
assert.equal(Parse.sourceCellText(source.wb,'1. CSKH',3493,1,'wrong fallback'),'');
const lazy=read(source.bytes);delete lazy.Sheets[sheet].C2026.w;delete lazy.Sheets[sheet].M2026.w;
assert.equal(Parse.sourceCellText(lazy,sheet,2026,2,46278),'13/09/2026');
assert.equal(Parse.sourceCellText(lazy,sheet,2026,12,0),'0.00');
const synthetic={Sheets:{Cells:{A1:{t:'n',v:46278,z:'General'},B1:{t:'s',v:'26793145563657353'},
  C1:{t:'n',v:0,z:';;;'},D1:{t:'n',v:.25,z:'0%'},E1:{t:'n',v:0,z:'dd/mm/yyyy'}}},Workbook:{WBProps:{date1904:true}}};
assert.equal(Parse.sourceCellText(synthetic,'Cells',1,0,''),'46278','Do not guess that any 46xxx number is a date');
assert.equal(Parse.sourceCellText(synthetic,'Cells',1,1,''),'26793145563657353');
assert.match(Parse.sourceCellText(synthetic,'Cells',1,2,''),/định dạng đang ẩn/);
assert.equal(Parse.sourceCellText(synthetic,'Cells',1,3,''),'25%');
assert.equal(Parse.sourceCellText(synthetic,'Cells',1,4,''),'01/01/1904');
assert.equal(Parse.sourceCellText(null,'Missing',1,0,'<script>'),'<'+'script>');
const entries=Pipeline.collectSpots([[{level:'info',message:'Đã xử lý',spots:[
  {file:source.name,sheet,row:2026,col:2,value:'46278',why:'Kiểm tra ngày'},
  {file:source.name,sheet,row:2026,col:12,value:'0',why:'Để trống số 0'},
  {file:source.name,sheet:'1. CSKH',row:3493,col:1,value:'',why:'Chưa nhập ngày'}]}]]);
let checked=0;
function verifySource(a,b){
  for(const name of a.SheetNames)for(const [addr,cell] of Object.entries(a.Sheets[name])){
    if(addr.startsWith('!'))continue;
    for(const field of ['v','t','f','z','w'])assert.deepEqual(b.Sheets[name][addr]?.[field],cell[field],name+'!'+addr+' '+field);
    checked++;
  }
}
for(const opts of [{},{sourceWb:source.wb}]){
  const result=Markup.build(source.bytes,entries[0],opts),out=read(result.bytes),list=out.Sheets[result.sheetTitle];
  assert.equal(list.D4.v,'Nội dung ô trong file gốc');assert.equal(list.E4.v,'Trạng thái');
  assert.equal(list.D5.v,'13/09/2026');assert.equal(list.D6.v,'0.00');assert.equal(list.D7.v,'(trống)');
  assert.equal(list.E5.v,'Đã xử lý — xem ghi chú');verifySource(source.wb,out);
  const repeated=read(Markup.build(result.bytes,entries[0]).bytes);verifySource(source.wb,repeated);
  assert.equal(out.Sheets[sheet].C2026.s.fgColor.rgb,'FFFFF2A8'.slice(2),'Still highlighted');
}
assert.equal(JSON.stringify(source.wb),before,'Formatting must not mutate the in-memory workbook');
const october=load('ExtraData/Daily Report 11.10.2026/Input'),cfg=AppConfig.defaults();
const op=Pipeline.processOpFile(october.wb,cfg,Parse.toDateSerial('2026-10-11'),october.name);
const swapped=op.warnings.find(w=>w.spots?.some(s=>s.row===7466&&s.sheet==='1. CSKH'));
const row=swapped.spots.find(s=>s.row===7466);assert.equal(row.value,'46336','Keep raw evidence for diagnostics');
assert.match(row.why,/ngày 10 tháng 11 năm 2026/);assert.match(row.why,/ngày 11 tháng 10 năm 2026/);
assert.doesNotMatch(swapped.detail,/46336/);assert.doesNotMatch(row.why,/46336/);
const entry=Pipeline.collectSpots([[swapped]])[0],result=Markup.build(october.bytes,entry),out=read(result.bytes);
const index=entry.sheets['1. CSKH'].findIndex(s=>s.row===7466)+5;
assert.equal(out.Sheets[result.sheetTitle]['D'+index].v,'11/10/26');
assert.equal(out.Sheets[result.sheetTitle]['E'+index].v,'Cần kiểm tra');
assert.equal(out.Sheets['1. CSKH'].B7466.w,'11/10/26','Yellow fill must not turn a date into 46336');
verifySource(october.wb,out);
const artifacts=path.join(__dirname,'artifacts');fs.mkdirSync(artifacts,{recursive:true});
fs.writeFileSync(path.join(artifacts,'warning-values-20261011.xlsx'),result.bytes);
console.log('PASS: readable dates/decimals, blank/hidden cells, 1904 dates, large IDs, XLSX list and repeated highlighting; '+checked+' source-cell checks preserve values, formulas and formatting.');
