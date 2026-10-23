// Exercise the real UI handlers against a minimal DOM; no browser/network needed.
const fs=require('fs'),path=require('path'),vm=require('vm'),assert=require('assert/strict');
const tool=path.resolve(__dirname,'..');
function element(){
  const listeners={},classes=new Set();
  return {value:'',style:{},children:[],listeners,disabled:false,
    set innerHTML(value){this.children=[];},get innerHTML(){return '';},
    classList:{add:c=>classes.add(c),remove:c=>classes.delete(c),toggle:(c,on)=>on?classes.add(c):classes.delete(c)},
    appendChild(child){this.children.push(child);},querySelector(){return element();},setAttribute(k,v){this[k]=v;},
    addEventListener(type,fn){listeners[type]=fn;}};
}
const ids=new Map([...fs.readFileSync(path.join(tool,'index.html'),'utf8').matchAll(/id="([^"]+)"/g)].map(m=>[m[1],element()]));
const storage=new Map();
const ctx=vm.createContext({console,setTimeout:()=>1,clearTimeout:()=>{},
  document:{getElementById:id=>{assert(ids.has(id),'Unknown HTML id '+id);return ids.get(id);},
    createElement:()=>element(),querySelectorAll:()=>[]},
  localStorage:{getItem:k=>storage.get(k),setItem:(k,v)=>storage.set(k,v)},
  addEventListener:()=>{},postMessage:()=>{}});
ctx.window=ctx;
for(const f of ['config.js','parse.js'])vm.runInContext(fs.readFileSync(path.join(tool,'src',f),'utf8'),ctx);
const old=ctx.AppConfig.defaults();delete old.session.startTimeMappingVersion;
old.rulesVersion=3;old.requestedByAliases['LUU NGOC QUANG']='NGUYEN THAO NGUYEN';
old.session.stagingColumns[3]={out:'Requested time',from:'Requested time'};
storage.set('ccvn-report-tool.config.v1',JSON.stringify(old));
const source=fs.readFileSync(path.join(tool,'src/app.js'),'utf8');
assert(source.includes("document.addEventListener('DOMContentLoaded', init);"));
vm.runInContext(source.replace("document.addEventListener('DOMContentLoaded', init);",'globalThis.testState=state; init();'),ctx);
assert.equal(ids.get('start-time-source').value,'Requested time');
assert.equal(ids.get('al-list').children.length,2,'Two confirmed aliases; disabled Hoai must not render');
assert.equal(ctx.testState.cfg.rulesVersion,6);
assert.equal(ctx.testState.cfg.requestedByAliases['LUU NGOC QUANG'],'NGUYEN THAO NGUYEN');
const select=ids.get('start-time-source');
// Preserve an unsaved JSON edit while using the quick selector.
const edited=JSON.parse(ids.get('cfg-text').value);edited.op.acceptSwappedDates=true;
ids.get('cfg-text').value=JSON.stringify(edited);
ctx.testState.blocks={};
select.value='Requested time';select.listeners.change({target:select});
assert.equal(ctx.testState.cfg.session.stagingColumns[3].from,'Requested time');
assert.equal(ctx.testState.cfg.op.acceptSwappedDates,true);
assert.equal(ctx.testState.stale,true);
assert.equal(ids.get('btn-all').disabled,true);
const saved=JSON.parse(storage.get('ccvn-report-tool.config.v1'));
assert.equal(ctx.AppConfig.migrate(saved).session.stagingColumns[3].from,'Requested time');
assert.equal(saved.session.startTimeMappingVersion,3);
ids.get('cfg-text').value='{invalid';select.value='Assign time';select.listeners.change({target:select});
assert.equal(select.value,'Requested time','Invalid JSON must not silently overwrite config');
assert.equal(ctx.testState.cfg.session.stagingColumns[3].from,'Requested time');
ids.get('btn-cfg-reset').listeners.click();
assert.equal(select.value,'Requested time');
assert.equal(JSON.parse(ids.get('cfg-text').value).session.stagingColumns[3].from,'Requested time');
ctx.testState.blocks=null;
ids.get('al-from').value='LUU NGOC QUANG';ids.get('al-to').value='CUSTOM PERSON';
ids.get('al-add').listeners.click();
ids.get('al-list').children.find(row=>row.children.some(c=>c.textContent==='LUU NGOC QUANG'))
  .children.find(c=>c.className==='al-del').listeners.click();
assert.equal(ctx.AppConfig.migrate(ctx.testState.cfg).requestedByAliases['LUU NGOC QUANG'],false,'Deleted default alias must stay disabled after reload');
ids.get('al-list').children[0].children.find(c=>c.className==='al-del').listeners.click();
ids.get('al-from').value='vu nhu quynh';ids.get('al-to').value='CUSTOM PERSON';
ids.get('al-add').listeners.click();
assert.equal(ctx.testState.cfg.requestedByAliases['VU NHU QUYNH'],'CUSTOM PERSON','A disabled alias can be replaced, including case-insensitive names');
assert.equal(ids.get('al-list').children.length,1);
console.log('PASS: UI config migration, source selector, unsaved JSON preservation, storage, stale-output lock, invalid JSON, reset, disabled aliases and persistent deletion.');
