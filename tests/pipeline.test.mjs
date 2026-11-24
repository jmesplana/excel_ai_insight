import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {webcrypto} from 'node:crypto';
import {defaultWorkflow, defaultNode, portableConfig, importWorkflow, seedRun, nextWorkIndex, applyResult, retryEntry, reviewCell, scheduleRateRetry, mergeRun, validateMappings, needsLegacyUpgrade, upgradeLegacyWorkflow, activeStep, locateEntry, rewindCursor, acceptReviewerAnswer, reviewerFields, restoreEntry, discardEntries, sourceSiblings} from '../static/js/pipeline-core.js';
import {readSheet, readMaster, buildWorkbook} from '../static/js/pipeline-workbook.js';
if (!globalThis.crypto) globalThis.crypto=webcrypto;
const context=vm.createContext({});vm.runInContext(fs.readFileSync('static/vendor/xlsx.full.min.js','utf8'),context);const xlsx=context.XLSX;
const source=()=>({filename:'team.xlsx',datasetId:'hash',sheet:'Sheet1',columns:['ID','Feedback'],rows:[{excelRow:2,values:{ID:1,Feedback:'Merci.Où ?Quand ?'}},{excelRow:4,values:{ID:2,Feedback:123}}]});
const pairs=[{feedback_excerpt:'Merci.',type_of_feedback:'Appréciations'},{feedback_excerpt:'Où ?',type_of_feedback:'Questions'},{feedback_excerpt:'Quand ?',type_of_feedback:'Questions'}];

test('one source becomes three linked rows and numeric input remains review, IDs unique',()=>{
 const run=seedRun(defaultWorkflow(),source());
 applyResult(run,0,{status:'ok',items:pairs});
 applyResult(run,3,{status:'review',reason:'Insufficient context'});
 assert.equal(run.work.length,4);
 assert.deepEqual(run.work.slice(0,3).map(e=>e.values.ID),[1,1,1]);
 assert.deepEqual(run.work.slice(0,3).map(e=>e.values._message_number),['1','2','3']);
 assert.equal(new Set(run.work.map(e=>e.values._feedback_id)).size,4);
 assert.equal(run.work[3].values._source_row,4);
 const merged=mergeRun(run);assert.equal(merged.rows.length,3);assert.equal(merged.review.length,1);
});

test('correction retains numeric ID and original source while retrying failed step',()=>{
 const run=seedRun(defaultWorkflow(),source());applyResult(run,0,{status:'review',reason:'Bad extraction'});
 const id=run.work[0].values._feedback_id;
 retryEntry(run,0,{Feedback:'Merci.'});applyResult(run,0,{status:'ok',items:[pairs[0]]});
 assert.equal(run.work[0].values._feedback_id,id);
 assert.equal(JSON.parse(run.work[0].values._original_record).Feedback,'Merci.Où ?Quand ?');
 assert.equal(run.work[0].history[1].status,'corrected');
});

test('duplicate imported records skipped, identical text from another row kept',()=>{
 const run=seedRun(defaultWorkflow(),source());applyResult(run,0,{status:'ok',items:pairs});applyResult(run,3,{status:'review',reason:'No context'});
 const merged=mergeRun(run), next=source();next.rows.push({excelRow:5,values:{ID:3,Feedback:'Merci.Où ?Quand ?'}});
 const nextRun=seedRun(defaultWorkflow(),next,merged);
 assert.equal(nextRun.skipped,2);assert.equal(nextRun.work.length,1);assert.equal(nextRun.nextId,5);
});

test('partial master cannot silently skip unfinished siblings on reimport',()=>{
 const config=defaultWorkflow();config.nodes.push({...defaultNode('analyze'),inputs:['feedback_excerpt'],fields:[{name:'translation',options:[]}]});
 const run=seedRun(config,source());applyResult(run,0,{status:'ok',items:pairs});applyResult(run,0,{status:'ok',items:[{translation:'Thanks'}]});
 assert.throws(()=>seedRun(config,source(),mergeRun(run)),/unfinished import/);
});

test('column remapping, collisions and secret-free portable configuration',()=>{
 const config=defaultWorkflow();config.mappings={Feedback:'Comment'};config.idColumn='Case';
 const src={...source(),columns:['Case','Comment'],rows:[{excelRow:7,values:{Case:'original',Comment:'Merci.'}}]};
 const run=seedRun(config,src);assert.equal(run.work[0].values.Feedback,'Merci.');assert.equal(run.work[0].values._original_id,'original');
 assert.throws(()=>validateMappings(config,['Case','Comment','Feedback']),/overwrite/);
 config.apiKey='secret';config.nodes[0].apiKey='secret';assert.equal(JSON.stringify(portableConfig(config)).includes('secret'),false);
 assert.throws(()=>validateMappings(defaultWorkflow(),['ID','Feedback','feedback_excerpt']),/overwrite/);
});

test('actual step 2 codebook converts to ordered classification and deterministic lookups',()=>{
 const raw=JSON.parse(fs.readFileSync('examples/examples_1/step2-ebola-feedback-coding.excerpt.input-column.config.json'));
 const {config}=importWorkflow(raw);
 assert.equal(config.nodes.filter(n=>n.type==='classify').length,7);
 assert.equal(config.nodes.filter(n=>n.type==='lookup').length,6);
 assert.equal(config.nodes[0].sourceBranches.length,7);
 assert.ok(config.nodes[1].inputs.includes('feedback_path'));
});

test('physical row numbers survive blank rows and offset headers',()=>{
 const wb=xlsx.utils.book_new();xlsx.utils.book_append_sheet(wb,xlsx.utils.aoa_to_sheet([['Title'],['ID','Feedback'],[1,'Thanks'],[],[2,'Why?']]),'Feedback');
 const read=readSheet(wb,'Feedback',2,xlsx);
 assert.deepEqual(read.rows.map(r=>r.excelRow),[3,5]);
 assert.equal(read.rows[1].values.Feedback,'Why?');
});

test('master append preserves old formula, other worksheet, review and audit; repeated export stable',()=>{
 const wb=xlsx.utils.book_new();const ws=xlsx.utils.aoa_to_sheet([['ID','Feedback'],[40,'Old']]);ws.C1={t:'s',v:'calculated'};ws.C2={t:'n',f:'1+1',v:2};ws['!ref']='A1:C2';
 xlsx.utils.book_append_sheet(wb,ws,'Master');xlsx.utils.book_append_sheet(wb,xlsx.utils.aoa_to_sheet([['notes'],['Keep me']]),'Notes');
 const master=readMaster(wb,'Master',xlsx), run=seedRun(defaultWorkflow(),source(),master);
 applyResult(run,0,{status:'ok',items:pairs});applyResult(run,3,{status:'review',reason:'Bad source'});
 for(let i=0;i<2;i++){
  const out=buildWorkbook(run,xlsx);
  assert.equal(out.Sheets.Master.C2.f,'1+1');assert.equal(out.Sheets.Notes.A2.v,'Keep me');
  const rows=readSheet(out,'Master',1,xlsx).rows;assert.equal(rows.length,4);assert.equal(rows[1].values._source_row,2);
  assert.equal(readSheet(out,'Pipeline review',1,xlsx).rows.length,1);
  assert.equal(readSheet(out,'Pipeline audit',1,xlsx).rows.length,4);
  assert.equal(readSheet(out,'Pipeline imports',1,xlsx).rows.length,1);
 }
 assert.equal(readSheet(wb,'Master',1,xlsx).rows.length,1);
});

test('unexpected result columns cannot overwrite provenance',()=>{
 const run=seedRun(defaultWorkflow(),source());
 assert.throws(()=>applyResult(run,0,{status:'ok',items:[{_feedback_id:999}]}),/unexpected/);
 assert.equal(run.work[0].status,'pending');
});

test('long classifier audits and original JSON survive Excel cell limits without truncation',()=>{
 const run=seedRun(defaultWorkflow(),source());const long='é'.repeat(40000);
 applyResult(run,0,{status:'ok',items:pairs,audit:{probabilities:long}});
 const wb=buildWorkbook(run,xlsx);
 const chunks=readSheet(wb,'Pipeline overflow',1,xlsx).rows.map(r=>r.values).filter(r=>r.record_key===run.work[0].values._record_key && r.column==='details');
 assert.equal(JSON.parse(chunks.map(c=>c.value).join('')).probabilities,long);
 assert.doesNotThrow(()=>xlsx.write(wb,{type:'buffer',bookType:'xlsx'}));
});


test('rate cooldown survives recovery, honors minimum wait and stops after three responses',()=>{
 const run=seedRun(defaultWorkflow(),source());
 const result={status:'error',errorCode:'rate_limit',retryAfterSeconds:20};
 assert.equal(scheduleRateRetry(run,0,result,1000,0),'retry');
 assert.equal(run.work[0].retryAt,21000);
 assert.equal(run.providerCooldowns.azure,21000);
 const restored=JSON.parse(JSON.stringify(run));
 assert.equal(scheduleRateRetry(restored,0,result,21000,0),'retry');
 assert.equal(scheduleRateRetry(restored,0,result,41000,0),'exhausted');
 applyResult(restored,0,{...result,reason:'Still limited'});
 const id=restored.work[0].values._feedback_id;
 retryEntry(restored,0);
 assert.equal(restored.work[0].retryAt,61000);
 assert.equal(restored.work[0].rateLimitAttempts,undefined);
 assert.equal(restored.work[0].values._feedback_id,id);
 assert.equal(scheduleRateRetry(restored,0,{...result,retryAfterSeconds:120},61000,0),'pause');
});

test('retrying a failed downstream step preserves the successful excerpt, translation and code',()=>{
 const config=defaultWorkflow();config.nodes.push({...defaultNode('analyze'),inputs:['feedback_excerpt'],fields:[{name:'translation_en',options:[]}]});
 const run=seedRun(config,source());applyResult(run,0,{status:'ok',items:[pairs[0]]});
 applyResult(run,0,{status:'error',errorCode:'rate_limit',reason:'Wait'});
 const id=run.work[0].values._feedback_id;
 retryEntry(run,0);assert.equal(run.work[0].node,1);assert.equal(run.work[0].values.feedback_excerpt,'Merci.');
 applyResult(run,0,{status:'ok',items:[{translation_en:'Thank you.'}]});
 assert.equal(run.work[0].values._feedback_id,id);assert.equal(run.work[0].status,'ok');
});


test('configured output columns remain in master and review sheets when extraction fails for every row',()=>{
 const config=defaultWorkflow();config.nodes.push({...defaultNode('analyze'),inputs:['feedback_excerpt'],fields:[{name:'translation_en',options:[]}]});
 const run=seedRun(config,source());
 run.work.forEach((e,i)=>applyResult(run,i,{status:'review',reason:'Unclear phrase'}));
 const wb=buildWorkbook(run,xlsx);
 for(const name of ['Master data','Pipeline review']) {
  const sheet=readSheet(wb,name,1,xlsx);
  for(const column of ['feedback_excerpt','type_of_feedback','translation_en'])assert.ok(sheet.columns.includes(column),name+': '+column);
 }
 assert.equal(readSheet(wb,'Master data',1,xlsx).rows.length,0);
 assert.equal(readSheet(wb,'Pipeline review',1,xlsx).rows.length,2);
});

test('tentative translation stays in review export until retry succeeds',()=>{
 const config=defaultWorkflow();
 config.nodes=[{...defaultNode('analyze'),inputs:['Feedback'],fields:[{name:'translation_en',options:[]}]}];
 const run=seedRun(config,source());
 applyResult(run,0,{status:'review',reason:'Ambiguous',drafts:{translation_en:'Thanks. Where? When?'}});
 const reviewRows=()=>readSheet(buildWorkbook(run,xlsx),'Pipeline review',1,xlsx).rows;
 assert.equal(reviewRows()[0].values.translation_en,'[Needs Review] | Thanks. Where? When?');
 assert.equal(run.work[0].values.translation_en,undefined);
 run.work[0].drafts.translation_en='null';
 run.work[0].values.translation_en='null';
 assert.equal(reviewRows()[0].values.translation_en,'[Needs Review] | Source: Merci.Où ?Quand ?');
 retryEntry(run,0);
 applyResult(run,0,{status:'review',reason:'Ambiguous'});
 assert.equal(reviewRows()[0].values.translation_en,'[Needs Review] | Source: Merci.Où ?Quand ?');
 retryEntry(run,0);
 applyResult(run,0,{status:'ok',items:[{translation_en:'Final text'}]});
 assert.equal(mergeRun(run).rows[0].translation_en,'Final text');
});

test('a source translation in review collects tentative codes for each message and retries as one source',()=>{
 const config=JSON.parse(fs.readFileSync('presets/drc-ebola-framework.pipeline.json','utf8'));
 config.nodes=config.nodes.filter(n=>!n.pairedSourceStep); // Legacy saved run, before the file-wide split.
 delete config.screening; // The numeric fixture row tests review handling here; screening is tested separately.
 const run=seedRun(config,source());
 applyResult(run,0,{status:'review',tentative:true,reason:'Uncertain wording',drafts:{source_translation_en:'Thank you. Where?'}});
 assert.equal(run.work[0].status,'pending');
 const messages=pairs.slice(0,2).map((pair,i)=>({...pair,translation_en:['Thank you.','Where?'][i]}));
 applyResult(run,0,{status:'ok',items:messages});
 const restored=structuredClone(run);
 const category=config.nodes.find(n=>n.output==='feedback_category').table[0];
 applyResult(restored,0,{status:'ok',items:[{feedback_path:category.when.feedback_path}]});
 applyResult(restored,0,{status:'review',reason:'Topic unclear'});
 const {rows,review}=mergeRun(restored);
 assert.equal(rows.length,0);
 assert.equal(review.length,1);
 assert.equal(review[0].translation_en,'[Needs Review] | 1. Thank you.\n2. Where?');
 assert.ok(review[0].feedback_category.includes(category.value));
 assert.ok(review[0].feedback_path.includes('2. Topic unclear'));
 const id=review[0]._feedback_id;
 retryEntry(restored,0);
 applyResult(restored,0,{status:'ok',items:[{source_translation_en:'Thank you. Where?'}]});
 applyResult(restored,0,{status:'ok',items:messages});
 assert.equal(restored.work.length,3);
 assert.equal(restored.work[0].values._feedback_id,id);
 assert.equal(restored.work[0].drafts,undefined);
});

test('DRC phases translate and classify every source, split intact pairs, then code each child',()=>{
 const config=JSON.parse(fs.readFileSync('presets/drc-ebola-framework.pipeline.json','utf8'));
 const input=source();
 input.rows[0].values={ID:123,Feedback:'Où ? Merci || beaucoup.'};
 input.rows[1].values={ID:456,Feedback:'Bravo.'};
 const run=seedRun(config,input);
 const cursor={phase:0,index:0};
 const next=()=>nextWorkIndex(run,false,cursor);
 for (const i of [0,1]) {
   assert.equal(next(),i);
   applyResult(run,i,{status:'ok',items:[{source_translation_en:'Full translation'}]});
 }
 const pair=(type,excerpt)=>({type_of_feedback:type,feedback_excerpt:excerpt,translation_en:'English '+excerpt});
 const first=[pair('Questions','Où ?'),pair('Appréciations','Merci || beaucoup.')];
 applyResult(run,next(),{status:'ok',items:first});
 assert.equal(run.work.length,2);
 assert.equal(run.work[0].values.type_of_feedback,'Questions || Appréciations');
 assert.equal(next(),1);
 applyResult(run,1,{status:'ok',items:[pair('Appréciations','Bravo.')]});
 const saved=structuredClone(run);
 applyResult(run,next(),{status:'ok',items:run.work[0].pairedItems});
 assert.equal(next(),2);
 applyResult(run,2,{status:'ok',items:run.work[2].pairedItems});
 assert.equal(next(),0);
 assert.deepEqual(run.work.map(e=>[e.values._original_id,e.values.type_of_feedback,e.values.feedback_excerpt]),[
  [123,'Questions','Où ?'],[123,'Appréciations','Merci || beaucoup.'],[456,'Appréciations','Bravo.']]);
 assert.equal(run.work.every(e=>config.nodes[e.node].name==='feedback_path'),true);
 assert.equal(nextWorkIndex(saved,false,{phase:0,index:0}),0);
 const sheet=readSheet(buildWorkbook(run,xlsx),'Pipeline classifications',1,xlsx);
 assert.equal(sheet.rows[0].values.type_of_feedback,'Questions || Appréciations');
 assert.equal(sheet.rows[0].values._original_id,123);
});

test('retry a flagged translation once for its split source and retain matched message IDs',()=>{
 const config=JSON.parse(fs.readFileSync('presets/drc-ebola-framework.pipeline.json','utf8'));
 config.nodes=config.nodes.slice(0,4);
 const input=source();input.rows=input.rows.slice(0,1);
 const run=seedRun(config,input);
 const items=pairs.slice(0,2).map((p,i)=>({...p,translation_en:['Thanks','Where?'][i]}));
 applyResult(run,0,{status:'review',tentative:true,reason:'Unclear nuance'});
 applyResult(run,0,{status:'ok',items});
 applyResult(run,0,{status:'ok',items:run.work[0].pairedItems});
 const code={status:'ok',items:[{feedback_path:'Example'}]};
 applyResult(run,0,code);applyResult(run,1,code);
 assert.deepEqual(run.work.map(e=>e.status),['review','review']);
 const ids=run.work.map(e=>e.values._feedback_id);
 retryEntry(run,1);
 assert.equal(run.work.length,1);
 applyResult(run,0,{status:'ok',items:[{source_translation_en:'Thank you. Where?'}]});
 applyResult(run,0,{status:'ok',items});
 applyResult(run,0,{status:'ok',items:run.work[0].pairedItems});
 applyResult(run,0,code);applyResult(run,1,code);
 assert.deepEqual(run.work.map(e=>e.values._feedback_id),ids);
 assert.deepEqual(run.work.map(e=>e.status),['ok','ok']);
 assert.equal(Object.keys(run.classifications).length,1);
});

test('a review path fills supported lookup drafts and keeps them outside accepted data',()=>{
 const config=JSON.parse(fs.readFileSync('presets/drc-ebola-framework.pipeline.json','utf8'));
 delete config.screening; // The numeric fixture row tests review handling here; screening is tested separately.
 const run=seedRun(config,source());
 const index=config.nodes.findIndex(n=>n.name==='feedback_path');
 run.work[0].node=index;
 const category=config.nodes.find(n=>n.output==='feedback_category').table[0];
 applyResult(run,0,{status:'review',reason:'Possible mismatch',drafts:{feedback_path:category.when.feedback_path}});
 assert.equal(mergeRun(run).review[0].feedback_category,'[Needs Review] | '+category.value);
 assert.equal(mergeRun(run).rows.length,0);
 retryEntry(run,0);
 applyResult(run,0,{status:'review',drafts:{feedback_path:config.nodes[index].reviewValues[0]}});
 assert.equal(run.work[0].drafts.feedback_category,undefined);
 assert.equal(reviewCell(config,config.nodes[index],run.work[0],'feedback_path'),'');
 run.master.review.push({_status:'review',feedback_path:'[Needs Review] | '+config.nodes[index].reviewValues[0]});
 const rows=readSheet(buildWorkbook(run,xlsx),'Pipeline review',1,xlsx).rows;
 assert.deepEqual(rows.map(r=>[r.values.feedback_path,r.values._status]),[[null,'review'],[null,'review']]);
});

const GUIDANCE=['useContext','systemNotes','interpretFirst','valuesBeforeStatus','recheck','recheckWithReason','fallbackOption','tentativeOnReview'];
const asVersion1=preset=>({...structuredClone(preset),version:1,context:undefined,roles:undefined,
 nodes:preset.nodes.map(n=>Object.fromEntries(Object.entries(n).filter(([k])=>!GUIDANCE.includes(k))))});

test('version 1 DRC workflows upgrade to exactly the preset settings they used to receive implicitly',()=>{
 for(const file of ['drc-ebola-framework','drc-feedback-messages']){
  const preset=JSON.parse(fs.readFileSync(`presets/${file}.pipeline.json`,'utf8'));
  const legacy=portableConfig(asVersion1(preset));
  assert.equal(needsLegacyUpgrade(legacy),true);
  const ebola=JSON.parse(fs.readFileSync('presets/drc-ebola-framework.pipeline.json','utf8'));
  const upgraded=upgradeLegacyWorkflow(legacy,ebola);
  assert.equal(upgraded.context,preset.context);
  assert.deepEqual(upgraded.nodes,portableConfig(preset).nodes,file);
  for(const [role,column] of Object.entries(preset.roles||{})) assert.deepEqual(upgraded.roles[role],column);
  assert.equal(needsLegacyUpgrade(upgraded),false);
 }
 const generic={...defaultWorkflow(),nodes:[{...defaultNode('analyze'),instructions:'Summarise.'}]};
 assert.equal(needsLegacyUpgrade(portableConfig(generic)),false);
});

test('without roles, a tentative review stops the record instead of using DRC column names',()=>{
 const config={...defaultWorkflow(),nodes:[{...defaultNode('analyze'),id:'t',inputs:['Feedback'],fields:[{name:'english',options:[]}],tentativeOnReview:true},defaultNode('extract')]};
 const run=seedRun(config,source());
 applyResult(run,0,{status:'review',tentative:true,reason:'Unclear',drafts:{english:'Thanks'}});
 assert.equal(run.work[0].status,'review');
 assert.equal(run.work[0].translationFlag,undefined);
 assert.equal(run.work[0].tentativeCoding,undefined);
});

test('workflow context and roles survive export, and malformed roles are rejected',()=>{
 const saved=portableConfig({...defaultWorkflow(),context:'Setting.',roles:{excerpt:'feedback_excerpt',tentativeColumns:['a']}});
 assert.equal(saved.version,2);
 assert.equal(saved.context,'Setting.');
 assert.deepEqual(saved.roles,{excerpt:'feedback_excerpt',tentativeColumns:['a']});
 for(const roles of [{excerpt:''},{unknown:'x'},{tentativeColumns:'a'},[]]) assert.throws(()=>portableConfig({...defaultWorkflow(),roles}),/roles/);
 assert.throws(()=>portableConfig({...defaultWorkflow(),context:5}),/context/);
});

// Mirrors continueRun's scheduling with simulated responses: limit requests in flight, completed in random order.
function simulate(config, rows, limit, seed) {
 let state=seed;const random=()=>(state=(state*1103515245+12345)%2147483648)/2147483648;
 const run=seedRun(config,{filename:'f.csv',datasetId:'d',sheet:'S',columns:['Feedback'],rows:rows.map((Feedback,i)=>({excelRow:i+2,values:{Feedback}}))});
 const respond=(node,values)=>{
  if(node.pairedSourceStep)return null;
  if(node.type==='analyze')return {status:'ok',items:[{en:'EN '+values.Feedback}]};
  if(node.type==='extract')return values.Feedback.includes('?')?{status:'review',reason:'unclear'}:
   {status:'ok',items:values.Feedback.split('.').filter(Boolean).map(t=>({feedback_excerpt:t,type_of_feedback:t.length%2?'Odd':'Even'}))};
  if(node.type==='classify')return {status:'ok',items:[{topic:values.feedback_excerpt.length>3?'Long':'Short'}]};
  return {status:'ok',items:[{}]};
 };
 const inFlight=new Map(),cursor={phase:0,index:0};
 for(;;){
  while(inFlight.size<limit){
   const index=nextWorkIndex(run,false,cursor,e=>inFlight.has(e),inFlight.size);if(index<0)break;
   const entry=run.work[index],{node,values}=activeStep(run,entry);
   inFlight.set(entry,{hint:index,result:respond(node,values)||{status:'ok',items:entry.pairedItems}});
  }
  if(!inFlight.size)break;
  const jobs=[...inFlight];const [entry,job]=jobs[Math.floor(random()*jobs.length)];
  inFlight.delete(entry);
  const index=locateEntry(run,entry,job.hint);rewindCursor(cursor,index);applyResult(run,index,job.result);
 }
 return run;
}
test('parallel scheduling reaches the same records as running one request at a time',()=>{
 const config={...defaultWorkflow(),idColumn:'',roles:{excerpt:'feedback_excerpt',messageType:'type_of_feedback'},nodes:[
  {...defaultNode('analyze'),id:'t',inputs:['Feedback'],fields:[{name:'en',options:[]}]},
  {...defaultNode('extract'),id:'x',inputs:['Feedback','en'],fields:[{name:'feedback_excerpt',options:[]},{name:'type_of_feedback',options:['Odd','Even']}]},
  {...defaultNode('split'),id:'s',inputs:['feedback_excerpt','type_of_feedback'],pairedSourceStep:'x'},
  {...defaultNode('classify'),id:'c',inputs:['feedback_excerpt'],fields:[{name:'topic',options:['Long','Short']}]},
  {...defaultNode('validate'),id:'v',inputs:['topic']}]};
 const rows=Array.from({length:40},(_,i)=>i%7===3?'Where? '+i:`Hi.Thanks ${i}.Water${'x'.repeat(i%4)}.`);
 const summary=run=>run.work.map(e=>[e.values._record_key,e.status,e.values.topic,e.values.feedback_excerpt]);
 const sequential=summary(simulate(config,rows,1,1));
 assert.equal(sequential.filter(r=>r[1]==='ok').length,102);
 assert.equal(sequential.filter(r=>r[1]==='review').length,6);
 for(const seed of [2,3,4,5,6]){
  const parallel=simulate(config,rows,5,seed);
  assert.deepEqual(summary(parallel),sequential,'seed '+seed);
  // Excel rows follow the same order: accepted rows in the master, flagged ones on the review sheet.
  const order=r=>{const m=mergeRun(r);return [m.rows.map(x=>x._record_key),m.review.map(x=>x._record_key)];};
  assert.deepEqual(order(parallel),order(simulate(config,rows,1,1)),'export order, seed '+seed);
  // Feedback IDs are the one thing assigned in completion order.
  const byKey=r=>r.work.map(e=>e.values._feedback_id);
  if(seed===2) assert.notDeepEqual(byKey(parallel),byKey(simulate(config,rows,1,1)));
  const ids=parallel.work.map(e=>e.values._feedback_id);
  assert.equal(new Set(ids).size,ids.length);
 }
});

test('a reviewer note stays with the record and its split messages, and is exported',()=>{
 const run=seedRun(defaultWorkflow(),source());
 applyResult(run,0,{status:'review',reason:'Unclear'});
 const entry=retryEntry(run,0,{},'  “maso” means words  ');
 assert.equal(entry.values._reviewer_note,'“maso” means words');
 assert.equal(entry.history.at(-1).note,'“maso” means words');
 applyResult(run,0,{status:'ok',items:pairs});
 assert.ok(run.work.slice(0,3).every(e=>e.values._reviewer_note==='“maso” means words'));
 assert.equal(mergeRun(run).rows[0]._reviewer_note,'“maso” means words');
 assert.throws(()=>{applyResult(run,3,{status:'review',reason:'x'});retryEntry(run,3,{},'x'.repeat(5001));},/under 5,000/);
});

test('a reviewer answer keeps to the codebook branch, then the record continues',()=>{
 const config={...defaultWorkflow(),idColumn:'',nodes:[
  {...defaultNode('classify'),id:'c',inputs:['Feedback','kind'],fields:[{name:'code',options:[]}],fallbackOption:'None',
   sourceBranches:[{when:{kind:'Q'},options:[{label:'Ask',description:'d'},'None']},{when:{kind:'R'},options:['Request']}]},
  {...defaultNode('lookup'),id:'l',inputs:['code'],output:'group',table:[{when:{code:'Ask'},value:'Questions'}]}]};
 const run=seedRun(config,{filename:'f',datasetId:'d',sheet:'S',columns:['Feedback','kind'],rows:[{excelRow:2,values:{Feedback:'Where?',kind:'Q'}}]});
 applyResult(run,0,{status:'review',reason:'Unsure',drafts:{code:'Ask'}});
 assert.deepEqual(reviewerFields(config.nodes[0],run.work[0].values),[{name:'code',options:['Ask']}]);
 assert.throws(()=>acceptReviewerAnswer(run,0,{code:'Request'}),/allowed values/);
 assert.throws(()=>acceptReviewerAnswer(run,0,{code:'None'}),/allowed values/);
 const at=acceptReviewerAnswer(run,0,{code:'Ask'},'Asked where the centre is');
 const entry=run.work[at];
 assert.equal(entry.values.code,'Ask');
 assert.equal(entry.status,'pending');
 assert.equal(entry.node,1);
 assert.equal(entry.values._reviewer_decisions,'Classify');
 assert.equal(entry.history.at(-1).provider,'reviewer');
 assert.throws(()=>acceptReviewerAnswer(run,at,{group:'x'}),/flagged/);
 const extract=seedRun(defaultWorkflow(),source());
 applyResult(extract,0,{status:'review',reason:'Unclear'});
 assert.throws(()=>acceptReviewerAnswer(extract,0,{}),/add a note/);
});

test('a reviewer translation reopens a flagged DRC source and codes its messages again',()=>{
 const config=JSON.parse(fs.readFileSync('presets/drc-ebola-framework.pipeline.json','utf8'));
 const run=seedRun(config,{filename:'f',datasetId:'d',sheet:'S',columns:['ID','Feedback'],rows:[{excelRow:2,values:{ID:1,Feedback:'Tondimi maso bolobi'}}]});
 applyResult(run,0,{status:'review',tentative:true,reason:'Unclear word',drafts:{source_translation_en:'We accept eyes'}});
 // The flagged translation still runs extraction and coding, then waits for review as one source.
 const cursor={phase:0,index:0};
 for(let index;(index=nextWorkIndex(run,false,cursor))>=0;){
  const entry=run.work[index],{node,values}=activeStep(run,entry);
  const field=reviewerFields(node,values)[0];
  const match=node.type==='lookup'&&node.table.find(r=>node.inputs.every(k=>r.when[k]===values[k]));
  applyResult(run,index,node.pairedSourceStep?{status:'ok',items:entry.pairedItems}
   :node.type==='extract'?{status:'ok',items:[{feedback_excerpt:'Tondimi maso bolobi',type_of_feedback:'Appréciations',translation_en:'We accept'}]}
   :node.type==='lookup'?(match?{status:'ok',items:[{[node.output]:match.value}]}:{status:'review',reason:'none'})
   :{status:'ok',items:[{[field.name]:field.options[0]}]});
 }
 assert.equal(run.work.length,1);
 assert.equal(run.work[0].status,'review');
 assert.ok(run.work[0].translationFlag);
 const at=acceptReviewerAnswer(run,0,{source_translation_en:'We accept your words.'});
 const entry=run.work[at];
 assert.equal(entry.translationFlag,undefined);
 assert.equal(entry.values.source_translation_en,'We accept your words.');
 assert.equal(config.nodes[entry.node].id,'extract-messages');
 assert.equal(entry.status,'pending');
});

test('screening skips rows that cannot be feedback before any AI step, and they can be processed anyway',()=>{
 const config=JSON.parse(fs.readFileSync('presets/drc-ebola-framework.pipeline.json','utf8'));
 const texts=['12345','','sdfgh qwerty','jjjjjj','Tondimi maso bolobi','Kama una cikia ile utapata dawa','polepole','MSF','x'];
 const run=seedRun(config,{filename:'f',datasetId:'d',sheet:'S',columns:['ID','Feedback'],rows:texts.map((Feedback,i)=>({excelRow:i+2,values:{ID:i,Feedback}}))});
 assert.deepEqual(run.work.map(e=>e.status),['skipped','skipped','skipped','skipped','pending','pending','pending','pending','skipped']);
 assert.equal(run.work[0].values._skip_reason,'No words: only numbers or symbols');
 assert.equal(run.work[2].values._skip_reason,'Looks like random typing');
 assert.equal(run.work[0].values._feedback_id,undefined);
 assert.equal(nextWorkIndex(run,false,{phase:0,index:0}),4);
 const merged=mergeRun(run);
 assert.equal(merged.review.filter(r=>r._status==='skipped').length,5);
 assert.equal(merged.logs.at(-1).skipped,5);
 assert.equal(restoreEntry(run,0),'pending');
 assert.equal(run.work[0].values._skip_reason,undefined);
 assert.equal(nextWorkIndex(run,false,{phase:0,index:0}),0);
 const again=seedRun({...config,screening:undefined},{filename:'f',datasetId:'d2',sheet:'S',columns:['ID','Feedback'],rows:[{excelRow:2,values:{ID:1,Feedback:'12345'}}]});
 assert.equal(again.work[0].status,'pending');
});

test('the model can skip a record; discarding needs a reason, can cover a whole source row, and can be restored',()=>{
 const run=seedRun(defaultWorkflow(),source());
 applyResult(run,0,{status:'skipped',reason:'Random keystrokes'});
 assert.equal(run.work[0].status,'skipped');
 assert.equal(run.work[0].values._skip_reason,'Model: Random keystrokes');
 restoreEntry(run,0);
 applyResult(run,0,{status:'ok',items:pairs});
 for(const i of [0,1,2]) run.work[i].status='review';
 assert.deepEqual(sourceSiblings(run,1),[0,1,2]);
 assert.throws(()=>discardEntries(run,[1],'  '),/reason/);
 assert.equal(discardEntries(run,sourceSiblings(run,1),'Test or random typing'),3);
 assert.ok(run.work.slice(0,3).every(e=>e.status==='discarded' && e.values._discard_reason==='Test or random typing'));
 assert.equal(discardEntries(run,[0],'again'),0);
 assert.equal(mergeRun(run).review.filter(r=>r._discard_reason).length,3);
 assert.equal(restoreEntry(run,1),'review');
 assert.equal(run.work[1].values._discard_reason,undefined);
 assert.throws(()=>restoreEntry(run,1),/Only skipped or discarded/);
});
