const {test,expect}=require('@playwright/test');
// The runner sends several requests at once by default. Tests about step-by-step order use one at a time.
async function setParallel(page,n){await page.getByText('Run options').click();await page.locator('#parallel-limit').fill(String(n));await page.locator('#parallel-limit').dispatchEvent('change');}

test('paired extraction, isolated review, correction, checkpoint restore and Excel merge',async({page})=>{
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.route('**/api/pipeline/step',async route=>{
  const {node,record,credentials}=route.request().postDataJSON();
  expect(node.provider).toBe('azure');expect(credentials).not.toHaveProperty('jevApiKey');
  if(record.Feedback==='Bof eh')return route.fulfill({json:{status:'review',reason:'Insufficient context'}});
  const items=record.Feedback==='Merci.'?[{feedback_excerpt:'Merci.',type_of_feedback:'Appréciations'}]:[{feedback_excerpt:'Merci pour votre aide.',type_of_feedback:'Appréciations'},{feedback_excerpt:'Où est le centre ?',type_of_feedback:'Questions'},{feedback_excerpt:'Quand ouvre-t-il ?',type_of_feedback:'Questions'}];
  return route.fulfill({json:{status:'ok',items,model:'azure-test'}});
 });
 await page.goto('/pipeline');
 await page.locator('#to-files').click();
 await page.locator('#source-file').setInputFiles({name:'team.csv',mimeType:'text/csv',buffer:Buffer.from('\ufeffID,Feedback\n1,Merci pour votre aide.Où est le centre ?Quand ouvre-t-il ?\n\n2,Bof eh\n')});
 await expect(page.locator('#source-summary')).toContainText('2 source rows');
 await page.screenshot({path:'/tmp/pipeline-redesign-files.png',fullPage:true});
 await page.locator('#sheet-options summary').click();
 await page.locator('#header-row').fill('0');await page.locator('#header-row').blur();
 await expect(page.locator('#notice')).toContainText('valid header row');
 await page.locator('#header-row').fill('1');await page.locator('#header-row').blur();
 await expect(page.locator('#source-summary')).toContainText('team.csv');
 await page.locator('#start-run').click();
 await expect(page.locator('#run-status')).toContainText('Processing complete');
 await expect(page.locator('#run-counts')).toContainText('3 accepted');
 await expect(page.locator('#run-counts')).toContainText('1 review');
 await expect(page.locator('#results tbody tr')).toHaveCount(4);
 await expect(page.locator('#results')).toContainText('row 4');
 await page.screenshot({path:'/tmp/pipeline-redesign-results.png',fullPage:true});
 await expect(page.locator('#split-summary')).toContainText('1 source split into multiple messages');
 await expect(page.locator('#run-time')).toContainText('Elapsed');
 await page.getByText('Choose download columns', {exact:true}).click();
 await page.locator('#export-columns input[value="Feedback"]').uncheck();
 const completedDownload = page.waitForEvent('download');
 await page.locator('#download-results').click();
 const completedFile = await completedDownload;
 expect(completedFile.suggestedFilename()).toBe('team_completed.xlsx');
 const chunks = [];
 for await (const chunk of await completedFile.createReadStream()) chunks.push(chunk);
 const exported = await page.evaluate(base64 => {
   const book = XLSX.read(base64, {type:'base64'});
   return {sheets:book.SheetNames, rows:XLSX.utils.sheet_to_json(book.Sheets.Results, {header:1})};
 }, Buffer.concat(chunks).toString('base64'));
 expect(exported.sheets).toEqual(['Results']);
 expect(exported.rows).toEqual([['ID','feedback_excerpt','type_of_feedback'],
   [1,'Merci pour votre aide.','Appréciations'],[1,'Où est le centre ?','Questions'],[1,'Quand ouvre-t-il ?','Questions']]);
 await page.locator('#export-format').selectOption('csv');
 const csvDownload=page.waitForEvent('download');
 await page.locator('#download-results').click();
 const csvFile=await csvDownload;
 expect(csvFile.suggestedFilename()).toBe('team_completed.csv');
 const csv=[];
 for await (const chunk of await csvFile.createReadStream()) csv.push(chunk);
 expect(Buffer.concat(csv).toString('utf8').trim()).toEqual('ID,feedback_excerpt,type_of_feedback\r\n1,Merci pour votre aide.,Appréciations\r\n1,Où est le centre ?,Questions\r\n1,Quand ouvre-t-il ?,Questions');
 const download=page.waitForEvent('download');await page.locator('#download-master').click();
 const auditFile=await download;
 expect(auditFile.suggestedFilename()).toBe('updated_master.xlsx');
 const audit=[];
 for await (const chunk of await auditFile.createReadStream()) audit.push(chunk);
 const auditSheets=await page.evaluate(data=>XLSX.read(data,{type:'base64',bookSheets:true}).SheetNames,Buffer.concat(audit).toString('base64'));
 expect(auditSheets).toContain('Pipeline audit');
 expect(auditSheets).toContain('Pipeline review');
 await page.reload();await page.locator('.workspace-menu summary').click();await page.locator('#restore-run').click();await expect(page.locator('#run-counts')).toContainText('1 review');
 await page.locator('[data-review]').click();await page.getByText('Correct the input text').click();await page.locator('#correction-0').fill('Merci.');await page.locator('#save-correction').click();
 await expect(page.locator('#run-counts')).toContainText('4 accepted');await expect(page.locator('#run-counts')).toContainText('0 review');
 expect(errors).toEqual([]);
});

test('builder imports step 2 as appended configurable steps and exports no keys',async({page})=>{
 await page.goto('/pipeline');
 await page.locator('#edit-workflow').click();
 await page.locator('#append-file').setInputFiles('examples/examples_1/step2-ebola-feedback-coding.excerpt.input-column.config.json');
 await expect(page.locator('.pipeline-node')).toHaveCount(14);
 await page.locator('#node-provider-1').selectOption('azure');await page.locator('#node-model-1').fill('my-deployment');
 await page.locator('#api-settings-btn').click();await expect(page.locator('#api-settings-modal')).toBeVisible();
 await page.locator('#modal-api-key').fill('secret-test-key');await page.locator('#save-api-settings').click();
 await page.locator('.workspace-menu summary').click();
 const download=page.waitForEvent('download');await page.locator('#export-config').click();const file=await download;const stream=await file.createReadStream();let text='';for await(const chunk of stream)text+=chunk;
 const config=JSON.parse(text);expect(config.nodes[1].provider).toBe('azure');expect(config.nodes[1].model).toBe('my-deployment');expect(text).not.toContain('secret-test-key');expect(config.nodes[1].sourceBranches).toHaveLength(7);
});

test('legacy paired lists with unequal counts go to review without API credentials',async({page})=>{
 await page.goto('/pipeline');
 const config={format:'aidstack-insights-pipeline',version:1,name:'Legacy split',idColumn:'ID',mappings:{},nodes:[{id:'split',name:'Split pairs',type:'split',inputs:['excerpt','type'],delimiter:' || '}]};
 await page.locator('#config-file').setInputFiles({name:'config.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
 await page.locator('#to-files').click();
 await page.locator('#source-file').setInputFiles({name:'source.csv',mimeType:'text/csv',buffer:Buffer.from('ID,excerpt,type\n1,One || Two,Question\n2,Three || Four,Question || Question\n')});
 await page.locator('#start-run').click();await expect(page.locator('#run-status')).toContainText('Processing complete');
 await expect(page.locator('#run-counts')).toContainText('2 accepted');await expect(page.locator('#run-counts')).toContainText('1 review');
 await expect(page.locator('#results')).toContainText('unequal counts');
});

test('Laya classification uses the local connector and keeps cloud credentials out',async({page})=>{
 let localCalls=0, cloudCalls=0;
 await page.route('**/api/pipeline/step',async route=>{cloudCalls++;await route.abort();});
 await page.route('http://127.0.0.1:8001/analyze_batch_jev',async route=>{
  localCalls++;const payload=route.request().postDataJSON();expect(payload.classificationProvider).toBe('laya');expect(payload).not.toHaveProperty('apiKey');expect(payload.configs[0].options).toEqual(['Question','Thanks']);
  await route.fulfill({json:{results:[{rowIndex:0,decisions:{category:{status:'ok',value:'Question',model:'laya-test'}}}]}});
 });
 await page.goto('/pipeline');
 const config={format:'aidstack-insights-pipeline',version:1,name:'Local classify',idColumn:'ID',mappings:{},nodes:[{id:'c',name:'Classify',type:'classify',inputs:['Feedback'],provider:'laya',instructions:'Classify',fields:[{name:'category',options:['Question','Thanks']}],maxOutputTokens:4096}]};
 await page.locator('#config-file').setInputFiles({name:'config.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
 await page.locator('#to-files').click();
 await page.locator('#source-file').setInputFiles({name:'source.csv',mimeType:'text/csv',buffer:Buffer.from('ID,Feedback\n1,Where?\n')});
 await page.locator('#start-run').click();await expect(page.locator('#run-status')).toContainText('Processing complete');
 await expect(page.locator('#run-counts')).toContainText('1 accepted');expect(localCalls).toBe(1);expect(cloudCalls).toBe(0);
});

test('guided setup and editor stay usable on narrow screens',async({page})=>{
 await page.setViewportSize({width:390,height:844});await page.goto('/pipeline');
 await expect(page.locator('#workflow-name')).not.toBeVisible();
 await expect(page.locator('#preset-select')).toBeVisible();
 await expect(page.locator('#start-run')).not.toBeVisible();
 expect(await page.evaluate(()=>document.documentElement.scrollWidth <= innerWidth+1)).toBe(true);
 await page.screenshot({path:'/tmp/pipeline-redesign-mobile.png',fullPage:true});
 await page.setViewportSize({width:1440,height:1000});await page.screenshot({path:'/tmp/pipeline-redesign-desktop.png',fullPage:true});
 await page.locator('#edit-workflow').click();await expect(page.locator('#workflow-name')).toBeVisible();
 await page.setViewportSize({width:390,height:844});
 expect(await page.evaluate(()=>document.documentElement.scrollWidth <= innerWidth+1)).toBe(true);
 await page.locator('#close-editor').click();await expect(page.locator('#workflow-name')).not.toBeVisible();
});

test('full DRC workflow loads with one click and shows a compact summary',async({page})=>{
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto('/pipeline');await page.locator('#preset-select').selectOption('drc-ebola-framework');
 await expect(page.locator('#selected-workflow-name')).toContainText('DRC');
 await expect(page.locator('#workflow-editor')).not.toBeVisible();
 await expect(page.locator('#workflow-outline .outline-step')).toHaveCount(8);
 await page.locator('#edit-workflow').click();
 await expect(page.locator('.node-body:visible')).toHaveCount(1);
 await page.locator('[data-index="3"] [data-action="toggle"]').click();
 await expect(page.locator('.node-body:visible')).toHaveCount(1);
 await expect(page.locator('#node-provider-3')).toBeVisible();
 await page.locator('#node-provider-3').selectOption('jev');
 await page.locator('#close-editor').click();
 await expect(page.locator('#workflow-summary')).toContainText('Jev');
 await page.locator('#to-files').click();
 await expect(page.locator('#stage-files')).toBeVisible();
 await expect(page.locator('#stage-workflow')).not.toBeVisible();
 await expect(page.locator('#start-run')).toBeDisabled();
 await expect(page.locator('#source-options')).not.toBeVisible();
 expect(errors).toEqual([]);
});

test('a rate-limited later step waits and retries without repeating extraction',async({page})=>{
 let extracts=0, analyses=0;
 await page.route('**/api/pipeline/step',async route=>{
  const {node}=route.request().postDataJSON();
  if(node.type==='extract'){extracts++;return route.fulfill({json:{status:'ok',items:[{excerpt:'Merci.',type:'Thanks'}]}});}
  analyses++;
  return route.fulfill({json:analyses===1?{status:'error',errorCode:'rate_limit',retryAfterSeconds:1,reason:'Slow down'}:{status:'ok',items:[{translation:'Thanks.'}]}});
 });
 await page.goto('/pipeline');
 const config={format:'aidstack-insights-pipeline',version:1,name:'Retry test',idColumn:'ID',mappings:{},nodes:[{id:'e',name:'Extract',type:'extract',provider:'azure',inputs:['Feedback'],instructions:'Extract',fields:[{name:'excerpt',options:[]},{name:'type',options:['Thanks','Question']}],verbatimInput:'Feedback',verbatimField:'excerpt'},{id:'t',name:'Translate',type:'analyze',provider:'azure',inputs:['excerpt'],instructions:'Translate',fields:[{name:'translation',options:[]}]}]};
 await page.locator('#config-file').setInputFiles({name:'config.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
 await page.locator('#to-files').click();await page.locator('#source-file').setInputFiles({name:'source.csv',mimeType:'text/csv',buffer:Buffer.from('ID,Feedback\n1,Merci.\n')});
 await page.locator('#start-run').click();
 await expect(page.locator('#run-status')).toContainText('asked us to slow down');
 await expect(page.locator('#run-status')).toContainText('Processing complete');
 expect(extracts).toBe(1);expect(analyses).toBe(2);
 await expect(page.locator('#run-counts')).toContainText('1 accepted');
});

test('quota failures pause the queue and identify the stopped step',async({page})=>{
 let calls=0;
 await page.route('**/api/pipeline/step',async route=>{calls++;return route.fulfill({json:{status:'error',errorCode:'quota_exhausted',retryable:false,reason:'Check account quota before retrying.'}});});
 await page.goto('/pipeline');await page.locator('#to-files').click();
 await page.locator('#source-file').setInputFiles({name:'source.csv',mimeType:'text/csv',buffer:Buffer.from('ID,Feedback\n1,Merci.\n2,Why?\n')});
 await setParallel(page,1);
 await page.locator('#start-run').click();await expect(page.locator('#run-status')).toContainText('Check account quota');
 await expect(page.locator('#pause-run')).not.toBeVisible();expect(calls).toBe(1);
 await expect(page.locator('#results')).toContainText('Quota / billing limit');
 await expect(page.locator('#results')).toContainText('Stopped at:');
 await expect(page.locator('#results')).toContainText('Source text:');
 await expect(page.locator('#run-counts')).toContainText('1 pending');
});

test('requested result columns stay visible when no source reaches translation',async({page})=>{
 let calls=0;
 await page.route('**/api/pipeline/step',async route=>{calls++;return route.fulfill({json:{status:'review',reason:'The phrase needs clarification.'}});});
 await page.goto('/pipeline');
 const config={format:'aidstack-insights-pipeline',version:1,name:'Multilingual review',idColumn:'ID',mappings:{},nodes:[{id:'e',name:'Extract',type:'extract',provider:'azure',inputs:['Feedback'],instructions:'French, Swahili, Lingala or mixtures.',fields:[{name:'feedback_excerpt',options:[]},{name:'type_of_feedback',options:['Thanks','Question']}],verbatimInput:'Feedback',verbatimField:'feedback_excerpt'},{id:'t',name:'Translate',type:'analyze',provider:'azure',inputs:['feedback_excerpt'],instructions:'Translate',fields:[{name:'translation_en',options:[]}]}]};
 await page.locator('#config-file').setInputFiles({name:'config.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
 await page.locator('#to-files').click();await page.locator('#source-file').setInputFiles({name:'source.csv',mimeType:'text/csv',buffer:Buffer.from('ID,Date,Week,Province,City,Feedback\n1,46152,41,Tshopo,Kabondo,Merci xyz\n')});
 await page.locator('#start-run').click();await expect(page.locator('#run-status')).toContainText('Processing complete');
 expect(calls).toBe(1);
 await expect(page.locator('#results thead')).toContainText('feedback excerpt');
 await expect(page.locator('#results thead')).toContainText('type of feedback');
 await expect(page.locator('#results thead')).toContainText('translation en');
 await expect(page.locator('#results thead')).not.toContainText('Province');
 await expect(page.locator('#results')).toContainText('Not run');
 await expect(page.locator('#results')).toContainText('Merci xyz');
 await expect(page.locator('#column-visibility-note')).toContainText('3 of 3');
 await page.locator('#all-result-columns').check();await expect(page.locator('#results thead')).toContainText('Province');
 const originalId=await page.locator('#results tbody tr td').nth(1).textContent();
 await page.locator('#retry-extraction-reviews').click();
 await expect(page.locator('#run-status')).toContainText('Processing complete');
 expect(calls).toBe(2);
 await expect(page.locator('#results tbody tr')).toHaveCount(1);
 await expect(page.locator('#results tbody tr td').nth(1)).toHaveText(originalId);
 await expect(page.locator('#run-counts')).toContainText('1 review');
});

test('retry review records resumes the stopped step and skips accepted records', async ({page})=>{
 const attempts=[];
 let retrying=false;
 await page.route('**/api/pipeline/step',route=>{
   const {node,record}=route.request().postDataJSON();
   attempts.push([node.id,record.Feedback]);
   const needsReview=node.id==='translate' && record.Feedback==='bonjour' && !retrying;
   return route.fulfill({json:needsReview?{status:'review',reason:'Ambiguous translation',drafts:{english:'Hello'}}:
     {status:'ok',items:[{[node.fields[0].name]:'Done'}]}});
 });
 await page.goto('/pipeline');
 const node=(id,field)=>({id,name:id,type:'analyze',inputs:['Feedback'],provider:'azure',instructions:'Analyze',fields:[{name:field,options:[]}]});
 const config={format:'aidstack-insights-pipeline',version:1,name:'Translate',nodes:[node('prep','first'),node('translate','english')]};
 await page.locator('#config-file').setInputFiles({name:'steps.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
 await page.locator('#to-files').click();
 await page.locator('#source-file').setInputFiles({name:'source.csv',mimeType:'text/csv',buffer:Buffer.from('Feedback\nmerci\nbonjour\n')});
 await page.locator('#start-run').click();
 await expect(page.locator('#results')).toContainText('[Needs Review] | Hello');
 await expect(page.locator('#run-status')).toContainText('Processing complete');
 retrying=true; attempts.length=0;
 await page.locator('#retry-reviews').click();
 await expect(page.locator('#run-counts')).toContainText('2 accepted');
 expect(attempts).toEqual([['translate','bonjour']]);
 await expect(page.locator('#results')).not.toContainText('[Needs Review]');
});

test('DRC translates a source row once before coding its distinct messages',async({page})=>{
 const source='Merci. Où est le centre ?';
 const translation='Thank you. Where is the centre?';
 const calls=[];
 await page.route('**/api/pipeline/step',route=>{
   const {node,record}=route.request().postDataJSON();
   calls.push({id:node.id,record});
   if(node.id==='translate-feedback')return route.fulfill({json:{status:'ok',items:[{source_translation_en:record.Feedback===source?translation:'Bravo.'}]}});
   if(node.id==='extract-messages' && record.Feedback==='Bravo.')return route.fulfill({json:{status:'ok',items:[
     {feedback_excerpt:'Bravo.',translation_en:'Bravo.',type_of_feedback:'Appréciations'}]}});
   if(node.id==='extract-messages')return route.fulfill({json:{status:'ok',items:[
     {feedback_excerpt:'Merci.',translation_en:'Thank you.',type_of_feedback:'Appréciations'},
     {feedback_excerpt:'Où est le centre ?',translation_en:'Where is the centre?',type_of_feedback:'Questions'}]}});
   return route.fulfill({json:{status:'review',reason:'Stop at codebook selection for this test.'}});
 });
 await page.goto('/pipeline');await page.locator('#preset-select').selectOption('drc-ebola-framework');
 await expect(page.locator('#selected-workflow-name')).toContainText('DRC');
 await page.locator('#to-files').click();
 await page.locator('#source-file').setInputFiles({name:'feedback.csv',mimeType:'text/csv',buffer:Buffer.from('ID,Feedback\n1,'+source+'\n2,Bravo.\n')});
 await page.locator('#start-run').click();
 await expect(page.locator('#run-status')).toContainText('Processing complete');
 expect(calls.map(c=>c.id)).toEqual(['translate-feedback','translate-feedback','extract-messages','extract-messages','codebook-1','codebook-1','codebook-1']);
 expect(calls[0].record).toEqual({Feedback:source});
 expect(calls[2].record).toEqual({Feedback:source,source_translation_en:translation});
 expect(calls.slice(4).map(c=>[c.record.translation_en,c.record.type_of_feedback])).toEqual([
   ['Thank you.','Appréciations'],['Where is the centre?','Questions'],['Bravo.','Appréciations']]);
});

test('a flagged DRC translation still receives an English excerpt and codebook path',async({page})=>{
 const calls=[];
 await page.route('**/api/pipeline/step',route=>{
   const {node}=route.request().postDataJSON();
   calls.push(node.id);
   if(node.id==='translate-feedback')return route.fulfill({json:{status:'review',tentative:true,reason:'One word unclear'}});
   if(node.type==='extract')return route.fulfill({json:{status:'ok',items:[{feedback_excerpt:'Tondimi maso bolobi',translation_en:'We agree with what you said.',type_of_feedback:'Observations, perceptions ou croyances'}]}});
   if(node.fields?.[0]?.name==='feedback_path')return route.fulfill({json:{status:'ok',items:[{feedback_path:'Observations sur la réponse > Information > Observations concernant l’information sur la réponse'}]}});
   if(node.type==='lookup')return route.continue();
   return route.fulfill({json:{status:'ok',items:[{[node.fields[0].name]:node.fields[0].name==='feedback_sector_suggested'?'RCCE/CEA':node.fields[0].name==='feedback_action'?'Document and monitor trend':'New'}]}});
 });
 await page.goto('/pipeline');await page.locator('#preset-select').selectOption('drc-ebola-framework');
 await expect(page.locator('#selected-workflow-name')).toContainText('DRC');
 await page.locator('#to-files').click();
 await page.locator('#source-file').setInputFiles({name:'test.csv',mimeType:'text/csv',buffer:Buffer.from('ID,Feedback\n1,Tondimi maso bolobi\n')});
 await page.locator('#start-run').click();
 await expect(page.locator('#run-status')).toContainText('Processing complete');
 expect(calls.slice(0,3)).toEqual(['translate-feedback','extract-messages','codebook-1']);
 expect(calls).not.toContain('codebook-10');
 await page.locator('#all-result-columns').check();
 const headings = await page.locator('#results th').allTextContents();
 await expect(page.locator('#results tbody tr td').nth(headings.indexOf('feedback code'))).toContainText('Observations concernant l’information sur la réponse');
 await expect(page.locator('#run-counts')).toContainText('1 review');
 await expect(page.locator('#results')).toContainText('[Needs Review] | We agree with what you said.');
 await expect(page.locator('#results')).toContainText('[Needs Review] | Observations sur la réponse');
});

test('completed sheet includes all 40000 accepted messages beyond the preview', async ({page}) => {
 test.setTimeout(180000);
 await page.goto('/pipeline');
 const config = {format:'aidstack-insights-pipeline',version:1,name:'Export',
   nodes:[{id:'check',name:'Check',type:'validate',inputs:['Feedback']}]};
 const work = Array.from({length:40000}, (_,i) => ({
   node:1,status:'ok',history:[],values:{ID:i+1,Feedback:'Feedback '+(i+1),
     _source_row:i+2,_source_record_id:'source:'+i,_record_key:'source:'+i,_feedback_id:i+1}
 }));
 const recovery = {format:'aidstack-pipeline-run',version:1,runId:'large-export',nextId:40001,
   skipped:0,source:{filename:'large.csv',count:40000,columns:['ID','Feedback']},
   master:{rows:[],review:[],logs:[]},config,work};
 await page.locator('#run-file').setInputFiles({name:'saved.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(recovery))});
 await expect(page.locator('#notice')).toContainText('Run restored', {timeout:90000});
 const downloaded = page.waitForEvent('download', {timeout:90000});
 const start = Date.now();
 await page.locator('#download-results').click();
 const file = await downloaded;
 const chunks = [];
 for await (const chunk of await file.createReadStream()) chunks.push(chunk);
 console.log('40000-row, 2-column completed sheet download took '+(Date.now()-start)+'ms');
 const result = await page.evaluate(data => {
   const book=XLSX.read(data,{type:'base64'});
   const sheet=book.Sheets.Results;
   return [book.SheetNames,sheet['!ref'],sheet.A40001.v,sheet.B40001.v];
 },Buffer.concat(chunks).toString('base64'));
 expect(result).toEqual([['Results'],'A1:B40001',40000,'Feedback 40000']);
});

test('shows live progress and holds the screen awake across pause and resume', async ({page}) => {
    await page.addInitScript(() => {
        window.wakeEvents = [];
        Object.defineProperty(navigator, 'wakeLock', {value: {request: async type => {
            window.wakeEvents.push(type);
            const lock = new EventTarget();
            lock.release = async () => { window.wakeEvents.push('release'); };
            return lock;
        }}});
    });
    const requests = [];
    await page.route('**/api/pipeline/step', route => { requests.push(route); });
    const finishRow = async index => {
        await expect.poll(() => requests.length).toBe(index + 1);
        await requests[index].fulfill({json: {status:'ok', items:[{}]}});
    };
    await page.goto('/pipeline');
    const config = {format:'aidstack-insights-pipeline', version:1, name:'Progress',
        nodes:[{id:'check', name:'Check', type:'validate', inputs:['Feedback']}]};
    await page.locator('#config-file').setInputFiles({name:'config.json', mimeType:'application/json', buffer:Buffer.from(JSON.stringify(config))});
    await page.locator('#to-files').click();
    await page.locator('#source-file').setInputFiles({name:'source.csv', mimeType:'text/csv', buffer:Buffer.from('Feedback\na\nb\nc\n')});
    await setParallel(page, 1);
    await page.locator('#test-run').click();
    const bar = page.getByRole('progressbar', {name:'Workflow progress'});
    await expect(bar).toHaveAttribute('value', '0');
    await expect(page.locator('#run-activity')).toBeVisible();
    await expect(page.locator('#wake-status')).toContainText('Keeping your screen awake');
    await expect(page.locator('#run-time')).toContainText('Estimating…');
    await expect(page.locator('#run-time')).toContainText(/Elapsed [1-9]\d*s/);
    await finishRow(0);
    await expect(bar).toHaveAttribute('value', '33');
    await expect(page.locator('#run-time')).toContainText('About ');
    await expect.poll(() => requests.length).toBe(2);
    await page.locator('#pause-run').click();
    await finishRow(1);
    await expect(page.locator('#run-status')).toContainText('Paused');
    await expect(page.locator('#run-percent')).toHaveText('66%');
    await expect(page.locator('#run-activity')).toBeHidden();
    await expect(page.locator('#wake-status')).toBeHidden();
    await expect(page.locator('#run-time')).not.toContainText('Estimated');
    await page.locator('#resume-run').click();
    await expect(bar).toHaveAttribute('value', '66');
    await finishRow(2);
    await expect(page.locator('#run-status')).toContainText('Test complete');
    await expect(bar).toHaveAttribute('value', '100');
    expect(await page.evaluate(() => window.wakeEvents)).toEqual(['screen','release','screen','release']);
});

test('recovers the last committed step after closing the tab and offers a 1,000-record backup', async ({page, context}) => {
    const config = {format:'aidstack-insights-pipeline', version:1, name:'Recovery',
        nodes:[{id:'check', name:'Check', type:'validate', inputs:['Feedback']}]};
    const work = Array.from({length:1001}, (_,i) => ({
        node:i < 999 ? 1 : 0, status:i < 999 ? 'ok' : 'pending', history:[],
        values:{Feedback:String(i+1), _source_record_id:'s'+i, _record_key:'s'+i, _source_row:i+2}
    }));
    const run = {format:'aidstack-pipeline-run', version:1, runId:'recovery', nextId:1,
        source:{filename:'source.csv', count:work.length, columns:['Feedback']}, skipped:0,
        master:{rows:[], review:[], logs:[]}, config, work};
    await page.goto('/pipeline');
    await page.locator('#run-file').setInputFiles({name:'run.json', mimeType:'application/json', buffer:Buffer.from(JSON.stringify(run))});
    await expect(page.locator('#notice')).toContainText('Run restored');
    await expect(page.locator('#backup-reminder')).toBeHidden();
    const attempted = [];
    await context.route('**/api/pipeline/step', route => {
        attempted.push(route.request().postDataJSON().record.Feedback);
        if (attempted.length !== 2) return route.fulfill({json:{status:'ok',items:[{}]}});
    });
    await page.locator('#resume-run').click();
    await expect(page.locator('#backup-milestone')).toHaveText('1,000 records processed and saved.');
    const download = page.waitForEvent('download');
    await page.locator('#backup-now').click();
    expect((await download).suggestedFilename()).toContain('1000');
    await expect(page.locator('#backup-reminder')).toBeHidden();
    await expect.poll(() => attempted.length).toBe(2);
    await page.close();
    const reopened = await context.newPage();
    await reopened.goto('/pipeline');
    await reopened.locator('#open-saved-run').click();
    await expect(reopened.locator('#checkpoint-status')).toContainText('1,000 processed records');
    await expect(reopened.locator('#run-counts')).toContainText('1 pending');
    await reopened.locator('#resume-run').click();
    await expect(reopened.locator('#run-status')).toContainText('Processing complete');
    expect(attempted).toEqual(['1000','1001','1001']);
});

test('sends up to the chosen number of requests at once and stops sending after a quota error',async({page})=>{
 let active=0,peak=0,calls=0;
 await page.route('**/api/pipeline/step',async route=>{
  calls++;active++;peak=Math.max(peak,active);
  await new Promise(r=>setTimeout(r,150));active--;
  const {record}=route.request().postDataJSON();
  if(record.Feedback==='stop')return route.fulfill({json:{status:'error',errorCode:'quota_exhausted',reason:'Check account quota before retrying.'}});
  return route.fulfill({json:{status:'ok',items:[{}]}});
 });
 const config={format:'aidstack-insights-pipeline',version:2,name:'Parallel',nodes:[{id:'check',name:'Check',type:'validate',inputs:['Feedback']}]};
 await page.goto('/pipeline');
 await page.locator('#config-file').setInputFiles({name:'config.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
 await page.locator('#to-files').click();
 const rows=Array.from({length:24},(_,i)=>'row '+i).join('\n');
 await page.locator('#source-file').setInputFiles({name:'source.csv',mimeType:'text/csv',buffer:Buffer.from('Feedback\n'+rows+'\n')});
 await setParallel(page,6);
 await page.locator('#start-run').click();
 await expect(page.locator('#run-status')).toContainText('Processing complete');
 expect(peak).toBe(6);expect(calls).toBe(24);
 await expect(page.locator('#run-counts')).toContainText('24 accepted');
 // A quota error stops new requests; only those already sent finish.
 await page.locator('[data-go="files"]').click();
 await page.locator('#source-file').setInputFiles({name:'quota.csv',mimeType:'text/csv',buffer:Buffer.from('Feedback\nstop\n'+rows+'\n')});
 calls=0;peak=0;
 await page.locator('#start-run').click();
 await expect(page.locator('#run-status')).toContainText('Check account quota');
 expect(calls).toBeLessThanOrEqual(6);
 await expect(page.locator('#resume-run')).toBeVisible();
});

test('a reviewer note goes with the retry, and a reviewer can enter the answer',async({page})=>{
 const bodies=[];
 await page.route('**/api/pipeline/step',async route=>{
  const body=route.request().postDataJSON();bodies.push(body);
  if(body.node.id==='translate' && !body.note)return route.fulfill({json:{status:'review',reason:'“maso” is unclear',drafts:{english:'We accept eyes'}}});
  if(body.node.id==='translate')return route.fulfill({json:{status:'ok',items:[{english:'We accept your words'}]}});
  return route.fulfill({json:{status:'review',reason:'Unsure of topic',drafts:{topic:'Health'}}});
 });
 const config={format:'aidstack-insights-pipeline',version:2,name:'Review',nodes:[
  {id:'translate',name:'Translate',type:'analyze',inputs:['Feedback'],provider:'azure',instructions:'Translate.',fields:[{name:'english',options:[]}]},
  {id:'topic',name:'Topic',type:'classify',inputs:['english'],provider:'azure',instructions:'Pick.',fields:[{name:'topic',options:['Health','Water','No match']}],fallbackOption:'No match'}]};
 await page.goto('/pipeline');
 await page.locator('#config-file').setInputFiles({name:'config.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
 await page.locator('#to-files').click();
 await page.locator('#source-file').setInputFiles({name:'source.csv',mimeType:'text/csv',buffer:Buffer.from('Feedback\nTondimi maso bolobi\n')});
 await page.locator('#start-run').click();
 await expect(page.locator('#run-counts')).toContainText('1 review');
 await page.locator('[data-review]').click();
 await expect(page.locator('#answer-0')).toHaveValue('We accept eyes');
 await page.locator('#review-note').fill('“maso” here means words, not eyes.');
 await page.locator('#save-correction').click();
 await expect(page.locator('#results')).toContainText('Unsure of topic');
 const retried=bodies.find(b=>b.node.id==='translate' && b.note);
 expect(retried.note).toBe('“maso” here means words, not eyes.');
 expect(retried.record).toEqual({Feedback:'Tondimi maso bolobi'});
 expect(bodies.at(-1).note).toBe('“maso” here means words, not eyes.');
 // The topic step is flagged: choose the answer instead of asking the model again.
 await page.locator('[data-review]').click();
 await expect(page.locator('#answer-0 option')).toHaveText(['Choose a value','Health','Water']);
 await expect(page.locator('#answer-0')).toHaveValue('Health');
 await page.locator('#answer-0').selectOption('Water');
 const calls=bodies.length;
 await page.locator('#accept-answer').click();
 await expect(page.locator('#run-counts')).toContainText('1 accepted');
 expect(bodies.length).toBe(calls);
 await page.locator('#all-result-columns').check();
 await expect(page.locator('#results')).toContainText('Water');
});

test('junk rows are skipped without AI calls, flagged rows are discarded in bulk with a reason, and both can be restored',async({page})=>{
 const sent=[];
 await page.route('**/api/pipeline/step',async route=>{
  const {record}=route.request().postDataJSON();sent.push(record.Feedback);
  return route.fulfill({json:{status:'review',reason:'Cannot interpret'}});
 });
 const config={format:'aidstack-insights-pipeline',version:2,name:'Screened',idColumn:'',screening:{requireLetters:true,minLetters:2,randomTyping:true},
  nodes:[{id:'t',name:'Translate',type:'analyze',inputs:['Feedback'],provider:'azure',instructions:'Translate.',fields:[{name:'english',options:[]}]}]};
 await page.goto('/pipeline');
 await page.locator('#config-file').setInputFiles({name:'config.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
 await page.locator('#to-files').click();
 await page.locator('#source-file').setInputFiles({name:'source.csv',mimeType:'text/csv',buffer:Buffer.from('Feedback\n12345\nsdfgh\nkdjfa lsk\nzz kdkd\n')});
 await page.locator('#start-run').click();
 await expect(page.locator('#run-status')).toContainText('Processing complete');
 expect(sent.sort()).toEqual(['kdjfa lsk','zz kdkd']);
 await expect(page.locator('#run-counts')).toContainText('2 skipped');
 await page.locator('#result-filter').selectOption('skipped');
 await expect(page.locator('#results')).toContainText('No words: only numbers or symbols');
 await expect(page.locator('#results')).toContainText('Looks like random typing');
 await page.locator('#result-filter').selectOption('review');
 await page.locator('#select-visible').check();
 await expect(page.locator('#bulk-count')).toHaveText('2 selected');
 await page.locator('#bulk-discard').click();
 await page.locator('#discard-other').fill('Team test entries');
 await page.locator('#confirm-discard').click();
 await expect(page.locator('#run-counts')).toContainText('2 discarded');
 await expect(page.locator('#bulk-actions')).toBeHidden();
 await page.locator('#result-filter').selectOption('discarded');
 await expect(page.locator('#results')).toContainText('Reason: Team test entries');
 await page.locator('[data-restore]').first().click();
 await expect(page.locator('#run-counts')).toContainText('1 discarded');
 // A skipped row can be processed anyway; it goes to the AI like any other row.
 await page.locator('#result-filter').selectOption('skipped');
 await page.locator('[data-restore]').first().click();
 await expect(page.locator('#run-counts')).toContainText('1 skipped');
 expect(sent).toContain(12345);
});

test('filtered downloads include every matching record and its review details', async ({page}) => {
 const statuses = [...Array(101).fill('review'), 'error', 'skipped', 'discarded', 'pending'];
 const config = {format:'aidstack-insights-pipeline',version:2,name:'Review export',
   nodes:[{id:'check',name:'Check',type:'analyze',provider:'azure',instructions:'Check feedback.',
     inputs:['Feedback'],fields:[{name:'answer',options:[]}]}]};
 const work = statuses.map((status,i) => ({
   node:0,status,history:[],reason:status+' reason',
   drafts:status==='review'?{answer:'Draft '+i}:undefined,
   values:{Feedback:'Message '+i,_source_row:i+2,_source_record_id:'s:'+i,_record_key:'s:'+i,
     _discard_reason:status==='discarded'?'Duplicate':undefined}
 }));
 const recovery = {format:'aidstack-pipeline-run',version:1,runId:'filter-export',nextId:1,skipped:0,
   source:{filename:'feedback.csv',count:work.length,columns:['Feedback']},
   master:{rows:[],review:[],logs:[]},config,work};
 await page.goto('/pipeline');
 await page.locator('#run-file').setInputFiles({name:'saved.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(recovery))});
 await expect(page.locator('#notice')).toContainText('Run restored');
 await expect(page.locator('#download-results')).toBeDisabled();
 await page.getByText('Choose download columns', {exact:true}).click();
 await page.locator('#export-columns input[value="Feedback"]').uncheck();
 for (const [filter,suffix] of [['all','all_records'],['review','needs_review'],['pending','pending'],['skipped','skipped'],['discarded','discarded']]) {
   await page.locator('#result-filter').selectOption(filter);
   const format = filter==='review'?'xlsx':'csv';
   await page.locator('#export-format').selectOption(format);
   const downloaded = page.waitForEvent('download');
   await page.locator('#download-filtered').click();
   const file = await downloaded;
   expect(file.suggestedFilename()).toBe('partial_feedback_'+suffix+'.'+format);
   const chunks = [];
   for await (const chunk of await file.createReadStream()) chunks.push(chunk);
   const rows = await page.evaluate(data => {
     const book = XLSX.read(data,{type:'base64'});
     return XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]]);
   },Buffer.concat(chunks).toString('base64'));
   const expected = work.filter(e => filter==='all' || e.status===filter || (filter==='review' && e.status==='error'));
   const labels = {review:'Needs review',error:'Needs review (provider error)',pending:'Pending',skipped:'Skipped',discarded:'Discarded'};
   expect(Object.keys(rows[0]).at(-1)).toBe('Pipeline status');
   expect(rows.map(r=>[r._source_row,r['Pipeline status']])).toEqual(expected.map(e=>[e.values._source_row,labels[e.status]]));
   expect(rows[0]).not.toHaveProperty('Feedback');
   if (filter==='review') expect(rows[0]).toMatchObject({answer:'[Needs Review] | Draft 0',_reason:'review reason'});
   if (filter==='discarded') expect(rows[0]._discard_reason).toBe('Duplicate');
 }
 await page.locator('#result-filter').selectOption('ok');
 await expect(page.locator('#download-filtered')).toBeDisabled();
});
