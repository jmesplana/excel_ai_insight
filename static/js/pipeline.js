import {savePipeline, loadPipeline, hasSavedPipeline} from './pipeline-checkpoint.js';
import {ScreenWakeLock} from './screen-wake-lock.js';
import {ProgressTimer, formatDuration} from './progress-timer.js';
import {defaultWorkflow, defaultNode, portableConfig, importWorkflow, requiredInputs, outputs, seedRun, activeStep, nextWorkIndex, applyResult, retryEntry, reviewCell, scheduleRateRetry, digest} from './pipeline-core.js?v=3';
import {readSheet, readMaster} from './pipeline-workbook.js?v=3';
import {classifierFetch} from './classifier-transport.js';
import {initProviderSettings} from './provider-settings.js';

const $ = id => document.getElementById(id);
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[c]));
const lines = value => value.split('\n').map(v => v.trim()).filter(Boolean);
const option = (value, selected, label = value) => `<option value="${esc(value)}"${value === selected ? ' selected' : ''}>${esc(label)}</option>`;
let config = defaultWorkflow(), source = null, sourceIdentity = null, sourceWorkbook = null, master = {rows: [], review: [], logs: []}, masterWorkbook = null, run = null, running = false, starting = false, stopRequested = false;
let sourceToken = 0, masterToken = 0;
let stage = 'workflow', expandedNode = 0, preset = 'paired', presetToken = 0;
const friendly = value => String(value || '').replaceAll('_', ' ');
const providerName = value => ({azure:'Azure',openai:'OpenAI',jev:'Jev',laya:'Laya',ollama:'Ollama'})[value] || 'No AI';
const jsonDownload = (data, name) => {
    const url = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], {type:'application/json'}));
    const a = document.createElement('a'); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
};
function notice(text, danger = false) { $('notice').hidden = !text; $('notice').textContent = text; $('notice').className = `alert alert-${danger ? 'warning' : 'info'}`; }
function safe(handler) { return async event => { try { await handler(event); } catch (e) { notice(e.message, true); } }; }
function setStage(next, focus = true) {
    if(next === 'results' && !run) return;
    stage = next;
    for(const name of ['workflow','files','results']) $('stage-'+name).hidden = name !== next;
    document.querySelectorAll('[data-stage]').forEach(button => {
        if(button.dataset.stage===next) button.setAttribute('aria-current','step');
        else button.removeAttribute('aria-current');
        button.disabled = button.dataset.stage==='results' && !run;
    });
    document.querySelector('.workspace-menu').open = false;
    if(focus) {
        const heading = $(next === 'workflow' ? 'workflow-heading' : next === 'files' ? 'files-heading' : 'results-heading');
        heading.tabIndex=-1;heading.focus({preventScroll:true});window.scrollTo({top:0,behavior:'smooth'});
    }
}
function showEditor(open) {
    $('workflow-editor').hidden=!open;
    $('edit-workflow').setAttribute('aria-expanded', String(open));
    $('edit-workflow').textContent=open?'Hide editor':'Edit steps & models';
}
function renderOverview() {
    $('selected-workflow-name').textContent=config.name;
    $('file-workflow-summary').textContent='Workflow: '+config.name;
    const titles={extract:'Separate messages',analyze:'Analyze / translate',classify:'Classify feedback',split:'Split lists',validate:'Check values',lookup:'Fill derived columns'};
    const icons={extract:'chat-square-text',analyze:'translate',classify:'tags',split:'layout-split',validate:'check2-circle',lookup:'table'};
    const groups=[];
    config.nodes.forEach((node,i)=>{
        const title = config.nodes[i+1]?.pairedSourceStep === node.id ? 'Classify feedback types' :
            node.pairedSourceStep ? 'Split classified rows' : node.id==='translate-feedback' ? 'Translate full feedback' : titles[node.type];
        if(groups.at(-1)?.title===title) groups.at(-1).count++;
        else groups.push({type:node.type,title,count:1});
    });
    $('workflow-outline').innerHTML=groups.map(g=>`<span class="outline-step"><i class="bi bi-${icons[g.type]}" aria-hidden="true"></i>${esc(g.title)}${g.count>1?` <small>(${g.count})</small>`:''}</span>`).concat('<span class="outline-step"><i class="bi bi-file-earmark-spreadsheet" aria-hidden="true"></i>Excel results</span>').join('<span class="outline-arrow" aria-hidden="true">→</span>');
    const providers=[...new Set(config.nodes.filter(n=>n.provider).map(n=>providerName(n.provider)))];
    $('workflow-output-list').textContent=[...new Set(config.nodes.flatMap(outputs))].join(', ');
    $('workflow-summary').textContent=`${config.nodes.length} configured step${config.nodes.length===1?'':'s'}${providers.length?' · '+providers.join(', '):' · No AI connection needed'}. Original IDs and source references are kept automatically.`;
    for(const key of ['paired','ebola']) {const button=$('preset-'+key);button.classList.toggle('selected',preset===key);button.setAttribute('aria-pressed',String(preset===key));}
    $('import-config').classList.toggle('selected',preset==='imported');
    $('to-files').disabled=!config.nodes.length || running || starting;
}
function refreshFileActions() {
    const ready=Boolean(source?.rows?.length && master && config.nodes.length && !running && !starting);
    $('test-run').disabled=!ready;$('start-run').disabled=!ready;
    $('test-run').textContent=`Test ${$('test-limit').value || 3} rows`;
}
async function choosePreset(key) {
    const token=++presetToken;
    if(key==='paired')config=defaultWorkflow();
    else {
        const response=await fetch('/api/pipeline/templates/ebola');
        if(!response.ok)throw new Error('Could not load the workflow. Try again or upload its configuration file.');
        const loaded=await response.json();if(token!==presetToken)return;
        config=portableConfig(loaded);
    }
    preset=key;expandedNode=0;renderNodes();showEditor(false);notice('');
    if(source)renderSource();
}
document.querySelectorAll('[data-stage]').forEach(button=>button.onclick=safe(()=>{readEditor();renderOverview();setStage(button.dataset.stage);}));
document.querySelectorAll('[data-go]').forEach(button=>button.onclick=()=>setStage(button.dataset.go));
$('to-files').onclick=safe(()=>{portableConfig(readEditor());renderOverview();setStage('files');});
$('edit-workflow').onclick=safe(()=>{readEditor();renderOverview();showEditor($('workflow-editor').hidden);} );
$('close-editor').onclick=safe(()=>{readEditor();renderOverview();showEditor(false);$('edit-workflow').focus();});
$('preset-paired').onclick=safe(()=>choosePreset('paired'));
$('preset-ebola').onclick=safe(()=>choosePreset('ebola'));
$('edit-run').onclick=()=>{setStage('workflow');showEditor(true);};
$('run-full-file').onclick=safe(()=>begin(false));
$('test-limit').oninput=refreshFileActions;
const wakeLock = new ScreenWakeLock({onStatus: message => { $('wake-status').textContent = message; }});
const timer = new ProgressTimer();
let timerInterval = null, previousElapsed = 0;
function renderProgress(remainingSteps) {
    // Splits add downstream work. Review records have finished their automated pass.
    const totalSteps = run ? run.work.length * run.config.nodes.length : 0;
    const percent = totalSteps ? Math.floor(100 * (totalSteps - remainingSteps) / totalSteps)
        : (run ? 100 : 0);
    $('run-progress').value = percent;
    $('run-percent').textContent = percent + '%';
}
function renderTime() {
    if (!run) return;
    const active = timerInterval !== null;
    const time = active ? timer.snapshot() : null;
    if (active) run.elapsedMs = previousElapsed + time.elapsedMs;
    const eta = time?.remainingMs;
    $('run-time').textContent = 'Elapsed ' + formatDuration(run.elapsedMs || 0) +
        (active ? ' · Estimated time left this pass: ' + (eta == null ? 'Estimating…'
            : eta === 0 ? 'Finishing…' : 'About ' + formatDuration(Math.ceil(eta / 1000) * 1000)) : '');
}
let checkpointFailed = false, savedRecords = 0, lastBackupRecords = 0;
const processedRecords = work => work.reduce((count, entry) => count + Number(entry.status !== 'pending'), 0);
function renderCheckpoint() {
    $('checkpoint-status').textContent = checkpointFailed
        ? 'Browser save failed. Download a recovery file to keep your latest progress.'
        : `Saved in this browser · ${savedRecords.toLocaleString()} processed records. You can resume from the last saved step.`;
    const milestone = Math.floor(savedRecords / 1000) * 1000;
    $('backup-reminder').hidden = checkpointFailed || milestone <= lastBackupRecords;
    $('backup-milestone').textContent = `${milestone.toLocaleString()} records processed and saved.`;
}
async function saveCheckpoint(index, replace, newlyProcessed = 0) {
    renderTime();
    try {
        const full = checkpointFailed || index === undefined;
        await savePipeline(run, full ? undefined : index, replace);
        savedRecords = full ? processedRecords(run.work) : savedRecords + newlyProcessed;
        checkpointFailed = false;
    } catch (error) {
        checkpointFailed = true;
        throw error;
    } finally {
        renderCheckpoint();
    }
}
function advanced(n) {
    const keys = ['sourceBranches', 'branches', 'minConfidence', 'responseMode', 'table', 'rules', 'reviewValues'];
    return Object.fromEntries(keys.filter(k => n[k] !== undefined).map(k => [k,n[k]]));
}
function renderNodes() {
    $('workflow-name').value = config.name;
    $('nodes').innerHTML = config.nodes.map((n, i) => {
        const ai = ['extract','analyze','classify'].includes(n.type);
        return `<article class="pipeline-node" data-index="${i}"><div class="node-top"><button type="button" class="node-toggle" data-action="toggle" aria-expanded="${expandedNode===i}" aria-controls="node-body-${i}"><span class="node-number">${i+1}</span><span class="node-label"><strong>${esc(friendly(n.name))}</strong><small>${esc(providerName(n.provider))}${n.model?' · '+esc(n.model):''} · ${outputs(n).length} output${outputs(n).length===1?'':'s'}</small></span><i class="bi bi-chevron-${expandedNode===i?'up':'down'} node-chevron" aria-hidden="true"></i></button><div class="node-actions"><button class="btn btn-outline-secondary" data-action="up" aria-label="Move step ${i+1} up" ${!i?'disabled':''}>↑</button><button class="btn btn-outline-secondary" data-action="down" aria-label="Move step ${i+1} down" ${i===config.nodes.length-1?'disabled':''}>↓</button><button class="btn btn-outline-danger" data-action="remove" aria-label="Remove step ${i+1}">×</button></div></div><div class="node-body" id="node-body-${i}" ${expandedNode===i?'':'hidden'}>
<label class="form-label small" for="node-name-${i}">Step name</label><input id="node-name-${i}" data-key="name" class="form-control mb-2" value="${esc(n.name)}">
<label class="form-label small" for="node-inputs-${i}">Input columns (one per line)</label><textarea id="node-inputs-${i}" data-key="inputs" class="form-control mb-2" rows="2">${esc(n.inputs.join('\n'))}</textarea>
${ai ? `<div class="row g-2 mb-2"><div class="col-sm-5"><label class="form-label small" for="node-provider-${i}">Provider</label><select id="node-provider-${i}" data-key="provider" class="form-select">${(n.type==='classify'?['azure','openai','jev','laya','ollama']:['azure','openai','ollama']).map(p=>option(p,n.provider,({azure:'Azure AI Foundry',openai:'OpenAI',jev:'Jev',laya:'Laya · local',ollama:'Ollama · local'})[p])).join('')}</select></div><div class="col-sm-7"><label class="form-label small" for="node-model-${i}">Model / Azure deployment</label><input id="node-model-${i}" data-key="model" class="form-control" value="${esc(n.model)}" placeholder="Use this provider’s connection settings"></div></div><label class="form-label small" for="node-instructions-${i}">Instructions</label><textarea id="node-instructions-${i}" data-key="instructions" class="form-control mb-2" rows="4">${esc(n.instructions)}</textarea>
<div class="fields">${n.fields.map((f,j)=>fieldHTML(f,i,j)).join('')}</div>${n.type!=='classify'?'<button class="btn btn-link btn-sm px-0" data-action="field">+ Add output field</button>':''}
<details class="subtle-details"><summary>Response length</summary><label class="form-label small d-block mt-2" for="node-tokens-${i}">Maximum output tokens</label><input id="node-tokens-${i}" data-key="maxOutputTokens" class="form-control mb-2" type="number" min="64" max="32768" value="${n.maxOutputTokens || 4096}"></details>`:''}
${n.type==='extract'?`<div class="row g-2"><div class="col-sm-6"><label class="form-label small" for="node-verbatim-input-${i}">Original text input</label><input id="node-verbatim-input-${i}" data-key="verbatimInput" class="form-control" value="${esc(n.verbatimInput)}"></div><div class="col-sm-6"><label class="form-label small" for="node-verbatim-field-${i}">Verbatim excerpt output</label><input id="node-verbatim-field-${i}" data-key="verbatimField" class="form-control" value="${esc(n.verbatimField)}"></div></div><p class="small text-muted mt-2">${config.nodes[i+1]?.pairedSourceStep === n.id ? 'Keeps each source row intact with aligned types and excerpts until the file-wide split.' : 'Returns paired objects and expands them into rows automatically.'} Every excerpt is checked against the original text.</p>`:''}
${n.type==='split'?`<label class="form-label small" for="node-delimiter-${i}">Exact separator</label><input id="node-delimiter-${i}" data-key="delimiter" class="form-control" value="${esc(n.delimiter)}"><p class="small text-muted mt-2">All selected columns must have equal item counts. Repeated items are kept. Mismatches go to review.</p>`:''}
${n.type==='lookup'?`<label class="form-label small" for="node-output-${i}">Output column</label><input id="node-output-${i}" data-key="output" class="form-control" value="${esc(n.output)}">`:''}
${n.type==='validate'?'<p class="small text-muted">Every input must be nonempty. Add allowed-value rules below if needed.</p>':''}
<details class="mt-2"><summary class="small">${n.type==='lookup'?'Lookup table':n.type==='validate'?'Allowed-value rules':'Advanced rules & response format'} (JSON)</summary><p class="small text-muted mt-2">${n.type==='lookup'?'table: [{"when": {"input": "value"}, "value": "mapped result"}]':n.type==='validate'?'rules: [{"column": "category", "allowed": ["A", "B"]}]':'Branches and option descriptions are preserved on import. minConfidence applies only to Jev/Laya. Set responseMode to "json" for models without strict schema support; outputs are still validated. reviewValues lists labels that must go to human review.'}</p><textarea class="form-control advanced-json" aria-label="Advanced rules for step ${i+1}" rows="5">${esc(JSON.stringify(advanced(n),null,2))}</textarea></details></div></article>`;
    }).join('');
    renderMappings();renderOverview();refreshFileActions();
}
function fieldHTML(f, i, j) {
    return `<div class="field-row" data-field="${j}"><div class="d-flex justify-content-between"><label class="form-label small" for="field-${i}-${j}">Output column</label><button class="btn btn-link btn-sm text-danger p-0" data-action="remove-field" data-field="${j}" aria-label="Remove output ${j+1} from step ${i+1}">Remove</button></div><input id="field-${i}-${j}" data-field-name class="form-control mb-2" value="${esc(f.name)}"><label class="form-label small" for="options-${i}-${j}">Allowed values (one per line; leave blank for text)</label><textarea id="options-${i}-${j}" data-field-options class="form-control" rows="2">${esc((f.options||[]).map(o=>typeof o==='string'?o:o.label).join('\n'))}</textarea></div>`;
}
function readEditor() {
    config.name = $('workflow-name').value;
    $('nodes').querySelectorAll('[data-index]').forEach(card => {
        const n = config.nodes[Number(card.dataset.index)];
        card.querySelectorAll('[data-key]').forEach(el => { const k = el.dataset.key; n[k] = k==='inputs'?lines(el.value):k==='maxOutputTokens'?Number(el.value):el.value; });
        if (n.fields) n.fields = [...card.querySelectorAll('.field-row')].map((el,j)=>({name: el.querySelector('[data-field-name]').value.trim(), options: lines(el.querySelector('[data-field-options]').value).map(label=>n.fields[j]?.options?.find(o => (typeof o==='string'?o:o.label)===label) || label)}));
        const extra = JSON.parse(card.querySelector('.advanced-json').value || '{}');
        const allowed = ['sourceBranches','branches','minConfidence','responseMode','table','rules','reviewValues'];
        if (!extra || typeof extra !== 'object' || Array.isArray(extra) || Object.keys(extra).some(k=>!allowed.includes(k))) throw new Error('Unknown advanced step setting.');
        allowed.forEach(k=>delete n[k]); Object.assign(n,extra);
    });
    return config;
}
function renderMappings() {
    $('column-mappings').innerHTML = requiredInputs(config.nodes).map((key, i) => {
        const mapped=config.mappings[key] || key;
        const example=source?.rows.find(r=>r.values[mapped]!==null && r.values[mapped]!==undefined && r.values[mapped]!=='')?.values[mapped];
        return `<label class="form-label small" for="mapping-${i}">Use this column for <strong>${esc(friendly(key))}</strong></label><select class="form-select" id="mapping-${i}" data-mapping="${esc(key)}">${option('', '', 'Choose a column')}${(source?.columns||[]).map(c=>option(c,mapped)).join('')}</select><p class="mapping-example">${example===undefined?'Choose the column containing this information.':'Example: '+esc(String(example).slice(0,160))}</p>`;
    }).join('');
}
$('nodes').addEventListener('click', safe(event => {
    const button=event.target.closest('[data-action]'); if(!button) return;
    readEditor(); const i=Number(button.closest('[data-index]').dataset.index), action=button.dataset.action;
    if(action==='toggle')expandedNode=expandedNode===i?-1:i;
    if(action==='remove') config.nodes.splice(i,1);
    if(action==='up' || action==='down') { const j=i+(action==='up'?-1:1); [config.nodes[i],config.nodes[j]]=[config.nodes[j],config.nodes[i]]; }
    if(action==='field') config.nodes[i].fields.push({name:'new_output',options:[]});
    if(action==='remove-field') config.nodes[i].fields.splice(Number(button.dataset.field),1);
    renderNodes();
}));
$('nodes').addEventListener('change', safe(()=>{readEditor();renderMappings();renderOverview();}));
$('column-mappings').addEventListener('change', event=>{if(event.target.dataset.mapping) {config.mappings[event.target.dataset.mapping]=event.target.value;renderMappings();}});
$('original-id').onchange=()=>{config.idColumn=$('original-id').value;};
$('add-step').onclick=safe(()=>{readEditor();config.nodes.push(defaultNode($('step-kind').value));expandedNode=config.nodes.length-1;preset='custom';renderNodes();});
$('new-workflow').onclick=()=>{config=defaultWorkflow();preset='custom';expandedNode=0;renderNodes();setStage('workflow');showEditor(true);notice('Change the first step or add steps to create your workflow.');};
$('export-config').onclick=safe(()=>{jsonDownload(portableConfig(readEditor()),'pipeline-config.json');});
$('import-config').onclick=()=>$('config-file').click();
$('append-config').onclick=()=>$('append-file').click();
async function importConfig(file, append=false) {
    if(!file) return;
    const imported=importWorkflow(JSON.parse(await file.text()));
    if(append) {
        readEditor();
        const ids=new Map(imported.config.nodes.map(n=>[n.id,crypto.randomUUID()]));
        for(const n of imported.config.nodes) {
            const added={...n,id:ids.get(n.id)};
            if(n.pairedSourceStep) added.pairedSourceStep=ids.get(n.pairedSourceStep);
            config.nodes.push(added);
        }
    }
    else config=imported.config;
    presetToken++;preset='imported';expandedNode=append?config.nodes.length-imported.config.nodes.length:0;
    renderNodes(); if(source) renderSource(); setStage('workflow');showEditor(append);notice(append?'Steps added. Open a step to check its settings.':'Workflow loaded. Review the summary, then continue to files.');
}
$('config-file').onchange=safe(e=>importConfig(e.target.files[0]));
$('append-file').onchange=safe(e=>importConfig(e.target.files[0],true));
async function loadSource(bytes, filename, token) {
    const datasetId=await digest(bytes);
    if(token!==sourceToken) return;
    let data = bytes;
    if (/\.csv$/i.test(filename)) {
        try { data = new TextDecoder('utf-8', {fatal:true}).decode(bytes); }
        catch { /* Let SheetJS detect the encoding of older CSV files. */ }
    }
    sourceWorkbook=XLSX.read(data,{type:typeof data === 'string' ? 'string' : 'array',cellDates:false});
    sourceIdentity={filename,datasetId};
    $('source-sheet').innerHTML=sourceWorkbook.SheetNames.map(n=>option(n,sourceWorkbook.SheetNames[0])).join('');
    renderSource();
}
function renderSource() {
    if(!sourceWorkbook) return;
    const parsed=readSheet(sourceWorkbook,$('source-sheet').value,Number($('header-row').value));
    source={...sourceIdentity,...parsed,sheet:$('source-sheet').value};
    $('original-id').innerHTML=option('',config.idColumn,'No existing ID')+source.columns.map(c=>option(c,config.idColumn)).join('');
    if(!source.columns.includes(config.idColumn)) config.idColumn='';
    $('source-summary').textContent=`${source.filename} · ${source.rows.length} source rows · ${source.columns.length} columns`;
    $('source-options').hidden=false;
    if(sourceWorkbook.SheetNames.length>1)$('sheet-options').open=true;
    const columns=source.columns.slice(0,4);
    $('source-preview').innerHTML='<table class="table table-sm"><thead><tr>'+columns.map(k=>'<th>'+esc(k)+'</th>').join('')+'</tr></thead><tbody>'+source.rows.slice(0,3).map(r=>'<tr>'+columns.map(k=>'<td>'+esc(String(r.values[k]??'').slice(0,240))+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
    renderMappings();refreshFileActions();
}
$('source-file').onchange=safe(async e=>{const f=e.target.files[0];if(!f)return;const token=++sourceToken;source=null;sourceWorkbook=null;refreshFileActions();await loadSource(await f.arrayBuffer(),f.name,token);});
$('source-sheet').onchange=$('header-row').onchange=safe(()=>{try{renderSource();}catch(e){source=null;refreshFileActions();throw e;}});
$('pull-source').onclick=safe(async()=>{
    const url=new URL($('source-url').value); if(!['https:','http:'].includes(url.protocol) || url.username || url.password) throw new Error('Use an HTTP(S) file URL without embedded credentials.');
    const token=++sourceToken;source=null;sourceWorkbook=null;refreshFileActions();
    notice('Fetching source file…');
    const response=await fetch(url,{credentials:'omit',signal:AbortSignal.timeout(30000)});
    if(!response.ok) throw new Error(`Source download failed (${response.status}). Use a direct file URL or upload the exported file.`);
    const bytes=await response.arrayBuffer();
    await loadSource(bytes,decodeURIComponent(url.pathname.split('/').pop()) || 'source.csv',token);notice('Source loaded. Map the columns before running.');
});
function renderMaster() {
    master=readMaster(masterWorkbook,$('master-sheet').value);$('master-options').hidden=false;refreshFileActions();
    $('master-summary').textContent=`${master.rows.length} existing rows; ${master.review.length} review records. Matching imported source records will be skipped.`;
}
$('master-file').onchange=safe(async e=>{const f=e.target.files[0];if(!f)return;const token=++masterToken;master=null;refreshFileActions();const bytes=await f.arrayBuffer();if(token!==masterToken)return;masterWorkbook=XLSX.read(bytes,{type:'array'});$('master-sheet').innerHTML=masterWorkbook.SheetNames.map(n=>option(n,masterWorkbook.SheetNames[0])).join('');renderMaster();});
$('master-sheet').onchange=safe(()=>{master=null;refreshFileActions();renderMaster();});
$('clear-master').onclick=()=>{masterToken++;masterWorkbook=null;master={rows:[],review:[],logs:[]};$('master-file').value='';$('master-sheet').innerHTML='';$('master-summary').textContent='A new master workbook will be created.';$('master-options').hidden=true;refreshFileActions();};
function credentials(provider) {
    const value=id=>$(id)?.value.trim() || '';
    if(provider==='azure') return {apiKey:value('modal-azure-api-key'),azureEndpoint:value('modal-azure-endpoint'),azureDeployment:value('modal-azure-deployment'),azureApiVersion:value('modal-azure-api-version')};
    if(provider==='jev') return {jevApiKey:value('modal-jev-api-key'),jevModel:value('modal-jev-model')};
    if(provider==='laya') return {layaApiKey:value('modal-laya-api-key'),layaEndpoint:value('modal-laya-endpoint'),layaModel:value('modal-laya-model')};
    if(provider==='ollama') return {ollamaEndpoint:value('modal-ollama-endpoint'),ollamaModel:value('modal-ollama-model')};
    return {apiKey:value('modal-api-key'),model:value('modal-openai-model')};
}
const lookupIndexes = new WeakMap();
async function callStep(node, values) {
    const record=Object.fromEntries(node.inputs.map(k=>[k,values[k]]));
    if (node.type === 'lookup' && node.inputs.every(key => typeof record[key] === 'string')) {
        let index = lookupIndexes.get(node);
        if (!index) {
            index = new Map(node.table.map(row => [JSON.stringify(node.inputs.map(key => row.when[key])),row.value]));
            lookupIndexes.set(node,index);
        }
        const key = JSON.stringify(node.inputs.map(key => record[key]));
        if (!index.has(key)) return {status:'review',reason:'No configured lookup matches this record.'};
        const value = index.get(key), items = [{[node.output]:value}];
        return (node.reviewValues || []).includes(value)
            ? {status:'review',items,reason:'The selected value is configured to require human review.'}
            : {status:'ok',items};
    }
    if(node.provider!=='laya') return post('/api/pipeline/step',{node,record,credentials:credentials(node.provider)});
    let options=node.fields[0].options || [];
    for(const key of ['sourceBranches','branches']) if(node[key]?.length) {
        const matches=node[key].filter(b=>Object.entries(b.when).every(([k,v])=>record[k]===v));
        if(matches.length!==1) return {status:'review',reason:'Missing, unrecognized or ambiguous value for a codebook branch.'};
        options=matches[0].options;
    }
    const name=node.fields[0].name;
    const payload={...credentials('laya'),classificationProvider:'laya',configs:[{outputColumnName:name,questionType:'choice',instructions:node.instructions,options,review:{minConfidence:node.minConfidence ?? .6}}],rows:[{rowIndex:0,state:{record,decisions:record}}]};
    if(node.model)payload.layaModel=node.model;
    const response=await classifierFetch('/analyze_batch_jev',payload);
    const result=await response.json();
    if(!response.ok)throw new Error(result.error || 'Local classification failed.');
    const decision=result.results?.[0]?.decisions?.[name];
    if(!decision)throw new Error('Local classifier returned an incomplete response.');
    const status=(node.reviewValues || []).includes(decision.value)?'review':['ok','review','error'].includes(decision.status)?decision.status:'review';
    return {status,reason:decision.reason || (status==='review'?'Selected value needs human review.':''),items:[{[name]:decision.value}],model:decision.model,audit:{[name]:decision}};
}
async function post(path, body) {
    const response=await fetch(path,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body),signal:AbortSignal.timeout(180000)});
    let result;try{result=await response.json();}catch{throw new Error('The server returned an incomplete response. Resume to retry the pending step.');}
    if(!response.ok) throw new Error(result.error || `Request failed (${response.status}).`);
    return result;
}
function lock(locked) {
    $('source-controls').disabled=locked;$('master-controls').disabled=locked;$('workflow-controls').disabled=locked;
    for(const id of ['new-workflow','import-config','export-config','test-run','start-run','test-limit','restore-run','load-run','apply-run-settings']) $(id).disabled=locked;
    for(const id of ['preset-paired','preset-ebola','to-files'])$(id).disabled=locked;
    $('pause-run').disabled=!locked;refreshFileActions();
}
async function begin(test) {
    if(running || starting) return;
    starting=true;lock(true);
    try {
    if(!source?.rows || !source.rows.length) throw new Error('Load a populated source file first.');
    if(!master) throw new Error('Select a valid master worksheet or start a new master.');
    const frozen=portableConfig(readEditor());
    await post('/api/pipeline/validate',{nodes:frozen.nodes});
    const limit=Number($('test-limit').value);
    if(test && (!Number.isInteger(limit)||limit<1||limit>100)) throw new Error('Test 1–100 source rows.');
    run=seedRun(frozen,{...source,rows:test?source.rows.slice(0,limit):source.rows},structuredClone(master));
    $('saved-run-banner').hidden=true;
    $('backup-reminder').hidden=true;
    $('checkpoint-status').textContent='Saving this run in your browser…';
    lastBackupRecords=0;
    run.source.columns=source.columns;run.source.uploadedCount=source.rows.length;run.test=!!test;$('result-filter').value='all';setStage('results');notice('');renderRun();
    $('run-status').textContent='Preparing the run…';
    await continueRun({newRun:true});
    } finally { starting=false;if(!running)lock(false);renderRun(); }
}
async function waitForProvider(entry, node) {
    const retryAt=Math.max(Number.isFinite(entry.retryAt)?entry.retryAt:0, Number.isFinite(run.providerCooldowns?.[node.provider])?run.providerCooldowns[node.provider]:0);
    while(retryAt > Date.now()) {
        if(stopRequested) return false;
        const seconds=Math.ceil((retryAt-Date.now())/1000);
        $('run-status').textContent=`${providerName(node.provider)} asked us to slow down. Retrying ${friendly(node.name)} in ${seconds}s. Completed steps are saved.`;
        await new Promise(resolve=>setTimeout(resolve,250));
    }
    return !stopRequested;
}
async function continueRun({newRun = false} = {}) {
    if(running||!run)return;
    running=true;stopRequested=false;delete run.pauseReason;setStage('results');lock(true);renderRun();
    try {
        wakeLock.start();
        const reviewsOnly = run.work.some(e => e.status === 'pending' && e.reviewRetry);
        const remaining = entry => entry.status === 'pending' && (!reviewsOnly || entry.reviewRetry) ? run.config.nodes.length - entry.node : 0;
        let unitsLeft = run.work.reduce((sum, entry) => sum + remaining(entry), 0), finished = 0;
        const otherSteps = reviewsOnly ? run.work.reduce((sum, entry) =>
            sum + (entry.status === 'pending' && !entry.reviewRetry ? run.config.nodes.length - entry.node : 0), 0) : 0;
        previousElapsed = run.elapsedMs || 0;
        timer.start({total:unitsLeft});
        timerInterval = setInterval(renderTime, 1000);
        renderTime();
        if (newRun || checkpointFailed) await saveCheckpoint(undefined,newRun);
        let lastPaint = performance.now();
        const cursor={phase:0,index:0};
        while(!stopRequested) {
            const index=nextWorkIndex(run,reviewsOnly,cursor);if(index<0)break;
            const entry=run.work[index], {node,values}=activeStep(run,entry);
            if(!await waitForProvider(entry,node)) break;
            $('run-status').textContent=`${run.test?'Test · ':''}${node.name}: original row ${entry.values._source_row}, message ${entry.values._message_number}`;
            const result=node.pairedSourceStep && entry.pairedItems ? {status:'ok',items:entry.pairedItems} : await callStep(node,values);
            const rateAction=scheduleRateRetry(run,index,result);
            if(rateAction==='retry' || rateAction==='pause') {
                if(rateAction==='pause') {
                    run.pauseReason=`${providerName(node.provider)} requested a longer wait. Continue processing after ${new Date(entry.retryAt).toLocaleTimeString()}. Completed steps are saved.`;
                    await saveCheckpoint(index);break;
                }
                await saveCheckpoint(index);renderRun();continue;
            }
            if(rateAction==='exhausted')result.reason='The provider is still limiting requests after three attempts. Check its limits or wait before retrying this step.';
            const before = remaining(entry), length = run.work.length;
            applyResult(run,index,result);
            unitsLeft -= before;
            let newlyProcessed = 0;
            for (let j = index; j <= index + run.work.length - length; j++) {
                unitsLeft += remaining(run.work[j]);
                if (run.work[j].status !== 'pending') newlyProcessed++;
            }
            renderProgress(unitsLeft + otherSteps);
            finished++;
            timer.update(finished,finished + unitsLeft);
            if(rateAction==='exhausted' || result.errorCode==='quota_exhausted' || result.errorCode==='authentication') {
                run.pauseReason=rateAction==='exhausted'
                    ? 'Paused after three rate-limit responses. Check the provider limits, then retry failed steps. Completed steps are saved.'
                    : result.reason;
                await saveCheckpoint(index,false,newlyProcessed);renderRun();break;
            }
            await saveCheckpoint(index,false,newlyProcessed);
            if (performance.now() - lastPaint >= 1000) { renderRun();lastPaint=performance.now(); }
        }
        $('run-status').textContent=run.pauseReason || (run.work.some(e=>e.status==='pending')?'Paused. Pending steps can be resumed.':`${run.test?'Test complete':'Processing complete'}. Review the flagged records, then download the workbook.`);
    } catch(e) { notice(e.message,true);$('run-status').textContent='Paused. Saved results are available to download.'; }
    finally {
        wakeLock.stop();
        renderTime();
        clearInterval(timerInterval);
        timerInterval=null;
        running=false;
        lock(false);
        renderRun();
    }
}
$('test-run').onclick=safe(()=>begin(true));$('start-run').onclick=safe(()=>begin(false));
$('pause-run').onclick=()=>{stopRequested=true;$('run-status').textContent='Pausing after the current step…';};
$('resume-run').onclick=safe(()=>continueRun());
$('apply-run-settings').onclick=safe(async()=>{
    const updated=portableConfig(readEditor());
    if(updated.nodes.length!==run.config.nodes.length || updated.nodes.some((n,i)=>n.id!==run.config.nodes[i].id || n.type!==run.config.nodes[i].type || n.pairedSourceStep!==run.config.nodes[i].pairedSourceStep || JSON.stringify(n.inputs)!==JSON.stringify(run.config.nodes[i].inputs) || JSON.stringify(outputs(n))!==JSON.stringify(outputs(run.config.nodes[i])))) throw new Error('Keep the same step order, types, inputs and output names when updating an existing run. Start a new run for structural changes.');
    await post('/api/pipeline/validate',{nodes:updated.nodes});
    run.revisions ||= [];run.revisions.push({at:new Date().toISOString(),config:run.config});
    for(const entry of run.work) if(entry.status!=='ok')entry.history.push({status:'settings_updated',at:new Date().toISOString(),revision:run.revisions.length});
    run.config=updated;await saveCheckpoint();notice('Updated settings for unfinished work. Accepted records are retained. Retry a review record or resume pending work.');
});

$('retry-reviews').onclick=safe(async()=>{
    if (running) return;
    run.work.forEach((entry,index)=>{
        delete entry.reviewRetry;
        if (entry.status === 'review') {
            const current=run.work[index]===entry ? index : run.work.indexOf(entry);
            if(current>=0) retryEntry(run,current).reviewRetry = true;
        }
    });
    await saveCheckpoint();await continueRun();
});
$('retry-extraction-reviews').onclick=safe(async()=>{
    $('run-status').textContent='Retrying extraction reviews…';
    run.work.forEach((entry,index)=>{if(entry.status==='review' && run.config.nodes[entry.node]?.type==='extract')retryEntry(run,index);});
    await saveCheckpoint();await continueRun();
});
$('retry-run').onclick=safe(async()=>{run.work.forEach((e,i)=>{if(e.status==='error')retryEntry(run,i);});await saveCheckpoint();await continueRun();});
let exportingAudit = false;
function downloadBlob(file, filename) {
    const url = URL.createObjectURL(file), link = document.createElement('a');
    link.href=url;link.download=filename;link.click();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
}
$('download-master').onclick=safe(async()=>{
    if (exportingAudit) return;
    const name = (run.test?'test_':run.work.some(e=>e.status==='pending')?'partial_':'')+'updated_master.xlsx';
    const worker = new Worker(new URL('./pipeline-results-worker.js',import.meta.url));
    exportingAudit=true;$('download-master').disabled=true;
    try {
        const file = await new Promise((resolve,reject)=>{
            worker.onerror = e => reject(new Error(e.message || 'Audit export failed.'));
            worker.onmessage = ({data}) => data.error ? reject(new Error(data.error)) : resolve(data.file);
            worker.postMessage({audit:run});
        });
        downloadBlob(file,name);
    } finally {worker.terminate();exportingAudit=false;renderRun();}
});
$('download-run').onclick=$('backup-now').onclick=()=>{
    jsonDownload(run,`pipeline-recovery_${run.source.filename}_${savedRecords}.json`);
    lastBackupRecords=savedRecords;
    renderCheckpoint();
};
async function restore(saved, imported=false) {
    if(!saved || saved.format!=='aidstack-pipeline-run' || saved.version!==1 || !Array.isArray(saved.work) || !saved.master || !Array.isArray(saved.master.rows) || !Array.isArray(saved.master.review) || !Array.isArray(saved.master.logs) || !Number.isSafeInteger(saved.nextId) || saved.nextId<1 || !saved.source) throw new Error('Invalid pipeline recovery file.');
    saved.config=portableConfig(saved.config);
    const keys=new Set(), ids=new Set();
    for(const e of saved.work) {
        if(!e.values || !Array.isArray(e.history) || !Number.isInteger(e.node) || e.node<0 || e.node>saved.config.nodes.length || !['ok','pending','review','error','discarded'].includes(e.status) || e.status==='pending' && e.node>=saved.config.nodes.length || typeof e.values._record_key!=='string' || keys.has(e.values._record_key)) throw new Error('Recovery file has invalid records.');
        keys.add(e.values._record_key);
        if(e.values._feedback_id) {if(!Number.isSafeInteger(e.values._feedback_id) || e.values._feedback_id>=saved.nextId || ids.has(e.values._feedback_id))throw new Error('Recovery file has invalid IDs.');ids.add(e.values._feedback_id);}
    }
    savedRecords=processedRecords(saved.work);
    lastBackupRecords=0;
    $('saved-run-banner').hidden=true;
    renderCheckpoint();
    run=saved;config=structuredClone(run.config);preset='imported';renderNodes();renderRun();setStage('results');$('run-status').textContent=run.pauseReason || 'Run restored. Continue pending work or retry a failed step.';if(imported || !saved.source.columns) await saveCheckpoint(undefined,imported);notice('Run restored. Credentials come from Connections & models. Resume pending steps or correct a review record.');
}
$('restore-run').onclick=$('open-saved-run').onclick=safe(async()=>{
    if (!running && !starting) await restore(await loadPipeline());
});
$('load-run').onclick=()=>$('run-file').click();
$('run-file').onchange=safe(async e=>{if(e.target.files[0])await restore(JSON.parse(await e.target.files[0].text()),true);});
function failureDetails(entry) {
    if(!run || !['review','error','pending'].includes(entry.status))return '';
    const {node}=activeStep(run,entry);
    if(!node)return '';
    const step=entry.status==='pending'?'Current step':(entry.translationFlag || entry.drafts?.feedback_path)?'Flagged at':'Stopped at';
    let html=`<div class="review-reason"><strong>${step}:</strong> ${esc(friendly(node.name))}${node.provider?' · '+esc(providerName(node.provider)):''}</div>`;
    if(node.type==='extract' && entry.status!=='pending') {
        let original=entry.values[node.verbatimInput];
        try { const saved=JSON.parse(entry.values._original_record);original=saved[run.config.mappings[node.verbatimInput] || node.verbatimInput] ?? original; } catch { /* Older checkpoints retain the input directly. */ }
        html+=`<div class="review-source"><strong>Source text:</strong> ${esc(original ?? '(empty)')}</div>`;
    }
    return html;
}
function statusLabel(entry) {
    if(entry.errorCode==='rate_limit')return 'Rate limited';
    if(entry.errorCode==='quota_exhausted')return 'Quota / billing limit';
    if(entry.errorCode==='authentication')return 'Check credentials';
    if(entry.status==='pending' && entry.retryAt>Date.now())return 'Waiting for provider';
    return ({ok:'Accepted',review:'Needs review',error:'Provider error',pending:'Processing',discarded:'Discarded'})[entry.status];
}
let exportRun, exporting = false;
function configureExport() {
    if (!run || exportRun === run) return;
    exportRun = run;
    const fields = run.config.nodes.flatMap(outputs);
    const source = run.source.columns || Object.keys(JSON.parse(run.work[0]?.values._original_record || '{}'));
    const columns = [...new Set([...source, ...fields])];
    $('export-columns').innerHTML = columns.map(key =>
        `<label><input type="checkbox" value="${esc(key)}" checked> ${esc(key)}</label>`).join('');
}
function selectedColumns() {
    return [...$('export-columns').querySelectorAll('input:checked')].map(input => input.value);
}
function completedDownloadState() {
    $('download-results').disabled = !run || exporting || !selectedColumns().length || !run.work.some(e => e.status === 'ok');
}
$('export-columns').onchange = completedDownloadState;
$('download-results').onclick = safe(async () => {
    const columns = selectedColumns(), records = run.work.filter(e => e.status === 'ok');
    if (!columns.length || !records.length || exporting) return;
    const format = $('export-format').value;
    const filename = (run.test ? 'test_' : run.work.some(e => e.status === 'pending') ? 'partial_' : '') +
        run.source.filename.replace(/\.[^.]+$/, '') + '_completed.' + format;
    const worker = new Worker(new URL('./pipeline-results-worker.js',import.meta.url));
    exporting = true;
    completedDownloadState();
    try {
        const file = await new Promise((resolve,reject) => {
            let offset = 0;
            worker.onerror = event => reject(new Error(event.message || 'Excel export failed.'));
            worker.onmessage = ({data}) => {
                if (data.error) { reject(new Error(data.error)); return; }
                if (data.file) { resolve(data.file); return; }
                if (offset >= records.length) { worker.postMessage({finish:true}); return; }
                const batch = records.slice(offset,offset+512);
                offset += batch.length;
                worker.postMessage({rows:batch.map(e => columns.map(key => e.values[key] ?? null))});
            };
            worker.postMessage({headers:columns,format});
        });
        downloadBlob(file, filename);
    } finally {worker.terminate();exporting=false;completedDownloadState();}
});
function renderRun() {
    const entries=run?.work||[];
    const counts = {ok:0,pending:0,review:0,error:0,discarded:0};
    const stoppedColumns=new Set();
    const pairedSplit = run?.config.nodes.findIndex(n => n.pairedSourceStep) ?? -1;
    let multiple=0, previousId, siblings=0, awaitingSplit=false;
    const steps = run?.config.nodes.length || 0;
    let remainingSteps = 0;
    for (const e of entries) {
        counts[e.status]++;
        if (e.status === 'pending') remainingSteps += steps - e.node;
        if (e.status==='pending' && e.node<=pairedSplit) awaitingSplit=true;
        const id=e.values._source_record_id;
        siblings = id === previousId ? siblings+1 : 1;
        previousId=id;
        if (siblings===2) multiple++;
        if (e.status==='review') {
            for (const key of [...outputs(run.config.nodes[e.node]), ...Object.keys(e.drafts || {})]) stoppedColumns.add(key);
            if (e.translationFlag) { stoppedColumns.add('feedback_path'); stoppedColumns.add('feedback_category'); }
        }
    }
    renderProgress(remainingSteps);
    $('run-activity').hidden = !running;
    const count=status=>counts[status];
    configureExport();completedDownloadState();renderTime();
    $('split-summary').textContent=run ? `${run.test?'Test · ':''}${run.source.count} source rows in run${run.source.uploadedCount ? ' of '+run.source.uploadedCount+' uploaded' : ''} · ${multiple} source${multiple===1?'':'s'} split into multiple messages · ${entries.length} current records${pairedSplit>=0 ? (awaitingSplit ? ' · Split in progress or pending' : ' · Split stage complete for eligible rows') : ''}` : '';
    $('run-counts').innerHTML=run?`<span>${count('ok')} accepted</span><span>${count('review')+count('error')} review</span><span>${count('pending')} pending</span><span>${run.skipped} sources skipped</span>`:'';
    $('resume-run').disabled=running||!count('pending');$('retry-run').disabled=running||!count('error');
    $('apply-run-settings').disabled=running||!run;
    $('download-master').disabled=exportingAudit || !run || !entries.some(e=>e.status!=='pending');$('download-run').disabled=!run;
    $('pause-run').hidden=!running;$('keep-open').hidden=!running;
    $('retry-reviews').hidden=running||!count('review');
    $('retry-extraction-reviews').hidden=running||!entries.some(e=>e.status==='review' && run.config.nodes[e.node]?.type==='extract');
    $('resume-run').hidden=running||!count('pending');$('retry-run').hidden=running||!count('error');
    $('run-full-file').hidden=running||!run?.test||!source;$('edit-run').hidden=running||!run;
    $('apply-run-settings').hidden=!run;
    $('result-source').textContent=run?`${run.source.filename} · ${run.config.name}`:'';
    document.querySelector('[data-stage=results]').disabled=!run;
    refreshFileActions();
    const filter=$('result-filter').value;
    const visible=[];
    for(let i=0;i<entries.length && visible.length<100;i++) {
        const e=entries[i];
        if(filter==='all'||filter==='ok'&&e.status==='ok'||filter==='review'&&['review','error'].includes(e.status)) visible.push({e,i});
    }
    const expected=run?[...new Set(run.config.nodes.flatMap(outputs))]:[];
    const sourceColumns=run?.source.columns || Object.keys(entries[0]?.values || {}).filter(k=>!k.startsWith('_'));
    const cols=$('all-result-columns').checked?[...new Set([...expected,...sourceColumns])]:expected.filter((column,i)=>i<4 || stoppedColumns.has(column));
    const shown=expected.filter(k=>cols.includes(k)).length;
    $('column-visibility-note').textContent=run?`Showing ${shown} of ${expected.length} configured result columns. “Not run” means an earlier step must finish first.`:'';
    const cell=(entry,column)=>{
        if(entry.status==='review') {
            const marked = reviewCell(run.config.nodes[entry.node],entry,column);
            if(marked !== null) return esc(marked);
        }
        if(entry.values[column]!==undefined && entry.values[column]!==null)return esc(typeof entry.values[column]==='object'?JSON.stringify(entry.values[column]):entry.values[column]);
        const stepIndex=run.config.nodes.findIndex(n=>outputs(n).includes(column));
        if(stepIndex===entry.node && entry.status==='error')return '<span class="text-muted small">Failed</span>';
        return stepIndex>=entry.node && entry.status!=='ok'?'<span class="text-muted small">Not run</span>':'';
    };
    $('results').innerHTML=run?`<p class="small text-muted">Showing ${visible.length} of ${entries.length} current records (preview limit 100).</p><table class="table table-sm align-middle"><thead><tr><th>Status / action</th><th>Feedback ID</th><th>Original ID</th><th>Source</th><th>Message</th>${cols.map(c=>`<th>${esc(friendly(c))}</th>`).join('')}</tr></thead><tbody>${visible.map(({e,i})=>`<tr><td><span class="status-label status-${esc(e.status)}">${esc(statusLabel(e))}</span><div class="review-reason">${esc(e.reason||'')}</div>${failureDetails(e)}${['review','error'].includes(e.status)?`<button class="btn btn-sm btn-outline-primary mt-2" data-review="${i}" ${running?'disabled':''}>Review</button><button class="btn btn-sm btn-link text-danger" data-discard="${i}" ${running?'disabled':''}>Discard</button>`:''}</td><td>${esc(e.values._feedback_id||'Pending')}</td><td>${esc(e.values._original_id)}</td><td>${esc(e.values._source_filename)}<br>${esc(e.values._source_sheet)} · row ${esc(e.values._source_row)}</td><td>${esc(e.values._message_number)}</td>${cols.map(c=>`<td>${cell(e,c)}</td>`).join('')}</tr>`).join('')}</tbody></table>`:'<p class="text-muted">Your results will appear here.</p>';
}
$('result-filter').onchange=$('all-result-columns').onchange=renderRun;
let reviewIndex=null;
$('results').onclick=safe(async event=>{
    if(running)return;
    const discard=event.target.closest('[data-discard]');if(discard){const index=Number(discard.dataset.discard),e=run.work[index];e.status='discarded';e.history.push({status:'discarded',at:new Date().toISOString()});await saveCheckpoint(index);renderRun();return;}
    const btn=event.target.closest('[data-review]');if(!btn)return;
    reviewIndex=Number(btn.dataset.review);const e=run.work[reviewIndex], n=run.config.nodes[e.node];
    $('review-context').textContent=`${e.values._source_filename}, row ${e.values._source_row}, message ${e.values._message_number} — ${n.name}. ${e.reason}`;
    $('review-fields').innerHTML=n.inputs.map((k,i)=>`<label class="form-label mt-2" for="correction-${i}">${esc(k)}</label><textarea id="correction-${i}" data-correction="${esc(k)}" class="form-control" rows="3">${esc(e.values[k])}</textarea>`).join('');
    $('review-dialog').showModal();
});
$('save-correction').onclick=safe(async()=>{
    const edits={};$('review-fields').querySelectorAll('[data-correction]').forEach(el=>{if(el.value!==String(run.work[reviewIndex].values[el.dataset.correction]??''))edits[el.dataset.correction]=el.value;});
    retryEntry(run,reviewIndex,edits);await saveCheckpoint();$('review-dialog').close();await continueRun();
});
initProviderSettings();
$('api-settings-btn').onclick=()=>bootstrap.Modal.getOrCreateInstance($('api-settings-modal')).show();
window.addEventListener('beforeunload',event=>{if(running){event.preventDefault();event.returnValue='';}});
renderNodes();renderRun();setStage('workflow',false);
hasSavedPipeline().then(exists=>{
    $('saved-run-banner').hidden = !exists || Boolean(run);
}).catch(()=>{ /* Starting or restoring a run reports storage errors in context. */ });
