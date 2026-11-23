/** Portable workflow format and pure runner transitions. Credentials never enter a run. */
export const FORMAT = 'aidstack-insights-pipeline';
export const TYPES = ['extract', 'analyze', 'classify', 'split', 'validate', 'lookup'];
export const PROVIDERS = ['openai', 'azure', 'ollama', 'jev', 'laya'];
const NODE_KEYS = ['id', 'name', 'type', 'inputs', 'provider', 'model', 'instructions', 'fields', 'verbatimField', 'verbatimInput', 'maxOutputTokens', 'responseMode', 'minConfidence', 'questionType', 'sourceBranches', 'branches', 'delimiter', 'pairedSourceStep', 'table', 'output', 'rules', 'reviewValues'];
export const REVIEW_SHEET = 'Pipeline review';
export const LOG_SHEET = 'Pipeline imports';
export const AUDIT_SHEET = 'Pipeline audit';
export const META = ['_feedback_id', '_source_dataset_id', '_source_filename', '_source_sheet', '_source_row', '_source_record_id', '_original_id', '_message_number', '_record_key', '_status', '_failed_step', '_reason', '_error_code', '_original_record'];
const clone = v => structuredClone(v);
const fail = message => { throw new Error(message); };
export const outputs = n => n.type === 'split' ? n.inputs : n.type === 'lookup' ? [n.output] : (n.fields || []).map(f => f.name);
export function requiredInputs(nodes) {
    const generated = new Set(), required = new Set();
    for (const n of nodes) { for (const k of n.inputs || []) if (!generated.has(k)) required.add(k); outputs(n).forEach(k => generated.add(k)); }
    return [...required];
}
export function defaultNode(type = 'analyze') {
    const node = {id: crypto.randomUUID(), name: ({extract: 'Extract paired messages', analyze: 'Analyze or translate', classify: 'Classify', split: 'Split aligned lists', validate: 'Validate fields', lookup: 'Map a value'})[type], type, inputs: ['Feedback']};
    if (['extract', 'analyze', 'classify'].includes(type)) Object.assign(node, {provider: 'azure', model: '', instructions: 'Use only the supplied evidence. Flag insufficient context for review.', fields: [{name: 'result', options: []}], maxOutputTokens: 4096, responseMode: 'schema'});
    if (type === 'extract') Object.assign(node, {verbatimInput: 'Feedback', verbatimField: 'feedback_excerpt', fields: [{name: 'feedback_excerpt', options: []}, {name: 'type_of_feedback', options: ['Appréciations', 'Questions', 'Demandes', 'Suggestions', 'Observations, perceptions ou croyances', 'Signalements et préoccupations', 'Allégations et Signalement d’incident']}], instructions: 'Extract every distinct message from the feedback, in source order. Pair each exact verbatim excerpt with its feedback type. Keep repeated types. Preserve the entire speech act and necessary context. Interpret French, Congolese Swahili, Lingala and mixtures within the same sentence. Language mixing alone is not a reason for review. Short or informal messages can still have a clear meaning. Do not invent meaning; if a material phrase cannot be understood, quote that phrase and explain the ambiguity. A request for an item is Demandes; a request for information is Questions. Flag incomplete or uninterpretable feedback for review.'});
    if (type === 'classify') Object.assign(node, {inputs: ['feedback_excerpt'], fields: [{name: 'category', options: ['Question', 'Request', 'Appreciation']}], minConfidence: .6});
    if (type === 'split') Object.assign(node, {inputs: ['feedback_excerpt', 'type_of_feedback'], delimiter: ' || '});
    if (type === 'validate') node.rules = [];
    if (type === 'lookup') Object.assign(node, {output: 'mapped_value', table: [{when: {Feedback: 'example'}, value: 'mapped'}]});
    return node;
}
export function defaultWorkflow() { return {format: FORMAT, version: 1, name: 'Community feedback intake', idColumn: 'ID', mappings: {}, nodes: [defaultNode('extract')]}; }
export function portableConfig(raw) {
    if (!raw || raw.format !== FORMAT || raw.version !== 1 || !Array.isArray(raw.nodes) || !raw.nodes.length || raw.nodes.length > 100) fail('This is not a supported pipeline configuration (version 1, 1–100 steps).');
    const ids = new Set();
    const nodes = raw.nodes.map(n => {
        if (!n || !TYPES.includes(n.type) || typeof n.id !== 'string' || !n.id || ids.has(n.id)) fail('Step IDs must be unique and step types supported.');
        ids.add(n.id);
        if (!Array.isArray(n.inputs) || !n.inputs.length || n.inputs.some(k => typeof k !== 'string' || !k.trim() || k.startsWith('_')) || new Set(n.inputs).size !== n.inputs.length) fail('Every step needs distinct input columns without a leading _.');
        if (outputs(n).some(k => typeof k !== 'string' || !k.trim() || k.startsWith('_') || ['__proto__', 'constructor', 'prototype'].includes(k))) fail('Output names cannot be empty or start with _.');
        if (['extract', 'analyze', 'classify'].includes(n.type) && !PROVIDERS.includes(n.provider)) fail('Choose a supported AI provider.');
        return clone(Object.fromEntries(NODE_KEYS.filter(k => n[k] !== undefined).map(k => [k, n[k]])));
    });
    nodes.forEach((n,i) => {
        if (!n.pairedSourceStep) return;
        const before = nodes[i-1];
        if (n.type !== 'split' || before?.type !== 'extract' || before.id !== n.pairedSourceStep || JSON.stringify(n.inputs) !== JSON.stringify(outputs(before)))
            fail('Keep the paired split directly after its extractor with the same output columns.');
    });
    return {format: FORMAT, version: 1, name: String(raw.name || 'Untitled workflow'), idColumn: String(raw.idColumn || ''), mappings: Object.fromEntries(Object.entries(raw.mappings || {}).filter(([k,v]) => typeof v === 'string' && !k.startsWith('_'))), nodes};
}
export function importWorkflow(raw) {
    if (raw?.format === FORMAT) return {config: portableConfig(raw), notice: 'Pipeline configuration loaded.'};
    if (raw?.format === 'aidstack-insights-analysis-config') {
        if (!Array.isArray(raw.columns) || !raw.columns.length) fail('No analyses in this configuration.');
        const nodes = raw.columns.map(c => ({...defaultNode('analyze'), name: c.outputColumnName, inputs: c.columns || [c.column], instructions: `${raw.generalInstructions || ''}\n\n${c.prompt || ''}`, fields: [{name: c.outputColumnName, options: []}], maxOutputTokens: Math.min(32768, Math.max(64, c.maxOutputTokens || 4096))}));
        return {config: portableConfig({...defaultWorkflow(), nodes}), notice: 'Imported analyses as ordered steps. For paired extraction, use the paired-message step instead of separate excerpt/type analyses.'};
    }
    if (raw?.format === 'aidstack-insights-jev-config') {
        if (!Array.isArray(raw.questions) || !raw.questions.length) fail('No questions in this codebook.');
        const nodes = raw.questions.map(q => {
            if (q.questionType !== 'choice' || q.sourceExcerpt) fail('This importer supports Choice codebooks with an existing excerpt input. Score/Noul and sourceExcerpt are supported in the original classification module.');
            const n = {...defaultNode('classify'), name: q.outputColumnName, provider: raw.classifier?.provider || 'jev', model: raw.classifier?.model || raw.model || '', instructions: typeof q.instructions === 'string' ? q.instructions : JSON.stringify(q.instructions), inputs: [...new Set([...(raw.sourceColumns || []), ...(q.dependsOn || [])])], fields: [{name: q.outputColumnName, options: q.options || []}], minConfidence: q.review?.minConfidence ?? .6};
            for (const key of ['sourceBranches', 'branches']) if (q[key]) n[key] = clone(q[key]);
            return n;
        });
        for (const d of raw.derived || []) {
            if ((d.type || 'lookup') !== 'lookup') fail('Composite scores are not supported by the pipeline importer.');
            nodes.push({...defaultNode('lookup'), name: d.outputColumnName, inputs: d.inputs, output: d.outputColumnName, table: clone(d.table)});
        }
        return {config: portableConfig({...defaultWorkflow(), nodes}), notice: 'Codebook imported with branches and lookups. Uncertain records stop for review; JEV confidence thresholds apply only to JEV/Laya. Select Azure/OpenAI on a classification step to change its provider.'};
    }
    fail('Choose a pipeline, AI Analysis, or Choice classification configuration.');
}
export function validateMappings(config, columns) {
    for (const key of requiredInputs(config.nodes)) {
        if (!columns.includes(config.mappings[key] || key)) fail(`Map the input “${key}” to a column in this file.`);
    }
    const generated = new Set();
    for (const n of config.nodes) {
        for (const name of outputs(n)) {
            if (n.type !== 'split' && (columns.includes(name) || requiredInputs(config.nodes).includes(name) || generated.has(name))) fail(`Output “${name}” would overwrite an existing column. Rename the output or remove that source column.`);
            generated.add(name);
        }
    }
    for (const key of requiredInputs(config.nodes)) {
        const mapped = config.mappings[key] || key;
        if (mapped !== key && columns.includes(key)) fail(`Mapping “${key}” would overwrite an existing source column. Rename that column first.`);
    }
    if (config.idColumn && !columns.includes(config.idColumn)) fail('Select the original ID column or choose no existing ID.');
}
export async function digest(bytes) { return [...new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))].map(b => b.toString(16).padStart(2, '0')).join(''); }
export function seedRun(config, source, master = {rows: [], review: [], logs: []}) {
    config = portableConfig(config);
    validateMappings(config, source.columns);
    if (source.columns.some(c => c.startsWith('_'))) fail('Source columns starting with _ are reserved for pipeline provenance. Rename them before importing.');
    if (master.logs.some(l => l.source_dataset_id === source.datasetId && (l.complete === false || l.complete === 'FALSE'))) fail('This dataset has an unfinished import in the master. Restore its recovery file to finish before importing it again.');
    const existing = [...master.rows, ...master.review];
    const known = new Set(existing.map(r => r._source_record_id).filter(Boolean));
    const ids = existing.map(r => Number(r._feedback_id)).filter(Number.isSafeInteger);
    let nextId = ids.reduce((max, id) => Math.max(max, id), 0) + 1;
    if (!Number.isSafeInteger(nextId)) fail('The master has exhausted safe numeric IDs.');
    let skipped = 0;
    const work = [];
    const inputs = requiredInputs(config.nodes);
    for (const row of source.rows) {
        const sourceId = `${source.datasetId}:${source.sheet}:${row.excelRow}`;
        if (known.has(sourceId)) { skipped++; continue; }
        const values = clone(row.values);
        for (const k of inputs) values[k] = row.values[config.mappings[k] || k];
        Object.assign(values, {_source_dataset_id: source.datasetId, _source_filename: source.filename, _source_sheet: source.sheet, _source_row: row.excelRow, _source_record_id: sourceId, _original_id: config.idColumn ? row.values[config.idColumn] : '', _message_number: '1', _record_key: sourceId, _original_record: JSON.stringify(row.values)});
        work.push({values, node: 0, status: 'pending', history: []});
    }
    return {format: 'aidstack-pipeline-run', version: 1, runId: crypto.randomUUID(), createdAt: new Date().toISOString(), config, source: {filename: source.filename, datasetId: source.datasetId, sheet: source.sheet, count: source.rows.length}, master, nextId, work, skipped};
}
function assignId(run, entry) {
    if (!entry.values._feedback_id) { if (!Number.isSafeInteger(run.nextId)) fail('Numeric ID limit reached.'); entry.values._feedback_id = run.nextId++; }
}
export function nextWorkIndex(run, reviewsOnly, cursor) {
    const split = run.config.nodes.findIndex(n => n.pairedSourceStep);
    const lastPhase = split + 1;
    while (cursor.phase <= lastPhase) {
        for (; cursor.index < run.work.length; cursor.index++) {
            const entry = run.work[cursor.index];
            if (entry.status === 'pending' && (!reviewsOnly || entry.reviewRetry) &&
                Math.min(entry.node,lastPhase) === cursor.phase) return cursor.index;
        }
        cursor.phase++;
        cursor.index = 0;
    }
    return -1;
}
// A translation review stays one source record. Its tentative message codes can be retried together.
export function activeStep(run, entry) {
    const queue = entry.tentativeCoding;
    if (!queue) {
        const node=run.config.nodes[entry.node];
        if (entry.translationFlag && node?.type === 'extract') return {node:{...node,instructions:node.instructions+' The full translation is tentative. Use the original to supply any missing English for each excerpt.'}, values:entry.values};
        return {node,values:entry.values};
    }
    const step = run.config.nodes[entry.node + (queue.messages ? 2 : 1)];
    const node = {...step, instructions: step.instructions + '\nThe English translation is tentative. Use the original wording as evidence. If the supplied English is missing or incomplete, give the best supported English for each excerpt. Assign a category when the meaning supports it; mark genuinely unresolved decisions for review.'};
    return {node, values: {...entry.values, source_translation_en: entry.drafts.source_translation_en || '', ...queue.messages?.[queue.index]}};
}

function collectTentativeCodes(run, entry, result) {
    const queue = entry.tentativeCoding;
    delete entry.rateLimitAttempts;
    delete entry.retryAt;
    if (!queue.messages && result.status === 'ok') {
        if (!Array.isArray(result.items) || !result.items.length) fail('Missing extracted messages.');
        queue.messages = clone(result.items);
        queue.index = 0;
        return;
    }
    if (queue.messages) {
        const message = queue.messages[queue.index];
        const path = result.items?.[0]?.feedback_path || result.drafts?.feedback_path;
        if (path) {
            message.feedback_path = path;
            const lookup = run.config.nodes.find(n => n.type === 'lookup' && n.output === 'feedback_category' && n.inputs.length === 1 && n.inputs[0] === 'feedback_path');
            const category = lookup?.table.find(row => row.when.feedback_path === path);
            if (category) message.feedback_category = category.value;
        }
        if (result.status !== 'ok') message.feedback_path ||= result.reason || 'Category unresolved';
        queue.index++;
        if (queue.index < queue.messages.length && result.status !== 'error') return;
        const keys = new Set(queue.messages.flatMap(Object.keys));
        for (const key of keys) {
            entry.drafts[key] = queue.messages.length === 1 ? queue.messages[0][key]
                : queue.messages.map((m,i) => `${i+1}. ${m[key] || 'Unresolved'}`).join('\n');
        }
        entry.drafts.source_translation_en ||= entry.drafts.translation_en;
    } else {
        Object.assign(entry.drafts, result.drafts || {});
    }
    if (result.status !== 'ok') entry.reason += ' · ' + (result.reason || 'Tentative coding incomplete.');
    entry.status = 'review';
    delete entry.tentativeCoding;
    delete entry.reviewRetry;
    assignId(run, entry);
}

export function applyResult(run, index, result) {
    const entry = run.work[index], {node} = activeStep(run, entry);
    if (!['ok', 'review', 'error'].includes(result?.status)) fail('Invalid step response; the record remains pending.');
    const audit = {step: node.id, name: node.name, provider: node.provider || 'local', model: result.model || '', at: new Date().toISOString(), status: result.status, reason: result.reason || '', errorCode: result.errorCode || '', details: result.audit || null};
    if (entry.tentativeCoding) {
        entry.history.push({...audit, tentative:true});
        collectTentativeCodes(run, entry, result);
        return;
    }
    const upcomingSplit = run.config.nodes[entry.node+2]?.pairedSourceStep;
    if (result.status === 'review' && result.tentativeTranslation && upcomingSplit && upcomingSplit === run.config.nodes[entry.node+1]?.id) {
        entry.translationFlag = {node:entry.node, reason:result.reason || 'Translation needs review.'};
        entry.values.source_translation_en = result.drafts?.source_translation_en || '';
        entry.history.push(audit);
        entry.node++;
        return;
    }
    if (result.status === 'review' && result.tentativeTranslation && node.id === 'translate-feedback' &&
        run.config.nodes[entry.node+1]?.id === 'extract-messages' && run.config.nodes[entry.node+2]?.fields?.[0]?.name === 'feedback_path') {
        entry.history.push(audit);
        entry.reason = result.reason || 'Translation needs review.';
        entry.drafts = {...result.drafts};
        entry.tentativeCoding = {};
        return;
    }
    if (result.status !== 'ok') {
        delete entry.reviewRetry;
        if (result.status === 'review') entry.drafts = result.drafts || {};
        if (Array.isArray(result.items) && result.items.length === 1) {
            const allowed = new Set(outputs(node));
            for (const [k,v] of Object.entries(result.items[0])) if (allowed.has(k)) entry.values[k] = v;
        }
        if (result.status === 'review' && outputs(node).includes('feedback_path')) {
            const path = entry.drafts.feedback_path || entry.values.feedback_path;
            if (path && !(node.reviewValues || []).includes(path)) {
                for (const lookup of run.config.nodes.slice(entry.node+1)) {
                    if (lookup.type !== 'lookup' || lookup.inputs.length !== 1 || lookup.inputs[0] !== 'feedback_path') continue;
                    const match = lookup.table.find(row => row.when.feedback_path === path);
                    if (match) entry.drafts[lookup.output] = match.value;
                }
            }
        }
        entry.errorCode = result.errorCode || '';
        entry.status = result.status; entry.reason = result.reason || 'This record needs review.'; entry.history.push(audit); assignId(run, entry); return;
    }
    if (!Array.isArray(result.items) || !result.items.length || result.items.length > 100 || result.items.some(r => !r || typeof r !== 'object' || Array.isArray(r))) fail('Incomplete step response; the record remains pending.');
    const allowed = new Set(outputs(node));
    for (const item of result.items) if (Object.keys(item).some(k => !allowed.has(k))) fail('The step returned unexpected output columns.');
    const split = run.config.nodes[entry.node+1];
    if (split?.pairedSourceStep === node.id) {
        const joined = Object.fromEntries(split.inputs.map(k=>[k,result.items.map(item=>item[k]).join(split.delimiter)]));
        if (entry.translationFlag && !entry.values.source_translation_en) entry.values.source_translation_en = result.items.map(item=>item.translation_en).join(' ');
        Object.assign(entry.values,joined);
        entry.pairedItems = result.items;
        run.classifications ||= {};
        const keys = ['_source_record_id','_original_id','_source_filename','_source_sheet','_source_row',run.config.idColumn,...node.inputs].filter(Boolean);
        run.classifications[entry.values._source_record_id] = {...Object.fromEntries(keys.map(k=>[k,entry.values[k]])),...joined,
            _status:entry.translationFlag?'translation needs review':'classified'};
        entry.history.push(audit);
        entry.node++;
        return;
    }
    const expanded = node.type === 'extract' || node.type === 'split';
    const previousMessages = entry.previousMessages?.slice();
    const children = result.items.map((item, i) => {
        const values = {...entry.values, ...item};
        if (expanded) {
            values._record_key = `${entry.values._record_key}/${node.id}/${i+1}`;
            values._message_number = entry.values._record_key === entry.values._source_record_id ? String(i+1) : `${entry.values._message_number}.${i+1}`;
            if (i) delete values._feedback_id;
        }
        const child = {values, node: entry.node + 1, status: entry.node + 1 === run.config.nodes.length ? 'ok' : 'pending', history: [...entry.history, audit]};
        if (entry.translationFlag) {
            child.translationFlag = entry.translationFlag;
            if (child.status === 'ok') {
                child.status = 'review';
                child.node = entry.translationFlag.node;
                child.reason = entry.translationFlag.reason;
            }
        }
        if (expanded && previousMessages) {
            const match=previousMessages.findIndex(m=>m.excerpt===values.feedback_excerpt && m.type===values.type_of_feedback);
            const previous=match>=0 ? previousMessages.splice(match,1)[0] : null;
            if(previous?.id) values._feedback_id=previous.id;
            if(previous?.discarded) child.status='discarded';
        } else if (entry.previousMessages) child.previousMessages=entry.previousMessages;
        if (entry.reviewRetry && child.status === 'pending') child.reviewRetry = true;
        if (child.status === 'ok' || child.status === 'review') assignId(run, child);
        return child;
    });
    run.work.splice(index, 1, ...children);
}
export function scheduleRateRetry(run, index, result, now = Date.now(), jitter = Math.random()) {
    const entry = run.work[index], {node} = activeStep(run, entry);
    if (result.errorCode !== 'rate_limit') return null;
    entry.rateLimitAttempts = (entry.rateLimitAttempts || 0) + 1;
    const supplied = result.retryAfterSeconds;
    const seconds = typeof supplied === 'number' && Number.isFinite(supplied) && supplied >= 0
        ? supplied : 15 * 2 ** (entry.rateLimitAttempts - 1);
    entry.retryAt = Math.min(8640000000000000, now + (seconds + Math.min(1, Math.max(0, jitter))) * 1000);
    run.providerCooldowns ||= {};
    run.providerCooldowns[node.provider] = Math.max(run.providerCooldowns[node.provider] || 0, entry.retryAt);
    entry.history.push({step:node.id,name:node.name,provider:node.provider,status:'rate_limited',
        at:new Date(now).toISOString(),attempt:entry.rateLimitAttempts,retryAt:entry.retryAt});
    if (entry.rateLimitAttempts >= 3) return 'exhausted';
    return seconds > 60 ? 'pause' : 'retry';
}
export function retryEntry(run, index, edits = {}) {
    let entry = run.work[index];
    if (!['review', 'error'].includes(entry.status)) fail('Only failed or review records can be retried.');
    const allowed = run.config.nodes[entry.node].inputs;
    if (Object.keys(edits).some(k => !allowed.includes(k))) fail('Only this step’s input fields can be corrected.');
    if (entry.translationFlag?.node === entry.node && entry.values._record_key !== entry.values._source_record_id) {
        const sourceId=entry.values._source_record_id;
        const siblings=run.work.filter(e=>e.values._source_record_id === sourceId);
        const previousMessages=siblings.map(e=>({excerpt:e.values.feedback_excerpt,type:e.values.type_of_feedback,id:e.values._feedback_id,discarded:e.status==='discarded'}));
        const values={...entry.values,_record_key:sourceId,_message_number:'1'};
        delete values._feedback_id;
        for(const n of run.config.nodes) for(const key of outputs(n)) delete values[key];
        const splitId=run.config.nodes.find(n=>n.pairedSourceStep)?.id;
        const splitIndex=siblings[0].history.findLastIndex(h=>h.step===splitId);
        const history=siblings[0].history.slice(0,splitIndex+1);
        history.push({at:new Date().toISOString(),status:'source_retry',details:{priorMessages:siblings.map(e=>({key:e.values._record_key,id:e.values._feedback_id,status:e.status,history:e.history.slice(splitIndex+1)}))}});
        entry={node:entry.node,status:'review',values,history,previousMessages};
        const first=run.work.indexOf(siblings[0]);
        if(run.classifications) delete run.classifications[sourceId];
        run.work=run.work.filter(e=>e.values._source_record_id!==sourceId);
        run.work.splice(first,0,entry);
    }
    entry.history.push({step: run.config.nodes[entry.node].id, at: new Date().toISOString(), status: 'corrected', before: Object.fromEntries(Object.keys(edits).map(k => [k, entry.values[k]])), after: clone(edits)});
    Object.assign(entry.values, edits); entry.status = 'pending'; delete entry.reason; delete entry.errorCode; delete entry.rateLimitAttempts;
    delete entry.drafts;
    const node = run.config.nodes[entry.node];
    for (const key of outputs(node)) if (!node.inputs.includes(key)) delete entry.values[key];
    return entry;
}
const noMatchingPath = 'À vérifier : aucun code ne correspond au type fourni';
function isNoMatchingPath(value) {
    const text = typeof value === 'string' ? value.trim() : '';
    return text === noMatchingPath || text === '[Needs Review] | ' + noMatchingPath;
}
export function reviewCell(node, entry, column) {
    const tentative = entry.translationFlag && ['source_translation_en','translation_en','type_of_feedback','feedback_excerpt','feedback_path','feedback_category','feedback_code','feedback_sous_code'].includes(column) && Object.hasOwn(entry.values,column);
    if (entry.status !== 'review' || !(tentative || outputs(node).includes(column) || Object.hasOwn(entry.drafts || {}, column))) return null;
    const usable = value => value != null && String(value).trim() !== '' && !/^(null|\[?needs review\]?)$/i.test(String(value).trim());
    let text = [entry.drafts?.[column], entry.values[column]].find(usable);
    if (column === 'feedback_path' && (text === undefined || isNoMatchingPath(text))) return '';
    if (text === undefined) {
        const values = entry.values;
        const source = [values[node.verbatimInput], values.feedback_excerpt, ...node.inputs.map(k => values[k])].find(usable);
        const translation = [values.translation_en, values.source_translation_en].find(usable);
        text = translation !== undefined ? 'Translation: ' + translation
            : source !== undefined ? 'Source: ' + source : entry.reason || 'No text available';
    }
    return '[Needs Review] | ' + text;
}
export function exportedRow(run, entry) {
    const node = run.config.nodes[entry.node];
    const row = {...entry.values, _status: entry.status, _failed_step: node?.name || '', _reason: entry.reason || '', _error_code: entry.errorCode || ''};
    if (entry.status === 'review') for (const key of new Set([...outputs(node), ...Object.keys(row), ...Object.keys(entry.drafts || {})])) {
        const marked=reviewCell(node,entry,key);
        if(marked==='') delete row[key];
        else if(marked!==null) row[key]=marked;
    }
    return row;
}
export function mergeRun(run) {
    const rows = clone(run.master.rows), review = clone(run.master.review);
    for (const row of review) if (isNoMatchingPath(row.feedback_path)) delete row.feedback_path;
    const seen = new Set([...rows, ...review].map(r => r._record_key).filter(Boolean));
    for (const entry of run.work) {
        if (entry.status === 'pending' || seen.has(entry.values._record_key)) continue;
        (entry.status === 'ok' ? rows : review).push(exportedRow(run, entry)); seen.add(entry.values._record_key);
    }
    const log = {run_id: run.runId, source_dataset_id: run.source.datasetId, source_file: run.source.filename, source_sheet: run.source.sheet, imported_at: run.createdAt, source_rows: run.source.count, complete: !run.work.some(e => e.status === 'pending'), accepted: run.work.filter(e => e.status === 'ok').length, review: run.work.filter(e => ['review', 'error', 'discarded'].includes(e.status)).length, skipped_sources: run.skipped, workflow: run.config.name};
    return {rows, review, logs: [...run.master.logs.filter(l => l.run_id !== run.runId), log]};
}
