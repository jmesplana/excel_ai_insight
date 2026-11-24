import {REVIEW_SHEET, LOG_SHEET, AUDIT_SHEET, META, mergeRun, outputs} from './pipeline-core.js?v=4';
/** Read physical Excel row numbers before skipping empty rows. */
export function readSheet(workbook, name, headerRow = 1, xlsx = XLSX) {
    const ws = workbook.Sheets[name];
    if (!ws?.['!ref']) throw new Error('Select a populated worksheet.');
    const range = xlsx.utils.decode_range(ws['!ref']);
    if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > range.e.r + 1) throw new Error('Choose a valid header row.');
    const header = headerRow - 1;
    const matrix = xlsx.utils.sheet_to_json(ws, {header: 1, defval: null, blankrows: true, range: {s: {r: header, c: 0}, e: range.e}});
    const columns = Array.from({length: range.e.c + 1}, (_, i) => String(matrix[0]?.[i] ?? '').trim());
    if (columns.some(c => !c) || new Set(columns).size !== columns.length) throw new Error('Column headers must be nonempty and unique.');
    if (columns.some(c => ['__proto__', 'constructor', 'prototype'].includes(c))) throw new Error('Rename reserved column headers.');
    const rows = [];
    for (let i = 1; i < matrix.length; i++) {
        if (!matrix[i].some(v => v !== null && v !== '')) continue;
        rows.push({excelRow: headerRow + i, values: Object.fromEntries(columns.map((k,j) => [k, matrix[i][j] ?? null]))});
    }
    return {columns, rows};
}
export function readMaster(workbook, sheet, xlsx = XLSX) {
    if ([REVIEW_SHEET, LOG_SHEET, AUDIT_SHEET, 'Pipeline classifications', 'Pipeline workflow', 'Pipeline overflow'].includes(sheet)) throw new Error('Choose the main master data worksheet.');
    const parsed = readSheet(workbook, sheet, 1, xlsx);
    const managed = (name, key) => {
        if (!workbook.Sheets[name]) return [];
        const p = readSheet(workbook, name, 1, xlsx);
        if (!p.columns.includes(key)) throw new Error(`“${name}” is reserved for pipeline output. Rename your existing sheet before using it as a master.`);
        return p.rows.map(r => r.values);
    };
    const review = managed(REVIEW_SHEET, '_record_key');
    const logs = managed(LOG_SHEET, 'run_id');
    managed(AUDIT_SHEET, 'record_key');
    managed('Pipeline workflow', 'run_id');
    managed('Pipeline classifications', '_source_record_id');
    managed('Pipeline overflow', 'run_id');
    const rows = parsed.rows.map(r => r.values);
    const ids = new Set(), keys = new Set();
    for (const r of [...rows, ...review]) {
        if (r._feedback_id !== null && r._feedback_id !== undefined && r._feedback_id !== '') {
            const id = Number(r._feedback_id);
            if (!Number.isSafeInteger(id) || id < 1 || ids.has(id)) throw new Error('Master pipeline IDs must be unique positive safe integers.');
            ids.add(id);
        }
        if (r._record_key) { if (keys.has(r._record_key)) throw new Error('The master contains duplicate pipeline record keys.'); keys.add(r._record_key); }
    }
    return {workbook, sheet, rows, review, logs, columns: parsed.columns};
}
export function buildWorkbook(run, xlsx = XLSX) {
    const merged = mergeRun(run);
    const wb = run.master.workbook ? structuredClone(run.master.workbook) : xlsx.utils.book_new();
    const sheetName = run.master.sheet || 'Master data';
    const priorOverflow = run.master.workbook?.Sheets['Pipeline overflow'] ? xlsx.utils.sheet_to_json(run.master.workbook.Sheets['Pipeline overflow']) : [];
    const overflow = priorOverflow.filter(r=>r.run_id!==run.runId);
    const fitCells = (name, rows) => rows.map((row,i) => Object.fromEntries(Object.entries(row).map(([key,value]) => {
        if (typeof value !== 'string' || value.length <= 32767) return [key,value];
        const record = row._record_key || row.record_key || `row:${i+1}`;
        const entryId = `${name}:${i+1}:${key}`;
        for(let p=0;p<value.length;p+=30000) overflow.push({run_id:run.runId,entry_id:entryId,record_key:record,sheet:name,column:key,part:1+p/30000,value:value.slice(p,p+30000)});
        return [key,`Full value retained in Pipeline overflow (run ${run.runId}, entry ${entryId}).`];
    })));
    const configuredColumns = run.config.nodes.flatMap(outputs);
    const sourceColumns = [...new Set(run.work.flatMap(e=>Object.keys(e.values)).filter(k=>!k.startsWith('_')))];
    const allColumns = [...new Set([...(run.master.columns || []), ...sourceColumns, ...configuredColumns, ...merged.rows.flatMap(Object.keys), ...META])];
    if (wb.Sheets[sheetName]) {
        const ws = wb.Sheets[sheetName];
        const nextRow = xlsx.utils.decode_range(ws['!ref']).e.r + 1;
        xlsx.utils.sheet_add_aoa(ws, [allColumns], {origin: 'A1'});
        const added = merged.rows.slice(run.master.rows.length);
        if (added.length) xlsx.utils.sheet_add_json(ws, fitCells(sheetName, added), {header: allColumns, skipHeader: true, origin: {r: nextRow, c: 0}});
    } else xlsx.utils.book_append_sheet(wb, xlsx.utils.json_to_sheet(fitCells(sheetName, merged.rows), {header: allColumns}), sheetName);
    const setSheet = (name, rows, header) => {
        const ws = xlsx.utils.json_to_sheet(fitCells(name, rows), {header});
        if (!wb.SheetNames.includes(name)) wb.SheetNames.push(name);
        wb.Sheets[name] = ws;
    };
    setSheet(REVIEW_SHEET, merged.review, [...new Set([...META, ...sourceColumns, ...configuredColumns, ...merged.review.flatMap(Object.keys)])]);
    if (run.classifications) {
        const name = 'Pipeline classifications';
        const prior = wb.Sheets[name] ? xlsx.utils.sheet_to_json(wb.Sheets[name]) : [];
        const records = [...prior.filter(r=>r.run_id!==run.runId), ...Object.values(run.classifications).map(row=>({...row,run_id:run.runId}))];
        const roles=run.config.roles || {};
        setSheet(name,records,['run_id','_source_record_id','_original_id',...[roles.messageType,roles.excerpt,roles.messageTranslation].filter(Boolean)]);
    }
    setSheet(LOG_SHEET, merged.logs, ['run_id', 'source_dataset_id', 'source_file', 'source_sheet', 'imported_at', 'source_rows', 'complete', 'accepted', 'review', 'skipped', 'skipped_sources', 'workflow']);
    const oldAudit = run.master.workbook?.Sheets[AUDIT_SHEET] ? xlsx.utils.sheet_to_json(run.master.workbook.Sheets[AUDIT_SHEET]) : [];
    const audit = run.work.flatMap(e => e.history.map(h => ({run_id: run.runId, record_key: e.values._record_key, feedback_id: e.values._feedback_id || '', ...h, details: JSON.stringify(h.details || {}), before: JSON.stringify(h.before || {}), after: JSON.stringify(h.after || {})})));
    setSheet(AUDIT_SHEET, [...oldAudit.filter(a => a.run_id !== run.runId), ...audit], ['run_id', 'record_key', 'feedback_id', 'step', 'name', 'provider', 'model', 'at', 'status', 'reason', 'details', 'before', 'after']);
    const oldConfigs = run.master.workbook?.Sheets['Pipeline workflow'] ? xlsx.utils.sheet_to_json(run.master.workbook.Sheets['Pipeline workflow']) : [];
    const configText = JSON.stringify({config:run.config,revisions:run.revisions || []});
    const configRows = Array.from({length:Math.ceil(configText.length/30000)},(_,i)=>({run_id:run.runId,part:i+1,configuration:configText.slice(i*30000,(i+1)*30000)}));
    setSheet('Pipeline workflow',[...oldConfigs.filter(r=>r.run_id!==run.runId),...configRows],['run_id','part','configuration']);
    if(overflow.length) setSheet('Pipeline overflow',overflow,['run_id','entry_id','record_key','sheet','column','part','value']);
    return wb;
}
