import {deserializeJevConfig, validateReportConfig, orderedJevOutputs} from './jev-config.js';

export const RECOVERY_FORMAT = 'aidstack-insights-recovery';
const pick = (value, keys) => Object.fromEntries(keys.filter(key => value?.[key] !== undefined).map(key => [key, value[key]]));
const common = ['version', 'runId', 'filename', 'sheetName', 'sourceColumns', 'columns', 'outputColumns',
    'rows', 'sourceRows', 'totalRows', 'isTestRun', 'cursor', 'results'];

/** An explicit allowlist keeps credentials, callbacks, and transient UI flags out of storage. */
export function snapshotRun(kind, run) {
    const saved = pick(run, common);
    if (kind === 'analysis') {
        Object.assign(saved, pick(run, ['batchSize', 'errors', 'columnConfigs', 'generalInstructions']));
        saved.providerConfig = pick(run.providerConfig, ['provider', 'model', 'azureEndpoint', 'azureDeployment',
            'azureApiVersion', 'ollamaEndpoint', 'ollamaModel']);
    } else {
        Object.assign(saved, pick(run, ['config', 'audit']));
        saved.classifier = pick(run.classifier || {classificationProvider: 'jev'},
            ['classificationProvider', 'jevModel', 'layaEndpoint', 'layaModel']);
    }
    return saved;
}

export function recoveryDocument(kind, run) {
    return {format: RECOVERY_FORMAT, version: 1, kind, exportedAt: new Date().toISOString(), run: snapshotRun(kind, run)};
}

const record = value => value && typeof value === 'object' && !Array.isArray(value);
const strings = value => Array.isArray(value) && value.length > 0 && value.every(s => typeof s === 'string' && s.length > 0)
    && new Set(value).size === value.length;
const rows = value => Array.isArray(value) && value.every(record);

/** Validate before replacing any saved work. Recovery never starts provider calls. */
export function parseRecovery(raw) {
    const fail = () => { throw new Error('Invalid or incomplete recovery file. Choose an Aidstack recovery JSON download.'); };
    if (raw?.format !== RECOVERY_FORMAT || raw.version !== 1 || !['analysis', 'jev'].includes(raw.kind)) fail();
    const r = raw.run;
    if (!record(r) || r.version !== 1 || typeof r.runId !== 'string' || !r.runId
        || typeof r.filename !== 'string' || typeof r.sheetName !== 'string' || !r.sheetName
        || !strings(r.sourceColumns) || !strings(r.columns) || !strings(r.outputColumns)
        || !rows(r.rows) || !r.rows.length || !rows(r.results)
        || !Number.isInteger(r.cursor) || r.cursor < 0 || r.cursor > r.rows.length || r.results.length !== r.cursor
        || !Number.isInteger(r.totalRows) || r.totalRows < r.rows.length
        || (r.sourceRows !== undefined && (!rows(r.sourceRows) || r.sourceRows.length !== r.totalRows))) fail();
    if (JSON.stringify(r.columns) !== JSON.stringify([...r.sourceColumns, ...r.outputColumns])) fail();
    if (r.rows.some(row => r.sourceColumns.some(col => !Object.hasOwn(row, col)))) fail();
    if (r.results.some((row, i) => r.sourceColumns.some(col =>
        JSON.stringify(row[col]) !== JSON.stringify(r.rows[i][col])))) fail();
    if (raw.kind === 'analysis') {
        if (!Number.isInteger(r.batchSize) || r.batchSize < 1 || r.batchSize > 100
            || !Number.isInteger(r.errors) || r.errors < 0 || typeof r.generalInstructions !== 'string'
            || !Array.isArray(r.columnConfigs) || r.columnConfigs.length !== r.outputColumns.length
            || !['openai', 'azure', 'ollama'].includes(r.providerConfig?.provider)) fail();
        for (const [i, c] of r.columnConfigs.entries()) {
            if (!record(c) || typeof c.id !== 'string' || !r.sourceColumns.includes(c.column)
                || !strings(c.columns) || c.columns.some(col => !r.sourceColumns.includes(col))
                || typeof c.prompt !== 'string' || !c.prompt.trim() || c.resolvedName !== r.outputColumns[i]
                || ![256, 1024, 4096].includes(c.maxOutputTokens)) fail();
        }
        if (new Set(r.columnConfigs.map(c => c.id)).size !== r.columnConfigs.length) fail();
    } else {
        if (!Array.isArray(r.audit) || r.audit.length !== r.cursor
            || r.audit.some((a, i) => !record(a) || a.rowIndex !== i || !record(a.decisions))
            || !['jev', 'laya'].includes(r.classifier?.classificationProvider)) fail();
        const config = deserializeJevConfig(r.config, r.sourceColumns);
        if (config.warnings.length || !config.sourceColumns.length) fail();
        validateReportConfig(r.config, r.sourceColumns);
        const outputs = orderedJevOutputs(r.config).flatMap(q => [q.outputColumnName,
            ...(r.config.includeConfidence && q.questionType ? [q.outputColumnName + '__confidence',
                ...(['score', 'noul'].includes(q.questionType) ? [q.outputColumnName + '__score'] : [])] : [])]);
        if (JSON.stringify(outputs) !== JSON.stringify(r.outputColumns)) fail();
    }
    return {kind: raw.kind, run: snapshotRun(raw.kind, r)};
}
