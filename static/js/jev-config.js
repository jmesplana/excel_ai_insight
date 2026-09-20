// Export/import of the Jev Classification setup. A controlled vocabulary is
// laborious to type (the option lists can run to dozens of entries), so the
// whole question set is portable across datasets.
// Everything here is pure: the DOM lives in app.js, the file format lives here.

export const JEV_CONFIG_FORMAT = 'aidstack-insights-jev-config';
export const JEV_CONFIG_VERSION = 2;

export const QUESTION_TYPES = ['choice', 'score', 'noul'];
export const DEFAULT_QUESTION_TYPE = 'choice';

// Mirrors the API limits enforced in jev_provider.py, so a configuration is
// rejected in the browser with a precise message instead of at the API.
export const MAX_CHOICE_OPTIONS = 255;
export const MIN_SCORE_LEVELS = 2;
export const MAX_SCORE_LEVELS = 10;

const str = value => (typeof value === 'string' ? value : '');

/** Split a textarea of options into a clean list, one per line. */
export function parseOptions(text) {
    if (str(text).trim().startsWith('[')) return JSON.parse(text);
    return str(text).split('\n').map(line => line.trim()).filter(Boolean);
}

/** Join an option list back into textarea content. */
export function formatOptions(options) {
    return (options || []).some(o => typeof o === 'object') ? JSON.stringify(options, null, 2) : (options || []).join('\n');
}

function normalizeType(value) {
    const type = str(value).toLowerCase();
    return QUESTION_TYPES.includes(type) ? type : DEFAULT_QUESTION_TYPE;
}

/**
 * Validate one question the way the API will, returning a human message or null.
 *
 * Checked in the browser because a bad option list would otherwise only surface
 * after the run starts, once a key has already been spent on the probe call.
 */
export function validateQuestion(question) {
    const name = str(question.outputColumnName).trim();
    if (!name) return 'Give every result column a name.';
    if (!question.instructions || (typeof question.instructions === 'string' && !question.instructions.trim())) {
        return `"${name}" needs instructions telling Jev what to decide.`;
    }
    if (question.branches) return null;
    const options = question.options || [];
    if (question.questionType === 'choice') {
        if (options.length < 2) {
            return `"${name}" needs at least 2 options, one per line.`;
        }
        if (options.length > MAX_CHOICE_OPTIONS) {
            return `"${name}" has ${options.length} options; Jev allows at most ${MAX_CHOICE_OPTIONS}.`;
        }
    }
    if (question.questionType === 'score') {
        if (options.length < MIN_SCORE_LEVELS || options.length > MAX_SCORE_LEVELS) {
            return `"${name}" needs between ${MIN_SCORE_LEVELS} and ${MAX_SCORE_LEVELS} levels, one per line (${options.length} given).`;
        }
    }
    return null;
}

/**
 * Build the portable document from the configuration currently on screen.
 * @param {object} state - { sourceColumns, includeConfidence, questions }
 */
const QUESTION_FIELDS = ['outputColumnName', 'questionType', 'instructions', 'options', 'review', 'dependsOn', 'branches', 'criteria'];
const ROOT_FIELDS = ['format', 'version', 'exportedAt', 'sheetName', 'sourceColumns', 'includeConfidence', 'questions', 'derived', 'report', 'execution', 'model'];
const copy = value => JSON.parse(JSON.stringify(value));
export function serializeJevConfig({ sourceColumns = [], includeConfidence = false,
                                     sheetName = '', questions = [], warnings, format, version, exportedAt, ...advanced } = {}) {
    const unknown = Object.keys(advanced).filter(k => !['derived', 'report', 'execution', 'model'].includes(k));
    if (unknown.length) throw new Error(`Unknown workflow fields: ${unknown.join(', ')}`);
    for (const q of questions) {
        const extras = Object.keys(q).filter(k => !QUESTION_FIELDS.includes(k));
        if (extras.length) throw new Error(`Unknown question fields: ${extras.join(', ')}`);
    }
    return {
        format: JEV_CONFIG_FORMAT, version: JEV_CONFIG_VERSION,
        exportedAt: new Date().toISOString(), sheetName,
        sourceColumns: sourceColumns.filter(name => str(name).trim() !== ''),
        includeConfidence: !!includeConfidence,
        ...Object.fromEntries(['derived', 'report', 'execution', 'model'].filter(k => advanced[k] !== undefined).map(k => [k, copy(advanced[k])])),
        questions: questions.map(q => ({
            ...Object.fromEntries(QUESTION_FIELDS.filter(k => q[k] !== undefined).map(k => [k, copy(q[k])])),
            outputColumnName: str(q.outputColumnName).trim(),
            questionType: normalizeType(q.questionType),
            instructions: q.instructions || '', options: q.options || []
        }))
    };
}

export function validateReportConfig(config, columns) {
    if (config.model !== undefined && (typeof config.model !== 'string' || !config.model.trim())) throw new Error('model must be a non-empty model name.');
    if (config.derived !== undefined && !Array.isArray(config.derived)) throw new Error('derived must be an array.');
    const report = config.report || {};
    if (typeof report !== 'object' || Array.isArray(report)) throw new Error('report must be an object.');
    const reportFields = ['title', 'instructions', 'groupBy', 'crossTabs', 'evidenceColumns', 'maxExamples', 'maxGroups', 'maxCategories'];
    if (Object.keys(report).some(k => !reportFields.includes(k))) throw new Error('Unknown report setting. Check the JSON configuration guide.');
    for (const field of ['title', 'instructions']) if (report[field] !== undefined && typeof report[field] !== 'string') throw new Error(`${field} must be text.`);
    for (const field of ['groupBy', 'evidenceColumns', 'crossTabs']) if (report[field] !== undefined && !Array.isArray(report[field])) throw new Error(`${field} must be an array.`);
    const execution = config.execution || {};
    if (typeof execution !== 'object' || Array.isArray(execution) || Object.keys(execution).some(k => !['batchSize', 'workers', 'requestsPerMinute', 'tokensPerSecond', 'maxAttempts', 'structuredState'].includes(k))) throw new Error('Unknown execution setting.');
    if (execution.structuredState !== undefined && typeof execution.structuredState !== 'boolean') throw new Error('structuredState must be true or false.');
    const known = new Set([...columns, ...config.questions.map(q => q.outputColumnName), ...(config.derived || []).map(d => d.outputColumnName)]);
    const fields = [...(report.groupBy || []), ...(report.evidenceColumns || []), ...(report.crossTabs || []).flat()];
    for (const field of fields) if (!known.has(field)) throw new Error(`Report column "${field}" is not in this dataset or workflow.`);
    if ((report.crossTabs || []).some(pair => !Array.isArray(pair) || pair.length !== 2)) throw new Error('Each crossTabs entry needs two columns.');
    for (const key of ['maxExamples', 'maxGroups', 'maxCategories']) {
        if (report[key] !== undefined && (!Number.isInteger(report[key]) || report[key] < 0 || report[key] > 1000)) throw new Error(`${key} must be an integer from 0 to 1000.`);
    }
    if (config.execution?.batchSize !== undefined && (!Number.isInteger(config.execution.batchSize) || config.execution.batchSize < 1 || config.execution.batchSize > 25)) throw new Error('batchSize must be from 1 to 25.');
}

/**
 * Validate and normalize an imported document, reporting which of its source
 * columns are missing from the sheet being analyzed now.
 *
 * As with the LLM config, a column mismatch is a warning rather than a failure:
 * the option lists are the expensive part and are always worth keeping.
 *
 * @param {unknown} raw - parsed JSON from the uploaded file
 * @param {string[]} availableColumns - columns of the sheet being configured
 * @returns {{ sourceColumns, includeConfidence, questions, warnings }}
 * @throws {Error} when the file is not a Jev configuration this app can apply
 */
export function deserializeJevConfig(raw, availableColumns = []) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
        throw new Error('This file is not an Aidstack Jev configuration.');
    }
    if (raw.format !== JEV_CONFIG_FORMAT) {
        throw new Error('This file is not an Aidstack Jev configuration.');
    }
    if (Number(raw.version) > JEV_CONFIG_VERSION) {
        throw new Error(`This configuration was saved by a newer version of Aidstack Insights (v${raw.version}). Update the app, or export it again from the older version.`);
    }
    if (!Array.isArray(raw.questions) || raw.questions.length === 0) {
        throw new Error('This configuration has no result columns in it.');
    }

    const available = new Set(availableColumns);
    const requested = Array.isArray(raw.sourceColumns) ? raw.sourceColumns : [];
    const names = requested.map(name => str(name).trim()).filter(Boolean);
    const missing = names.filter(name => !available.has(name));

    if (Number(raw.version) >= 2) {
        const unknown = Object.keys(raw).filter(k => !ROOT_FIELDS.includes(k));
        if (unknown.length) throw new Error(`Unknown configuration fields: ${unknown.join(', ')}`);
        for (const q of raw.questions) {
            const extras = Object.keys(q).filter(k => !QUESTION_FIELDS.includes(k));
            if (extras.length) throw new Error(`Unknown question fields: ${extras.join(', ')}`);
            if (!QUESTION_TYPES.includes(q.questionType)) throw new Error('Unknown question type.');
        }
    }
    const questions = raw.questions.map(entry => ({
        ...Object.fromEntries(QUESTION_FIELDS.filter(k => entry[k] !== undefined).map(k => [k, copy(entry[k])])),
        outputColumnName: str(entry.outputColumnName).trim(),
        questionType: normalizeType(entry.questionType), instructions: entry.instructions || '',
        options: Array.isArray(entry.options) ? copy(entry.options) : []
    }));

    const warnings = [];
    if (missing.length) {
        warnings.push(`Not in this sheet, so left unselected: ${missing.join(', ')}. Pick the matching columns before running.`);
    }
    return {
        sourceColumns: names.filter(name => available.has(name)),
        includeConfidence: !!raw.includeConfidence,
        sheetName: str(raw.sheetName),
        questions,
        ...Object.fromEntries(['derived', 'report', 'execution', 'model'].filter(k => raw[k] !== undefined).map(k => [k, copy(raw[k])])),
        warnings
    };
}

/** Filename for a downloaded configuration, e.g. aidstack-jev-config-2026-09-20.json */
export function jevConfigFilename(date = new Date()) {
    return `aidstack-jev-config-${date.toISOString().slice(0, 10)}.json`;
}

/**
 * Build the state string sent to Jev for one row.
 *
 * Values are labelled with their column name so question instructions can
 * reference a field by name, matching the docs' guidance to give each question
 * only the context it needs in a structured shape.
 */
export function buildJevState(row, sourceColumns, structured = false) {
    const notEmpty = v => v !== null && v !== undefined && String(v).trim() !== '';
    if (structured) {
        const entries = (sourceColumns || []).filter(col => notEmpty(row[col])).map(col => [col, row[col]]);
        return entries.length ? Object.fromEntries(entries) : null;
    }
    const parts = (sourceColumns || [])
        .filter(col => notEmpty(row[col]))
        .map(col => `${col}: ${row[col]}`);
    return parts.length ? parts.join('\n') : null;
}
