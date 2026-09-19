// Export/import of the Jev Classification setup. A controlled vocabulary is
// laborious to type (the option lists can run to dozens of entries), so the
// whole question set is portable across datasets.
// Everything here is pure: the DOM lives in app.js, the file format lives here.

export const JEV_CONFIG_FORMAT = 'aidstack-insights-jev-config';
export const JEV_CONFIG_VERSION = 1;

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
    return str(text).split('\n').map(line => line.trim()).filter(Boolean);
}

/** Join an option list back into textarea content. */
export function formatOptions(options) {
    return (options || []).join('\n');
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
    if (!str(question.instructions).trim()) {
        return `"${name}" needs instructions telling Jev what to decide.`;
    }
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
export function serializeJevConfig({ sourceColumns = [], includeConfidence = false,
                                     sheetName = '', questions = [] } = {}) {
    return {
        format: JEV_CONFIG_FORMAT,
        version: JEV_CONFIG_VERSION,
        exportedAt: new Date().toISOString(),
        // Recorded for the human reading the file and for the import warning
        // about columns the new sheet does not have.
        sheetName,
        sourceColumns: sourceColumns.filter(name => str(name).trim() !== ''),
        includeConfidence: !!includeConfidence,
        questions: questions.map(q => ({
            outputColumnName: str(q.outputColumnName).trim(),
            questionType: normalizeType(q.questionType),
            instructions: str(q.instructions),
            options: (q.options || []).map(o => str(o).trim()).filter(Boolean)
        }))
    };
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

    const questions = raw.questions.map(entry => {
        const source = entry && typeof entry === 'object' ? entry : {};
        return {
            outputColumnName: str(source.outputColumnName).trim(),
            questionType: normalizeType(source.questionType),
            instructions: str(source.instructions),
            options: Array.isArray(source.options)
                ? source.options.map(o => str(o).trim()).filter(Boolean)
                : []
        };
    });

    const warnings = [];
    if (missing.length) {
        warnings.push(`Not in this sheet, so left unselected: ${missing.join(', ')}. Pick the matching columns before running.`);
    }
    return {
        sourceColumns: names.filter(name => available.has(name)),
        includeConfidence: !!raw.includeConfidence,
        sheetName: str(raw.sheetName),
        questions,
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
export function buildJevState(row, sourceColumns) {
    const notEmpty = v => v !== null && v !== undefined && String(v).trim() !== '';
    const parts = (sourceColumns || [])
        .filter(col => notEmpty(row[col]))
        .map(col => `${col}: ${row[col]}`);
    return parts.length ? parts.join('\n') : null;
}
