// Export/import of the Configure Analysis setup so a laborious configuration
// (long category lists, several result columns) can be reused on the next file.
// Everything here is pure: the DOM lives in app.js, the file format lives here.

export const CONFIG_FORMAT = 'aidstack-insights-analysis-config';
export const CONFIG_VERSION = 1;

// Output lengths the UI offers; an imported value outside this set is snapped
// back to the default so the <select> can never be left on a phantom option.
export const OUTPUT_TOKEN_CHOICES = [256, 1024, 4096];
export const DEFAULT_OUTPUT_TOKENS = 1024;

const str = value => (typeof value === 'string' ? value : '');

/**
 * Build the portable document from the configuration currently on screen.
 * @param {object} state - { generalInstructions, sheetName, columns }
 * @param {Array} state.columns - [{ columns: string[], outputColumnName, prompt, maxOutputTokens }]
 */
export function serializeConfig({ generalInstructions = '', sheetName = '', columns = [] } = {}) {
    return {
        format: CONFIG_FORMAT,
        version: CONFIG_VERSION,
        exportedAt: new Date().toISOString(),
        // Recorded for the human reading the file and for the import warning
        // about columns that the new sheet does not have.
        sheetName,
        generalInstructions,
        columns: columns.map(col => ({
            columns: (col.columns || []).filter(name => str(name).trim() !== ''),
            outputColumnName: str(col.outputColumnName).trim(),
            prompt: str(col.prompt),
            maxOutputTokens: normalizeTokens(col.maxOutputTokens)
        }))
    };
}

function normalizeTokens(value) {
    const n = Number(value);
    return OUTPUT_TOKEN_CHOICES.includes(n) ? n : DEFAULT_OUTPUT_TOKENS;
}

/**
 * Validate and normalize an imported document, reporting which of its source
 * columns are missing from the sheet being analyzed now.
 *
 * Column names rarely match across files, so a mismatch is a warning, not a
 * failure: everything that does match is applied and the rest is left for the
 * user to re-pick, which beats discarding a long prompt they cannot retype.
 *
 * @param {unknown} raw - parsed JSON from the uploaded file
 * @param {string[]} availableColumns - columns of the sheet being configured
 * @returns {{ generalInstructions: string, columns: Array, warnings: string[] }}
 * @throws {Error} when the file is not a configuration this app can apply
 */
export function deserializeConfig(raw, availableColumns = []) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
        throw new Error('This file is not an Aidstack analysis configuration.');
    }
    if (raw.format !== CONFIG_FORMAT) {
        throw new Error('This file is not an Aidstack analysis configuration.');
    }
    if (Number(raw.version) > CONFIG_VERSION) {
        throw new Error(`This configuration was saved by a newer version of Aidstack Insights (v${raw.version}). Update the app, or export it again from the older version.`);
    }
    if (!Array.isArray(raw.columns) || raw.columns.length === 0) {
        throw new Error('This configuration has no column analyses in it.');
    }

    const available = new Set(availableColumns);
    const missing = new Set();
    const columns = raw.columns.map(entry => {
        const source = entry && typeof entry === 'object' ? entry : {};
        // Accept the single-column shorthand as well as the list form.
        const requested = Array.isArray(source.columns)
            ? source.columns
            : [source.column];
        const names = requested.map(name => str(name).trim()).filter(Boolean);
        names.forEach(name => { if (!available.has(name)) missing.add(name); });
        return {
            columns: names.filter(name => available.has(name)),
            outputColumnName: str(source.outputColumnName).trim(),
            prompt: str(source.prompt),
            maxOutputTokens: normalizeTokens(source.maxOutputTokens)
        };
    });

    const warnings = [];
    if (missing.size) {
        warnings.push(`Not in this sheet, so left unselected: ${[...missing].join(', ')}. Pick the matching columns before analyzing.`);
    }
    return {
        generalInstructions: str(raw.generalInstructions),
        sheetName: str(raw.sheetName),
        columns,
        warnings
    };
}

/** Filename for a downloaded configuration, e.g. aidstack-config-2026-09-19.json */
export function configFilename(date = new Date()) {
    return `aidstack-config-${date.toISOString().slice(0, 10)}.json`;
}
