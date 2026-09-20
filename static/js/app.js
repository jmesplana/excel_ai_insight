import { initResults } from './results.js';
let resultsUI;
import { API_KEY_STORAGE_KEY, PROVIDER_STORAGE_KEY, AZURE_ENDPOINT_STORAGE_KEY, AZURE_DEPLOYMENT_STORAGE_KEY, AZURE_API_VERSION_STORAGE_KEY, OPENAI_MODEL_STORAGE_KEY, getLLMConfig, hasLocalCredentials, syncProviderUI } from './provider-settings.js';
import { escapeHtml, renderMarkdown } from './rendering.js';
import { parseWorkbook, exportResults, explodeColumn } from './spreadsheet.js';
import { BatchRun } from './batch-runner.js';
import { serializeConfig, deserializeConfig, configFilename } from './analysis-config.js';
import { getJevConfig, hasJevCredentials, JEV_API_KEY_STORAGE_KEY, JEV_MODEL_STORAGE_KEY } from './provider-settings.js';
import { serializeJevConfig, deserializeJevConfig, jevConfigFilename, parseOptions,
    formatOptions, validateQuestion, buildJevState, DEFAULT_QUESTION_TYPE } from './jev-config.js';

/* Global variables */
let availableColumns = [];
let fileData = {};
let currentStep = 1;
let resultChart = null;
const ICD_CLIENT_ID_STORAGE_KEY = 'excel_ai_insight_icd_client_id';
const ICD_CLIENT_SECRET_STORAGE_KEY = 'excel_ai_insight_icd_client_secret';

// Application mode: 'analysis' (default AI analysis) | 'jev' (Jev typed
// classification against a controlled list) | 'icd' (Medical Translation
// ICD-11) | 'clean' (split multi-value cells; no AI, no credentials, no step 4/5)
let appMode = 'analysis';

// Holds the ICD translation dataset in memory: { sheetName, columns: [...], data: [...] }
let icdResult = null;

// ICD-11 language options (release 2026-01 MMS) shared by all language dropdowns.
const ICD_LANGUAGES = [
    ['ar', 'Arabic'], ['zh', 'Chinese'], ['cs', 'Czech'], ['en', 'English'],
    ['fr', 'French'], ['de', 'German'], ['kk', 'Kazakh'], ['la', 'Latin'],
    ['pt', 'Portuguese'], ['ru', 'Russian'], ['sk', 'Slovak'], ['es', 'Spanish'],
    ['sv', 'Swedish'], ['tr', 'Turkish'], ['uz', 'Uzbek']
];

function icdLangOptionsHtml(defaultCode) {
    return ICD_LANGUAGES.map(([code, name]) =>
        `<option value="${code}"${code === defaultCode ? ' selected' : ''}>${name}</option>`
    ).join('');
}

// Holds the analyzed dataset in memory: { sheetName, columns: [...], data: [...] }
let analyzedResult = null;
const ANALYZE_BATCH_SIZE = 20;     // rows per /analyze_batch request
// Jev spends one request per row rather than per cell, and the API allows
// 1,200 requests/minute, so a batch can safely be larger than the LLM path's.
const JEV_BATCH_SIZE = 25;         // rows per /analyze_batch_jev request
const MAX_CHAT_ROWS = 5000;        // cap rows sent to /chat_with_data (Vercel body limit)

// Build the input text for one (row, config) cell, mirroring the server's
// previous process_row logic (single value, or "col: val" join for multi-column).
function buildAnalysisInput(row, config) {
    const cols = (config.columns && config.columns.length) ? config.columns : [config.column];
    const notEmpty = v => v !== null && v !== undefined && String(v).trim() !== '';
    if (cols.length > 1) {
        const parts = cols.filter(c => notEmpty(row[c])).map(c => `${c}: ${row[c]}`);
        return parts.length ? parts.join('\n') : null;
    }
    return notEmpty(row[config.column]) ? String(row[config.column]) : null;
}

// Return { columns, rows } for the chat endpoint from the analyzed data if
// present, otherwise from the currently selected sheet of the uploaded file.
function getChatDataset() {
    if (analyzedResult) {
        return { columns: analyzedResult.columns, rows: analyzedResult.data.slice(0, MAX_CHAT_ROWS) };
    }
    if (fileData && fileData.sheets) {
        const sel = document.getElementById('sheet-select');
        const name = (sel && sel.value && fileData.sheets[sel.value])
            ? sel.value : Object.keys(fileData.sheets)[0];
        const sd = fileData.sheets[name];
        if (sd) return { columns: sd.columns, rows: sd.data.slice(0, MAX_CHAT_ROWS) };
    }
    return null;
}

// POST a chat question and consume the SSE stream. Calls onToken(fullText)
// as content arrives; resolves with the full accumulated answer.
async function streamChat(payload, onToken) {
    const response = await fetch('/chat_with_data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });
    if (!response.ok) {
        const err = await readJson(response).catch(() => ({}));
        throw new Error(err.error || `HTTP error! status: ${response.status}`);
    }
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '', full = '';
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop();
        for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed.startsWith('data:')) continue;
            let data;
            try { data = JSON.parse(trimmed.slice(5).trim()); } catch (e) { continue; }
            if (data.error) throw new Error(data.error);
            if (data.content) { full += data.content; if (onToken) onToken(full); }
        }
    }
    return full;
}

// Build an .xlsx from the in-memory analyzed dataset and trigger a download.
// =========================================================================
// CLEAN DATA: split a multi-value column into one row per value
// =========================================================================

// "|| | ;" -> ['||','|',';']. Whitespace separates the separators, so a
// literal space cannot itself be one; that is the intended trade-off.
function parseSeparators(text) {
    return (text || '').trim().split(/\s+/).filter(Boolean);
}

/**
 * Wire one Clean Data panel.
 *
 * The same panel serves two places with different data sources: step 3 acts
 * on the uploaded sheet (no API key, no spend), step 5 on the analyzed
 * results. The caller supplies get/set so this code never needs to know which.
 *
 * @param {string} prefix - element id prefix ('explode' or 'pre-explode').
 * @param {() => object|null} getSource - current {columns, data} to split.
 * @param {(result: object) => void} onApply - install the exploded dataset.
 * @param {() => string[]} [preferredColumns] - columns to default the picker to.
 */
function createCleanDataPanel({ prefix, getSource, onApply, preferredColumns = () => [] }) {
    const el = suffix => document.getElementById(`${prefix}-${suffix}`);
    let preview = null;      // staged {result, column}, not yet applied
    let beforeApply = null;  // source snapshot, for Undo
    // The split dataset currently offered for download. Set by Preview and
    // kept across Apply, which clears `preview` but must not strip the only
    // way to get the file out (clean mode has no other download button).
    let downloadable = null;

    const setMessage = (text, type = 'info') => {
        const box = el('message');
        if (box) box.innerHTML = text
            ? `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(text)}</div>` : '';
    };

    // Keep the current selection when possible; otherwise fall back to the
    // caller's preferred column (the newest AI output, on step 5).
    function refreshColumns() {
        const select = el('column');
        const source = getSource();
        if (!select || !source) return;
        const previous = select.value;
        select.innerHTML = source.columns
            .map(col => `<option value="${escapeHtml(col)}">${escapeHtml(col)}</option>`).join('');
        if (source.columns.includes(previous)) select.value = previous;
        else {
            const preferred = preferredColumns().filter(c => source.columns.includes(c));
            if (preferred.length) select.value = preferred[preferred.length - 1];
        }
    }

    function renderPreview(result, column) {
        const rows = result.data.slice(0, 20);
        let html = '<thead class="table-light"><tr>';
        result.columns.forEach(col => {
            html += `<th${col === column ? ' class="table-warning"' : ''}>${escapeHtml(col)}</th>`;
        });
        html += '</tr></thead><tbody>';
        rows.forEach(row => {
            html += '<tr>';
            result.columns.forEach(col => {
                const value = row[col];
                html += `<td>${escapeHtml(value === null || value === undefined ? '' : value)}</td>`;
            });
            html += '</tr>';
        });
        el('preview-table').innerHTML = html + '</tbody>';
        const info = el('preview-info');
        if (info) info.textContent = `— showing ${rows.length} of ${result.data.length} rows`;
        el('preview').classList.remove('hidden');
    }

    function buildPreview() {
        const source = getSource();
        if (!source) { setMessage('Load a file first.', 'warning'); return; }
        const column = el('column').value;
        const separators = parseSeparators(el('separator').value);
        const rename = (el('output').value || '').trim();
        const dedupe = el('dedupe').checked;
        if (!separators.length) { setMessage('Enter at least one separator.', 'warning'); return; }

        try {
            const exploded = explodeColumn(source, column, separators,
                { outputColumn: rename || column, dedupe });
            preview = { result: exploded, column: rename || column };
            const before = source.data.length, after = exploded.data.length;
            if (after === before) {
                setMessage(`No cell in "${column}" contained ${separators.join(' or ')}. `
                    + 'Check the separator and retry.', 'warning');
            } else {
                setMessage(`${before} rows become ${after} rows.`, 'success');
            }
            renderPreview(exploded, rename || column);
            downloadable = exploded;
            el('download-btn').classList.remove('hidden');
        } catch (error) {
            preview = null;
            downloadable = null;
            el('download-btn').classList.add('hidden');
            el('preview').classList.add('hidden');
            setMessage(error.message, 'danger');
        }
    }

    function apply() {
        if (!preview) return;
        beforeApply = getSource();
        const applied = preview.result;
        const producedColumn = preview.column;
        onApply(applied);
        refreshColumns();
        // Keep the picker on the column the split just produced, rather than
        // letting it fall back to the first column in the sheet.
        const select = el('column');
        if (select && applied.columns.includes(producedColumn)) select.value = producedColumn;
        el('undo-btn').classList.remove('hidden');
        el('preview').classList.add('hidden');
        // Applying replaces the working sheet; the download stays armed so the
        // split file is still one click away.
        downloadable = applied;
        el('download-btn').classList.remove('hidden');
        setMessage(`Applied — now ${applied.data.length} rows. `
            + 'Use "Download split file" to save it.', 'success');
        preview = null;
    }

    function undo() {
        if (!beforeApply) return;
        onApply(beforeApply);
        setMessage(`Reverted — back to ${beforeApply.data.length} rows.`, 'info');
        beforeApply = null;
        refreshColumns();
        el('undo-btn').classList.add('hidden');
        // Nothing split is in effect any more, so offer nothing to download.
        downloadable = null;
        el('download-btn').classList.add('hidden');
    }

    // Clear staged state so a new file or a new analysis never inherits a
    // preview or an Undo that points at the previous dataset.
    function reset() {
        preview = null;
        beforeApply = null;
        downloadable = null;
        el('preview')?.classList.add('hidden');
        el('undo-btn')?.classList.add('hidden');
        el('download-btn')?.classList.add('hidden');
        setMessage('');
    }

    const previewBtn = el('preview-btn');
    if (!previewBtn) return { refreshColumns() {}, reset() {} };
    previewBtn.addEventListener('click', buildPreview);
    el('apply-btn').addEventListener('click', apply);
    el('undo-btn').addEventListener('click', undo);
    el('download-btn').addEventListener('click', () => {
        if (downloadable) exportResults(downloadable, fileData.filename || 'data', 'split');
    });
    return { refreshColumns, reset };
}

// Step 5 panel: splits the analyzed results.
let cleanDataResults = null;
// Step 3 panel: splits the uploaded sheet before any analysis.
let cleanDataPreview = null;

function initCleanData() {
    cleanDataResults = createCleanDataPanel({
        prefix: 'explode',
        getSource: () => analyzedResult,
        onApply: result => {
            analyzedResult = result;
            resultsUI.load({ sheets: { [result.sheetName]: result } });
        },
        preferredColumns: () => (analyzedResult && analyzedResult.outputColumns) || [],
    });

    cleanDataPreview = createCleanDataPanel({
        prefix: 'pre-explode',
        getSource: () => currentSheet(),
        onApply: result => {
            // Write the split sheet back into fileData so the preview table,
            // the column pickers and the analysis itself all see the new rows.
            const name = document.getElementById('sheet-select').value;
            fileData.sheets[name] = { columns: result.columns, data: result.data };
            updatePreviewTable(name);
        },
    });
}

// The sheet currently selected in the step 3 preview, or null before upload.
function currentSheet() {
    const select = document.getElementById('sheet-select');
    if (!select || !fileData.sheets) return null;
    const sheet = fileData.sheets[select.value];
    return sheet ? { ...sheet, sheetName: select.value } : null;
}


function downloadAnalyzedFile() {
    if (!analyzedResult) return;
    exportResults(analyzedResult, fileData.filename || 'data', analyzedResult.partial ? 'partial' : 'analyzed');
}

function downloadIcdFile() {
    exportResults(icdResult, fileData.filename || 'data', 'icd_mapped');
}

function renderIcdResults() {
    const container = document.getElementById('icd-result-preview');
    const info = document.getElementById('icd-table-info');
    if (!container || !icdResult) return;

    const esc = v => (v === null || v === undefined) ? '' :
        String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    let html = '<table class="table table-striped table-bordered"><thead class="table-light"><tr>';
    icdResult.columns.forEach(col => { html += `<th>${esc(col)}</th>`; });
    html += '</tr></thead><tbody>';
    const reviewOnly = document.getElementById('icd-review-only').checked;
    const visibleRows = icdResult.data.filter(row => !reviewOnly || row['Review status'] !== 'WHO code found');
    visibleRows.forEach(row => {
        html += '<tr>';
        icdResult.columns.forEach(col => { html += `<td>${esc(row[col])}</td>`; });
        html += '</tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
    if (info) info.textContent = `${visibleRows.length} of ${icdResult.data.length} row(s)`;
}

// Run the Medical Translation (ICD-11) workflow. Mirrors analyzeColumns():
// validates credentials, fail-fast token check, batches rows to the backend,
// accumulates results by rowIndex, builds icdResult and renders step 5.
const ICD_TEST_ROWS = 10;

/**
 * Rows to use for a test run, read from that mode's "Test rows" input.
 *
 * The field is a free number box, so it is clamped to what the sheet actually
 * has: a typo of 500 on a 20-row file tests 20 rows rather than erroring, and
 * a blank or junk value falls back to the mode's default instead of running
 * zero rows (which would look like a silent failure).
 *
 * @param {string} inputId - id of the number input for this mode
 * @param {number} fallback - default when the field is empty or unusable
 * @param {number} available - rows in the sheet being analyzed
 */
function testRowCount(inputId, fallback, available) {
    const requested = Math.floor(Number(document.getElementById(inputId)?.value));
    const rows = Number.isFinite(requested) && requested > 0 ? requested : fallback;
    return Math.max(1, Math.min(rows, available));
}

async function runIcdTranslation(isTestRun = false) {
    const openaiApiKey = (document.getElementById('modal-api-key').value || '').trim() ||
                         (localStorage.getItem(API_KEY_STORAGE_KEY) || '').trim();
    const clientId = (document.getElementById('modal-icd-client-id').value || '').trim() ||
                     (localStorage.getItem(ICD_CLIENT_ID_STORAGE_KEY) || '').trim();
    const clientSecret = (document.getElementById('modal-icd-client-secret').value || '').trim() ||
                         (localStorage.getItem(ICD_CLIENT_SECRET_STORAGE_KEY) || '').trim();

    // WHO ICD credentials are mandatory.
    if (!clientId || !clientSecret) {
        showAlert('icd-config-message',
            'WHO ICD API credentials are required. Open API Settings (top navigation) to add your Client ID and Client Secret.',
            'danger');
        return;
    }

    const inputType = document.querySelector('input[name="icd-input-type"]:checked').value; // text | code
    const sourceSystem = document.getElementById('icd-source-system').value;                // mms | icd10
    const sourceLang = document.getElementById('icd-source-lang').value;
    const targetLang = document.getElementById('icd-target-lang').value;
    const sourceColumn = document.getElementById('icd-source-column').value;

    const wantIcd11 = document.getElementById('icd-target-icd11').checked;
    const wantIcd10 = document.getElementById('icd-target-icd10').checked;
    const targets = [];
    if (wantIcd11) targets.push('icd11');
    if (wantIcd10) targets.push('icd10');

    if (!sourceColumn) {
        showAlert('icd-config-message', 'Please choose a source column.', 'danger');
        return;
    }
    if (targets.length === 0) {
        showAlert('icd-config-message', 'Please select at least one output (ICD-11 term and/or ICD-10).', 'danger');
        return;
    }

    const selectedSheet = document.getElementById('sheet-select').value;
    const sheet = fileData.sheets[selectedSheet];
    if (!sheet || !sheet.data) {
        showAlert('icd-config-message', 'No data available for the selected sheet.', 'danger');
        return;
    }
    const allRows = sheet.data;
    const totalRows = allRows.length;
    if (!totalRows) {
        showAlert('icd-config-message', 'This sheet has no data rows.', 'warning');
        return;
    }
    const rowCount = isTestRun ? testRowCount('icd-test-rows', ICD_TEST_ROWS, totalRows) : totalRows;

    showSpinner(true, 'Validating WHO ICD credentials...', false);

    try {
        // Fail-fast token exchange.
        const validateResp = await fetch('/icd_validate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ clientId: clientId, clientSecret: clientSecret })
        });
        const validateResult = await readJson(validateResp);
        if (!validateResp.ok || !validateResult.ok) {
            showSpinner(false);
            showAlert('icd-config-message',
                `WHO ICD credential check failed: ${validateResult.error || ('HTTP ' + validateResp.status)}`,
                'danger');
            return;
        }

        // Batch translate.
        showSpinner(true, isTestRun ? `Running test on first ${rowCount} rows...` : 'Translating to ICD-11...', true);
        updateProgress(0, 0, rowCount, 'Starting translation...');

        const resultsByIndex = {};
        let totalErrors = 0;

        for (let start = 0; start < rowCount; start += ANALYZE_BATCH_SIZE) {
            const end = Math.min(start + ANALYZE_BATCH_SIZE, rowCount);

            const batchRows = [];
            for (let i = start; i < end; i++) {
                const cell = allRows[i][sourceColumn];
                batchRows.push({ rowIndex: i, text: (cell === null || cell === undefined) ? '' : String(cell) });
            }

            const response = await fetch('/icd_translate_batch', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    clientId: clientId,
                    clientSecret: clientSecret,
                    ...getLLMConfig(),
                    inputType: inputType,
                    sourceSystem: sourceSystem,
                    sourceLang: sourceLang,
                    targetLang: targetLang,
                    targets: targets,
                    rows: batchRows
                })
            });

            const result = await readJson(response);
            if (!response.ok) throw new Error(result.error || `HTTP error! status: ${response.status}`);
            if (result.error) throw new Error(result.error);

            totalErrors += result.errors || 0;
            (result.results || []).forEach(r => { resultsByIndex[r.rowIndex] = r; });

            const done = end;
            updateProgress(Math.round((done / rowCount) * 100), done, rowCount,
                `Translated ${done} of ${rowCount} rows...`);
        }

        // Build the output columns (original + added, in selection order).
        const addedColumns = [];
        if (wantIcd11) {
            addedColumns.push('ICD-11 Code');
            addedColumns.push(`ICD-11 Term (${targetLang.toUpperCase()})`);
        }
        if (wantIcd10) {
            addedColumns.push('ICD-10 Code');
            addedColumns.push('ICD-10 Match');
        }
        addedColumns.push('Type');
        addedColumns.push('Source Term');
        addedColumns.push('Confidence');
        addedColumns.push('Source');
        addedColumns.push('Notes', 'Review status', 'WHO reference', 'ICD-11 code source', 'ICD-11 term source', 'ICD-10 code source');

        // De-duplicate: if the source sheet already has any of these
        // columns (e.g. a previously-exported file was re-uploaded, or a
        // re-run), reuse them and overwrite the values in place instead of
        // appending empty duplicate columns.
        const outColumns = sheet.columns.slice();
        addedColumns.forEach(c => { if (!outColumns.includes(c)) outColumns.push(c); });
        // Explain, in plain language, why a row did not map to a clean ICD
        // diagnosis code. Procedures/administrative entries are not codeable
        // in ICD (which classifies diseases), and unverified LLM answers are
        // flagged so the user knows to double-check them.
        const deriveIcdNote = (r) => {
            const o = r.outputs || {};
            const hasCode = !!((o.icd11 && o.icd11.code) || (o.icd10 && o.icd10.code));
            const parts = [];
            if (r.kind === 'procedure') {
                parts.push('Procedure, not a diagnosis — ICD-11 classifies diseases, not interventions, so there is no true diagnosis code. Use the WHO ICHI classification for procedures.');
            } else if (r.kind === 'other') {
                parts.push('Not a codeable diagnosis (e.g. an administrative, symptom, or non-clinical entry).');
            }
            if (!hasCode) {
                parts.push(r.note || 'No ICD match found.');
            } else if (r.note) {
                // e.g. "Provided by LLM (not found in WHO API)".
                parts.push(r.note);
            }
            return parts.join(' ');
        };
        const outData = allRows.slice(0, rowCount).map((row, i) => {
            const out = Object.assign({}, row);
            const r = resultsByIndex[i] || {};
            const outputs = r.outputs || {};
            if (wantIcd11) {
                const o11 = outputs.icd11 || {};
                out['ICD-11 Code'] = o11.code || '';
                out[`ICD-11 Term (${targetLang.toUpperCase()})`] = o11.term || '';
            }
            if (wantIcd10) {
                const o10 = outputs.icd10 || {};
                out['ICD-10 Code'] = o10.code || '';
                out['ICD-10 Match'] = o10.relationship || '';
            }
            const kindLabel = {
                diagnosis: 'Diagnosis',
                procedure: 'Procedure (not codeable in ICD — see ICHI)',
                other: 'Other'
            }[r.kind] || (r.kind || '');
            out['Type'] = kindLabel;
            out['Source Term'] = r.sourceTitle || '';
            out['Confidence'] = r.confidence || '';
            out['Source'] = r.source || '';
            out['Notes'] = deriveIcdNote(r);
            out['Review status'] = !r.entityUri ? 'Unverified / unmatched' :
                r.confidence !== 'high' || (r.source || '').includes('LLM') ? 'Needs review' : 'WHO code found';
            out['WHO reference'] = r.entityUri || '';
            out['ICD-11 code source'] = outputs.icd11?.codeSource || '';
            out['ICD-11 term source'] = outputs.icd11?.termSource || '';
            out['ICD-10 code source'] = outputs.icd10?.codeSource || '';
            return out;
        });

        icdResult = { sheetName: selectedSheet, columns: outColumns, data: outData, isTest: isTestRun };

        const icdMsg = document.getElementById('icd-result-message');
        if (icdMsg) {
            const errNote = totalErrors ? ` ${totalErrors} row(s) had no/low match.` : '';
            icdMsg.querySelector('span').innerHTML = isTestRun
                ? `<i class="bi bi-lightning-charge"></i> <strong>Test run</strong> on the first ${rowCount} of ${totalRows} row(s).${errNote} Review the results below, then go <strong>Back</strong> and click <strong>Translate All Rows</strong> to process everything.`
                : `<i class="bi bi-check-circle"></i> Translation complete! Mapped ${rowCount} row(s).${errNote}`;
        }

        document.getElementById('icd-review-only').onchange = renderIcdResults;
        renderIcdResults();

        const dlLink = document.getElementById('icd-download-link');
        if (dlLink) {
            // Hide the download for a test sample to avoid confusing it with the full output.
            dlLink.style.display = isTestRun ? 'none' : '';
            dlLink.href = '#';
            dlLink.onclick = (ev) => { ev.preventDefault(); downloadIcdFile(); };
        }

        goToStep(5);
    } catch (error) {
        showAlert('icd-config-message', `Error during translation: ${error.message}`, 'danger');
    } finally {
        showSpinner(false);
    }
}

/* Utility functions */
function showSpinner(show, message = 'Processing...', showProgress = false) {
    const spinner = document.getElementById('spinner-overlay');
    const progressContainer = document.getElementById('progress-container');
    
    document.getElementById('spinner-message').textContent = message;
    
    if (show) {
        spinner.classList.remove('hidden');
        
        // Show or hide progress tracking UI
        if (showProgress) {
            progressContainer.classList.remove('hidden');
            resetProgress();
        } else {
            progressContainer.classList.add('hidden');
        }
    } else {
        spinner.classList.add('hidden');
        progressContainer.classList.add('hidden');
    }
}

function resetProgress() {
    document.getElementById('progress-percentage').textContent = '0%';
    document.getElementById('operations-count').textContent = '0/0';
    document.getElementById('progress-bar').style.width = '0%';
    document.getElementById('current-operation').textContent = 'Initializing...';
}

function updateProgress(percentage, completedOperations, totalOperations, currentOperation = null) {
    document.getElementById('progress-percentage').textContent = `${percentage}%`;
    document.getElementById('operations-count').textContent = `${completedOperations}/${totalOperations}`;
    document.getElementById('progress-bar').style.width = `${percentage}%`;
    
    if (currentOperation) {
        document.getElementById('current-operation').textContent = currentOperation;
    }
}

function showAlert(id, message, type = 'success') {
    const alertElement = document.getElementById(id);
    alertElement.textContent = message;
    alertElement.className = `alert alert-${type} mt-3`;
    alertElement.classList.remove('hidden');
    
    // Auto-hide success messages after 5 seconds
    if (type === 'success') {
        setTimeout(() => {
            alertElement.classList.add('hidden');
        }, 5000);
    }
}

function goToStep(step) {
    resultsUI?.onStep(step);
    // The Clean Data picker lists the columns of the current results, which
    // only exist once the analysis has produced them.
    if (step === 5 && appMode === 'analysis') cleanDataResults?.refreshColumns();
    // Hide all steps
    document.querySelectorAll('.step-content').forEach(el => el.classList.add('hidden'));
    
    // Show the target step
    document.getElementById(`step-${step}`).classList.remove('hidden');
    
    // Update the stepper
    document.querySelectorAll('.stepper-item').forEach(el => {
        const stepNum = parseInt(el.dataset.step);
        
        if (stepNum < step) {
            el.classList.add('completed');
            el.classList.remove('active');
        } else if (stepNum === step) {
            el.classList.add('active');
            el.classList.remove('completed');
        } else {
            el.classList.remove('active', 'completed');
        }
    });

    // Branch the per-step panels based on the active application mode.
    applyModePanels(step);

    currentStep = step;
}

// Show/hide the analysis vs ICD panels for the given step according to appMode.
function applyModePanels(step) {
    const isIcd = (appMode === 'icd');
    const isClean = (appMode === 'clean');
    const isJev = (appMode === 'jev');
    const setHidden = (id, hidden) => {
        const el = document.getElementById(id);
        if (el) el.classList.toggle('hidden', hidden);
    };

    if (step === 1) {
        // Clean mode needs no instructions and no credentials; the mode cards
        // are the whole of step 1 for it.
        // Jev takes its instructions per question on step 4, so the shared
        // general-instructions box does not apply to it either.
        setHidden('analysis-instructions-block', isIcd || isClean || isJev);
        setHidden('icd-intro-block', !isIcd);
        setHidden('jev-intro-block', !isJev);
        setHidden('clean-intro-block', !isClean);
        // The default banner and title speak for the AI flows; clean mode
        // needs no credentials, so neither should mention them.
        setHidden('config-intro-alert', isClean);
        // Jev uses its own key, so the banner must not send the user to the
        // OpenAI/Azure settings the LLM path needs.
        const intro = document.getElementById('config-intro-alert');
        if (intro && !isClean) {
            intro.innerHTML = isJev
                ? '<i class="bi bi-info-circle"></i> Pick <strong>Jev Classification</strong> below, then define your result columns and their option lists on the <strong>Configure Analysis</strong> step. Set your <strong>Jev API Key</strong> in the <strong>API Settings</strong> button in the top navigation bar, unless your server already provides one.'
                : '<i class="bi bi-info-circle"></i> Before starting your analysis, configure general instructions that will apply to all analyzed columns. Configure OpenAI or Azure in the <strong>API Settings</strong> button in the top navigation bar, unless your server already provides credentials.';
        }
        const title = document.getElementById('config-card-title');
        if (title) title.textContent = isClean ? 'Choose a Mode' : 'Analysis Configuration';
    }

    if (step === 3) {
        // Clean mode ends at step 3: the panel is the whole workflow, so open
        // it by default and drop the "Next" that leads into analysis config.
        const body = document.getElementById('pre-clean-data-body');
        if (body) body.classList.toggle('show', isClean);
        setHidden('preview-next-btn', isClean);
        const chatTip = document.getElementById('preview-chat-tip');
        if (chatTip) chatTip.classList.toggle('hidden', isClean);
    }

    if (step === 4) {
        setHidden('analysis-config', isIcd || isJev);
        setHidden('icd-config', !isIcd);
        setHidden('jev-config', !isJev);
        if (isIcd) {
            populateIcdConfig();
        }
        if (isJev) {
            populateJevConfig();
        }
    }

    if (step === 5) {
        // Jev writes into analyzedResult like the LLM path, so it shares the
        // analysis results view (preview, chart, chat and download).
        setHidden('analysis-results', isIcd);
        setHidden('icd-results', !isIcd);
    }
}

// Populate the ICD config controls (source column + language dropdowns) from
// the currently selected sheet. Called when entering step 4 in ICD mode.
function populateIcdConfig() {
    const sheetSelect = document.getElementById('sheet-select');
    const selectedSheet = sheetSelect ? sheetSelect.value : null;
    const sheet = (fileData.sheets && selectedSheet) ? fileData.sheets[selectedSheet] : null;
    const columns = (sheet && sheet.columns) ? sheet.columns : (availableColumns || []);

    const colSelect = document.getElementById('icd-source-column');
    if (colSelect) {
        const prev = colSelect.value;
        colSelect.innerHTML = '';
        columns.forEach(col => {
            const opt = document.createElement('option');
            opt.value = col;
            opt.textContent = col;
            colSelect.appendChild(opt);
        });
        if (prev && columns.includes(prev)) colSelect.value = prev;
    }

    // Populate language dropdowns once (defaults: source es, target en).
    const srcLang = document.getElementById('icd-source-lang');
    if (srcLang && !srcLang.options.length) srcLang.innerHTML = icdLangOptionsHtml('es');
    const tgtLang = document.getElementById('icd-target-lang');
    if (tgtLang && !tgtLang.options.length) tgtLang.innerHTML = icdLangOptionsHtml('en');

    updateIcdInputTypeToggle();
}

// Toggle the source-lang (free text) vs source-system (existing code) controls.
function updateIcdInputTypeToggle() {
    const codeRadio = document.getElementById('icd-input-type-code');
    const isCode = codeRadio ? codeRadio.checked : false;
    const langGroup = document.getElementById('icd-source-lang-group');
    const sysGroup = document.getElementById('icd-source-system-group');
    if (langGroup) langGroup.classList.toggle('hidden', isCode);
    if (sysGroup) sysGroup.classList.toggle('hidden', !isCode);
}

/* Theme toggle */
function toggleDarkMode() {
    const body = document.body;
    const themeIcon = document.querySelector('#theme-toggle i');
    
    body.classList.toggle('dark-mode');
    
    if (body.classList.contains('dark-mode')) {
        themeIcon.classList.remove('bi-sun-fill');
        themeIcon.classList.add('bi-moon-fill');
        localStorage.setItem('theme', 'dark');
    } else {
        themeIcon.classList.remove('bi-moon-fill');
        themeIcon.classList.add('bi-sun-fill');
        localStorage.setItem('theme', 'light');
    }
}

/* Initialize tooltips.
 *
 * Scoped to `root` so re-rendering one card does not touch triggers elsewhere,
 * and guarded by getOrCreateInstance: constructing a second Tooltip on an
 * element orphans the first, which then has no working hide handler and leaves
 * its popup stuck on screen. */
function initTooltips(root = document) {
    const tooltipTriggerList = root.querySelectorAll('[data-bs-toggle="tooltip"]');
    [...tooltipTriggerList].forEach(el => bootstrap.Tooltip.getOrCreateInstance(el));
}

/* Dispose tooltips inside `root` before its markup is discarded.
 *
 * Bootstrap appends the visible popup to document.body, not next to the
 * trigger, so removing the trigger alone strands the popup on screen. */
function disposeTooltips(root) {
    if (!root) return;
    root.querySelectorAll('[data-bs-toggle="tooltip"]').forEach(el => {
        bootstrap.Tooltip.getInstance(el)?.dispose();
    });
}

/* Navigation between main content and about page */
function showMainContent() {
    document.getElementById('main-content').classList.remove('hidden');
    document.getElementById('about-content').classList.add('hidden');
}

function showAbout() {
    document.getElementById('main-content').classList.add('hidden');
    document.getElementById('about-content').classList.remove('hidden');
}

/* Chart initialization */
function initResultChart(data, labels, columnName) {
    // Ensure resultChart is declared
    if (typeof resultChart === 'undefined') {
        window.resultChart = null;
    }
    
    // Destroy existing chart if it exists
    if (resultChart) {
        resultChart.destroy();
        resultChart = null;
    }
    
    // Create a simple bar chart showing frequency of analysis results
    const ctx = document.getElementById('result-chart').getContext('2d');
    
    // Count occurrences of each result category
    const counts = {};
    data.forEach(item => {
        if (!counts[item]) {
            counts[item] = 1;
        } else {
            counts[item]++;
        }
    });
    
    // Prepare data for chart
    const chartLabels = Object.keys(counts);
    const chartData = Object.values(counts);
    
    // Create chart
    resultChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: chartLabels,
            datasets: [{
                label: `Analysis Results: ${columnName}`,
                data: chartData,
                backgroundColor: 'rgba(76, 175, 80, 0.6)',
                borderColor: 'rgba(76, 175, 80, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Frequency'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Categories'
                    }
                }
            }
        }
    });
}

/* API Key Management */
function loadSavedApiKey() {
    const savedKey = localStorage.getItem(API_KEY_STORAGE_KEY);
    if (savedKey) {
        // Update main form field if it exists
        const apiKeyInput = document.getElementById('modal-api-key');
        if (apiKeyInput) {
            apiKeyInput.value = savedKey;
            const saveKeyCheckbox = document.getElementById('save-api-key');
            if (saveKeyCheckbox) {
                saveKeyCheckbox.checked = true;
            }
        }
        
        // Update modal fields if they exist
        const modalApiKeyField = document.getElementById('modal-api-key');
        const modalSaveKeyCheckbox = document.getElementById('modal-save-api-key');
        if (modalApiKeyField) {
            modalApiKeyField.value = savedKey;
            if (modalSaveKeyCheckbox) {
                modalSaveKeyCheckbox.checked = true;
            }
        }
    }

    // Load saved WHO ICD credentials (used by Medical Translation mode)
    const savedIcdId = localStorage.getItem(ICD_CLIENT_ID_STORAGE_KEY);
    const savedIcdSecret = localStorage.getItem(ICD_CLIENT_SECRET_STORAGE_KEY);
    const icdIdField = document.getElementById('modal-icd-client-id');
    const icdSecretField = document.getElementById('modal-icd-client-secret');
    if (icdIdField && savedIcdId) icdIdField.value = savedIcdId;
    if (icdSecretField && savedIcdSecret) icdSecretField.value = savedIcdSecret;
    if ((savedIcdId || savedIcdSecret) && modalSaveKeyCheckbox) {
        modalSaveKeyCheckbox.checked = true;
    }
}

function toggleApiKeyVisibility() {
    const apiKeyInput = document.getElementById('modal-api-key');
    const toggleBtn = document.getElementById('toggle-api-key');
    
    if (apiKeyInput && toggleBtn) {
        const iconElement = toggleBtn.querySelector('i');
        
        if (apiKeyInput.type === 'password') {
            apiKeyInput.type = 'text';
            if (iconElement) {
                iconElement.classList.remove('bi-eye');
                iconElement.classList.add('bi-eye-slash');
            }
        } else {
            apiKeyInput.type = 'password';
            if (iconElement) {
                iconElement.classList.remove('bi-eye-slash');
                iconElement.classList.add('bi-eye');
            }
        }
    }
}

function handleSaveApiKeyChange(e) {
    if (e && e.target) {
        if (e.target.checked) {
            const apiKeyInput = document.getElementById('modal-api-key');
            if (apiKeyInput && apiKeyInput.value) {
                localStorage.setItem(API_KEY_STORAGE_KEY, apiKeyInput.value);
            }
        } else {
            // Use the local function to avoid reference errors
            localStorage.removeItem(API_KEY_STORAGE_KEY);
            const saveApiKeyCheckbox = document.getElementById('save-api-key');
            if (saveApiKeyCheckbox) {
                saveApiKeyCheckbox.checked = false;
            }
        }
    }
}

function clearSavedApiKey() {
    localStorage.removeItem(API_KEY_STORAGE_KEY);
    const saveApiKeyCheckbox = document.getElementById('save-api-key');
    if (saveApiKeyCheckbox) {
        saveApiKeyCheckbox.checked = false;
    }
    
    // Also clear the modal fields if they exist
    const modalSaveApiKeyCheckbox = document.getElementById('modal-save-api-key');
    if (modalSaveApiKeyCheckbox) {
        modalSaveApiKeyCheckbox.checked = false;
    }
}

/* File handling helpers */

// Robust response reader: handles non-JSON / timeout responses gracefully
// instead of crashing on `response.json()` with "Unexpected token".
async function readJson(response) {
    const text = await response.text();
    try {
        return JSON.parse(text);
    } catch (e) {
        const snippet = text.slice(0, 200).trim();
        throw new Error(response.ok
            ? `Unexpected non-JSON response: ${snippet}`
            : `Server error ${response.status}: ${snippet || 'request failed'}`);
    }
}

// Read the uploaded File into an ArrayBuffer.
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error('Could not read the file.'));
        reader.readAsArrayBuffer(file);
    });
}

/* File Upload — parsed entirely in the browser (no server round-trip). */
async function uploadFile(e) {
    e.preventDefault();

    const fileInput = document.getElementById('file');
    const progressBar = document.getElementById('upload-progress');

    if (!fileInput.files.length) {
        showAlert('upload-message', 'Please select a file to upload', 'danger');
        return;
    }

    // Show progress
    progressBar.classList.remove('hidden');
    progressBar.querySelector('.progress-bar').style.width = '0%';

    try {
        // Show spinner with parsing message
        showSpinner(true, 'Reading file...');

        const file = fileInput.files[0];
        const buffer = await readFileAsArrayBuffer(file);
        progressBar.querySelector('.progress-bar').style.width = '60%';

        const workbook = XLSX.read(buffer, { type: 'array' });

        const sheets = parseWorkbook(workbook);

        fileData = { filename: file.name, sheets: sheets };
        activeAnalysis = null;
        analyzedResult = null;
        icdResult = null;
        document.getElementById('analysis-recovery').classList.add('hidden');
        progressBar.querySelector('.progress-bar').style.width = '100%';

        showAlert('upload-message', 'File loaded successfully!', 'success');

        // Hide spinner
        showSpinner(false);

        // Move to next step after a short delay
        setTimeout(() => {
            displayFilePreview();
            goToStep(3);
        }, 600);

    } catch (error) {
        // Hide spinner
        showSpinner(false);

        // Show error
        showAlert('upload-message', `Error reading file: ${error.message}`, 'danger');
        progressBar.classList.add('hidden');
    }
}

/* Display File Preview */
function displayFilePreview() {
    // A new file invalidates any split staged against the previous one.
    cleanDataPreview?.reset();
    cleanDataResults?.reset();
    const sheetSelect = document.getElementById('sheet-select');
    sheetSelect.innerHTML = '';
    
    Object.keys(fileData.sheets).forEach(sheet => {
        const option = document.createElement('option');
        option.value = sheet;
        option.textContent = sheet;
        sheetSelect.appendChild(option);
    });
    
    sheetSelect.onchange = () => updatePreviewTable(sheetSelect.value);
    updatePreviewTable(sheetSelect.value);
}

function updatePreviewTable(sheetName) {
    const previewTable = document.getElementById('preview-table');
    const sheetData = fileData.sheets[sheetName];
    availableColumns = sheetData.columns;
    
    // Create Bootstrap table
    let tableHTML = '<table class="table table-striped table-bordered"><thead class="table-light"><tr>';
    sheetData.columns.forEach(column => {
        tableHTML += `<th>${escapeHtml(column)}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';
    
    // Only render the first 10 rows for preview (data now holds ALL rows).
    sheetData.data.slice(0, 10).forEach(row => {
        tableHTML += '<tr>';
        sheetData.columns.forEach(column => {
            const cell = row[column];
            tableHTML += `<td>${escapeHtml(cell === null || cell === undefined ? '' : cell)}</td>`;
        });
        tableHTML += '</tr>';
    });

    tableHTML += '</tbody></table>';
    if (sheetData.data.length > 10) {
        tableHTML += `<p class="text-muted small">Showing first 10 of ${sheetData.data.length} rows.</p>`;
    }
    previewTable.innerHTML = tableHTML;

    // Store the total row count for progress tracking
    const sheetSelect = document.getElementById('sheet-select');
    sheetSelect.dataset.rowCount = sheetData.data.length;
    
    // Update column selection in pattern detection
    const patternColumnSelect = document.getElementById('pattern-column');
    patternColumnSelect.innerHTML = '';
    availableColumns.forEach(column => {
        const option = document.createElement('option');
        option.value = column;
        option.textContent = column;
        patternColumnSelect.appendChild(option);
    });
    
    updateColumnConfigs();

    // Keep the Clean Data picker in step with the sheet being previewed.
    cleanDataPreview?.refreshColumns();
}

/* Column Configuration */
function updateColumnConfigs() {
    const columnConfigs = document.getElementById('column-configs');
    disposeTooltips(columnConfigs);
    columnConfigs.innerHTML = '';
    addColumnConfig();
}

/**
 * Append a column-analysis card.
 * @param {object} [preset] - values from an imported configuration:
 *   { columns: string[], outputColumnName, prompt, maxOutputTokens }
 */
function addColumnConfig(preset = null) {
    const columnConfigs = document.getElementById('column-configs');
    const configId = crypto.randomUUID();
    
    // Create column config card
    const configCard = document.createElement('div');
    configCard.className = 'column-selection-container mb-3';
    configCard.dataset.id = configId;
    
    // Create column selection
    const columnSelectionHTML = `
        <div class="row mb-2">
            <div class="col-md-10">
                <label class="form-label">Select Columns to Analyze</label>
                <div class="main-column-selector">
                    <select class="form-select main-column" name="column">
                        <option value="">Select a column</option>
                        ${availableColumns.map(col => `<option value="${escapeHtml(col)}">${escapeHtml(col)}</option>`).join('')}
                    </select>
                    <button type="button" class="btn btn-sm btn-success add-column-btn-small ms-2">
                        <i class="bi bi-plus"></i>
                    </button>
                </div>
                <div class="additional-columns">
                    <!-- Additional columns will be added here -->
                </div>
            </div>
            <div class="col-md-2 d-flex align-items-end">
                <button type="button" class="btn btn-sm btn-outline-danger remove-config-btn mb-2">
                    <i class="bi bi-trash"></i>
                </button>
            </div>
        </div>
        <div class="mb-3">
            <label class="form-label">Result Column Name
                <i class="bi bi-question-circle help-icon" data-bs-toggle="tooltip"
                   title="Name for the new column that will contain analysis results. E.g., 'Sentiment Score', 'Category', 'Translation'"></i>
            </label>
            <input type="text" class="form-control" name="output-column-name"
                   placeholder="E.g., Sentiment Score, Category, Translation...">
        </div>
        <div class="mb-3">
            <label class="form-label">Output length
                <select class="form-select" name="max-output-tokens">
                    <option value="256">Short labels (256 tokens)</option>
                    <option value="1024" selected>Standard (1,024 tokens)</option>
                    <option value="4096">Long translations (4,096 tokens)</option>
                </select>
            </label>
        </div>
        <div class="mb-3">
            <label class="form-label">Analysis Instructions
                <i class="bi bi-question-circle help-icon" data-bs-toggle="tooltip"
                   title="Specific instructions for analyzing this column. Be clear about what insights you want."></i>
            </label>
            <div class="input-group">
                <input type="text" class="form-control" name="prompt"
                       placeholder="Enter specific instructions for analyzing this column...">
                <button class="btn btn-outline-secondary instruction-template-btn" type="button">
                    <i class="bi bi-lightning"></i>
                </button>
            </div>
        </div>
    `;
    
    configCard.innerHTML = columnSelectionHTML;
    
    // Add event listeners for the buttons
    const addColumnBtn = configCard.querySelector('.add-column-btn-small');
    addColumnBtn.addEventListener('click', function() {
        addAdditionalColumn(configCard);
    });
    
    const removeConfigBtn = configCard.querySelector('.remove-config-btn');
    removeConfigBtn.addEventListener('click', function() {
        if (document.querySelectorAll('.column-selection-container').length > 1) {
            disposeTooltips(configCard);
            configCard.remove();
        } else {
            // Don't remove if it's the only config
            showAlert('analyze-message', 'You need at least one column configuration', 'warning');
        }
    });
    
    // Add event listener for the instruction template button
    const templateBtn = configCard.querySelector('.instruction-template-btn');
    templateBtn.addEventListener('click', function() {
        // Create dropdown menu for common templates
        const menu = document.createElement('div');
        menu.className = 'dropdown-menu p-2 shadow';
        menu.style.width = '300px';
        menu.innerHTML = `
            <h6 class="dropdown-header">Quick Templates</h6>
            <button class="dropdown-item" data-template="Analyze this column and provide a sentiment score (1-5) where 1 is very negative and 5 is very positive. Explain your reasoning.">
                Sentiment Analysis (1-5)
            </button>
            <button class="dropdown-item" data-template="Categorize this data into one of these types: [type1, type2, type3]. Explain your classification.">
                Categorization Template
            </button>
            <button class="dropdown-item" data-template="Identify any errors, inconsistencies, or unusual values in this data. If issues are found, suggest corrections.">
                Data Quality Check
            </button>
            <button class="dropdown-item" data-template="Extract key entities (people, organizations, locations, dates) mentioned in this text.">
                Entity Extraction
            </button>
            <button class="dropdown-item" data-template="Provide a concise 1-2 sentence summary of the key points in this text.">
                Text Summarization
            </button>
        `;
        
        // Position the menu
        menu.style.position = 'absolute';
        menu.style.zIndex = '1000';
        
        // Add event listeners to template items
        menu.querySelectorAll('.dropdown-item').forEach(item => {
            item.addEventListener('click', function() {
                const template = this.dataset.template;
                configCard.querySelector('input[name="prompt"]').value = template;
                document.body.removeChild(menu);
            });
        });
        
        // Add to document body, position, and show
        document.body.appendChild(menu);
        const rect = templateBtn.getBoundingClientRect();
        menu.style.top = `${rect.bottom + 5}px`;
        menu.style.left = `${rect.left - 250}px`;
        
        // Close when clicking outside
        document.addEventListener('click', function closeMenu(e) {
            if (!menu.contains(e.target) && e.target !== templateBtn) {
                if (document.body.contains(menu)) {
                    document.body.removeChild(menu);
                }
                document.removeEventListener('click', closeMenu);
            }
        });
    });
    
    if (preset) applyColumnPreset(configCard, preset);

    columnConfigs.appendChild(configCard);
    initTooltips(configCard);
    return configCard;
}

/* Fill a freshly built config card from an imported configuration. */
function applyColumnPreset(configCard, preset) {
    const [mainColumn, ...extras] = preset.columns || [];
    if (mainColumn) configCard.querySelector('select[name="column"]').value = mainColumn;
    extras.forEach(name => {
        addAdditionalColumn(configCard);
        const selects = configCard.querySelectorAll('.additional-column select');
        selects[selects.length - 1].value = name;
    });
    configCard.querySelector('input[name="output-column-name"]').value = preset.outputColumnName || '';
    configCard.querySelector('input[name="prompt"]').value = preset.prompt || '';
    configCard.querySelector('[name="max-output-tokens"]').value = String(preset.maxOutputTokens);
}

function addAdditionalColumn(configCard) {
    const additionalColumnsDiv = configCard.querySelector('.additional-columns');
    
    const columnDiv = document.createElement('div');
    columnDiv.className = 'additional-column d-flex align-items-center mt-2';
    columnDiv.innerHTML = `
        <select class="form-select form-select-sm" name="additional-column">
            <option value="">Select additional column</option>
            ${availableColumns.map(col => `<option value="${escapeHtml(col)}">${escapeHtml(col)}</option>`).join('')}
        </select>
        <button type="button" class="btn btn-sm btn-danger remove-column-btn-small ms-2">
            <i class="bi bi-dash"></i>
        </button>
    `;
    
    // Add event listener for remove button
    columnDiv.querySelector('.remove-column-btn-small').addEventListener('click', function() {
        columnDiv.remove();
    });
    
    additionalColumnsDiv.appendChild(columnDiv);
}

/* Read every column-analysis card off the page, in display order. */
function readColumnConfigs() {
    return Array.from(document.querySelectorAll('.column-selection-container')).map(config => {
        const mainColumn = config.querySelector('select[name="column"]').value;
        const additionalColumns = Array.from(config.querySelectorAll('.additional-column select'))
            .map(select => select.value)
            .filter(col => col !== "");
        const columns = [mainColumn, ...additionalColumns].filter(col => col !== "");
        const outputColumnName = config.querySelector('input[name="output-column-name"]').value.trim();

        return {
            column: mainColumn,
            columns: columns,
            prompt: config.querySelector('input[name="prompt"]').value,
            maxOutputTokens: Number(config.querySelector('[name="max-output-tokens"]').value),
            outputColumnName: outputColumnName || null, // Use null if not provided
            id: config.dataset.id || Date.now().toString()
        };
    });
}

/* Configuration Export / Import */
// "analysis" pluralizes to "analyses", not "analysises".
function pluralizeAnalyses(count) {
    return `${count} column ${count === 1 ? 'analysis' : 'analyses'}`;
}

function exportAnalysisConfig() {
    const columns = readColumnConfigs();
    // An empty form would export a file that restores nothing.
    if (!columns.some(cfg => cfg.columns.length || cfg.prompt.trim())) {
        showAlert('analyze-message', 'Nothing to export yet — configure at least one column first.', 'warning');
        return;
    }
    const doc = serializeConfig({
        generalInstructions: document.getElementById('general-instructions').value,
        sheetName: document.getElementById('sheet-select').value,
        columns
    });
    const url = URL.createObjectURL(new Blob([JSON.stringify(doc, null, 2)], { type: 'application/json' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = configFilename();
    link.click();
    URL.revokeObjectURL(url);
    showAlert('analyze-message', `Configuration exported (${pluralizeAnalyses(doc.columns.length)}).`, 'success');
}

async function importAnalysisConfig(event) {
    const input = event.target;
    const file = input.files && input.files[0];
    if (!file) return;
    // Reset first so re-picking the same file fires change again.
    input.value = '';

    let imported;
    try {
        imported = deserializeConfig(JSON.parse(await file.text()), availableColumns);
    } catch (error) {
        const message = error instanceof SyntaxError
            ? 'That file is not valid JSON.'
            : error.message;
        showAlert('analyze-message', `Could not import configuration: ${message}`, 'danger');
        return;
    }

    document.getElementById('general-instructions').value = imported.generalInstructions;
    const container = document.getElementById('column-configs');
    container.innerHTML = '';
    imported.columns.forEach(preset => addColumnConfig(preset));

    const applied = `Imported ${pluralizeAnalyses(imported.columns.length)}.`;
    if (imported.warnings.length) {
        showAlert('analyze-message', `${applied} ${imported.warnings.join(' ')}`, 'warning');
    } else {
        showAlert('analyze-message', `${applied} General instructions were restored too.`, 'success');
    }
}

/* Pattern Detection */
function togglePatternDetection() {
    const patternDetection = document.getElementById('pattern-detection');
    const toggleBtn = document.getElementById('toggle-pattern-btn');
    
    if (patternDetection.classList.contains('hidden')) {
        patternDetection.classList.remove('hidden');
        toggleBtn.innerHTML = '<i class="bi bi-eye-slash"></i> Hide Pattern Detection';
    } else {
        patternDetection.classList.add('hidden');
        toggleBtn.innerHTML = '<i class="bi bi-search"></i> Detect Patterns';
    }
}

async function detectPatterns() {
    const column = document.getElementById('pattern-column').value;
    const numCategories = document.getElementById('num-categories').value;
    const patternPrompt = document.getElementById('pattern-prompt').value;
    const selectedSheet = document.getElementById('sheet-select').value;
    
    if (!column || !patternPrompt) {
        showAlert('pattern-message', 'Please fill in all required fields for pattern detection', 'warning');
        return;
    }
    
    // Show loading spinner
    showSpinner(true, 'Detecting patterns in your data...');

    try {
        // Sample non-empty values for this column from the in-memory data.
        const rows = (fileData.sheets[selectedSheet] || {}).data || [];
        const sampleValues = rows
            .map(r => r[column])
            .filter(v => v !== null && v !== undefined && String(v).trim() !== '')
            .slice(0, 100);

        const response = await fetch('/detect_patterns', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                ...getLLMConfig(),
                column: column,
                patternPrompt: patternPrompt,
                numCategories: parseInt(numCategories),
                sampleValues: sampleValues
            })
        });

        const result = await readJson(response);
        if (!response.ok) {
            throw new Error(result.error || `HTTP error! status: ${response.status}`);
        }
        if (result.error) {
            throw new Error(result.error);
        }
        
        // Parse the JSON string returned
        const categoriesData = JSON.parse(result.result);
        displayCategories(categoriesData);
        
        showAlert('pattern-message', 'Categories detected successfully!', 'success');
    } catch (error) {
        showAlert('pattern-message', `Error detecting patterns: ${error.message}`, 'danger');
    } finally {
        showSpinner(false);
    }
}

function displayCategories(data) {
    const categoriesList = document.getElementById('categories-list');
    const explanation = document.getElementById('explanation');
    
    // Generate category items
    categoriesList.innerHTML = data.categories.map(category => 
        `<span class="category-item">${escapeHtml(category)}</span>`
    ).join(' ');
    
    // Display explanation
    explanation.textContent = data.explanation;
    
    // Show results section
    document.getElementById('pattern-results').classList.remove('hidden');
}

function copyCategoriesToClipboard() {
    const categoriesList = document.getElementById('categories-list');
    const categories = Array.from(categoriesList.querySelectorAll('.category-item'))
        .map(item => item.textContent)
        .join(', ');
        
    // Create temporary element for copying
    const tempInput = document.createElement('textarea');
    tempInput.value = categories;
    document.body.appendChild(tempInput);
    tempInput.select();
    document.execCommand('copy');
    document.body.removeChild(tempInput);
    
    showAlert('pattern-message', 'Categories copied to clipboard!', 'success');
}

function useCategoriesInPrompt() {
    const categories = Array.from(document.querySelectorAll('.category-item'))
        .map(item => item.textContent)
        .join(', ');
        
    // Find an empty config or create a new one
    const configs = document.querySelectorAll('.column-selection-container');
    let targetConfig = null;
    
    // Find an empty config to use
    for (const config of configs) {
        const promptInput = config.querySelector('input[name="prompt"]');
        if (!promptInput.value) {
            targetConfig = config;
            break;
        }
    }
    
    // If no empty config found, create a new one
    if (!targetConfig) {
        addColumnConfig();
        targetConfig = document.querySelector('.column-selection-container:last-child');
    }
    
    // Set the same column as was used for pattern detection
    const columnSelect = targetConfig.querySelector('select[name="column"]');
    columnSelect.value = document.getElementById('pattern-column').value;
    
    // Set the prompt with categories
    const promptInput = targetConfig.querySelector('input[name="prompt"]');
    promptInput.value = `Categorize each item into one of these categories: ${categories}. Explain the reason for your categorization.`;
    
    // Hide pattern detection
    togglePatternDetection();
    
    showAlert('analyze-message', 'Categories added to analysis prompt!', 'success');
}

/* Column Analysis */
async function analyzeColumns(isTestRun = false) {
    if (activeAnalysis?.running) return;
    // Clear previous results before starting a new analysis, but only if elements exist
    const resultPreview = document.getElementById('result-preview');
    const chartContent = document.getElementById('chart-content');
    const chartColumnSelect = document.getElementById('chart-column-select');
    
    if (resultPreview) resultPreview.innerHTML = '<div class="alert alert-info">Processing your data...</div>';
    if (chartContent) chartContent.innerHTML = '<div class="alert alert-info">Chart will appear after analysis is complete.</div>';
    if (chartColumnSelect) chartColumnSelect.innerHTML = '';
    
    const generalInstructions = document.getElementById('general-instructions').value;
    const selectedSheet = document.getElementById('sheet-select').value;
    
    // Validate inputs
    // Credentials may also come from the server's environment, so only
    // warn when nothing is configured locally; the server has the final say.
    if (!hasLocalCredentials()) {
        showAlert('analyze-message',
            'No AI credentials configured locally — attempting to use the server configuration. Open API Settings if this fails.',
            'info');
    }
    
    // Get column configurations
    const columnConfigs = readColumnConfigs().filter(config => config.column && config.prompt);
    
    // Check if we have valid configurations
    if (columnConfigs.length === 0) {
        showAlert('analyze-message', 'Please configure at least one column for analysis', 'danger');
        return;
    }
    
    // Source rows from the in-memory sheet.
    const sheet = fileData.sheets[selectedSheet];
    if (!sheet || !sheet.data) {
        showAlert('analyze-message', 'No data available for the selected sheet', 'danger');
        return;
    }
    const allRows = sheet.data;
    if (!allRows.length) { showAlert('analyze-message', 'This sheet has no data rows.', 'warning'); return; }
    const rowCount = isTestRun ? testRowCount('test-rows', 5, allRows.length) : allRows.length;

    // Resolve each output column name once, ensuring uniqueness against
    // existing columns (previously done per-row on the server).
    const existing = new Set(sheet.columns);
    columnConfigs.forEach(cfg => {
        let name = cfg.outputColumnName || `${cfg.column}_analysis_${cfg.id}`;
        const base = name;
        let counter = 1;
        while (existing.has(name)) { name = `${base}_${counter++}`; }
        existing.add(name);
        cfg.resolvedName = name;
    });

    const serverConfigs = columnConfigs.map(c => ({
        id: c.id, column: c.column, prompt: c.prompt, outputColumnName: c.resolvedName, maxOutputTokens: c.maxOutputTokens
    }));
    const providerConfig = getLLMConfig();
    const outColumns = sheet.columns.concat(columnConfigs.map(c => c.resolvedName));
    activeAnalysis = new BatchRun({
        rows: allRows.slice(0, rowCount), batchSize: ANALYZE_BATCH_SIZE,
        processBatch: async (batch, start) => {
            const rows = batch.map((row, offset) => ({ rowIndex: start + offset,
                inputs: Object.fromEntries(columnConfigs.map(cfg => [cfg.id, buildAnalysisInput(row, cfg)])) }));
            const response = await fetch('/analyze_batch', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ...providerConfig, generalInstructions,
                    isFirstBatch: start === 0, configs: serverConfigs, rows })
            });
            const result = await readJson(response);
            if (!response.ok || result.error) throw new Error(result.error || 'Analysis failed.');
            const values = new Map((result.results || []).map(r => [r.rowIndex, r.values]));
            if (rows.some(row => !values.has(row.rowIndex))) throw new Error('Incomplete batch response. Resume to retry.');
            return { rows: batch.map((row, i) => ({ ...row, ...values.get(start + i) })), errors: result.errors };
        },
        onProgress: run => {
            analyzedResult = { sheetName: selectedSheet, columns: outColumns,
                data: run.results, partial: !run.complete, outputColumns: serverConfigs.map(c => c.outputColumnName) };
            updateProgress(Math.round(run.cursor / rowCount * 100), run.cursor, rowCount,
                `Analyzed ${run.cursor} of ${rowCount} rows`);
        }
    });
    activeAnalysis.isTestRun = isTestRun;
    analyzedResult = null;
    // A new run invalidates any split staged against the previous results.
    cleanDataResults?.reset();
    await continueAnalysis();
}

let activeAnalysis = null;
async function continueAnalysis() {
    const run = activeAnalysis;
    if (!run || run.running) return;
    document.querySelectorAll('.test-run-preview-panel').forEach(panel => panel.remove());
    showSpinner(true, 'Analyzing rows…', true);
    const stopButton = document.getElementById('stop-analysis-btn');
    stopButton.classList.remove('hidden');
    stopButton.disabled = false;
    stopButton.textContent = 'Stop after current batch';
    stopButton.onclick = () => { run.stop(); stopButton.disabled = true; stopButton.textContent = 'Stopping after current batch…'; };
    document.getElementById('analysis-recovery').classList.add('hidden');
    let failure = null;
    try { await run.run(); } catch (error) { failure = error; }
    finally { showSpinner(false); stopButton.classList.add('hidden'); }
    const message = `${run.complete ? 'Analysis complete' : 'Analysis paused'}: ${run.cursor} of ${run.rows.length} rows. ${run.errors} cell(s) had errors.`;
    showAlert('analyze-message', failure ? `${message} ${failure.message}` : message,
        failure ? 'danger' : run.complete ? 'success' : 'info');
    const recovery = document.getElementById('analysis-recovery');
    recovery.classList.toggle('hidden', run.complete);
    document.getElementById('partial-download-btn').disabled = !run.cursor;
    document.getElementById('resume-analysis-btn').onclick = continueAnalysis;
    document.getElementById('partial-download-btn').onclick = downloadAnalyzedFile;
    if (!analyzedResult) return;
    document.getElementById('result-message').textContent = message;
    const downloadLink = document.getElementById('download-link');
    downloadLink.href = '#';
    downloadLink.onclick = event => { event.preventDefault(); downloadAnalyzedFile(); };
    if (run.isTestRun && run.complete) previewAnalyzedData(run.results, true);
    else if (run.complete) {
        resultsUI.load({ sheets: { [analyzedResult.sheetName]: analyzedResult } });
        goToStep(5);
    }
}

/* Jev Classification (appMode === 'jev')
 *
 * Unlike the LLM path, a Jev result column is a typed *question*: the answer is
 * always one of the options configured here, so the column holds a clean label
 * rather than prose to be tidied afterwards. All questions for a row travel in
 * one request, which is why the row -- not the cell -- is the unit of work.
 */

// Option lists people reach for most often, so a first run needs no typing.
const JEV_QUESTION_TEMPLATES = {
    sentiment: {
        outputColumnName: 'Sentiment',
        questionType: 'choice',
        instructions: 'Classify the overall sentiment expressed in the text.',
        options: ['Positive', 'Neutral', 'Negative', 'Mixed']
    },
    urgency: {
        outputColumnName: 'Urgency',
        questionType: 'score',
        instructions: 'Rate how urgently this requires a response, independently of how it is worded.',
        options: [
            'Low: general comment, appreciation or non-urgent suggestion',
            'Medium: a question or concern needing a reply, but no immediate danger',
            'High: a serious problem needing prompt action',
            'Critical: immediate risk to safety requiring escalation now'
        ]
    },
    actionable: {
        outputColumnName: 'Needs follow-up',
        questionType: 'noul',
        instructions: 'Does this entry require someone to take a follow-up action?',
        options: ['Yes', 'No']
    }
};

/** Fill the source-column checkboxes from the sheet being analyzed. */
function populateJevConfig() {
    const sheet = currentSheet();
    const columns = (sheet && sheet.columns) ? sheet.columns : (availableColumns || []);
    const container = document.getElementById('jev-source-columns');
    if (!container) return;

    // Preserve the current selection across re-entry into step 4.
    const checked = new Set(readJevSourceColumns());
    container.innerHTML = columns.map((col, i) => `
        <div class="form-check">
            <input class="form-check-input jev-source-column" type="checkbox"
                   id="jev-src-${i}" value="${escapeHtml(col)}"${checked.has(col) ? ' checked' : ''}>
            <label class="form-check-label" for="jev-src-${i}">${escapeHtml(col)}</label>
        </div>`).join('');

    // Default to the first column so a quick run needs no extra clicks.
    if (!checked.size) {
        const first = container.querySelector('.jev-source-column');
        if (first) first.checked = true;
    }
    if (!document.querySelector('.jev-question-container')) addJevQuestion();
}

function readJevSourceColumns() {
    return Array.from(document.querySelectorAll('.jev-source-column:checked'))
        .map(input => input.value);
}

/**
 * Append a Jev question card.
 * @param {object} [preset] - { outputColumnName, questionType, instructions, options }
 */
function addJevQuestion(preset = null) {
    const container = document.getElementById('jev-question-configs');
    if (!container) return null;
    const card = document.createElement('div');
    card.className = 'column-selection-container jev-question-container mb-3';
    card.innerHTML = `
        <div class="row mb-2">
            <div class="col-md-10">
                <label class="form-label">Result Column Name
                    <i class="bi bi-question-circle help-icon" data-bs-toggle="tooltip"
                       title="Name of the new column that will hold the chosen label."></i>
                </label>
                <input type="text" class="form-control" name="jev-output-column-name"
                       placeholder="E.g., Feedback type, Urgency, Sector...">
            </div>
            <div class="col-md-2 d-flex align-items-end">
                <button type="button" class="btn btn-sm btn-outline-danger jev-remove-btn mb-2">
                    <i class="bi bi-trash"></i>
                </button>
            </div>
        </div>
        <div class="mb-3">
            <label class="form-label">Question type
                <i class="bi bi-question-circle help-icon" data-bs-toggle="tooltip"
                   title="Choice picks one label from your list. Score rates against ordered levels. Yes/No answers a single question."></i>
            </label>
            <select class="form-select" name="jev-question-type">
                <option value="choice">Choice — pick one label from a list</option>
                <option value="score">Score — rate against ordered levels</option>
                <option value="noul">Yes / No — answer a single question</option>
            </select>
        </div>
        <div class="mb-3">
            <label class="form-label">Instructions
                <i class="bi bi-question-circle help-icon" data-bs-toggle="tooltip"
                   title="Tell Jev what to decide. Keep it narrow and specific — one judgement per result column."></i>
            </label>
            <textarea class="form-control" name="jev-instructions" rows="2"
                      placeholder="E.g., Determine the dominant type of feedback."></textarea>
        </div>
        <div class="mb-2">
            <label class="form-label jev-options-label">Options <span class="text-muted">(one per line)</span>
                <i class="bi bi-question-circle help-icon" data-bs-toggle="tooltip"
                   title="Jev can only answer with one of these. Paste your existing codebook here."></i>
            </label>
            <textarea class="form-control jev-options" name="jev-options" rows="5"
                      placeholder="Question&#10;Suggestion&#10;Complaint"></textarea>
            <div class="form-text jev-options-help"></div>
        </div>
        <div class="d-flex flex-wrap gap-2">
            <select class="form-select form-select-sm w-auto jev-template-select">
                <option value="">Start from a template…</option>
                <option value="sentiment">Sentiment (Choice)</option>
                <option value="urgency">Urgency (Score)</option>
                <option value="actionable">Needs follow-up (Yes/No)</option>
            </select>
            <button type="button" class="btn btn-sm btn-outline-info jev-use-categories-btn"
                    data-bs-toggle="tooltip"
                    title="Fill the options from the categories found by Detect Patterns in AI Analysis mode.">
                <i class="bi bi-magic"></i> Use detected categories
            </button>
        </div>
    `;

    const typeSelect = card.querySelector('[name="jev-question-type"]');
    const optionsField = card.querySelector('.jev-options');
    const optionsLabel = card.querySelector('.jev-options-label');
    const optionsHelp = card.querySelector('.jev-options-help');

    // Each question type wants a different shape of list, so the same textarea
    // is relabelled rather than shown as three separate controls.
    function syncType() {
        const type = typeSelect.value;
        const count = parseOptions(optionsField.value).length;
        if (type === 'score') {
            optionsLabel.innerHTML = 'Levels <span class="text-muted">(one per line, lowest first)</span>';
            optionsHelp.textContent = `Between 2 and 10 ordered levels, lowest first. Describe each one — "Low: a general comment" works better than "Low". ${count} entered.`;
            optionsField.placeholder = 'Low: a general comment\nMedium: needs a reply\nHigh: needs prompt action';
        } else if (type === 'noul') {
            optionsLabel.innerHTML = 'Labels <span class="text-muted">(optional: yes label, then no label)</span>';
            optionsHelp.textContent = 'Jev answers with a probability. Leave blank to write "Yes"/"No", or give two lines to use your own wording.';
            optionsField.placeholder = 'Yes\nNo';
        } else {
            optionsLabel.innerHTML = 'Options <span class="text-muted">(one per line)</span>';
            optionsHelp.textContent = `Jev can only answer with one of these. At least 2, at most 255. ${count} entered.`;
            optionsField.placeholder = 'Question\nSuggestion\nComplaint';
        }
    }
    typeSelect.addEventListener('change', syncType);
    optionsField.addEventListener('input', syncType);

    card.querySelector('.jev-remove-btn').addEventListener('click', () => {
        if (document.querySelectorAll('.jev-question-container').length > 1) {
            disposeTooltips(card);
            card.remove();
        } else {
            showAlert('jev-message', 'You need at least one result column.', 'warning');
        }
    });

    card.querySelector('.jev-template-select').addEventListener('change', function () {
        const template = JEV_QUESTION_TEMPLATES[this.value];
        if (template) applyJevPreset(card, template);
        this.value = '';
    });

    // Detect Patterns lives in AI Analysis mode but produces exactly the kind
    // of controlled list a Choice question needs, so it is reusable here.
    card.querySelector('.jev-use-categories-btn').addEventListener('click', () => {
        const categories = Array.from(document.querySelectorAll('#categories-list .badge'))
            .map(el => el.textContent.trim()).filter(Boolean);
        if (!categories.length) {
            showAlert('jev-message', 'No detected categories yet. Run Detect Patterns in AI Analysis mode first.', 'warning');
            return;
        }
        optionsField.value = formatOptions(categories);
        typeSelect.value = 'choice';
        syncType();
        showAlert('jev-message', `Filled ${categories.length} options from the detected categories.`, 'success');
    });

    if (preset) applyJevPreset(card, preset);
    syncType();
    container.appendChild(card);
    initTooltips(card);
    return card;
}

/* Fill a Jev question card from a template or an imported configuration. */
function applyJevPreset(card, preset) {
    card.querySelector('[name="jev-output-column-name"]').value = preset.outputColumnName || '';
    card.querySelector('[name="jev-question-type"]').value = preset.questionType || DEFAULT_QUESTION_TYPE;
    card.querySelector('[name="jev-instructions"]').value = preset.instructions || '';
    card.querySelector('.jev-options').value = formatOptions(preset.options);
    card.querySelector('.jev-options').dispatchEvent(new Event('input'));
}

/* Read every Jev question card off the page, in display order. */
function readJevQuestions() {
    return Array.from(document.querySelectorAll('.jev-question-container')).map(card => ({
        outputColumnName: card.querySelector('[name="jev-output-column-name"]').value.trim(),
        questionType: card.querySelector('[name="jev-question-type"]').value,
        instructions: card.querySelector('[name="jev-instructions"]').value,
        options: parseOptions(card.querySelector('.jev-options').value)
    }));
}

function exportJevConfig() {
    const questions = readJevQuestions();
    if (!questions.some(q => q.outputColumnName || q.instructions.trim() || q.options.length)) {
        showAlert('jev-message', 'Nothing to export yet — configure at least one result column first.', 'warning');
        return;
    }
    const doc = serializeJevConfig({
        sourceColumns: readJevSourceColumns(),
        includeConfidence: document.getElementById('jev-include-confidence')?.checked,
        sheetName: document.getElementById('sheet-select').value,
        questions
    });
    const url = URL.createObjectURL(new Blob([JSON.stringify(doc, null, 2)], { type: 'application/json' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = jevConfigFilename();
    link.click();
    URL.revokeObjectURL(url);
    showAlert('jev-message', `Configuration exported (${doc.questions.length} result column(s)).`, 'success');
}

async function importJevConfig(event) {
    const input = event.target;
    const file = input.files && input.files[0];
    if (!file) return;
    // Reset first so re-picking the same file fires change again.
    input.value = '';

    let imported;
    try {
        imported = deserializeJevConfig(JSON.parse(await file.text()), availableColumns);
    } catch (error) {
        const message = error instanceof SyntaxError ? 'That file is not valid JSON.' : error.message;
        showAlert('jev-message', `Could not import configuration: ${message}`, 'danger');
        return;
    }

    const selected = new Set(imported.sourceColumns);
    document.querySelectorAll('.jev-source-column').forEach(input => {
        input.checked = selected.has(input.value);
    });
    const confidence = document.getElementById('jev-include-confidence');
    if (confidence) confidence.checked = imported.includeConfidence;

    const container = document.getElementById('jev-question-configs');
    disposeTooltips(container);
    container.innerHTML = '';
    imported.questions.forEach(preset => addJevQuestion(preset));

    const applied = `Imported ${imported.questions.length} result column(s).`;
    showAlert('jev-message', imported.warnings.length ? `${applied} ${imported.warnings.join(' ')}` : applied,
        imported.warnings.length ? 'warning' : 'success');
}

let activeJevRun = null;

async function runJevClassification(isTestRun = false) {
    if (activeJevRun?.running) return;

    if (!hasJevCredentials()) {
        showAlert('jev-message',
            'No Jev API key configured locally — attempting to use the server configuration. Open API Settings if this fails.',
            'info');
    }

    const sourceColumns = readJevSourceColumns();
    if (!sourceColumns.length) {
        showAlert('jev-message', 'Select at least one column to send to Jev.', 'danger');
        return;
    }

    const questions = readJevQuestions();
    if (!questions.length) {
        showAlert('jev-message', 'Configure at least one result column.', 'danger');
        return;
    }
    // Validate here rather than at the API so a long option list is corrected
    // before any request is spent.
    for (const question of questions) {
        const problem = validateQuestion(question);
        if (problem) { showAlert('jev-message', problem, 'danger'); return; }
    }
    const names = questions.map(q => q.outputColumnName);
    if (new Set(names).size !== names.length) {
        showAlert('jev-message', 'Result column names must be unique.', 'danger');
        return;
    }

    const selectedSheet = document.getElementById('sheet-select').value;
    const sheet = fileData.sheets[selectedSheet];
    if (!sheet || !sheet.data) {
        showAlert('jev-message', 'No data available for the selected sheet', 'danger');
        return;
    }
    const allRows = sheet.data;
    if (!allRows.length) { showAlert('jev-message', 'This sheet has no data rows.', 'warning'); return; }
    const rowCount = isTestRun ? testRowCount('jev-test-rows', 5, allRows.length) : allRows.length;

    // Resolve output names against the sheet's existing columns, as the LLM
    // path does, so a question never silently overwrites a source column.
    const existing = new Set(sheet.columns);
    questions.forEach(question => {
        let name = question.outputColumnName;
        const base = name;
        let counter = 1;
        while (existing.has(name)) { name = `${base}_${counter++}`; }
        existing.add(name);
        question.resolvedName = name;
    });

    const includeConfidence = !!document.getElementById('jev-include-confidence')?.checked;
    const serverConfigs = questions.map(q => ({
        outputColumnName: q.resolvedName,
        questionType: q.questionType,
        instructions: q.instructions,
        options: q.options
    }));

    // Confidence columns are appended next to the value they describe.
    const outColumns = sheet.columns.slice();
    const outputColumns = [];
    questions.forEach(q => {
        outColumns.push(q.resolvedName);
        outputColumns.push(q.resolvedName);
        if (includeConfidence) {
            outColumns.push(`${q.resolvedName}__confidence`);
            outputColumns.push(`${q.resolvedName}__confidence`);
            if (q.questionType === 'score' || q.questionType === 'noul') {
                outColumns.push(`${q.resolvedName}__score`);
                outputColumns.push(`${q.resolvedName}__score`);
            }
        }
    });

    const jevConfig = getJevConfig();
    activeJevRun = new BatchRun({
        rows: allRows.slice(0, rowCount), batchSize: JEV_BATCH_SIZE,
        processBatch: async (batch, start) => {
            const rows = batch.map((row, offset) => ({
                rowIndex: start + offset,
                state: buildJevState(row, sourceColumns)
            }));
            const response = await fetch('/analyze_batch_jev', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ...jevConfig, includeConfidence,
                    isFirstBatch: start === 0, configs: serverConfigs, rows })
            });
            const result = await readJson(response);
            if (!response.ok || result.error) throw new Error(result.error || 'Classification failed.');
            const values = new Map((result.results || []).map(r => [r.rowIndex, r.values]));
            if (rows.some(row => !values.has(row.rowIndex))) throw new Error('Incomplete batch response. Resume to retry.');
            return { rows: batch.map((row, i) => ({ ...row, ...values.get(start + i) })), errors: result.errors };
        },
        onProgress: run => {
            analyzedResult = { sheetName: selectedSheet, columns: outColumns,
                data: run.results, partial: !run.complete, outputColumns };
            updateProgress(Math.round(run.cursor / rowCount * 100), run.cursor, rowCount,
                `Classified ${run.cursor} of ${rowCount} rows`);
        }
    });
    activeJevRun.isTestRun = isTestRun;
    analyzedResult = null;
    // A new run invalidates any split staged against the previous results.
    cleanDataResults?.reset();
    await continueJevRun();
}

async function continueJevRun() {
    const run = activeJevRun;
    if (!run || run.running) return;
    document.querySelectorAll('.test-run-preview-panel').forEach(panel => panel.remove());
    showSpinner(true, 'Classifying rows…', true);
    const stopButton = document.getElementById('stop-analysis-btn');
    stopButton.classList.remove('hidden');
    stopButton.disabled = false;
    stopButton.textContent = 'Stop after current batch';
    stopButton.onclick = () => { run.stop(); stopButton.disabled = true; stopButton.textContent = 'Stopping after current batch…'; };
    document.getElementById('jev-recovery').classList.add('hidden');
    let failure = null;
    try { await run.run(); } catch (error) { failure = error; }
    finally { showSpinner(false); stopButton.classList.add('hidden'); }
    const message = `${run.complete ? 'Classification complete' : 'Classification paused'}: ${run.cursor} of ${run.rows.length} rows. ${run.errors} row(s) had errors.`;
    showAlert('jev-message', failure ? `${message} ${failure.message}` : message,
        failure ? 'danger' : run.complete ? 'success' : 'info');
    const recovery = document.getElementById('jev-recovery');
    recovery.classList.toggle('hidden', run.complete);
    document.getElementById('jev-partial-download-btn').disabled = !run.cursor;
    if (!analyzedResult) return;
    document.getElementById('result-message').textContent = message;
    const downloadLink = document.getElementById('download-link');
    downloadLink.href = '#';
    downloadLink.onclick = event => { event.preventDefault(); downloadAnalyzedFile(); };
    if (run.isTestRun && run.complete) previewAnalyzedData(run.results, true);
    else if (run.complete) {
        resultsUI.load({ sheets: { [analyzedResult.sheetName]: analyzedResult } });
        goToStep(5);
    }
}

function previewAnalyzedData(analyzedRows, isTestRun = false) {
    try {
        // Clear any existing data first to ensure we display fresh results
        const resultPreview = document.getElementById('result-preview');
        const chartContent = document.getElementById('chart-content');
        const chartColumnSelect = document.getElementById('chart-column-select');
        
        // Remove any existing test run panels
        const existingPanels = document.querySelectorAll('.test-run-preview-panel');
        existingPanels.forEach(panel => {
            try {
                document.body.removeChild(panel);
            } catch (e) {
                console.error('Error removing panel:', e);
            }
        });
        
        if (resultPreview) resultPreview.innerHTML = '<div class="alert alert-info">Loading analysis results...</div>';
        if (chartContent) chartContent.innerHTML = '<div class="alert alert-info">Preparing chart visualization...</div>';
        if (chartColumnSelect) chartColumnSelect.innerHTML = '';

        // Render directly from the in-memory analyzed rows (no server fetch).
        const jsonData = analyzedRows || [];
        const previewData = jsonData.slice(0, 10);
        const headers = (analyzedResult && analyzedResult.columns) || Object.keys(previewData[0] || {});

        // Create the table HTML
        let tableHTML = `
            <table class="table table-striped table-bordered table-hover">
                <thead class="table-light">
                    <tr>
                        ${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}
                    </tr>
                </thead>
                <tbody>
        `;

        // Add the data rows
        previewData.forEach(row => {
            tableHTML += '<tr>';
            headers.forEach(header => {
                const cell = row[header];
                tableHTML += `<td>${escapeHtml(cell === null || cell === undefined ? '' : cell)}</td>`;
            });
            tableHTML += '</tr>';
        });

        tableHTML += '</tbody></table>';

        // Update the preview
        if (resultPreview) resultPreview.innerHTML = tableHTML;

        // Set up chart visualization
        setupChartVisualization(jsonData, headers);

        // If this is a test run, create a floating preview panel
        if (isTestRun) {
            // Create a floating preview panel for test run results
            const previewPanel = document.createElement('div');
            previewPanel.className = 'card position-fixed bottom-0 end-0 mb-4 me-4 shadow test-run-preview-panel';
            previewPanel.style.width = '90%';
            previewPanel.style.maxWidth = '800px';
            previewPanel.style.maxHeight = '70vh';
            previewPanel.style.overflow = 'auto';
            previewPanel.style.zIndex = '1050';
            previewPanel.innerHTML = `
                <div class="card-header bg-success text-white d-flex justify-content-between align-items-center">
                    <span><i class="bi bi-lightning"></i> Test Run Results (${previewData.length} Rows)</span>
                    <div>
                        <button class="btn btn-sm btn-outline-light me-2" id="go-to-results-btn">
                            <i class="bi bi-arrows-fullscreen"></i> Full View
                        </button>
                        <button class="btn btn-sm btn-outline-light" id="close-preview-btn">
                            <i class="bi bi-x-lg"></i>
                        </button>
                    </div>
                </div>
                <div class="card-body">
                    <div class="test-result-preview overflow-auto" style="max-height: 50vh;">
                        ${tableHTML}
                    </div>
                    <div class="d-flex justify-content-end mt-3">
                        <button type="button" class="btn btn-sm btn-success" id="download-test-results-btn">
                            <i class="bi bi-download"></i> Download Test Results
                        </button>
                    </div>
                </div>
            `;
            document.body.appendChild(previewPanel);

            // Add event listeners to the preview panel buttons
            document.getElementById('close-preview-btn').addEventListener('click', () => {
                document.body.removeChild(previewPanel);
            });

            document.getElementById('go-to-results-btn').addEventListener('click', () => {
                document.body.removeChild(previewPanel);
                const dataObject = { sheets: { [analyzedResult.sheetName]: analyzedResult } };
                resultsUI.load(dataObject);
                goToStep(5);
            });

            document.getElementById('download-test-results-btn').addEventListener('click', downloadAnalyzedFile);
        }
    } catch (error) {
        console.error('Error previewing analyzed data:', error);
    }
}

function setupChartVisualization(data, headers) {
    // Get references to chart elements
    const chartContent = document.getElementById('chart-content');
    const chartColumnSelect = document.getElementById('chart-column-select');
    const chartContainer = chartContent ? chartContent.querySelector('.chart-container') : null;
    
    // Clear the select options but keep the chart container structure
    if (chartColumnSelect) chartColumnSelect.innerHTML = '';
    
    // Filter headers to only include analysis columns
    const analysisColumns = headers.filter(h => h.includes('_analysis_'));
    
    if (analysisColumns.length === 0) {
        if (chartContainer) chartContainer.innerHTML = '<div class="alert alert-info">No analysis columns found for visualization.</div>';
        return;
    }
    
    // Always recreate the canvas element to ensure a fresh chart
    if (chartContainer) {
        // Remove existing canvas if it exists
        const existingCanvas = chartContainer.querySelector('canvas');
        if (existingCanvas) {
            chartContainer.removeChild(existingCanvas);
        }
        
        // Create a new canvas element
        const canvas = document.createElement('canvas');
        canvas.id = 'result-chart';
        chartContainer.appendChild(canvas);
    }
    
    // Already have a reference to chartColumnSelect
    // Just ensure it exists before manipulating it
    if (chartColumnSelect) {
        // Clear again to be safe
        chartColumnSelect.innerHTML = '';
        
        analysisColumns.forEach(column => {
            const option = document.createElement('option');
            option.value = column;
            
            // Try to get a more user-friendly name
            const originalColName = column.split('_analysis_')[0];
            option.textContent = `Analysis of ${originalColName}`;
            
            chartColumnSelect.appendChild(option);
        });
    }
    
    // Set up event listener for chart selection
    if (chartColumnSelect) {
        chartColumnSelect.addEventListener('change', function() {
            const selectedColumn = this.value;
            if (selectedColumn) {
                // Extract data for the selected column
                const chartData = data.map(row => row[selectedColumn]);
                initResultChart(chartData, data.map(row => ''), selectedColumn);
            }
        });
        
        // Initialize with first column
        if (analysisColumns.length > 0) {
            chartColumnSelect.value = analysisColumns[0];
            const chartData = data.map(row => row[analysisColumns[0]]);
            initResultChart(chartData, data.map(row => ''), analysisColumns[0]);
        }
    }
}

/* Document Ready */
document.addEventListener('DOMContentLoaded', function() {
    // Load saved API key
    loadSavedApiKey();
    
    // Initialize tooltips
    initTooltips();
    
    // Load theme preference
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'dark') {
        toggleDarkMode();
    }
    
    
    // Initialize UI state - show landing page, hide workflow
    const landingPageEl = document.getElementById('landing-page');
    const workflowStepperEl = document.getElementById('workflow-stepper');
    
    if (landingPageEl && workflowStepperEl) {
        // Show landing page, hide workflow stepper
        landingPageEl.classList.remove('hidden');
        workflowStepperEl.classList.add('hidden');
        
        // Hide all step content except step 1 (since we want that visible if user clicks "Get Started")
        const stepContents = document.querySelectorAll('.step-content');
        stepContents.forEach(content => {
            if (content.id !== 'step-1') {
                content.classList.add('hidden');
            }
        });
    }
    
    // Set up event listeners
    
    // Theme toggle
    const themeToggle = document.getElementById('theme-toggle');
    if (themeToggle && typeof toggleDarkMode === 'function') {
        themeToggle.addEventListener('click', toggleDarkMode);
    }
    
    // Navigation
    const homeLink = document.getElementById('home-link');
    const aboutLink = document.getElementById('about-link');
    const backToAppLink = document.getElementById('back-to-app');
    
    if (homeLink && typeof showHome === 'function') {
        homeLink.addEventListener('click', function(e) {
            e.preventDefault();
            showHome();
        });
    }
    
    if (aboutLink && typeof showAbout === 'function') {
        aboutLink.addEventListener('click', function(e) {
            e.preventDefault();
            showAbout();
        });
    }
    
    if (backToAppLink) {
        backToAppLink.addEventListener('click', function(e) {
            e.preventDefault();
            // Use appropriate function if it exists, otherwise fallback
            if (typeof showMainContent === 'function') {
                showMainContent();
            } else if (typeof goToStep === 'function') {
                goToStep(1);
            }
        });
    }
    
    // API key management - check if elements exist first (we've moved these to modal)
    const toggleApiKey = document.getElementById('toggle-api-key');
    const saveApiKey = document.getElementById('save-api-key');
    const clearSavedKey = document.getElementById('clear-saved-key');
    const apiKey = document.getElementById('modal-api-key');
    
    if (toggleApiKey && typeof toggleApiKeyVisibility === 'function') {
        toggleApiKey.addEventListener('click', toggleApiKeyVisibility);
    }
    
    if (saveApiKey && typeof handleSaveApiKeyChange === 'function') {
        saveApiKey.addEventListener('change', handleSaveApiKeyChange);
    }
    
    if (clearSavedKey) {
        clearSavedKey.addEventListener('click', function() {
            localStorage.removeItem(API_KEY_STORAGE_KEY);
            if (saveApiKey) saveApiKey.checked = false;
        });
    }
    
    if (apiKey && saveApiKey) {
        apiKey.addEventListener('input', function(e) {
            if (saveApiKey.checked) {
                localStorage.setItem(API_KEY_STORAGE_KEY, e.target.value);
            }
        });
    }

    // Persist WHO ICD credentials live while the save checkbox is ticked.
    const modalIcdClientId = document.getElementById('modal-icd-client-id');
    const modalIcdClientSecret = document.getElementById('modal-icd-client-secret');
    const modalSaveCheckbox = document.getElementById('modal-save-api-key');
    if (modalIcdClientId && modalSaveCheckbox) {
        modalIcdClientId.addEventListener('input', function(e) {
            if (modalSaveCheckbox.checked) {
                localStorage.setItem(ICD_CLIENT_ID_STORAGE_KEY, e.target.value.trim());
            }
        });
    }
    if (modalIcdClientSecret && modalSaveCheckbox) {
        modalIcdClientSecret.addEventListener('input', function(e) {
            if (modalSaveCheckbox.checked) {
                localStorage.setItem(ICD_CLIENT_SECRET_STORAGE_KEY, e.target.value.trim());
            }
        });
    }

    // Mode chooser (Step 1): AI Analysis, Jev Classification, Medical
    // Translation (ICD-11), or Clean Data (split multi-value cells; no AI).
    const MODE_RADIOS = { analysis: 'mode-analysis', jev: 'mode-jev', icd: 'mode-icd', clean: 'mode-clean' };
    function setAppMode(mode) {
        appMode = MODE_RADIOS[mode] ? mode : 'analysis';
        const radio = document.getElementById(MODE_RADIOS[appMode]);
        if (radio) radio.checked = true;
        document.querySelectorAll('.mode-card').forEach(card => {
            card.classList.toggle('border-primary', card.dataset.mode === appMode);
        });
        applyStepperLabels();
        applyModePanels(currentStep);
    }
    document.querySelectorAll('input[name="app-mode"]').forEach(r => {
        r.addEventListener('change', function() { setAppMode(this.value); });
    });
    document.querySelectorAll('.mode-card').forEach(card => {
        card.addEventListener('click', function() { setAppMode(this.dataset.mode); });
    });
    setAppMode('analysis');

    // ICD config: toggle source-lang vs source-system on input-type change.
    document.querySelectorAll('input[name="icd-input-type"]').forEach(r => {
        r.addEventListener('change', updateIcdInputTypeToggle);
    });

    // ICD config / results navigation + run button.
    const icdConfigureBackBtn = document.getElementById('icd-configure-back-btn');
    const icdResultsBackBtn = document.getElementById('icd-results-back-btn');
    const runIcdBtn = document.getElementById('run-icd-btn');
    if (icdConfigureBackBtn) icdConfigureBackBtn.addEventListener('click', () => goToStep(3));
    if (icdResultsBackBtn) icdResultsBackBtn.addEventListener('click', () => goToStep(4));
    if (runIcdBtn && typeof runIcdTranslation === 'function') {
        runIcdBtn.addEventListener('click', () => runIcdTranslation(false));
    }
    const runIcdTestBtn = document.getElementById('run-icd-test-btn');
    if (runIcdTestBtn && typeof runIcdTranslation === 'function') {
        runIcdTestBtn.addEventListener('click', () => runIcdTranslation(true));
    }


    // Step navigation - with null checks
    const configNextBtn = document.getElementById('config-next-btn');
    const uploadBackBtn = document.getElementById('upload-back-btn');
    const previewBackBtn = document.getElementById('preview-back-btn');
    const previewNextBtn = document.getElementById('preview-next-btn');
    const configureBackBtn = document.getElementById('configure-back-btn');
    const resultsBackBtn = document.getElementById('results-back-btn');
    
    if (configNextBtn && typeof goToStep === 'function') configNextBtn.addEventListener('click', () => goToStep(2));
    if (uploadBackBtn && typeof goToStep === 'function') uploadBackBtn.addEventListener('click', () => goToStep(1));
    if (previewBackBtn && typeof goToStep === 'function') previewBackBtn.addEventListener('click', () => goToStep(2));
    if (previewNextBtn && typeof goToStep === 'function') previewNextBtn.addEventListener('click', () => goToStep(4));
    if (configureBackBtn && typeof goToStep === 'function') configureBackBtn.addEventListener('click', () => goToStep(3));
    if (resultsBackBtn && typeof goToStep === 'function') resultsBackBtn.addEventListener('click', () => goToStep(4));
    
    // File upload
    const uploadForm = document.getElementById('upload-form');
    if (uploadForm && typeof uploadFile === 'function') {
        uploadForm.addEventListener('submit', uploadFile);
    }
    
    // Analysis workflow - with null checks
    const addColumnBtn = document.getElementById('add-column-btn');
    const togglePatternBtn = document.getElementById('toggle-pattern-btn');
    const detectPatternsBtn = document.getElementById('detect-patterns-btn');
    const copyCategoriesBtn = document.getElementById('copy-categories-btn');
    const useCategoriesBtn = document.getElementById('use-categories-btn');
    const testRunBtn = document.getElementById('test-run-btn');
    const analyzeBtn = document.getElementById('analyze-btn');
    
    if (addColumnBtn && typeof addColumnConfig === 'function') {
        // Wrapped: the click Event must not be taken as an imported preset.
        addColumnBtn.addEventListener('click', () => addColumnConfig());
    }
    
    const exportConfigBtn = document.getElementById('export-config-btn');
    const importConfigBtn = document.getElementById('import-config-btn');
    const importConfigInput = document.getElementById('import-config-input');

    if (exportConfigBtn) exportConfigBtn.addEventListener('click', exportAnalysisConfig);
    if (importConfigBtn && importConfigInput) {
        importConfigBtn.addEventListener('click', () => importConfigInput.click());
        importConfigInput.addEventListener('change', importAnalysisConfig);
    }

    // Jev Classification workflow
    const jevAddQuestionBtn = document.getElementById('jev-add-question-btn');
    const jevExportBtn = document.getElementById('jev-export-config-btn');
    const jevImportBtn = document.getElementById('jev-import-config-btn');
    const jevImportInput = document.getElementById('jev-import-config-input');
    const jevRunBtn = document.getElementById('jev-run-btn');
    const jevTestRunBtn = document.getElementById('jev-test-run-btn');
    const jevBackBtn = document.getElementById('jev-configure-back-btn');
    const jevResumeBtn = document.getElementById('jev-resume-btn');
    const jevPartialDownloadBtn = document.getElementById('jev-partial-download-btn');

    // Wrapped: the click Event must not be taken as an imported preset.
    if (jevAddQuestionBtn) jevAddQuestionBtn.addEventListener('click', () => addJevQuestion());
    if (jevExportBtn) jevExportBtn.addEventListener('click', exportJevConfig);
    if (jevImportBtn && jevImportInput) {
        jevImportBtn.addEventListener('click', () => jevImportInput.click());
        jevImportInput.addEventListener('change', importJevConfig);
    }
    if (jevRunBtn) jevRunBtn.addEventListener('click', () => runJevClassification(false));
    if (jevTestRunBtn) jevTestRunBtn.addEventListener('click', () => runJevClassification(true));
    if (jevBackBtn) jevBackBtn.addEventListener('click', () => goToStep(3));
    if (jevResumeBtn) jevResumeBtn.addEventListener('click', continueJevRun);
    if (jevPartialDownloadBtn) jevPartialDownloadBtn.addEventListener('click', downloadAnalyzedFile);

    if (togglePatternBtn && typeof togglePatternDetection === 'function') {
        togglePatternBtn.addEventListener('click', togglePatternDetection);
    }
    
    if (detectPatternsBtn && typeof detectPatterns === 'function') {
        detectPatternsBtn.addEventListener('click', detectPatterns);
    }
    
    if (copyCategoriesBtn && typeof copyCategoriesToClipboard === 'function') {
        copyCategoriesBtn.addEventListener('click', copyCategoriesToClipboard);
    }
    
    if (useCategoriesBtn && typeof useCategoriesInPrompt === 'function') {
        useCategoriesBtn.addEventListener('click', useCategoriesInPrompt);
    }
    
    if (testRunBtn && typeof analyzeColumns === 'function') {
        testRunBtn.addEventListener('click', () => analyzeColumns(true));
    }
    
    if (analyzeBtn && typeof analyzeColumns === 'function') {
        analyzeBtn.addEventListener('click', () => analyzeColumns(false));
    }
    
    // Template selection with null check
    document.querySelectorAll('.prompt-template').forEach(template => {
        template.addEventListener('click', function() {
            const generalInstructions = document.getElementById('general-instructions');
            if (generalInstructions && this.dataset.template) {
                generalInstructions.value = this.dataset.template;
            }
        });
    });
    
    // Landing page functionality
    // These variables are already defined in the DOMContentLoaded event listener
    // Remove the duplicate window.addEventListener('DOMContentLoaded') that was causing issues
    
    function startAnalysis() {
        const landingPage = document.getElementById('landing-page');
        const workflowStepper = document.getElementById('workflow-stepper');
        landingPage.classList.add('hidden');
        document.getElementById('main-content').classList.remove('hidden');
        workflowStepper.classList.remove('hidden');
        goToStep(1);
    }

    // The stepper is written for the analysis flow; clean mode ends at step 3,
    // so hide the analysis-only steps and relabel the ones it does use.
    function applyStepperLabels() {
        const clean = (appMode === 'clean');
        document.querySelectorAll('.stepper-item').forEach(item => {
            item.classList.toggle('hidden', clean && parseInt(item.dataset.step) > 3);
        });
        const names = { 1: clean ? 'Choose Mode' : 'Configuration',
                        3: clean ? 'Split & Download' : 'Preview & Chat' };
        Object.entries(names).forEach(([step, label]) => {
            const el = document.querySelector(`.stepper-item[data-step="${step}"] .step-name`);
            if (el) el.textContent = label;
        });
    }
    
    function showHome() {
        const landingPage = document.getElementById('landing-page');
        const workflowStepper = document.getElementById('workflow-stepper');
        landingPage.classList.remove('hidden');
        document.getElementById('main-content').classList.add('hidden');
        workflowStepper.classList.add('hidden');
    }
    
    function showAbout() {
        // Create a modal to display the About information
        const aboutModal = new bootstrap.Modal(document.createElement('div'));
        aboutModal.element.innerHTML = `
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">About Aidstack Insights</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div class="row">
                        <div class="col-md-4 text-center mb-4 mb-md-0">
                            <img src="/static/excel_ai_insight_logo.webp" alt="Aidstack Insights" class="img-fluid" style="max-height: 180px;">
                        </div>
                        <div class="col-md-8">
                            <h4>Turn Spreadsheets into Insights in Minutes</h4>
                            <p>Aidstack Insights is a powerful tool that leverages advanced AI to analyze Excel and CSV data, automatically extracting insights that would take hours to find manually.</p>
                            <p>Our mission is to make data analysis accessible to everyone, regardless of technical expertise.</p>
                        </div>
                    </div>
                    
                    <hr class="my-4">
                    
                    <h5>Key Features</h5>
                    <ul>
                        <li><strong>AI-Powered Analysis:</strong> Generate meaningful insights from your data in seconds</li>
                        <li><strong>Pattern Detection:</strong> Automatically identify patterns and categorize data</li>
                        <li><strong>Custom Instructions:</strong> Tailor the analysis to your specific needs</li>
                        <li><strong>Multi-Column Analysis:</strong> Analyze relationships between different data points</li>
                        <li><strong>Privacy-First:</strong> All processing happens on secure AI servers, no data storage</li>
                    </ul>
                    
                    <h5 class="mt-4">Use Cases</h5>
                    <ul>
                        <li><strong>Business Intelligence:</strong> Extract actionable insights from sales, marketing, or financial data</li>
                        <li><strong>Data Cleaning:</strong> Identify inconsistencies and errors in your datasets</li>
                        <li><strong>Customer Analysis:</strong> Understand patterns in customer feedback and behavior</li>
                        <li><strong>Research Analysis:</strong> Quickly process and extract meaning from research data</li>
                        <li><strong>Report Generation:</strong> Create summaries and highlights from large data sets</li>
                    </ul>

                    <h5 class="mt-4"><i class="bi bi-question-circle"></i> Frequently Asked Questions</h5>
                    <div class="accordion" id="modalFaqAccordion">
                        <div class="accordion-item">
                            <h2 class="accordion-header">
                                <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#modalOfflineCollapse">
                                    Can I use it offline?
                                </button>
                            </h2>
                            <div id="modalOfflineCollapse" class="accordion-collapse collapse" data-bs-parent="#modalFaqAccordion">
                                <div class="accordion-body">
                                    <p><strong>Short answer:</strong> Partially, but it requires technical setup.</p>
                                    <p><strong>For offline LLM processing:</strong><br>
                                    Yes, if you have a powerful machine and are comfortable setting up tools like <a href="https://ollama.ai" target="_blank">Ollama</a> to run local LLMs.</p>
                                    <div class="alert alert-info">
                                        <strong><i class="bi bi-shield-check"></i> OpenAI API Data Privacy:</strong>
                                        <ul class="mb-0 mt-2 small">
                                            <li>✅ <strong>Not used for training</strong> - Your data is NOT used to train or improve their models</li>
                                            <li>⏰ <strong>30-day retention</strong> - Retained for abuse monitoring only</li>
                                            <li>🗑️ <strong>Auto-deleted</strong> - Deleted after 30 days</li>
                                            <li>📄 <a href="https://openai.com/policies/api-data-usage-policies" target="_blank">Full Policy</a></li>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="accordion-item">
                            <h2 class="accordion-header">
                                <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#modalCostCollapse">
                                    How much does it cost?
                                </button>
                            </h2>
                            <div id="modalCostCollapse" class="accordion-collapse collapse" data-bs-parent="#modalFaqAccordion">
                                <div class="accordion-body">
                                    <p><strong>The tool is free.</strong> You only pay for OpenAI API usage (typically cents for hundreds of rows). <a href="https://openai.com/api/pricing/" target="_blank">View Pricing</a></p>
                                </div>
                            </div>
                        </div>
                        <div class="accordion-item">
                            <h2 class="accordion-header">
                                <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#modalDataCollapse">
                                    Do you store my data?
                                </button>
                            </h2>
                            <div id="modalDataCollapse" class="accordion-collapse collapse" data-bs-parent="#modalFaqAccordion">
                                <div class="accordion-body">
                                    <p><strong>No.</strong> Your API key is stored locally in your browser only. Files are processed temporarily and automatically deleted.</p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-primary" data-bs-dismiss="modal">Close</button>
                </div>
            </div>
        </div>
        `;
        
        aboutModal.element.classList.add('modal', 'fade');
        document.body.appendChild(aboutModal.element);
        aboutModal.show();
        
        // Clean up modal after it's hidden
        aboutModal.element.addEventListener('hidden.bs.modal', function() {
            document.body.removeChild(aboutModal.element);
        });
    }
    
    // Button handlers
    const getStartedBtn = document.getElementById('get-started-btn');
    const startAnalyzingBtn = document.getElementById('start-analyzing-btn');

    if (getStartedBtn) {
        getStartedBtn.addEventListener('click', startAnalysis);
    }

    if (startAnalyzingBtn) {
        startAnalyzingBtn.addEventListener('click', startAnalysis);
    }

    
    // Load saved API key from localStorage if available
    function loadSavedApiKey() {
        const savedKey = localStorage.getItem(API_KEY_STORAGE_KEY);
        if (savedKey) {
            // Update modal fields - need to get fresh references here
            const modalApiKeyField = document.getElementById('modal-api-key');
            const modalSaveKeyCheckbox = document.getElementById('modal-save-api-key');
            
            if (modalApiKeyField) {
                modalApiKeyField.value = savedKey;
                if (modalSaveKeyCheckbox) {
                    modalSaveKeyCheckbox.checked = true;
                }
            }
            
            // Update main form field if it exists 
            // (we've removed it from UI but keeping compatibility with old code)
            const apiKeyInput = document.getElementById('modal-api-key');
            if (apiKeyInput) {
                apiKeyInput.value = savedKey;
                const saveKeyCheckbox = document.getElementById('save-api-key');
                if (saveKeyCheckbox) {
                    saveKeyCheckbox.checked = true;
                }
            }
        }
    }
    
    // Clear saved API key
    function clearSavedApiKey() {
        localStorage.removeItem(API_KEY_STORAGE_KEY);
        const modalSaveKeyCheckbox = document.getElementById('modal-save-api-key');
        if (modalSaveKeyCheckbox) {
            modalSaveKeyCheckbox.checked = false;
        }
    }
    
    // API Settings Modal functionality
    const apiSettingsBtn = document.getElementById('api-settings-btn');
    const modalApiKey = document.getElementById('modal-api-key');
    const modalToggleApiKey = document.getElementById('modal-toggle-api-key');
    const modalSaveApiKey = document.getElementById('modal-save-api-key');
    const saveApiSettings = document.getElementById('save-api-settings');
    
    // Initialize Modal
    const apiSettingsModal = new bootstrap.Modal(document.getElementById('api-settings-modal'));

    /** Restore provider choice and Azure/model fields into the modal. */
    function loadSavedProviderSettings() {
        const provider = localStorage.getItem(PROVIDER_STORAGE_KEY) || 'openai';
        const radio = document.getElementById(
            provider === 'azure' ? 'provider-azure' : 'provider-openai');
        if (radio) radio.checked = true;

        const restore = (id, key) => {
            const el = document.getElementById(id);
            if (el) el.value = localStorage.getItem(key) || '';
        };
        restore('modal-azure-endpoint', AZURE_ENDPOINT_STORAGE_KEY);
        restore('modal-azure-deployment', AZURE_DEPLOYMENT_STORAGE_KEY);
        restore('modal-azure-api-version', AZURE_API_VERSION_STORAGE_KEY);
        restore('modal-openai-model', OPENAI_MODEL_STORAGE_KEY);
        // Jev credentials are independent of the provider radio above.
        restore('modal-jev-api-key', JEV_API_KEY_STORAGE_KEY);
        restore('modal-jev-model', JEV_MODEL_STORAGE_KEY);

        syncProviderUI();
    }

    // Switch the visible fields when the provider changes.
    document.querySelectorAll('input[name="llm-provider"]').forEach(radio => {
        radio.addEventListener('change', syncProviderUI);
    });

    // Test Connection: verify credentials before running a whole file.
    const testConnBtn = document.getElementById('test-connection-btn');
    if (testConnBtn) {
        testConnBtn.addEventListener('click', async function() {
            const out = document.getElementById('test-connection-result');
            testConnBtn.disabled = true;
            out.innerHTML = '<span class="text-muted"><i class="bi bi-hourglass-split"></i> Testing connection...</span>';
            try {
                const response = await fetch('/test_connection', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(getLLMConfig())
                });
                const result = await response.json();
                if (response.ok && result.ok) {
                    out.innerHTML = '<span class="text-success"><i class="bi bi-check-circle-fill"></i> '
                        + 'Connected to ' + result.provider + ' (model: ' + result.model + ')</span>';
                } else {
                    out.innerHTML = '<span class="text-danger"><i class="bi bi-x-circle-fill"></i> '
                        + escapeHtml(result.error || 'Connection failed') + '</span>';
                }
            } catch (e) {
                out.innerHTML = '<span class="text-danger"><i class="bi bi-x-circle-fill"></i> ' + escapeHtml(e.message) + '</span>';
            } finally {
                testConnBtn.disabled = false;
            }
        });
    }
    
    // Test Jev Connection: verify the Jev key independently of the LLM one.
    const testJevConnBtn = document.getElementById('test-jev-connection-btn');
    if (testJevConnBtn) {
        testJevConnBtn.addEventListener('click', async function() {
            const out = document.getElementById('test-jev-connection-result');
            testJevConnBtn.disabled = true;
            out.innerHTML = '<span class="text-muted"><i class="bi bi-hourglass-split"></i> Testing connection...</span>';
            try {
                const response = await fetch('/test_jev_connection', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(getJevConfig())
                });
                const result = await response.json();
                if (response.ok && result.ok) {
                    out.innerHTML = '<span class="text-success"><i class="bi bi-check-circle-fill"></i> '
                        + 'Connected to Jev (model: ' + escapeHtml(result.model) + ')</span>';
                } else {
                    out.innerHTML = '<span class="text-danger"><i class="bi bi-x-circle-fill"></i> '
                        + escapeHtml(result.error || 'Connection failed') + '</span>';
                }
            } catch (e) {
                out.innerHTML = '<span class="text-danger"><i class="bi bi-x-circle-fill"></i> ' + escapeHtml(e.message) + '</span>';
            } finally {
                testJevConnBtn.disabled = false;
            }
        });
    }

    // Toggle Jev key visibility in modal
    const modalJevApiKey = document.getElementById('modal-jev-api-key');
    const modalToggleJevApiKey = document.getElementById('modal-toggle-jev-api-key');
    if (modalToggleJevApiKey && modalJevApiKey) {
        modalToggleJevApiKey.addEventListener('click', function() {
            if (modalJevApiKey.type === 'password') {
                modalJevApiKey.type = 'text';
                modalToggleJevApiKey.innerHTML = '<i class="bi bi-eye-slash"></i>';
            } else {
                modalJevApiKey.type = 'password';
                modalToggleJevApiKey.innerHTML = '<i class="bi bi-eye"></i>';
            }
        });
    }

    // Add event listeners with null checks
    if (apiSettingsBtn) {
        apiSettingsBtn.addEventListener('click', function() {
            // Load current API key from localStorage if available
            loadSavedApiKey();
            loadSavedProviderSettings();
            apiSettingsModal.show();
        });
    }
    
    // Toggle API key visibility in modal
    if (modalToggleApiKey && modalApiKey) {
        modalToggleApiKey.addEventListener('click', function() {
            if (modalApiKey.type === 'password') {
                modalApiKey.type = 'text';
                modalToggleApiKey.innerHTML = '<i class="bi bi-eye-slash"></i>';
            } else {
                modalApiKey.type = 'password';
                modalToggleApiKey.innerHTML = '<i class="bi bi-eye"></i>';
            }
        });
    }
    
    // Save API settings
    if (saveApiSettings && modalApiKey && modalSaveApiKey) {
        saveApiSettings.addEventListener('click', function() {
            const apiKey = modalApiKey.value.trim();
            const saveKey = modalSaveApiKey.checked;

            // Read WHO ICD credentials (used by Medical Translation mode).
            const icdIdField = document.getElementById('modal-icd-client-id');
            const icdSecretField = document.getElementById('modal-icd-client-secret');
            const icdId = icdIdField ? icdIdField.value.trim() : '';
            const icdSecret = icdSecretField ? icdSecretField.value.trim() : '';

            // Read provider + Azure fields.
            const providerEl = document.querySelector('input[name="llm-provider"]:checked');
            const provider = (providerEl && providerEl.value) || 'openai';
            const fieldVal = (id) => {
                const el = document.getElementById(id);
                return el ? el.value.trim() : '';
            };
            const azureEndpoint = fieldVal('modal-azure-endpoint');
            const azureDeployment = fieldVal('modal-azure-deployment');
            const azureApiVersion = fieldVal('modal-azure-api-version');
            const openaiModel = fieldVal('modal-openai-model');
            const jevApiKey = fieldVal('modal-jev-api-key');
            const jevModel = fieldVal('modal-jev-model');

            if (!apiKey && !icdId && !icdSecret && !azureEndpoint && !azureDeployment && !jevApiKey) {
                alert('Please enter your AI provider credentials, your Jev API key, and/or your WHO ICD credentials.');
                return;
            }

            // Azure needs all three parts to work; warn early rather than
            // failing on the first analysis request.
            if (provider === 'azure' && apiKey && !(azureEndpoint && azureDeployment)) {
                alert('Azure AI Foundry requires an Endpoint and a Deployment Name in addition to the API key.');
                return;
            }

            // The provider choice itself is always remembered.
            localStorage.setItem(PROVIDER_STORAGE_KEY, provider);

            // Persist (or clear) provider details based on the save checkbox.
            const persist = (key, value) => {
                if (saveKey && value) localStorage.setItem(key, value);
                else localStorage.removeItem(key);
            };
            persist(AZURE_ENDPOINT_STORAGE_KEY, azureEndpoint);
            persist(AZURE_DEPLOYMENT_STORAGE_KEY, azureDeployment);
            persist(AZURE_API_VERSION_STORAGE_KEY, azureApiVersion);
            persist(OPENAI_MODEL_STORAGE_KEY, openaiModel);
            persist(JEV_API_KEY_STORAGE_KEY, jevApiKey);
            persist(JEV_MODEL_STORAGE_KEY, jevModel);

            // Persist (or clear) OpenAI key based on the save checkbox.
            if (apiKey) {
                if (saveKey) {
                    localStorage.setItem(API_KEY_STORAGE_KEY, apiKey);
                } else {
                    localStorage.removeItem(API_KEY_STORAGE_KEY);
                }
                const apiKeyInput = document.getElementById('modal-api-key');
                if (apiKeyInput) {
                    apiKeyInput.value = apiKey;
                    const saveKeyCheckbox = document.getElementById('save-api-key');
                    if (saveKeyCheckbox) {
                        saveKeyCheckbox.checked = saveKey;
                    }
                }
            }

            // Persist (or clear) WHO ICD credentials based on the same save checkbox.
            if (saveKey) {
                if (icdId) localStorage.setItem(ICD_CLIENT_ID_STORAGE_KEY, icdId);
                else localStorage.removeItem(ICD_CLIENT_ID_STORAGE_KEY);
                if (icdSecret) localStorage.setItem(ICD_CLIENT_SECRET_STORAGE_KEY, icdSecret);
                else localStorage.removeItem(ICD_CLIENT_SECRET_STORAGE_KEY);
            } else {
                localStorage.removeItem(ICD_CLIENT_ID_STORAGE_KEY);
                localStorage.removeItem(ICD_CLIENT_SECRET_STORAGE_KEY);
            }

            apiSettingsModal.hide();
            alert('API settings saved successfully!');
        });
    }
    
    // Handle the old toggle-api-key button (which we've removed from UI)
    const oldToggleBtn = document.getElementById('toggle-api-key');
    if (oldToggleBtn) {
        oldToggleBtn.addEventListener('click', function() {
            const apiKeyInput = document.getElementById('modal-api-key');
            if (apiKeyInput) {
                if (apiKeyInput.type === 'password') {
                    apiKeyInput.type = 'text';
                    this.innerHTML = '<i class="bi bi-eye-slash"></i>';
                } else {
                    apiKeyInput.type = 'password';
                    this.innerHTML = '<i class="bi bi-eye"></i>';
                }
            }
        });
    }

    resultsUI = initResults({ getFileData: () => fileData, getChatDataset, streamChat, getLLMConfig, showAlert });
    initCleanData();
});
