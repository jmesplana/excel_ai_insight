import {exportCheckpointCsv} from './checkpoint-export.js';

const kind = document.getElementById('kind');
const requested = new URLSearchParams(location.search).get('kind');
if (['jev', 'analysis'].includes(requested)) kind.value = requested;
const button = document.getElementById('download');
const status = document.getElementById('status');
const progress = document.getElementById('progress');

button.onclick = async () => {
    button.disabled = true;
    kind.disabled = true;
    progress.hidden = true;
    status.textContent = 'Opening saved results…';
    try {
        let sink;
        const filename = `recovered_${kind.value}_results.csv`;
        if (typeof window.showSaveFilePicker === 'function') {
            // Invoke the picker during the click's user activation. Streaming
            // avoids holding the full CSV, XLSX workbook or audit in memory.
            const handle = await window.showSaveFilePicker({suggestedName:filename,
                types:[{description:'CSV results',accept:{'text/csv':['.csv']}}]});
            sink = await handle.createWritable();
        } else {
            // Other browsers retain only CSV blobs, never a full workbook or
            // the audit. No giant joined string is constructed.
            const chunks = [];
            sink = {
                write: async text => { chunks.push(new Blob([text])); },
                close: async () => {
                    const url = URL.createObjectURL(new Blob(chunks, {type:'text/csv;charset=utf-8'}));
                    const link = document.createElement('a');
                    link.href = url; link.download = filename; link.click();
                    chunks.length = 0;
                    setTimeout(() => URL.revokeObjectURL(url), 60000);
                },
                abort: async () => { chunks.length = 0; },
            };
        }
        const result = await exportCheckpointCsv(kind.value, sink, (done, total) => {
            progress.hidden = false; progress.max = total; progress.value = done;
            if (done % 250 === 0 || done === total) status.textContent = `Exporting ${done.toLocaleString()} of ${total.toLocaleString()} saved rows…`;
        });
        status.textContent = `Exported ${result.exportedRows.toLocaleString()} of ${result.totalRows.toLocaleString()} rows from ${result.filename || 'the saved run'}. Your checkpoint is unchanged.`;
    } catch (error) {
        status.textContent = error.name === 'AbortError' ? 'Export canceled. Your checkpoint is unchanged.'
            : `Could not export: ${error.message}`;
    } finally {
        button.disabled = false;
        kind.disabled = false;
    }
};
