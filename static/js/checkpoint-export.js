/** Export saved result rows without restoring the dataset or collecting its audits. */
import {csvLine} from './csv-export.js';

function openSavedDatabase(kind) {
    if (!['jev', 'analysis'].includes(kind)) throw new Error('Choose classification or general analysis.');
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(kind === 'jev' ? 'aidstack-jev' : 'aidstack-analysis');
        request.onupgradeneeded = () => {
            request.transaction.abort();
            reject(new Error('No saved run at this address in this browser profile.'));
        };
        request.onerror = () => reject(request.error);
        request.onsuccess = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains('runs')) {
                db.close(); reject(new Error('No saved run found.')); return;
            }
            resolve(db);
        };
    });
}

function exportInfo(db) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction('runs', 'readonly');
        const store = tx.objectStore('runs');
        const progress = store.get('progress');
        const id = store.get('runId');
        const info = store.get('exportInfo');
        let metadata, baseline = null;
        info.onsuccess = () => {
            if (info.result) { metadata = info.result; return; }
            // Older checkpoints have no small metadata record. Their initial
            // source/settings record is needed once, but never getAll(audits).
            const source = store.get('latest');
            source.onsuccess = () => {
                const saved = source.result;
                if (!saved) return;
                metadata = {runId:saved.runId, filename:saved.filename, columns:saved.columns,
                    totalRows:saved.rows?.length, cursor:saved.cursor};
                baseline = saved.results;
            };
        };
        tx.oncomplete = () => {
            if (!metadata) { reject(new Error('No saved run found.')); return; }
            const result = {...metadata, cursor:progress.result?.cursor ?? metadata.cursor,
                storedRunId:id.result, baseline};
            if (id.result !== undefined && id.result !== metadata.runId) {
                reject(new Error('The checkpoint changed. Stop other runs and try the download again.')); return;
            }
            if (!Array.isArray(result.columns) || !result.columns.length
                || !result.columns.every(c => typeof c === 'string')
                || !Number.isInteger(result.cursor) || result.cursor < 0
                || !Number.isInteger(result.totalRows) || result.cursor > result.totalRows) {
                reject(new Error('The saved checkpoint has invalid metadata. It has not been modified.')); return;
            }
            if (!result.cursor) { reject(new Error('The checkpoint contains no completed rows yet.')); return; }
            resolve(result);
        };
        tx.onerror = () => reject(tx.error);
        tx.onabort = () => reject(tx.error || new Error('Could not read the checkpoint.'));
    });
}

function readCsvBatch(db, info, start, end) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction('runs', 'readonly');
        const store = tx.objectStore('runs');
        const id = store.get('runId');
        const lines = [];
        let index = start, failure;
        const fail = message => { failure = new Error(message); tx.abort(); };
        const next = () => {
            if (index === end) return;
            // Numerical lookups preserve original order (cursor keys would put
            // row:10 before row:2). At most one result/audit is decoded at a time.
            const request = store.get(`row:${index}`);
            request.onsuccess = () => {
                const record = request.result;
                if (record && record.index !== index) { fail(`Invalid saved record ${index + 1}.`); return; }
                const row = record?.row ?? info.baseline?.[index];
                if (!row || typeof row !== 'object' || Array.isArray(row)) {
                    // Imported recovery files can hold their initial results in
                    // 'latest' rather than individual row keys.
                    if (info.baseline === null) {
                        const source = store.get('latest');
                        source.onsuccess = () => {
                            if (source.result?.runId !== info.runId) { fail('The saved run changed during export.'); return; }
                            info.baseline = source.result.results || [];
                            next();
                        };
                        return;
                    }
                    fail(`Saved row ${index + 1} is missing. The checkpoint has not been changed.`); return;
                }
                lines.push(csvLine(info.columns.map(column => row[column])));
                index++;
                next();
            };
        };
        id.onsuccess = () => {
            if (id.result !== info.storedRunId) { fail('A different run replaced the checkpoint during export.'); return; }
            next();
        };
        tx.oncomplete = () => resolve(lines.join(''));
        tx.onerror = () => reject(failure || tx.error);
        tx.onabort = () => reject(failure || tx.error || new Error('Checkpoint read interrupted.'));
    });
}

/** The sink may stream to a user-selected file; no checkpoint writes or API calls. */
export async function exportCheckpointCsv(kind, sink, onProgress = () => {}) {
    let db;
    try {
        db = await openSavedDatabase(kind);
        const info = await exportInfo(db);
        onProgress(0, info.cursor);
        await sink.write('\ufeff' + csvLine(info.columns));
        for (let start = 0; start < info.cursor; start += 25) {
            const end = Math.min(start + 25, info.cursor);
            const csv = await readCsvBatch(db, info, start, end);
            // Finish the IndexedDB transaction before awaiting disk I/O.
            await sink.write(csv);
            onProgress(end, info.cursor);
        }
        await sink.close();
        return {filename:info.filename, exportedRows:info.cursor, totalRows:info.totalRows};
    } catch (error) {
        try { await sink.abort?.(); } catch { /* Retain the original read/write error. */ }
        throw error;
    } finally { db?.close(); }
}
