/** Incremental, atomic checkpoints. Keep the existing Jev database compatible. */
export async function saveCheckpoint(kind, action, value, changedIndices = []) {
    const db = await new Promise((resolve, reject) => {
        const request = indexedDB.open(kind === 'jev' ? 'aidstack-jev' : 'aidstack-analysis', 1);
        request.onupgradeneeded = () => request.result.createObjectStore('runs');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
    try {
        return await new Promise((resolve, reject) => {
            const tx = db.transaction('runs', action === 'get' ? 'readonly' : 'readwrite',
                action === 'get' ? {} : {durability: 'strict'});
            const store = tx.objectStore('runs');
            let source, progress, records, conflict;
            if (action === 'get') {
                source = store.get('latest');
                progress = store.get('progress');
                records = store.getAll(IDBKeyRange.bound('row:', 'row:\uffff'));
            } else if (action === 'delete') {
                store.clear();
            } else if (action === 'start') {
                try {
                    store.clear();
                    store.put(value, 'latest');
                    store.put(value.runId, 'runId');
                } catch (error) {
                    tx.abort();
                    reject(error);
                    return;
                }
            } else {
                const writeBatch = activeId => {
                    try {
                        if (activeId !== value.runId) {
                            throw new Error('Another run replaced this checkpoint. Download a recovery file for this run before leaving.');
                        }
                        store.put(value.runId, 'runId');
                        store.put({cursor: value.cursor, config: value.config, errors: value.errors || 0}, 'progress');
                        for (const i of changedIndices) {
                            store.put({index: i, row: value.results[i], ...(value.audit ? {audit: value.audit[i]} : {})}, `row:${i}`);
                        }
                    } catch (error) {
                        conflict = error;
                        tx.abort();
                    }
                };
                // Never deserialize the entire source dataset on each batch write.
                const active = store.get('runId');
                active.onsuccess = () => {
                    if (active.result !== undefined) { writeBatch(active.result); return; }
                    // Migrate a checkpoint created before the lightweight run ID existed.
                    const legacy = store.get('latest');
                    legacy.onsuccess = () => writeBatch(legacy.result?.runId);
                };
            }
            tx.oncomplete = () => {
                if (action !== 'get') { resolve(); return; }
                const saved = source.result;
                if (saved && progress.result) {
                    Object.assign(saved, progress.result);
                    for (const record of records.result) {
                        saved.results[record.index] = record.row;
                        if (saved.audit) saved.audit[record.index] = record.audit;
                    }
                }
                resolve(saved);
            };
            tx.onerror = () => reject(conflict || tx.error);
            tx.onabort = () => reject(conflict || tx.error || new Error('Checkpoint transaction aborted.'));
        });
    } finally { db.close(); }
}
