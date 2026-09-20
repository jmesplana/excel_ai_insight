/** Latest local run, no credentials. Source data is written once, then changed rows only. */
export async function checkpoint(action, value, changedIndices = []) {
    const db = await new Promise((resolve, reject) => {
        const request = indexedDB.open('aidstack-jev', 1);
        request.onupgradeneeded = () => request.result.createObjectStore('runs');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
    try {
        return await new Promise((resolve, reject) => {
            const tx = db.transaction('runs', action === 'get' ? 'readonly' : 'readwrite');
            const store = tx.objectStore('runs');
            let source, progress, records;
            if (action === 'get') {
                source = store.get('latest');
                progress = store.get('progress');
                records = store.getAll(IDBKeyRange.bound('row:', 'row:\uffff'));
            } else if (action === 'delete') {
                store.clear();
            } else if (action === 'start') {
                store.clear();
                store.put({...value, running: false, stopped: false}, 'latest');
            } else {
                store.put({cursor: value.cursor, config: value.config}, 'progress');
                for (const i of changedIndices) store.put({index: i, row: value.results[i], audit: value.audit[i]}, `row:${i}`);
            }
            tx.oncomplete = () => {
                if (action !== 'get') { resolve(); return; }
                const saved = source.result;
                if (saved && progress.result) {
                    Object.assign(saved, progress.result);
                    for (const record of records.result) {
                        saved.results[record.index] = record.row;
                        saved.audit[record.index] = record.audit;
                    }
                }
                resolve(saved);
            };
            tx.onerror = () => reject(tx.error);
            tx.onabort = () => reject(tx.error || new Error('Checkpoint transaction aborted.'));
        });
    } finally { db.close(); }
}
