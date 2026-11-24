// Replace a source and all its split messages together so retries restore atomically.
let connection, revision;
async function database() {
    if (!connection) connection = new Promise((resolve, reject) => {
        const request = indexedDB.open('aidstack-pipeline', 2);
        request.onupgradeneeded = () => {
            const db = request.result;
            for (const name of ['runs', 'sources']) {
                if (!db.objectStoreNames.contains(name)) db.createObjectStore(name);
            }
        };
        request.onsuccess = () => {
            const db = request.result;
            db.onversionchange = () => { db.close(); connection = undefined; };
            resolve(db);
        };
        request.onerror = () => { connection = undefined; reject(request.error); };
    });
    return connection;
}

// Check only keys on startup so large saved workbooks are loaded on demand.
export async function hasSavedPipeline() {
    const db = await database();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('runs');
        const runs = tx.objectStore('runs');
        const header = runs.getKey('header'), legacy = runs.getKey('latest');
        tx.oncomplete = () => resolve(Boolean(header.result || legacy.result));
        tx.onabort = () => reject(tx.error);
    });
}

// Saves run one at a time, in call order. A partial save captures its source group when called, because
// parallel requests can insert rows (and shift indexes) before the write starts.
let queue = Promise.resolve();
export function savePipeline(run, index, replace = false) {
    let group;
    if (index !== undefined) {
        const work = run.work, id = work[index].values._source_record_id;
        let start = index, end = index + 1;
        while (start > 0 && work[start-1].values._source_record_id === id) start--;
        while (end < work.length && work[end].values._source_record_id === id) end++;
        group = {id, entries: work.slice(start, end)};
    }
    const write = queue.then(() => writePipeline(run, group, replace));
    queue = write.catch(() => {});
    return write;
}
async function writePipeline(run, partial, replace) {
    const db = await database();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(['runs', 'sources'], 'readwrite', {durability:'strict'});
        const runs = tx.objectStore('runs'), sources = tx.objectStore('sources');
        const check = runs.get('revision');
        const next = crypto.randomUUID();
        let conflict = false;
        check.onsuccess = () => {
            if (!replace && revision !== check.result) { conflict = true; tx.abort(); return; }
            const {work, classifications, nextId, pauseReason, providerCooldowns, elapsedMs, ...header} = run;
            runs.put({nextId, pauseReason, providerCooldowns, elapsedMs}, 'progress');
            runs.put(next, 'revision');
            if (partial) {
                sources.put({entries: partial.entries, classification: classifications?.[partial.id]}, partial.id);
                return;
            }
            sources.clear();
            runs.delete('latest');
            runs.put(header, 'header');
            let block = [], blockNumber = 0;
            let start = 0;
            while (start < work.length) {
                const id = work[start].values._source_record_id;
                let end = start + 1;
                while (end < work.length && work[end].values._source_record_id === id) end++;
                block.push({entries:work.slice(start,end), classification:classifications?.[id]});
                if (block.length === 250 || end === work.length) {
                    sources.put(block, blockNumber++);
                    block = [];
                }
                start = end;
            }
        };
        tx.oncomplete = () => { revision = next; resolve(); };
        tx.onabort = () => reject(new Error(conflict ? 'Saved progress changed in another tab. Restore the latest run before continuing.' : 'Could not save progress. Download a recovery file before leaving.'));
    });
}

export async function loadPipeline() {
    const db = await database();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(['runs', 'sources']);
        const runs = tx.objectStore('runs');
        const header = runs.get('header'), progress = runs.get('progress'), token = runs.get('revision'), legacy = runs.get('latest');
        const groups = tx.objectStore('sources').getAll();
        tx.oncomplete = () => {
            revision = token.result;
            const work = [], classifications = {};
            // Numeric baseline blocks precede string-keyed updates in IndexedDB order.
            const ordered = new Map();
            for (const value of groups.result) {
                for (const group of Array.isArray(value) ? value : [value]) {
                    ordered.set(group.entries[0].values._source_record_id, group);
                }
            }
            for (const {entries, classification} of ordered.values()) {
                for (const entry of entries) work.push(entry);
                if (classification) classifications[entries[0].values._source_record_id] = classification;
            }
            resolve(header.result ? {...header.result, ...progress.result, work, classifications} : legacy.result);
        };
        tx.onabort = () => reject(new Error('Could not read saved progress.'));
    });
}
