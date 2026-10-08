const {test, expect} = require('@playwright/test');
const fs = require('fs');

async function seed(page, {kind = 'jev', count = 101, legacy = false, imported = false} = {}) {
    await page.goto(`/static/checkpoint-recovery.html?kind=${kind}`);
    await page.evaluate(async ({kind,count,legacy,imported}) => {
        const {saveCheckpoint} = await import('/static/js/run-checkpoint.js');
        const rows = Array.from({length:count}, (_, ID) => ({ID,Feedback:ID === 0 ? 'Bonjour, "ami"\nÉbola' : `row${ID}`}));
        const results = rows.map(row => ({...row,Category:'A',Empty:null}));
        const audit = rows.map((_, rowIndex) => ({rowIndex,decisions:{largeAudit:{labels:'Not exported '.repeat(40)}}}));
        const state = {runId:'saved-run',filename:'original.csv',columns:['ID','Feedback','Category','Empty'],
            rows,cursor:imported ? count : 0,results:imported ? results : [],audit:imported ? audit : []};
        await saveCheckpoint(kind,'start',state);
        if (!imported) await saveCheckpoint(kind,'put',{...state,cursor:count,results,audit},rows.map((_, i) => i));
        if (legacy) {
            const db = await new Promise(resolve => {
                const r = indexedDB.open(kind === 'jev' ? 'aidstack-jev' : 'aidstack-analysis');
                r.onsuccess = () => resolve(r.result);
            });
            await new Promise(resolve => {
                const tx = db.transaction('runs','readwrite');
                tx.objectStore('runs').delete('exportInfo');
                tx.oncomplete = resolve;
            });
            db.close();
        }
    }, {kind,count,legacy,imported});
}

for (const legacy of [true, false]) {
    test(`${legacy ? 'legacy 35,000-row' : 'new'} checkpoint streams ordered CSV without bulk audit reads`, async ({page}) => {
        test.setTimeout(120000);
        const count = legacy ? 35000 : 101;
        await seed(page,{count,legacy});
        const result = await page.evaluate(async ({count,legacy}) => {
            const {exportCheckpointCsv} = await import('/static/js/checkpoint-export.js');
            IDBObjectStore.prototype.getAll = () => { throw new Error('Bulk reads are forbidden'); };
            IDBObjectStore.prototype.put = () => { throw new Error('Checkpoint writes are forbidden'); };
            const get = IDBObjectStore.prototype.get;
            let latestReads = 0;
            IDBObjectStore.prototype.get = function(key) {
                if (key === 'latest') {
                    latestReads++;
                    if (!legacy) throw new Error('New checkpoints should use only small export metadata');
                }
                return get.call(this,key);
            };
            let header = true, next = 0, closed = false, aborted = false, maxRows = 0;
            const result = await exportCheckpointCsv('jev',{
                write: async text => {
                    if (header) {
                        if (text !== '\ufeffID,Feedback,Category,Empty\r\n') throw new Error('Wrong CSV header');
                        header = false; return;
                    }
                    const lines = text.split('\r\n').slice(0,-1);
                    maxRows = Math.max(maxRows,lines.length);
                    for (const line of lines) {
                        const expected = next === 0 ? '0,"Bonjour, ""ami""\nÉbola",A,' : `${next},row${next},A,`;
                        if (line !== expected) throw new Error(`Row ${next} is out of order or incorrectly escaped`);
                        next++;
                    }
                },
                close: async () => { closed = true; }, abort: async () => { aborted = true; },
            });
            return {next,closed,aborted,maxRows,latestReads,result};
        },{count,legacy});
        expect(result).toEqual({next:count,closed:true,aborted:false,maxRows:25,latestReads:legacy ? 1 : 0,
            result:{filename:'original.csv',exportedRows:count,totalRows:count}});
    });
}

test('CSV recovery supports imported checkpoints and general analysis', async ({page}) => {
    await seed(page,{kind:'analysis',count:3,imported:true});
    const exported = await page.evaluate(async () => {
        const {exportCheckpointCsv} = await import('/static/js/checkpoint-export.js');
        const parts = [];
        return {result:await exportCheckpointCsv('analysis',{
            write:async text => parts.push(text), close:async () => {},
        }), csv:parts.join('')};
    });
    expect(exported.result.exportedRows).toBe(3);
    expect(exported.csv).toContain('2,row2,A,\r\n');
});

test('missing rows and interrupted writes abort the file instead of silently dropping data', async ({page}) => {
    await seed(page,{count:51});
    const result = await page.evaluate(async () => {
        const {exportCheckpointCsv} = await import('/static/js/checkpoint-export.js');
        let aborted = false, closed = false, writes = 0, error;
        try {
            await exportCheckpointCsv('jev',{
                write:async () => { if (++writes === 2) throw new Error('Disk full'); },
                close:async () => { closed = true; }, abort:async () => { aborted = true; },
            });
        } catch (e) { error = e.message; }
        const db = await new Promise(resolve => { const r=indexedDB.open('aidstack-jev'); r.onsuccess=()=>resolve(r.result); });
        const cursor = await new Promise(resolve => {
            const r=db.transaction('runs','readonly').objectStore('runs').get('progress');
            r.onsuccess=()=>resolve(r.result.cursor);
        });
        await new Promise(resolve => {
            const tx=db.transaction('runs','readwrite'); tx.objectStore('runs').delete('row:30'); tx.oncomplete=resolve;
        });
        db.close();
        let missing;
        try { await exportCheckpointCsv('jev',{write:async()=>{},close:async()=>{throw new Error('Must not close');}}); }
        catch (e) { missing=e.message; }
        return {aborted,closed,error,cursor,missing};
    });
    expect(result).toEqual({aborted:true,closed:false,error:'Disk full',cursor:51,
        missing:'Saved row 31 is missing. The checkpoint has not been changed.'});
});

test('replacement checkpoints cannot be mixed into an export', async ({page}) => {
    await seed(page,{count:51});
    const result = await page.evaluate(async () => {
        const {exportCheckpointCsv} = await import('/static/js/checkpoint-export.js');
        const {saveCheckpoint} = await import('/static/js/run-checkpoint.js');
        let writes = 0, aborted = false;
        try {
            await exportCheckpointCsv('jev',{
                write:async () => {
                    if (++writes === 2) await saveCheckpoint('jev','start',{
                        runId:'new-run',filename:'new.csv',rows:[],results:[],cursor:0,columns:['New'],
                    });
                },close:async()=>{},abort:async()=>{aborted=true;},
            });
        } catch (e) { return {error:e.message,aborted}; }
    });
    expect(result).toEqual({error:'A different run replaced the checkpoint during export.',aborted:true});
});

test('recovery page downloads CSV without restoring results or calling a provider', async ({page}) => {
    const posts = [];
    page.on('request', request => { if (request.method() === 'POST') posts.push(request.url()); });
    await seed(page,{count:35});
    // Exercise the browser-download fallback used without File System Access.
    await page.evaluate(() => { window.showSaveFilePicker = undefined; });
    const download = page.waitForEvent('download');
    await page.locator('#download').click();
    const csv = fs.readFileSync(await (await download).path(),'utf8');
    expect(csv).toContain('34,row34,A,\r\n');
    await expect(page.locator('#status')).toContainText('Exported 35 of 35');
    expect(posts).toEqual([]);
    expect(await page.evaluate(() => typeof XLSX)).toBe('undefined');
});

test('missing checkpoints and canceled save dialogs leave the page usable', async ({page}) => {
    await page.goto('/static/checkpoint-recovery.html');
    await page.evaluate(() => { window.showSaveFilePicker = async () => { throw new DOMException('Canceled','AbortError'); }; });
    await page.locator('#download').click();
    await expect(page.locator('#status')).toContainText('Export canceled');
    await expect(page.locator('#download')).toBeEnabled();
    await page.evaluate(() => { window.showSaveFilePicker = undefined; });
    await page.locator('#download').click();
    await expect(page.locator('#status')).toContainText('No saved run');
    await expect(page.locator('#download')).toBeEnabled();
});

test('recovery page streams to the selected file when File System Access is available', async ({page}) => {
    await seed(page,{count:35});
    await page.evaluate(() => {
        window.savedFile = {chunks:[],closed:false,aborted:false};
        window.showSaveFilePicker = async () => ({createWritable:async () => ({
            write:async text => savedFile.chunks.push(text),
            close:async () => { savedFile.closed=true; },
            abort:async () => { savedFile.aborted=true; },
        })});
    });
    await page.locator('#download').click();
    await expect(page.locator('#status')).toContainText('Exported 35 of 35');
    const result = await page.evaluate(() => savedFile);
    expect(result.closed).toBe(true);
    expect(result.aborted).toBe(false);
    expect(result.chunks).toHaveLength(3);
    expect(result.chunks.join('')).toContain('34,row34,A,\r\n');
});
