const {test, expect} = require('@playwright/test');
const fs = require('fs');

async function setup(page, count = 25, jev = false) {
    await page.goto('/');
    await page.locator('#start-analyzing-btn').click();
    if (jev) await page.locator('#mode-card-jev').click();
    else await page.locator('#general-instructions').fill('Preserve meaning and use the original settings.');
    await page.locator('#config-next-btn').click();
    const csv = 'Text,Number\n' + Array.from({length: count}, (_, i) => `row${i},${i}`).join('\n');
    await page.locator('#file').setInputFiles({name: 'recovery.csv', mimeType: 'text/csv', buffer: Buffer.from(csv)});
    await page.locator('#upload-form button[type=submit]').click();
    await page.locator('#preview-next-btn').click();
    if (jev) {
        const config = {format:'aidstack-insights-jev-config', version:2, sourceColumns:['Text'],
            questions:[{outputColumnName:'Category', questionType:'choice', instructions:'Classify', options:['A','B']}],
            execution:{batchSize:2}};
        await page.locator('#jev-import-config-input').setInputFiles({name:'config.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
    } else {
        await page.locator('select[name="column"]').first().selectOption('Text');
        await page.locator('input[name="prompt"]').first().fill('Original prompt');
        await page.locator('input[name="output-column-name"]').first().fill('Answer');
    }
}

async function downloadJson(page) {
    const pending = page.waitForEvent('download');
    await page.locator('[data-recovery-download]:visible').click();
    return JSON.parse(fs.readFileSync(await (await pending).path(), 'utf8'));
}

function mockAnalysis(page, calls, failAt = () => false) {
    return page.route('**/analyze_batch', async route => {
        const body = route.request().postDataJSON(); calls.push(body);
        if (failAt(body.rows[0].rowIndex)) return route.fulfill({status:503, json:{error:'Interrupted'}});
        return route.fulfill({json:{errors:0, results:body.rows.map(r => ({rowIndex:r.rowIndex,
            values:{[body.configs[0].outputColumnName]:`answer${r.rowIndex}`}}))}});
    });
}

for (const recovery of ['browser', 'file']) {
    test(`general analysis resumes from ${recovery} with original settings and ordered exports`, async ({page}) => {
        const calls = []; let fail = true;
        await mockAnalysis(page, calls, start => start === 20 && fail);
        await setup(page);
        await page.locator('#modal-api-key').evaluate(el => { el.value = 'secret-not-in-backup'; });
        await page.locator('#analyze-btn').click();
        await expect(page.locator('#analysis-recovery')).toBeVisible();
        const doc = await downloadJson(page);
        expect(doc.run.cursor).toBe(20);
        expect(doc.run.rows).toHaveLength(25);
        expect(JSON.stringify(doc)).not.toContain('secret-not-in-backup');
        if (recovery === 'file') await page.evaluate(async () => {
            const {saveCheckpoint} = await import('/static/js/run-checkpoint.js');
            await saveCheckpoint('analysis', 'delete');
        });
        await page.reload();
        await page.locator('#start-analyzing-btn').click();
        if (recovery === 'browser') await page.locator('#analysis-restore-start-btn').click();
        else await page.locator('#recovery-import-input').setInputFiles({name:'backup.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(doc))});
        await expect(page.locator('#analysis-recovery')).toBeVisible();
        expect(calls).toHaveLength(2);
        await page.locator('input[name="prompt"]').first().fill('Changed UI must not affect resume');
        fail = false;
        await page.locator('#resume-analysis-btn').click();
        await expect(page.locator('#step-5')).toBeVisible();
        expect(calls.map(c => c.rows[0].rowIndex)).toEqual([0,20,20]);
        expect(calls[2].configs).toEqual(calls[0].configs);
        expect(calls[2].generalInstructions).toBe(calls[0].generalInstructions);
        expect(calls[2].apiKey).toBeNull();
        const completed = await downloadJson(page);
        expect(completed.run.results.map(r => r.Answer)).toEqual(Array.from({length:25},(_,i)=>`answer${i}`));
        const download = page.waitForEvent('download');
        await page.locator('[data-recovery-csv]:visible').click();
        const csv = fs.readFileSync(await (await download).path(),'utf8');
        expect(csv.trim().split('\n')).toHaveLength(26);
        expect(csv).toContain('row24,24,answer24');
        const excel = page.waitForEvent('download');
        await page.locator('#download-link').click();
        const bytes = Array.from(fs.readFileSync(await (await excel).path()));
        const data = await page.evaluate(bytes => {
            const wb = XLSX.read(new Uint8Array(bytes),{type:'array'});
            return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        }, bytes);
        expect(data.map(r => r.Number)).toEqual(Array.from({length:25},(_,i)=>i));
        expect(data[24].Answer).toBe('answer24');
    });
}

test('failed storage retains a downloadable batch and retries saving before more analysis', async ({page}) => {
    await setup(page);
    await page.evaluate(() => {
        const original = IDBDatabase.prototype.transaction;
        IDBDatabase.prototype.transaction = function(...args) {
            if (window.failStorage && args[1] === 'readwrite') throw new DOMException('Storage full', 'QuotaExceededError');
            return original.apply(this,args);
        };
    });
    let calls = 0;
    await page.route('**/analyze_batch', async route => {
        calls++;
        if (calls === 1) await page.evaluate(() => { window.failStorage = true; });
        const body = route.request().postDataJSON();
        await route.fulfill({json:{errors:0,results:body.rows.map(r=>({rowIndex:r.rowIndex,values:{Answer:'ok'}}))}});
    });
    await page.locator('#analyze-btn').click();
    await expect(page.locator('#analyze-message')).toContainText('Could not save progress');
    expect(calls).toBe(1);
    expect((await downloadJson(page)).run.results).toHaveLength(20);
    await page.locator('#resume-analysis-btn').click();
    await expect(page.locator('#spinner-overlay')).toBeHidden();
    expect(calls).toBe(1);
    await page.evaluate(() => { window.failStorage = false; });
    await page.locator('#resume-analysis-btn').click();
    await expect(page.locator('#step-5')).toBeVisible();
    expect(calls).toBe(2);
    await page.reload();
    await page.locator('#start-analyzing-btn').click();
    await page.locator('#analysis-restore-start-btn').click();
    await expect(page.locator('#result-message')).toContainText('25 of 25');
});

test('Jev recovery file restores audit, settings and only unfinished rows', async ({page}) => {
    let fail = true; const calls = [];
    await page.route('**/analyze_batch_jev', async route => {
        const body = route.request().postDataJSON(); calls.push(body.rows[0].rowIndex);
        if (body.rows[0].rowIndex === 2 && fail) return route.fulfill({status:503,json:{error:'Interrupted'}});
        await route.fulfill({json:{results:body.rows.map(r=>({rowIndex:r.rowIndex, values:{Category:'A'},
            decisions:{Category:{status:'ok',value:'A',confidence:.99}},calls:[{model:'pinned-model'}]}))}});
    });
    await setup(page,5,true);
    await page.locator('#jev-run-btn').click();
    await expect(page.locator('#step-5')).toBeVisible();
    const doc = await downloadJson(page);
    expect(doc.run.cursor).toBe(2);
    expect(doc.run.audit).toHaveLength(2);
    await page.evaluate(async () => { const {checkpoint} = await import('/static/js/jev-checkpoint.js'); await checkpoint('delete'); });
    await page.reload();
    await page.locator('#start-analyzing-btn').click();
    await page.locator('#recovery-import-input').setInputFiles({name:'backup.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(doc))});
    await expect(page.locator('#jev-coverage')).toContainText('2 of 5');
    expect(calls).toEqual([0,2]);
    fail = false;
    await page.locator('#jev-continue-results').click();
    await expect(page.locator('#jev-coverage')).toContainText('5 of 5');
    expect(calls).toEqual([0,2,2,4]);
    const final = await downloadJson(page);
    expect(final.run.audit.map(a=>a.rowIndex)).toEqual([0,1,2,3,4]);
});

test('corrupt recovery cannot replace a saved run', async ({page}) => {
    await mockAnalysis(page,[]);
    await setup(page,2);
    await page.locator('#analyze-btn').click();
    await expect(page.locator('#step-5')).toBeVisible();
    const doc = await downloadJson(page);
    doc.run.cursor = 999;
    await page.reload();
    await page.locator('#start-analyzing-btn').click();
    await page.locator('#recovery-import-input').setInputFiles({name:'bad.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(doc))});
    await expect(page.locator('#recovery-import-message')).toContainText('Invalid or incomplete');
    await page.locator('#analysis-restore-start-btn').click();
    await expect(page.locator('#result-message')).toContainText('2 of 2');
});

test('1,000-row reminder offers a self-contained backup without automatic downloads', async ({page}) => {
    const calls = []; const downloads = [];
    page.on('download', download => downloads.push(download));
    await mockAnalysis(page, calls, start => start === 1000);
    await setup(page,1005);
    await page.locator('#analyze-btn').click();
    await expect(page.locator('#analysis-recovery')).toBeVisible({timeout:30000});
    await expect(page.locator('[data-recovery-reminder]:visible')).toContainText('1,000 rows saved');
    expect(downloads).toHaveLength(0);
    const doc = await downloadJson(page);
    expect(doc.run.results).toHaveLength(1000);
    expect(doc.run.rows).toHaveLength(1005);
    await expect(page.locator('[data-recovery-reminder]:visible')).toHaveCount(0);
});

test('30,000-row checkpoints restore in order; aborted writes and stale runs cannot advance saved progress', async ({page}) => {
    test.slow();
    await page.goto('/');
    const result = await page.evaluate(async () => {
        const {saveCheckpoint} = await import('/static/js/run-checkpoint.js');
        const {BatchRun} = await import('/static/js/batch-runner.js');
        const rows = Array.from({length:30000}, (_,i)=>({Text:`row${i}`}));
        const state = {runId:'large', rows, cursor:0, results:[]};
        await saveCheckpoint('analysis','start',state);
        const run = new BatchRun({rows,batchSize:20,
            processBatch:async (batch,start)=>({rows:batch.map((r,i)=>({...r,Answer:start+i}))}),
            onCheckpoint:async (run,indices)=>saveCheckpoint('analysis','put',{...state,cursor:run.cursor,results:run.results},indices)});
        await run.run();
        const restored = await saveCheckpoint('analysis','get');
        // A synchronous DataCloneError must abort both cursor and row changes.
        let aborted = false;
        try { await saveCheckpoint('analysis','put',{...state,cursor:1,results:[{bad:()=>{}}]},[0]); }
        catch { aborted = true; }
        const afterAbort = await saveCheckpoint('analysis','get');
        await saveCheckpoint('analysis','start',{runId:'replacement',rows:[],results:[],cursor:0});
        let conflict = false;
        try { await saveCheckpoint('analysis','put',state,[]); } catch { conflict = true; }
        const replacement = await saveCheckpoint('analysis','get');
        return {count:restored.cursor, ordered:restored.results.every((r,i)=>r.Text===`row${i}` && r.Answer===i),
            aborted, unchanged:afterAbort.cursor===30000 && afterAbort.results[0].Answer===0,
            conflict, replacement:replacement.runId};
    });
    expect(result).toEqual({count:30000,ordered:true,aborted:true,unchanged:true,conflict:true,replacement:'replacement'});
});
