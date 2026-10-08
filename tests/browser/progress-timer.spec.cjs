const {test, expect} = require('@playwright/test');

async function setup(page, mode, count) {
    await page.goto('/');
    await page.locator('#start-analyzing-btn').click();
    if (mode !== 'analysis') await page.locator(`#mode-card-${mode}`).click();
    await page.locator('#config-next-btn').click();
    const csv = 'Text\n' + Array.from({length:count}, (_, i) => `row ${i}`).join('\n');
    await page.locator('#file').setInputFiles({name:'timing.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
    await page.locator('#upload-form button[type=submit]').click();
    await page.locator('#preview-next-btn').click();
    if (mode === 'jev') {
        const config = {format:'aidstack-insights-jev-config',version:2,sourceColumns:['Text'],
            questions:[{outputColumnName:'Category',questionType:'choice',instructions:'Classify',options:['A','B']}],
            execution:{batchSize:2}};
        await page.locator('#jev-import-config-input').setInputFiles({name:'timer.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
    } else if (mode === 'analysis') {
        await page.locator('select[name="column"]').first().selectOption('Text');
        await page.locator('input[name="prompt"]').first().fill('Classify');
    } else {
        await page.evaluate(() => {
            document.getElementById('modal-icd-client-id').value = 'test-client';
            document.getElementById('modal-icd-client-secret').value = 'test-secret';
        });
    }
    await page.clock.install();
}

async function holdRequests(page, url) {
    const pending = [];
    await page.route(url, route => new Promise(resolve => pending.push({route, resolve})));
    return async function take() {
        await expect.poll(() => pending.length).toBeGreaterThan(0);
        const {route, resolve} = pending.shift();
        return {body:route.request().postDataJSON(), respond:async response => {
            await route.fulfill(response); resolve();
        }};
    };
}

function jevResponse(body, failed = false) {
    return {json:{results:body.rows.map(row => ({rowIndex:row.rowIndex,values:{Category:failed ? '' : 'A'},
        decisions:{Category:{status:failed ? 'error' : 'ok',value:failed ? null : 'A'}}})),errors:failed ? body.rows.length : 0}};
}

test('classification timer follows batches, pause, resume and the retry subset', async ({page}) => {
    const errors = []; page.on('pageerror', error => errors.push(error.message));
    await setup(page, 'jev', 6);
    const take = await holdRequests(page, '**/analyze_batch_jev');
    await page.locator('#jev-run-btn').click();
    let request = await take();
    await expect(page.locator('#operations-count')).toHaveText('0/6');
    await expect(page.locator('#progress-remaining')).toHaveText('Estimating…');
    await page.clock.fastForward(6000);
    await expect(page.locator('#progress-elapsed')).not.toHaveText('0s');
    await request.respond(jevResponse(request.body, true));
    request = await take();
    await expect(page.locator('#operations-count')).toHaveText('2/6');
    await expect(page.locator('#progress-remaining')).toContainText('About');
    await page.locator('#stop-analysis-btn').click();
    await request.respond(jevResponse(request.body));
    await expect(page.locator('#spinner-overlay')).toBeHidden();
    const elapsed = await page.locator('#progress-elapsed').textContent();
    await page.clock.fastForward(60000);
    await expect(page.locator('#progress-elapsed')).toHaveText(elapsed);

    await page.locator('#jev-continue-results').click();
    request = await take();
    await expect(page.locator('#operations-count')).toHaveText('4/6');
    await expect(page.locator('#progress-remaining')).toHaveText('Estimating…');
    await expect(page.locator('#progress-elapsed')).toHaveText('0s');
    await request.respond(jevResponse(request.body));
    await expect(page.locator('#spinner-overlay')).toBeHidden();

    await page.locator('#jev-retry-results').click();
    request = await take();
    await expect(page.locator('#operations-count')).toHaveText('0/2');
    await expect(page.locator('#progress-percentage')).toHaveText('0%');
    await expect(page.locator('#progress-remaining')).toHaveText('Estimating…');
    await request.respond(jevResponse(request.body));
    await expect(page.locator('#spinner-overlay')).toBeHidden();
    await expect(page.locator('#operations-count')).toHaveText('2/2');
    expect(errors).toEqual([]);
});

test('general analysis timer resets its rate after a failed batch and resume', async ({page}) => {
    await setup(page, 'analysis', 25);
    const take = await holdRequests(page, '**/analyze_batch');
    const success = body => ({json:{errors:0,results:body.rows.map(row => ({rowIndex:row.rowIndex,
        values:{[body.configs[0].outputColumnName]:'A'}}))}});
    await page.locator('#analyze-btn').click();
    let request = await take();
    await expect(page.locator('#operations-count')).toHaveText('0/25');
    await expect(page.locator('#progress-remaining')).toHaveText('Estimating…');
    await page.clock.fastForward(6000);
    await request.respond(success(request.body));
    request = await take();
    await expect(page.locator('#operations-count')).toHaveText('20/25');
    await expect(page.locator('#progress-remaining')).toContainText('About');
    await request.respond({status:503,json:{error:'Temporary failure'}});
    await expect(page.locator('#spinner-overlay')).toBeHidden();
    await page.clock.fastForward(60000);
    await page.locator('#resume-analysis-btn').click();
    request = await take();
    await expect(page.locator('#operations-count')).toHaveText('20/25');
    await expect(page.locator('#progress-remaining')).toHaveText('Estimating…');
    await request.respond(success(request.body));
    await expect(page.locator('#spinner-overlay')).toBeHidden();
    await expect(page.locator('#result-message')).toContainText('Analysis complete');
});

test('ICD shows elapsed during validation and an ETA during row translation', async ({page}) => {
    await setup(page, 'icd', 25);
    const validate = await holdRequests(page, '**/icd_validate');
    const translate = await holdRequests(page, '**/icd_translate_batch');
    await page.locator('#run-icd-btn').click();
    const validation = await validate();
    await expect(page.locator('#progress-timing')).toBeVisible();
    await expect(page.locator('#progress-remaining-row')).toBeHidden();
    await page.clock.fastForward(6000);
    await expect(page.locator('#progress-elapsed')).not.toHaveText('0s');
    await validation.respond({json:{ok:true}});
    let request = await translate();
    await expect(page.locator('#operations-count')).toHaveText('0/25');
    await expect(page.locator('#progress-remaining')).toHaveText('Estimating…');
    await page.clock.fastForward(6000);
    await request.respond({json:{results:request.body.rows.map(row => ({rowIndex:row.rowIndex})),errors:0}});
    request = await translate();
    await expect(page.locator('#operations-count')).toHaveText('20/25');
    await expect(page.locator('#progress-remaining')).toContainText('About');
    await request.respond({json:{results:request.body.rows.map(row => ({rowIndex:row.rowIndex})),errors:0}});
    await expect(page.locator('#spinner-overlay')).toBeHidden();
    await expect(page.locator('#icd-result-message')).toContainText('Translation complete');
});
