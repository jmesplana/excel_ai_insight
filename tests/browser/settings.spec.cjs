const {test, expect} = require('@playwright/test');
const fs = require('fs');

async function open(page) {
  if (!await page.locator('#api-settings-btn').isVisible()) await page.getByRole('button', {name:'Toggle navigation'}).click();
  await page.locator('#api-settings-btn').click();
}
async function save(page) { await page.locator('#save-api-settings').click(); await expect(page.locator('#api-settings-modal')).toBeHidden(); }

test('independent keys survive switching, saving and reload; requests send only selected profile', async ({page}) => {
  const errors=[]; page.on('pageerror', e=>errors.push(e.message));
  await page.goto('/'); await open(page);
  await page.locator('#modal-api-key').fill('openai-test-secret');
  await page.locator('label[for=provider-azure]').click();
  await page.locator('#modal-azure-api-key').fill('azure-test-secret');
  await page.locator('#modal-azure-endpoint').fill('https://example.openai.azure.com');
  await page.locator('#modal-azure-deployment').fill('my-deployment');
  await page.locator('#settings-tab-classify').click();
  await page.locator('#modal-jev-api-key').fill('jev-test-secret');
  await page.locator('#modal-save-api-key').check(); await save(page);
  await page.reload(); await open(page);
  await expect(page.locator('#modal-azure-api-key')).toHaveValue('azure-test-secret');
  await page.locator('label[for=provider-openai]').click();
  await expect(page.locator('#modal-api-key')).toHaveValue('openai-test-secret');
  await page.locator('#settings-tab-classify').click();
  await expect(page.locator('#modal-jev-api-key')).toHaveValue('jev-test-secret');
  await page.locator('#settings-tab-text').click();
  await page.locator('label[for=provider-ollama]').click();
  await page.route('**/ollama_models', route => route.fulfill({json:{models:['gemma3:1b']}}));
  await page.locator('#ollama-models-btn').click();
  await expect(page.locator('#modal-ollama-model')).toHaveValue('gemma3:1b');
  let request;
  await page.route('**/test_connection', route => {request=route.request().postDataJSON(); return route.fulfill({json:{ok:true,provider:'Ollama',model:'gemma3:1b'}});});
  await page.locator('#test-connection-btn').click();
  await expect(page.locator('#test-connection-result')).toContainText('Connected to Ollama');
  expect(request).toEqual({provider:'ollama',ollamaEndpoint:null,ollamaModel:'gemma3:1b'});
  await save(page); await page.reload(); await open(page);
  await expect(page.locator('#provider-ollama')).toBeChecked();
  await page.locator('label[for=provider-openai]').click();
  await page.locator('#modal-api-key').fill(''); await save(page); await page.reload(); await open(page);
  await expect(page.locator('#modal-api-key')).toHaveValue('');
  await page.locator('label[for=provider-azure]').click();
  await expect(page.locator('#modal-azure-api-key')).toHaveValue('azure-test-secret');
  expect(errors).toEqual([]);
});

test('legacy Azure key migrates only to Azure; session-only values survive reopening but not reload', async ({page}) => {
  await page.goto('/');
  await page.evaluate(()=>{localStorage.setItem('excel_ai_insight_api_key','legacy-azure');localStorage.setItem('excel_ai_insight_provider','azure');});
  await page.reload(); await open(page);
  await expect(page.locator('#modal-azure-api-key')).toHaveValue('legacy-azure');
  await page.locator('label[for=provider-openai]').click();
  await expect(page.locator('#modal-api-key')).toHaveValue('');
  await page.locator('#modal-api-key').fill('session-secret');
  await page.locator('#modal-save-api-key').uncheck(); await save(page); await open(page);
  await expect(page.locator('#modal-api-key')).toHaveValue('session-secret');
  await save(page); await page.reload(); await open(page);
  await expect(page.locator('#modal-api-key')).toHaveValue('');
  expect(await page.evaluate(()=>localStorage.getItem('excel_ai_insight_azure_api_key'))).toBeNull();
});

test('offline interface and Laya configuration export/import and run', async ({page}) => {
  const external=[]; const errors=[]; page.on('pageerror', e=>errors.push(e.message));
  await page.route('**/*', route => {
    if (route.request().url() === 'http://127.0.0.1:8001/validate_jev_config') return route.fulfill({json:{ok:true}});
    if (!route.request().url().startsWith('http://127.0.0.1:8091')) {external.push(route.request().url());return route.abort();}
    return route.continue();
  });
  await page.goto('/'); await open(page);
  await page.locator('#settings-tab-classify').click();
  await page.locator('label[for=classifier-laya]').click();
  await page.locator('#modal-laya-api-key').fill('local-secret');
  await save(page);
  await page.locator('#start-analyzing-btn').click(); await page.locator('#mode-card-jev').click();
  await page.locator('#config-next-btn').click();
  await page.locator('#file').setInputFiles({name:'feedback.csv',mimeType:'text/csv',buffer:Buffer.from('Feedback\nBonjour\n')});
  await page.locator('#upload-form button[type=submit]').click(); await page.locator('#preview-next-btn').click();
  await page.locator('[name=jev-output-column-name]').first().fill('Category');
  await page.locator('[name=jev-instructions]').first().fill('Choose a category');
  await page.locator('.jev-options').first().fill('A\nB');
  const download=page.waitForEvent('download'); await page.locator('#jev-export-config-btn').click();
  const path=await (await download).path(); const doc=JSON.parse(fs.readFileSync(path,'utf8'));
  expect(doc.classifier).toEqual({provider:'laya',model:'multilingual'});
  expect(JSON.stringify(doc)).not.toContain('local-secret');
  await page.locator('.jev-options').first().fill('C\nD');
  await page.locator('#jev-import-config-input').setInputFiles(path);
  await expect(page.locator('.jev-options').first()).toHaveValue('A\nB');
  let payload;
  await page.route('http://127.0.0.1:8001/analyze_batch_jev', route=>{payload=route.request().postDataJSON();return route.fulfill({json:{results:[{rowIndex:0,values:{Category:'A'},decisions:{},calls:[]}],errors:0}});});
  await page.locator('#jev-run-btn').click();
  await expect(page.locator('#result-message')).toContainText('Classification complete');
  expect(payload.classificationProvider).toBe('laya'); expect(payload.jevApiKey).toBeUndefined();
  expect(payload.layaModel).toBe('multilingual');
  expect(external).toEqual([]); expect(errors).toEqual([]);
});

test('settings fits narrow screens', async ({page}) => {
  await page.setViewportSize({width:390,height:844}); await page.goto('/'); await open(page);
  await page.locator('#settings-tab-classify').click(); await page.locator('label[for=classifier-laya]').click();
  await expect(page.locator('#save-api-settings')).toBeVisible();
  expect(await page.evaluate(()=>document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  await expect(page.locator('#settings-classify')).toHaveCSS('opacity', '1');
  await page.screenshot({path:'test-results/settings-mobile.png', animations:'disabled'});
  await page.setViewportSize({width:1280,height:900});
  await page.screenshot({path:'test-results/settings-desktop.png', animations:'disabled'});
});
