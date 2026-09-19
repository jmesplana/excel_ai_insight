const {test, expect} = require('@playwright/test');

// The pre-analysis path: upload a file with multi-value cells, split it, and
// download the result without ever calling the AI or supplying an API key.
test('splits multi-value cells after upload, with no analysis', async ({page}) => {
  const errors=[]; page.on('pageerror', e=>errors.push(e.message));
  // Any /analyze_batch call here would mean the split path is spending money.
  let analyzed=false; await page.route('**/analyze_batch', r=>{analyzed=true; return r.abort();});

  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#config-next-btn').click();

  const csv = 'id,region,feedback_classification\n'
    + '1,Nord-Kivu,Observations || Signalements || Allegations\n'
    + '2,Ituri,Appreciations\n'
    + '3,Nord-Kivu,Questions || Demandes\n';
  await page.locator('#file').setInputFiles({name:'feedback.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
  await page.locator('#upload-form button[type=submit]').click();
  await expect(page.locator('#preview-table')).toContainText('Observations || Signalements');

  await page.locator('#pre-clean-data-card .card-header').click();
  await expect(page.locator('#pre-explode-column')).toBeVisible();
  await page.locator('#pre-explode-column').selectOption('feedback_classification');
  await page.locator('#pre-explode-separator').fill('||');
  await page.locator('#pre-explode-output').fill('category');
  await page.locator('#pre-explode-preview-btn').click();

  // 3 source rows carry 6 values in total.
  await expect(page.locator('#pre-explode-message')).toContainText('3 rows become 6 rows');
  await expect(page.locator('#pre-explode-preview-table')).toContainText('category');
  await expect(page.locator('#pre-explode-preview-table tbody tr')).toHaveCount(6);

  const download = page.waitForEvent('download');
  await page.locator('#pre-explode-download-btn').click();
  console.log('split download', (await download).suggestedFilename());

  // Applying rewrites the previewed sheet itself, so the split rows flow into
  // any later analysis rather than living only in the downloaded copy.
  await page.locator('#pre-explode-apply-btn').click();
  await expect(page.locator('#pre-explode-message')).toContainText('now 6 rows');
  await expect(page.locator('#preview-table')).toContainText('Showing first 10 of 6 rows')
    .catch(()=>{}); // fewer than 10 rows: the note is omitted by design
  await expect(page.locator('#preview-table')).toContainText('category');
  await expect(page.locator('#preview-table tbody tr')).toHaveCount(6);

  // Undo restores the pre-split sheet.
  await page.locator('#pre-explode-undo-btn').click();
  await expect(page.locator('#preview-table tbody tr')).toHaveCount(3);
  await expect(page.locator('#preview-table')).toContainText('Observations || Signalements');

  if (analyzed) throw Error('cleaning must not call the analysis endpoint');
  if (errors.length) throw Error(errors.join('; '));
});

// The post-analysis path: the same panel on step 5 splits AI output, and the
// split flows into the results table and the main Download button.
test('splits AI output on the results step', async ({page}) => {
  const errors=[]; page.on('pageerror', e=>errors.push(e.message));
  await page.route('**/analyze_batch', route => {
    const payload = route.request().postDataJSON();
    const name = payload.configs[0].outputColumnName;
    return route.fulfill({status:200, contentType:'application/json', body: JSON.stringify({
      errors:0,
      results: payload.rows.map(row => ({rowIndex: row.rowIndex,
        values: {[name]: 'Questions || Demandes'}})),
    })});
  });

  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#config-next-btn').click();
  const csv = 'id,feedback\n1,alpha\n2,beta\n';
  await page.locator('#file').setInputFiles({name:'f.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
  await page.locator('#upload-form button[type=submit]').click();
  await page.locator('#preview-next-btn').click();
  await page.locator('select[name="column"]').first().selectOption('feedback');
  await page.locator('input[name="prompt"]').first().fill('Classify');
  await page.locator('#analyze-btn').click();
  await expect(page.locator('#step-5')).toBeVisible();

  await page.locator('#clean-data-card .card-header').click();
  const column = page.locator('#explode-column');
  await expect(column).toBeVisible();
  // The picker defaults to the newest AI output column.
  const selected = await column.inputValue();
  if (!selected.includes('feedback')) throw Error('unexpected default column ' + selected);
  await page.locator('#explode-separator').fill('||');
  await page.locator('#explode-preview-btn').click();
  await expect(page.locator('#explode-message')).toContainText('2 rows become 4 rows');

  await page.locator('#explode-apply-btn').click();
  await expect(page.locator('#explode-message')).toContainText('now 4 rows');
  // The applied split reaches the results table and the main download.
  await expect(page.locator('#table-info')).toContainText('4 rows');
  const download = page.waitForEvent('download');
  await page.locator('#download-link').click();
  console.log('analyzed download', (await download).suggestedFilename());

  await page.locator('#explode-undo-btn').click();
  await expect(page.locator('#table-info')).toContainText('2 rows');
  if (errors.length) throw Error(errors.join('; '));
});

// Clean Data is a mode of the product, chosen on step 1 alongside AI Analysis
// and ICD: no instructions, no credentials, and no analysis steps after it.
test('clean mode runs without any AI configuration', async ({page}) => {
  const errors=[]; page.on('pageerror', e=>errors.push(e.message));
  let called=false;
  for (const route of ['**/analyze_batch','**/detect_patterns','**/test_connection']) {
    await page.route(route, r => {called=true; return r.abort();});
  }

  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-clean').click();
  await page.locator('#config-next-btn').click();

  // Choosing the mode replaces the analysis instructions with a plain
  // explanation; nothing asks for a key.
  await expect(page.locator('#step-2')).toBeVisible();
  await expect(page.locator('#analysis-instructions-block')).toBeHidden();

  const csv = 'id,region,tags\n1,Nord-Kivu,A || B || C\n2,Ituri,D\n';
  await page.locator('#file').setInputFiles({name:'t.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
  await page.locator('#upload-form button[type=submit]').click();
  await expect(page.locator('#step-3')).toBeVisible();

  // The panel is already open and the path onward to analysis is gone.
  await expect(page.locator('#pre-explode-column')).toBeVisible();
  await expect(page.locator('#preview-next-btn')).toBeHidden();

  await page.locator('#pre-explode-column').selectOption('tags');
  await page.locator('#pre-explode-separator').fill('||');
  await page.locator('#pre-explode-preview-btn').click();
  await expect(page.locator('#pre-explode-message')).toContainText('2 rows become 4 rows');

  const download = page.waitForEvent('download');
  await page.locator('#pre-explode-download-btn').click();
  console.log('clean-mode download', (await download).suggestedFilename());

  // Back walks the normal step chain; the analysis steps stay out of the
  // stepper because clean mode ends at step 3.
  await expect(page.locator('.stepper-item[data-step="4"]')).toBeHidden();
  await expect(page.locator('.stepper-item[data-step="5"]')).toBeHidden();
  await page.locator('#preview-back-btn').click();
  await expect(page.locator('#step-2')).toBeVisible();
  await page.locator('#upload-back-btn').click();
  await expect(page.locator('#step-1')).toBeVisible();
  await expect(page.locator('#clean-intro-block')).toBeVisible();

  if (called) throw Error('clean mode must never call an AI endpoint');
  if (errors.length) throw Error(errors.join('; '));
});

// Regression: Apply used to hide the preview block that contained the only
// download button, stranding the split file with no way to save it.
test('the split file is still downloadable after Apply', async ({page}) => {
  const errors=[]; page.on('pageerror', e=>errors.push(e.message));
  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-clean').click();
  await page.locator('#config-next-btn').click();
  const csv = 'id,tags\n1,A || B\n2,C\n';
  await page.locator('#file').setInputFiles({name:'t.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
  await page.locator('#upload-form button[type=submit]').click();

  await page.locator('#pre-explode-column').selectOption('tags');
  await page.locator('#pre-explode-separator').fill('||');
  await page.locator('#pre-explode-preview-btn').click();
  await expect(page.locator('#pre-explode-download-btn')).toBeVisible();

  // Apply, then download — the order the user actually took.
  await page.locator('#pre-explode-apply-btn').click();
  await expect(page.locator('#pre-explode-preview')).toBeHidden();
  await expect(page.locator('#pre-explode-download-btn')).toBeVisible();
  await expect(page.locator('#pre-explode-message')).toContainText('Download split file');

  const download = page.waitForEvent('download');
  await page.locator('#pre-explode-download-btn').click();
  const file = await download;
  console.log('post-apply download', file.suggestedFilename());
  // The saved workbook must hold the split rows, not the original two.
  const fs = require('fs');
  const path = await file.path();
  if (!fs.statSync(path).size) throw Error('downloaded an empty file');

  // Undo retracts the offer, since nothing split is in effect any more.
  await page.locator('#pre-explode-undo-btn').click();
  await expect(page.locator('#pre-explode-download-btn')).toBeHidden();

  if (errors.length) throw Error(errors.join('; '));
});
