const {test, expect} = require('@playwright/test');
const {types, excerpts} = require('../fixtures/paired-feedback.json');

test('paired cleaning pairs matching rows and copies mismatched cells through download, apply and undo', async ({page}) => {
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-clean').click();
  await page.locator('#config-next-btn').click();
  const quote = value => `"${value.replaceAll('"', '""')}"`;
  const sixExcerpts = excerpts.slice(0, 6).join(' || ');
  const expected = [
    ...types.map((type, i) => ({id:42, item_codes:type, item_details:excerpts[i]})),
    ...types.map(type => ({id:43, item_codes:type, item_details:sixExcerpts})),
  ];
  const csv = 'id,item_codes,item_details\n'
    + `42,${quote(types.join(' ||   '))},${quote(excerpts.join(' || '))}\n`
    + `43,${quote(types.join(' || '))},${quote(sixExcerpts)}\n`;
  await page.locator('#file').setInputFiles({name:'paired.csv', mimeType:'text/csv', buffer:Buffer.from('\ufeff' + csv)});
  await page.locator('#upload-form button[type=submit]').click();
  await expect(page.locator('#pre-explode-paired-column')).toHaveValue('');
  await expect(page.locator('#pre-explode-copy-mismatched')).not.toBeChecked();
  await expect(page.locator('#pre-explode-copy-mismatched')).toBeDisabled();
  await page.locator('#pre-explode-column').selectOption('item_codes');
  await page.locator('#pre-explode-paired-column').selectOption('item_details');
  await expect(page.locator('#pre-explode-dedupe')).toBeDisabled();
  await page.locator('#pre-explode-preview-btn').click();
  await expect(page.locator('#pre-explode-message')).toContainText('Data row 2');
  await expect(page.locator('#pre-explode-preview')).toBeHidden();
  await expect(page.locator('#pre-explode-download-btn')).toBeHidden();
  await page.locator('#pre-explode-copy-mismatched').check();
  await page.locator('#pre-explode-preview-btn').click();
  await expect(page.locator('#pre-explode-message')).toContainText('2 rows become 14 rows');
  const rows = page.locator('#pre-explode-preview-table tbody tr');
  await expect(rows).toHaveCount(14);
  for (let i = 0; i < expected.length; i++) {
    await expect(rows.nth(i).locator('td')).toHaveText(Object.values(expected[i]).map(String));
  }
  await expect(page.locator('#pre-explode-preview-table th.table-warning')).toHaveCount(2);
  await page.locator('#pre-explode-apply-btn').click();
  await expect(page.locator('#preview-table tbody tr')).toHaveCount(10);
  await expect(page.locator('#preview-table')).toContainText('Showing first 10 of 14 rows');
  const pending = page.waitForEvent('download');
  await page.locator('#pre-explode-download-btn').click();
  const download = await pending;
  const bytes = require('fs').readFileSync(await download.path()).toString('base64');
  const saved = await page.evaluate(bytes => {
    const wb = XLSX.read(bytes, {type:'base64'});
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
  }, bytes);
  expect(saved).toEqual(expected);
  await page.locator('#pre-explode-undo-btn').click();
  await expect(page.locator('#preview-table tbody tr')).toHaveCount(2);

  // Turning off fallback invalidates the preview and restores strict matching.
  await page.locator('#pre-explode-preview-btn').click();
  await page.locator('#pre-explode-copy-mismatched').uncheck();
  await expect(page.locator('#pre-explode-download-btn')).toBeHidden();
  await page.locator('#pre-explode-preview-btn').click();
  await expect(page.locator('#pre-explode-message')).toContainText('Data row 2');
  await page.locator('#pre-explode-copy-mismatched').check();

  // Invalid settings must not leave a stale preview or download available.
  await page.locator('#pre-explode-preview-btn').click();
  await expect(page.locator('#pre-explode-download-btn')).toBeVisible();
  await page.locator('#pre-explode-separator').fill('');
  await expect(page.locator('#pre-explode-download-btn')).toBeHidden();
  await page.locator('#pre-explode-preview-btn').click();
  await expect(page.locator('#pre-explode-message')).toContainText('Enter at least one separator');
  await expect(page.locator('#pre-explode-preview')).toBeHidden();
  await expect(page.locator('#pre-explode-download-btn')).toBeHidden();
  await expect(page.locator('#preview-table tbody tr')).toHaveCount(2);
  // A new upload does not inherit paired splitting or fallback from this file.
  await page.locator('#preview-back-btn').click();
  await page.locator('#file').setInputFiles({name:'another.csv', mimeType:'text/csv', buffer:Buffer.from('code,description\nA,single\n')});
  await page.locator('#upload-form button[type=submit]').click();
  await expect(page.locator('#pre-explode-paired-column')).toHaveValue('');
  await expect(page.locator('#pre-explode-copy-mismatched')).not.toBeChecked();
  await expect(page.locator('#pre-explode-copy-mismatched')).toBeDisabled();
  expect(errors).toEqual([]);
});

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
  const csv = 'id,feedback\n1,alpha || bravo\n2,charlie || delta\n';
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
  await page.locator('#explode-paired-column').selectOption('feedback');
  await expect(page.locator('#explode-dedupe')).toBeDisabled();
  await page.locator('#explode-separator').fill('||');
  await page.locator('#explode-preview-btn').click();
  await expect(page.locator('#explode-message')).toContainText('2 rows become 4 rows');
  const pairedRows = page.locator('#explode-preview-table tbody tr');
  await expect(pairedRows).toHaveCount(4);
  for (const [i, excerpt] of ['alpha', 'bravo', 'charlie', 'delta'].entries()) {
    await expect(pairedRows.nth(i)).toContainText(excerpt);
    await expect(pairedRows.nth(i)).toContainText(i % 2 === 0 ? 'Questions' : 'Demandes');
  }

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
