const {test, expect} = require('@playwright/test');
const fs = require('fs');

// Walks the real UI: build a two-column configuration, export it, then import
// it against a second file whose columns only partly match.
test('exported configuration reimports onto another file and warns about missing columns', async ({page}) => {
  const errors=[]; page.on('pageerror', e=>errors.push(e.message));
  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#general-instructions').fill('Respond in French only.');
  await page.locator('#config-next-btn').click();

  const csv='Feedback,type_of_feedback,region\nEbola existe,Observations,Nord-Kivu\nMerci,Appreciations,Ituri\n';
  await page.locator('#file').setInputFiles({name:'feedback.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
  await page.locator('#upload-form button[type=submit]').click();
  await page.locator('#preview-next-btn').click();

  // First analysis: two source columns combined, long output.
  await page.locator('select[name="column"]').first().selectOption('Feedback');
  await page.locator('.add-column-btn-small').first().click();
  await page.locator('.additional-column select').first().selectOption('type_of_feedback');
  await page.locator('input[name="output-column-name"]').first().fill('feedback_code');
  await page.locator('input[name="prompt"]').first().fill('Code the combination; do not explain.');
  await page.locator('[name="max-output-tokens"]').first().selectOption('4096');

  // Second analysis on the same columns, different result column.
  await page.locator('#add-column-btn').click();
  await page.locator('select[name="column"]').nth(1).selectOption('Feedback');
  await page.locator('input[name="output-column-name"]').nth(1).fill('feedback_dimension');
  await page.locator('input[name="prompt"]').nth(1).fill('Pick one dimension.');

  const download = page.waitForEvent('download');
  await page.locator('#export-config-btn').click();
  const file = await download;
  expect(file.suggestedFilename()).toMatch(/^aidstack-config-\d{4}-\d{2}-\d{2}\.json$/);
  const saved = await file.path();
  const doc = JSON.parse(fs.readFileSync(saved, 'utf8'));
  expect(doc.generalInstructions).toBe('Respond in French only.');
  expect(doc.columns).toHaveLength(2);
  expect(doc.columns[0].columns).toEqual(['Feedback','type_of_feedback']);
  expect(doc.columns[0].maxOutputTokens).toBe(4096);

  // Start over with a file that lacks type_of_feedback.
  await page.reload();
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#config-next-btn').click();
  await page.locator('#file').setInputFiles({name:'other.csv',mimeType:'text/csv',
    buffer:Buffer.from('Feedback,country\nBonjour,DRC\n')});
  await page.locator('#upload-form button[type=submit]').click();
  await page.locator('#preview-next-btn').click();

  await page.locator('#import-config-input').setInputFiles(saved);
  await expect(page.locator('#analyze-message')).toContainText('Imported 2 column analyses');
  await expect(page.locator('#analyze-message')).toContainText('type_of_feedback');
  await expect(page.locator('#general-instructions')).toHaveValue('Respond in French only.');
  await expect(page.locator('.column-selection-container')).toHaveCount(2);
  // The matching column is selected; the missing one is simply not added.
  await expect(page.locator('select[name="column"]').first()).toHaveValue('Feedback');
  await expect(page.locator('.additional-column select')).toHaveCount(0);
  await expect(page.locator('input[name="prompt"]').first())
    .toHaveValue('Code the combination; do not explain.');
  await expect(page.locator('[name="max-output-tokens"]').first()).toHaveValue('4096');
  await expect(page.locator('input[name="output-column-name"]').nth(1)).toHaveValue('feedback_dimension');

  // A file that is not a configuration is rejected, leaving the form intact.
  await page.locator('#import-config-input').setInputFiles(
    {name:'junk.json',mimeType:'application/json',buffer:Buffer.from('{"hello":1}')});
  await expect(page.locator('#analyze-message')).toContainText('not an Aidstack');
  await expect(page.locator('.column-selection-container')).toHaveCount(2);

  if(errors.length) throw Error(errors.join('; '));
});
