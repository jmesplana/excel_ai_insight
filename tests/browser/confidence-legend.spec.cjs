const {test, expect} = require('@playwright/test');

const CSV = 'Feedback\nEst-ce que le vaccin viendra ?\nMerci pour les informations\n';

const CONFIG = (includeConfidence) => ({
  format: 'aidstack-insights-jev-config',
  version: 1,
  sheetName: 'Sheet1',
  sourceColumns: ['Feedback'],
  includeConfidence,
  questions: [{
    outputColumnName: 'framework2_type',
    questionType: 'choice',
    instructions: 'Choisis le type.',
    options: ['Questions', 'Appréciations'],
  }],
});

function stubJev(page) {
  return page.route('**/analyze_batch_jev', async route => {
    const body = route.request().postDataJSON();
    const results = body.rows.map(row => {
      const values = {};
      body.configs.forEach(c => {
        values[c.outputColumnName] = c.options[0];
        if (body.includeConfidence) values[`${c.outputColumnName}__confidence`] = 0.42;
      });
      return {rowIndex: row.rowIndex, values};
    });
    await route.fulfill({json: {results, errors: 0}});
  });
}

async function runJevTest(page, includeConfidence) {
  await stubJev(page);
  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-jev').click();
  await page.locator('#config-next-btn').click();
  await page.locator('#file').setInputFiles(
    {name: 'feedback.csv', mimeType: 'text/csv', buffer: Buffer.from(CSV)});
  await page.locator('#upload-form button[type=submit]').click();
  await page.locator('#preview-next-btn').click();
  await page.locator('#jev-import-config-input').setInputFiles(
    {name: 'cfg.json', mimeType: 'application/json',
     buffer: Buffer.from(JSON.stringify(CONFIG(includeConfidence)))});
  await page.locator('#jev-source-columns input[value="Feedback"]').check();
  await page.locator('#jev-test-run-btn').click();
  await expect(page.locator('.test-run-preview-panel')).toBeVisible({timeout: 20000});
}

test('the test run explains the confidence columns when they are present', async ({page}) => {
  await runJevTest(page, true);

  const panel = page.locator('.test-run-preview-panel');
  await expect(panel).toContainText('Reading the confidence columns');
  // The point users get wrong: it is not a probability of being right.
  await expect(panel).toContainText('not');
  await expect(panel).toContainText('below 0.5');
  await expect(panel).toContainText('above 0.9');
});

test('the explanation is absent when no confidence columns were produced', async ({page}) => {
  await runJevTest(page, false);

  const panel = page.locator('.test-run-preview-panel');
  await expect(panel).toBeVisible();
  await expect(panel).not.toContainText('Reading the confidence columns');
});
