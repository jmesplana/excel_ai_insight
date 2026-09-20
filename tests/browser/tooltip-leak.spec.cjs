const {test, expect} = require('@playwright/test');

const CSV = 'Feedback\nEst-ce que le vaccin viendra ?\nMerci pour les informations\n';

// A config with several questions, so importing it rebuilds the card list
// repeatedly -- the path that used to stack duplicate Tooltip instances.
const CONFIG = {
  format: 'aidstack-insights-jev-config',
  version: 1,
  sheetName: 'Sheet1',
  sourceColumns: ['Feedback'],
  includeConfidence: true,
  questions: [1, 2, 3, 4, 5].map(n => ({
    outputColumnName: `col_${n}`,
    questionType: 'choice',
    instructions: `Question ${n}`,
    options: ['Alpha', 'Beta'],
  })),
};

async function openJevConfig(page) {
  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-jev').click();
  await page.locator('#config-next-btn').click();
  await page.locator('#file').setInputFiles(
    {name: 'feedback.csv', mimeType: 'text/csv', buffer: Buffer.from(CSV)});
  await page.locator('#upload-form button[type=submit]').click();
  await page.locator('#preview-next-btn').click();
}

test('importing a configuration leaves no orphaned tooltip on screen', async ({page}) => {
  await openJevConfig(page);

  const importBtn = page.locator('#jev-import-config-btn');
  await expect(importBtn).toBeVisible();

  // Hover the Import button to show its tooltip, as a user reaching for it does.
  await importBtn.hover();
  await expect(page.locator('.tooltip.show')).toHaveCount(1);

  // Import the configuration: this wipes and rebuilds every question card.
  await page.locator('#jev-import-config-input').setInputFiles(
    {name: 'cfg.json', mimeType: 'application/json',
     buffer: Buffer.from(JSON.stringify(CONFIG))});
  await expect(page.locator('.jev-question-container')).toHaveCount(5);

  // Move the pointer away. Every tooltip must go with it.
  await page.mouse.move(0, 0);
  await page.locator('#jev-include-confidence').hover();

  await expect(page.locator('.tooltip.show')).toHaveCount(0);
});

test('each tooltip trigger holds exactly one Bootstrap instance after import', async ({page}) => {
  await openJevConfig(page);

  await page.locator('#jev-import-config-input').setInputFiles(
    {name: 'cfg.json', mimeType: 'application/json',
     buffer: Buffer.from(JSON.stringify(CONFIG))});
  await expect(page.locator('.jev-question-container')).toHaveCount(5);

  // Import a second time: the old failure mode stacked another instance per call.
  await page.locator('#jev-import-config-input').setInputFiles(
    {name: 'cfg.json', mimeType: 'application/json',
     buffer: Buffer.from(JSON.stringify(CONFIG))});
  await expect(page.locator('.jev-question-container')).toHaveCount(5);

  // Showing then hiding must dispose cleanly. hide() animates, so wait for the
  // popup to actually leave the DOM rather than checking synchronously.
  await page.evaluate(() => {
    const el = document.getElementById('jev-import-config-btn');
    const tip = bootstrap.Tooltip.getOrCreateInstance(el);
    tip.show();
    tip.hide();
  });
  await expect(page.locator('.tooltip')).toHaveCount(0);
});
