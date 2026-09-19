// The Test Run row count is user-settable per mode; these cover the three
// ways the number box can be filled in: exact, too large, and empty.
const {test, expect} = require('@playwright/test');

// 12 data rows, so a default of 5 and a custom 3 / 50 are all distinguishable.
const CSV = 'Feedback,type_of_feedback\n'
  + Array.from({length:12},(_,i)=>`ligne ${i+1},Signalements`).join('\n') + '\n';

function stubJev(page, seen) {
  return page.route('**/analyze_batch_jev', async route => {
    const body = route.request().postDataJSON();
    body.rows.forEach(r => seen.add(r.rowIndex));
    await route.fulfill({json:{results: body.rows.map(r=>({rowIndex:r.rowIndex,
      values:Object.fromEntries(body.configs.map(c=>[c.outputColumnName,c.options[0]]))})), errors:0}});
  });
}

async function setup(page) {
  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-jev').click();
  await page.locator('#config-next-btn').click();
  await page.locator('#file').setInputFiles({name:'f.csv',mimeType:'text/csv',buffer:Buffer.from(CSV)});
  await page.locator('#upload-form button[type=submit]').click();
  await page.locator('#preview-next-btn').click();
  await page.locator('[name="jev-output-column-name"]').first().fill('type');
  await page.locator('[name="jev-instructions"]').first().fill('Classer.');
  await page.locator('.jev-options').first().fill('A\nB');
}

test('the test run uses the number of rows the user asks for', async ({page}) => {
  const errors=[]; page.on('pageerror',e=>errors.push(e.message));
  const seen = new Set();
  await stubJev(page, seen);
  await setup(page);
  await expect(page.locator('#jev-test-rows')).toHaveValue('5');
  await page.locator('#jev-test-rows').fill('3');
  await page.locator('#jev-test-run-btn').click();
  await expect(page.locator('.test-run-preview-panel')).toBeVisible({timeout:20000});
  expect([...seen].sort((a,b)=>a-b)).toEqual([0,1,2]);
  await expect(page.locator('.test-run-preview-panel')).toContainText('3 Rows');
  if(errors.length) throw Error(errors.join('; '));
});

test('a count larger than the sheet is clamped to the rows that exist', async ({page}) => {
  const errors=[]; page.on('pageerror',e=>errors.push(e.message));
  const seen = new Set();
  await stubJev(page, seen);
  await setup(page);
  await page.locator('#jev-test-rows').fill('500');
  await page.locator('#jev-test-run-btn').click();
  await expect(page.locator('.test-run-preview-panel')).toBeVisible({timeout:20000});
  // 12 rows in the file, not 500 -- and no error.
  expect(seen.size).toBe(12);
  if(errors.length) throw Error(errors.join('; '));
});

test('a blank count falls back to the default instead of running nothing', async ({page}) => {
  const errors=[]; page.on('pageerror',e=>errors.push(e.message));
  const seen = new Set();
  await stubJev(page, seen);
  await setup(page);
  await page.locator('#jev-test-rows').fill('');
  await page.locator('#jev-test-run-btn').click();
  await expect(page.locator('.test-run-preview-panel')).toBeVisible({timeout:20000});
  expect(seen.size).toBe(5);
  if(errors.length) throw Error(errors.join('; '));
});
