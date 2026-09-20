const {test, expect} = require('@playwright/test');
const fs = require('fs');

const CSV = 'Feedback,type_of_feedback\n'
  + "Nous avons peur d'aller a l'hopital,Signalements\n"
  + 'Merci pour les informations,Appreciations\n'
  + ',Appreciations\n';

// Walk to step 4 in Jev mode with a file loaded.
async function openJevConfig(page) {
  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-jev').click();
  await page.locator('#config-next-btn').click();
  await page.locator('#file').setInputFiles(
    {name:'feedback.csv', mimeType:'text/csv', buffer:Buffer.from(CSV)});
  await page.locator('#upload-form button[type=submit]').click();
  await page.locator('#preview-next-btn').click();
}

// Answer the way the real API does: keys are the slugified option labels.
function stubJev(page, onRequest = () => {}) {
  return page.route('**/analyze_batch_jev', async route => {
    const body = route.request().postDataJSON();
    onRequest(body);
    const results = body.rows.map(row => {
      const values = {};
      if (row.state === null || String(row.state).trim() === '') {
        body.configs.forEach(c => { values[c.outputColumnName] = 'No data (empty cell)'; });
        return {rowIndex: row.rowIndex, values};
      }
      body.configs.forEach(c => {
        if (c.questionType === 'score') {
          values[c.outputColumnName] = c.options[1];
          if (body.includeConfidence) {
            values[`${c.outputColumnName}__confidence`] = 0.55;
            values[`${c.outputColumnName}__score`] = 1.2;
          }
        } else {
          values[c.outputColumnName] = c.options[0];
          if (body.includeConfidence) values[`${c.outputColumnName}__confidence`] = 0.97;
        }
      });
      return {rowIndex: row.rowIndex, values};
    });
    await route.fulfill({json: {results, errors: 0}});
  });
}

test('Jev mode classifies every row and writes the chosen labels into the sheet', async ({page}) => {
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  const requests = [];
  await stubJev(page, body => requests.push(body));
  await openJevConfig(page);

  // The Jev panel replaces the LLM one, which has no options editor.
  await expect(page.locator('#jev-config')).toBeVisible();
  await expect(page.locator('#analysis-config')).toBeHidden();

  // Both source columns become the row state.
  await page.locator('.jev-source-column[value="Feedback"]').check();
  await page.locator('.jev-source-column[value="type_of_feedback"]').check();

  await page.locator('[name="jev-output-column-name"]').first().fill('feedback_type');
  await page.locator('[name="jev-instructions"]').first().fill('Determine le type dominant.');
  await page.locator('.jev-options').first().fill('Question\nPlainte\nRefus');

  // A second result column, this time an ordered scale.
  await page.locator('#jev-add-question-btn').click();
  await page.locator('[name="jev-output-column-name"]').nth(1).fill('criticalite');
  await page.locator('[name="jev-question-type"]').nth(1).selectOption('score');
  await page.locator('[name="jev-instructions"]').nth(1).fill("Evaluer l'urgence.");
  await page.locator('.jev-options').nth(1).fill('Faible\nMoyenne\nElevee');
  await page.locator('#jev-include-confidence').check();

  await page.locator('#jev-run-btn').click();
  await expect(page.locator('#result-message')).toContainText('Classification complete', {timeout: 20000});

  // Every question travelled in one request per batch, with the labelled state.
  expect(requests).toHaveLength(1);
  expect(requests[0].configs).toHaveLength(2);
  expect(requests[0].includeConfidence).toBe(true);
  expect(requests[0].rows[0].state)
    .toBe("Feedback: Nous avons peur d'aller a l'hopital\ntype_of_feedback: Signalements");
  // The blank Feedback cell still has a type, so the row is not skipped.
  expect(requests[0].rows[2].state).toBe('type_of_feedback: Appreciations');

  // Results land in the shared analysis view, confidence columns included.
  await expect(page.locator('#step-5')).toBeVisible();
  const header = await page.locator('#result-preview table thead').first().innerText();
  expect(header).toContain('feedback_type');
  expect(header).toContain('criticalite');
  expect(header).toContain('feedback_type__confidence');
  expect(header).toContain('criticalite__score');
  const firstRow = await page.locator('#result-preview table tbody tr').first().innerText();
  expect(firstRow).toContain('Question');
  expect(firstRow).toContain('Moyenne');

  if (errors.length) throw Error(errors.join('; '));
});

test('a Choice question with too few options is refused before any request', async ({page}) => {
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  let called = false;
  await stubJev(page, () => { called = true; });
  await openJevConfig(page);

  await page.locator('[name="jev-output-column-name"]').first().fill('feedback_type');
  await page.locator('[name="jev-instructions"]').first().fill('Determine le type.');
  await page.locator('.jev-options').first().fill('Question');
  await page.locator('#jev-run-btn').click();

  await expect(page.locator('#jev-message')).toContainText('at least 2 options');
  expect(called).toBe(false);
  // The user stays on the configuration step with their work intact.
  await expect(page.locator('#jev-config')).toBeVisible();
  await expect(page.locator('.jev-options').first()).toHaveValue('Question');

  if (errors.length) throw Error(errors.join('; '));
});

test('a Jev configuration survives export and reimport onto another file', async ({page}) => {
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  await openJevConfig(page);

  await page.locator('.jev-source-column[value="Feedback"]').check();
  await page.locator('.jev-source-column[value="type_of_feedback"]').check();
  await page.locator('[name="jev-output-column-name"]').first().fill('feedback_type');
  await page.locator('[name="jev-instructions"]').first().fill('Determine le type dominant.');
  await page.locator('.jev-options').first().fill('Question\nPlainte\nRefus');
  await page.locator('#jev-include-confidence').check();

  const download = page.waitForEvent('download');
  await page.locator('#jev-export-config-btn').click();
  const file = await download;
  expect(file.suggestedFilename()).toMatch(/^aidstack-jev-config-\d{4}-\d{2}-\d{2}\.json$/);
  const saved = await file.path();
  const doc = JSON.parse(fs.readFileSync(saved, 'utf8'));
  expect(doc.questions[0].options).toEqual(['Question', 'Plainte', 'Refus']);

  // Reimport against a file missing one of the source columns.
  await page.reload();
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-jev').click();
  await page.locator('#config-next-btn').click();
  await page.locator('#file').setInputFiles(
    {name:'other.csv', mimeType:'text/csv', buffer:Buffer.from('Feedback,country\nBonjour,DRC\n')});
  await page.locator('#upload-form button[type=submit]').click();
  await page.locator('#preview-next-btn').click();

  await page.locator('#jev-import-config-input').setInputFiles(saved);
  await expect(page.locator('#jev-message')).toContainText('type_of_feedback');
  // The option list -- the laborious part -- is restored in full.
  await expect(page.locator('.jev-options').first()).toHaveValue('Question\nPlainte\nRefus');
  await expect(page.locator('#jev-include-confidence')).toBeChecked();
  await expect(page.locator('.jev-source-column[value="Feedback"]')).toBeChecked();

  if (errors.length) throw Error(errors.join('; '));
});

test('configurable report, failed-decision retry and checkpoint restore', async ({page}) => {
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  let requests = 0;
  await page.route('**/analyze_batch_jev', async route => {
    const body = route.request().postDataJSON(); requests++;
    const results = body.rows.map(row => {
      const failure = row.rowIndex === 1 && !row.previous;
      const value = row.rowIndex === 0 ? 'A' : 'B';
      return {rowIndex: row.rowIndex, values: {Category: failure ? 'Error: retry' : value},
        decisions: {Category: failure ? {status: 'error', value: null, reason: 'Temporary error'} :
          {status: 'ok', value, confidence: .99, labels: {a:'A',b:'B'}, raw: {type:'choice',choice:value.toLowerCase(),confidence:.99,probabilities:{a:value==='A'?1:0,b:value==='B'?1:0}}, model:'jev-pinned'}},
        calls:[{model:'jev-pinned', usage:{input_tokens:10}}]};
    });
    await route.fulfill({json:{results,errors:requests===1?1:0}});
  });
  await openJevConfig(page);
  const config = {format:'aidstack-insights-jev-config', version:2, sourceColumns:['Feedback'], questions:[
    {outputColumnName:'Category', questionType:'choice', instructions:{question:'Which category?'},
      options:[{label:'A', description:{examples:['first']}},'B'], review:{minConfidence:.8}}],
    report:{title:'Custom dataset report',groupBy:['type_of_feedback'],crossTabs:[['type_of_feedback','Category']],maxExamples:2},
    execution:{batchSize:4}};
  await page.locator('#jev-import-config-input').setInputFiles({name:'config.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
  const exported = page.waitForEvent('download');
  await page.locator('#jev-export-config-btn').click();
  const doc = JSON.parse(fs.readFileSync(await (await exported).path(),'utf8'));
  expect(doc.report).toEqual(config.report);
  expect(doc.questions).toEqual(config.questions);
  await page.locator('#jev-run-btn').click();
  await expect(page.locator('#jev-dataset-report')).toContainText('Custom dataset report');
  await expect(page.locator('#jev-coverage')).toContainText('1 failed');
  await page.locator('#jev-retry-results').click();
  await expect(page.locator('#jev-coverage')).toContainText('0 failed');
  expect(requests).toBe(2);
  // The narrative receives full aggregates and explicit coverage, not rows.
  await page.route('**/jev_report', async route => {
    const body = route.request().postDataJSON();
    expect(body.report.coverage.processed).toBe(3);
    expect(body.report.distributions[0].counts.find(c=>c.value==='B').count).toBe(2);
    expect(body.rows).toBeUndefined();
    await route.fulfill({contentType:'text/event-stream',body:'data: {"content":"Three records analyzed."}\n\ndata: {"done":true}\n\n'});
  });
  await page.locator('#jev-narrative-generate').click();
  await expect(page.locator('#jev-narrative')).toContainText('Three records analyzed.');
  await page.reload();
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-jev').click();
  await page.locator('#jev-restore-start-btn').click();
  await expect(page.locator('#jev-coverage')).toContainText('3 of 3');
  await expect(page.locator('#jev-coverage')).toContainText('0 failed');
  expect(requests).toBe(2);
  const downloaded = page.waitForEvent('download');
  await page.locator('#download-link').click();
  const file = await (await downloaded).path();
  const buffer = fs.readFileSync(file);
  const sheets = await page.evaluate(bytes => XLSX.read(new Uint8Array(bytes),{type:'array'}).SheetNames, Array.from(buffer));
  expect(sheets).toContain('Summary'); expect(sheets).toContain('Decision audit'); expect(sheets).toContain('Run metadata');
  expect(errors).toEqual([]);
});
