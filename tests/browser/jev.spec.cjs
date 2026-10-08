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

  // Large-result CSV downloads must never allocate an intermediate workbook.
  await page.evaluate(() => { XLSX.utils.json_to_sheet = () => { throw new Error('CSV must not build a workbook'); }; });
  for (const button of ['#download-results-csv', '[data-recovery-csv]:visible']) {
    const download = page.waitForEvent('download');
    await page.locator(button).click();
    const csv = fs.readFileSync(await (await download).path(), 'utf8');
    expect(csv).toContain('feedback_type');
    expect(csv).toContain('Merci pour les informations');
    expect(csv).toContain('Question');
  }

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

test('outputOrder imports and places calculated columns between question columns', async ({page}) => {
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  await openJevConfig(page);
  const config = {format: 'aidstack-insights-jev-config', version: 2, sourceColumns: ['Feedback'],
    questions: [
      {outputColumnName: 'feedback_type', questionType: 'choice', instructions: 'Choose type.', options: ['Question', 'Request']},
      {outputColumnName: 'feedback_code', questionType: 'choice', instructions: 'Choose code.', options: ['Vaccine', 'Treatment']}
    ],
    derived: [
      {outputColumnName: 'feedback_category', type: 'lookup', inputs: ['feedback_code'],
        table: [{when: {feedback_code: 'Vaccine'}, value: 'Health'}, {when: {feedback_code: 'Treatment'}, value: 'Health'}]},
      {outputColumnName: 'feedback_subcode', type: 'lookup', inputs: ['feedback_code'],
        table: [{when: {feedback_code: 'Vaccine'}, value: 'Prevention'}, {when: {feedback_code: 'Treatment'}, value: 'Care'}]}
    ], outputOrder: ['feedback_type', 'feedback_category', 'feedback_code', 'feedback_subcode']};
  await page.locator('#jev-import-config-input').setInputFiles({name: 'ordered.json', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(config))});
  await expect(page.locator('#jev-message')).toContainText('Imported');
  expect(JSON.parse(await page.locator('#jev-workflow-json').inputValue()).outputOrder).toEqual(config.outputOrder);
  const download = page.waitForEvent('download');
  await page.locator('#jev-export-config-btn').click();
  const saved = JSON.parse(fs.readFileSync(await (await download).path(), 'utf8'));
  expect(saved.outputOrder).toEqual(config.outputOrder);
  await page.route('**/analyze_batch_jev', async route => {
    const body = route.request().postDataJSON();
    const results = body.rows.map(row => ({rowIndex: row.rowIndex,
      values: {feedback_type: 'Question', feedback_category: 'Health', feedback_code: 'Vaccine', feedback_subcode: 'Prevention'},
      decisions: Object.fromEntries(config.outputOrder.map(name => [name, {status: 'ok', value: 'Example'}])), calls: []}));
    await route.fulfill({json: {results, errors: 0}});
  });
  await page.locator('#jev-run-btn').click();
  await expect(page.locator('#result-message')).toContainText('Classification complete');
  const headers = await page.locator('#results-data-table thead th').evaluateAll(cells => cells.map(cell => cell.dataset.column));
  expect(headers.filter(name => config.outputOrder.includes(name))).toEqual(config.outputOrder);
  expect(errors).toEqual([]);
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

test('source-routed config reuses identical text and type across batches and recovery', async ({page}) => {
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  await page.goto('/');
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-jev').click();
  await page.locator('#config-next-btn').click();
  const csv = 'ID,Feedback,type_of_feedback\n1,Question and request,Questions\n1,Question and request,Demandes\n2,Question and request,Questions\n3,Question and request,Questions\n4,Question and request,Demandes\n';
  await page.locator('#file').setInputFiles({name:'repeated.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
  await page.locator('#upload-form button[type=submit]').click();
  await page.locator('#preview-next-btn').click();
  const config = {format:'aidstack-insights-jev-config',version:2,sourceColumns:['Feedback','type_of_feedback'],
    questions:[{outputColumnName:'Code',questionType:'choice',instructions:'Classify the assigned type only.',options:[],
      sourceBranches:[{when:{type_of_feedback:'Questions'},options:['Q1','Q2']},{when:{type_of_feedback:'Demandes'},options:['D1','D2']}]}],
    execution:{batchSize:2,workers:8,structuredState:true,reuseIdenticalRows:true}};
  await page.locator('#jev-import-config-input').setInputFiles({name:'typed.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(config))});
  await expect(page.locator('#jev-message')).toContainText('Imported');
  let requests = 0;
  await page.route('**/analyze_batch_jev', async route => {
    requests++;
    const body = route.request().postDataJSON();
    expect(body.configs[0].sourceBranches).toEqual(config.questions[0].sourceBranches);
    expect(body.rows.map(r => r.state.type_of_feedback)).toEqual(['Questions','Demandes']);
    await route.fulfill({json:{results:body.rows.map(row => {
      const value = row.state.type_of_feedback === 'Questions' ? 'Q1' : 'D1';
      return {rowIndex:row.rowIndex,values:{Code:value},decisions:{Code:{status:'ok',value}},
        calls:[{model:'jev-pinned',usage:{input_tokens:10},elapsedSeconds:.1}]};
    }),errors:0}});
  });
  await page.locator('#jev-run-btn').click();
  await expect(page.locator('#result-message')).toContainText('Classification complete: 5 of 5');
  expect(requests).toBe(1);
  const download = page.waitForEvent('download');
  await page.locator('#download-link').click();
  const bytes = fs.readFileSync(await (await download).path());
  const exported = await page.evaluate(bytes => {
    const wb=XLSX.read(new Uint8Array(bytes),{type:'array'});
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    const audit=XLSX.utils.sheet_to_json(wb.Sheets['Decision audit']);
    const meta=JSON.parse(XLSX.utils.sheet_to_json(wb.Sheets['Run metadata']).map(r=>r.text).join(''));
    return {rows,audit,meta};
  },Array.from(bytes));
  expect(exported.rows.map(r=>r.ID)).toEqual([1,1,2,3,4]);
  expect(exported.rows.map(r=>r.Code)).toEqual(['Q1','D1','Q1','Q1','D1']);
  expect(exported.meta.inputTokens).toBe(20);
  expect(exported.meta.execution.reusedRows).toBe(3);
  expect(exported.audit[2].reusedFromRow).toBe(1);
  await page.reload();
  await page.locator('#start-analyzing-btn').click();
  await page.locator('#mode-card-jev').click();
  await page.locator('#jev-restore-start-btn').click();
  await expect(page.locator('#jev-coverage')).toContainText('5 of 5');
  expect(requests).toBe(1);
  expect(errors).toEqual([]);
});

test('refined Ebola configuration imports, exports and validates dependent assessments', async ({page}) => {
  const path = require('path').join(__dirname, '../../sample_jev/ebola-feedback-coding.refined.config.json');
  test.skip(!fs.existsSync(path), 'Local reviewed Ebola configuration not present');
  const config=JSON.parse(fs.readFileSync(path,'utf8'));
  await openJevConfig(page);
  await page.locator('#jev-import-config-input').setInputFiles(path);
  await expect(page.locator('#jev-message')).toContainText('Imported 7 result column(s)');
  const download=page.waitForEvent('download');
  await page.locator('#jev-export-config-btn').click();
  const roundTrip=JSON.parse(fs.readFileSync(await (await download).path(),'utf8'));
  expect(roundTrip.questions).toEqual(config.questions);
  expect(roundTrip.derived).toEqual(config.derived);
  expect(roundTrip.questions[1].allowReviewDependencies).toBe(true);
  const validation=await page.request.post('/validate_jev_config', {data:{configs:roundTrip.questions,derived:roundTrip.derived}});
  expect(validation.ok()).toBe(true);
  expect(await validation.json()).toEqual({ok:true});
});

test('excerpt configuration round trips its source selector and dependency graph', async ({page}) => {
  const path = require('path').join(__dirname, '../../sample_jev/ebola-feedback-coding.excerpt.config.json');
  test.skip(!fs.existsSync(path), 'Local excerpt configuration not present');
  const config=JSON.parse(fs.readFileSync(path,'utf8'));
  await openJevConfig(page);
  await page.locator('#jev-import-config-input').setInputFiles(path);
  await expect(page.locator('#jev-message')).toContainText('Imported 8 result column(s)');
  const download=page.waitForEvent('download');
  await page.locator('#jev-export-config-btn').click();
  const saved=JSON.parse(fs.readFileSync(await (await download).path(),'utf8'));
  expect(saved.questions).toEqual(config.questions);
  expect(saved.derived).toEqual(config.derived);
  expect(saved.questions[0].sourceExcerpt.column).toBe('Feedback');
  expect(saved.questions[1].dependsOn).toEqual(['feedback_excerpt']);
  expect(saved.questions[2].questionType).toBe('choice');
  expect(saved.outputOrder).toEqual(config.outputOrder);
  expect(saved.outputOrder.indexOf('feedback_sector_suggested')).toBe(saved.outputOrder.indexOf('feedback_sector') + 1);
  expect(saved.questions.some(q => q.outputColumnName === 'feedback_sector_suggested')).toBe(true);
  const routing = saved.derived.find(d => d.outputColumnName === 'feedback_sector');
  expect(routing.inputs).toEqual(['feedback_action', 'feedback_sector_suggested']);
  expect(routing.table.find(row => row.when.feedback_action === 'Trigger public health alert'
    && row.when.feedback_sector_suggested === 'Clinical care').value).toBe('Surveillance');
  const response=await page.request.post('/validate_jev_config',{data:{configs:saved.questions,derived:saved.derived}});
  expect(await response.json()).toEqual({ok:true});
});
