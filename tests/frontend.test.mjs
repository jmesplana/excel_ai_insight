import test from 'node:test';
import assert from 'node:assert/strict';
import { BatchRun } from '../static/js/batch-runner.js';
import { parseWorkbook, explodeColumn, splitCell } from '../static/js/spreadsheet.js';
import { serializeConfig, deserializeConfig, configFilename, DEFAULT_OUTPUT_TOKENS } from '../static/js/analysis-config.js';

test('failed batch preserves earlier work and resume skips paid completed batches', async () => {
    const calls = []; let fail = true;
    const run = new BatchRun({ rows:[1,2,3,4], batchSize:2, processBatch:async (rows,start) => {
        calls.push(start); if (start === 2 && fail) throw new Error('offline'); return { rows };
    }});
    await assert.rejects(run.run());
    assert.deepEqual(run.results,[1,2]); assert.equal(run.cursor,2);
    fail = false; await run.run();
    assert.deepEqual(calls,[0,2,2]); assert.deepEqual(run.results,[1,2,3,4]);
});
test('stop completes the active batch and resumes without duplication', async () => {
    const run = new BatchRun({rows:[1,2,3], batchSize:2, processBatch:async rows => ({rows}), onProgress:r=>r.stop()});
    await run.run(); assert.equal(run.cursor,2); await run.run(); assert.equal(run.complete,true);
    assert.deepEqual(run.results,[1,2,3]);
});
test('incomplete response never advances the checkpoint', async () => {
    const run = new BatchRun({rows:[1,2],batchSize:2,processBatch:async()=>({rows:[1]})});
    await assert.rejects(run.run()); assert.equal(run.cursor,0);
});
const workbook = {SheetNames:['Data'],Sheets:{Data:{}}};
const xlsx = matrix => ({utils:{sheet_to_json:()=>matrix}});
test('rejects duplicate and blank headers before analysis', () => {
    assert.throws(()=>parseWorkbook(workbook,xlsx([['Name','Name'],['a','b']])),/duplicate/);
    assert.throws(()=>parseWorkbook(workbook,xlsx([['Name',''],['a','b']])),/blank/);
});
test('preserves typed values and literal HTML text', () => {
    const parsed=parseWorkbook(workbook,xlsx([['Name','Value'],['<img onerror=alert(1)>',42]]));
    assert.equal(parsed.Data.data[0].Value,42);
    assert.equal(parsed.Data.data[0].Name,'<img onerror=alert(1)>');
});

const feedback = { columns:['id','feedback','region','date','feedback_classification'], data:[
    {id:1,feedback:'Ebola existe',region:'Nord-Kivu',date:'2026-09-10',
     feedback_classification:"Observations || Signalements || Allégations"},
    {id:2,feedback:'Merci',region:'Ituri',date:'2026-09-11',feedback_classification:'Appréciations'},
    {id:3,feedback:'Avez-vous le médicament ?',region:'Nord-Kivu',date:'2026-09-12',
     feedback_classification:'Questions || Demandes'},
]};
test('explodes a multi-value column into one row per value and copies the rest down', () => {
    const out = explodeColumn(feedback,'feedback_classification',['||'],{outputColumn:'category'});
    assert.deepEqual(out.columns,['id','feedback','region','date','category']);
    assert.equal(out.data.length,6);
    assert.deepEqual(out.data.map(r=>r.category),
        ['Observations','Signalements','Allégations','Appréciations','Questions','Demandes']);
    // Non-split columns are copied verbatim onto every emitted row.
    assert.deepEqual(out.data.slice(0,3).map(r=>r.id),[1,1,1]);
    assert.equal(out.data[0].region,'Nord-Kivu');
    // The renamed source column leaves no stray key behind.
    assert.deepEqual(Object.keys(out.data[0]),['id','feedback','region','date','category']);
});
test('longest separator wins so "||" never splits as two empty "|" fields', () => {
    assert.deepEqual(splitCell('a || b',['|','||']),['a','b']);
    assert.deepEqual(splitCell('a | b',['||','|']),['a','b']);
});
test('separators are literal, not regex', () => {
    assert.deepEqual(splitCell('a.b',['.']),['a','b']);
    assert.deepEqual(splitCell('a+b',['+']),['a','b']);
});
test('blank fragments and empty cells never become rows', () => {
    assert.deepEqual(splitCell('a || || b',['||']),['a','b']);
    const sparse = explodeColumn({columns:['id','cat'],data:[{id:1,cat:'  '},{id:2,cat:null}]},
        'cat',['||']);
    assert.deepEqual(sparse.data,[{id:1,cat:''},{id:2,cat:''}]);
});
test('dedupe drops values repeated inside one cell but keeps them across rows', () => {
    const dupes = {columns:['id','cat'],data:[{id:1,cat:'A || A || B'},{id:2,cat:'A'}]};
    assert.equal(explodeColumn(dupes,'cat',['||']).data.length,3);
    assert.equal(explodeColumn(dupes,'cat',['||'],{dedupe:false}).data.length,4);
});
test('rejects an unknown column and a rename that would collide', () => {
    assert.throws(()=>explodeColumn(feedback,'nope',['||']),/not in this sheet/);
    assert.throws(()=>explodeColumn(feedback,'feedback_classification',['||'],{outputColumn:'region'}),
        /already exists/);
});
test('explode works on a freshly parsed sheet, with no analysis columns', () => {
    // The step 3 path: parse a workbook, split a column, export — no AI involved.
    const parsed = parseWorkbook(workbook, xlsx([
        ['id','region','feedback_classification'],
        [1,'Nord-Kivu','Observations || Signalements'],
        [2,'Ituri','Appréciations'],
    ]));
    const out = explodeColumn(parsed.Data, 'feedback_classification', ['||'], { outputColumn:'category' });
    assert.deepEqual(out.columns, ['id','region','category']);
    assert.equal(out.data.length, 3);
    assert.deepEqual(out.data.map(r => r.category),
        ['Observations','Signalements','Appréciations']);
    assert.equal(out.data[1].region, 'Nord-Kivu');
});
test('exploding twice splits on a second separator without losing rows', () => {
    // Mixed separators are also reachable in one pass via ['||',';'].
    const once = explodeColumn({columns:['id','cat'],data:[{id:1,cat:'A || B;C'}]}, 'cat', ['||']);
    assert.deepEqual(once.data.map(r=>r.cat), ['A','B;C']);
    const twice = explodeColumn(once, 'cat', [';']);
    assert.deepEqual(twice.data.map(r=>r.cat), ['A','B','C']);
    assert.deepEqual(twice.data.map(r=>r.id), [1,1,1]);
});

// --- Configure Analysis export / import ---------------------------------
const cfgColumns = ['Feedback', 'type_of_feedback', 'region'];
const savedConfig = serializeConfig({
    generalInstructions: 'Answer in French.',
    sheetName: 'Data',
    columns: [
        { columns: ['Feedback', 'type_of_feedback'], outputColumnName: 'feedback_code',
          prompt: 'Code the combination; do not explain.', maxOutputTokens: 1024 },
        { columns: ['Feedback', 'type_of_feedback'], outputColumnName: 'feedback_dimension',
          prompt: 'Pick one dimension.', maxOutputTokens: 4096 },
    ]
});

test('a round trip restores every field the user typed', () => {
    const back = deserializeConfig(JSON.parse(JSON.stringify(savedConfig)), cfgColumns);
    assert.equal(back.generalInstructions, 'Answer in French.');
    assert.deepEqual(back.columns.map(c => c.outputColumnName), ['feedback_code', 'feedback_dimension']);
    assert.deepEqual(back.columns[0].columns, ['Feedback', 'type_of_feedback']);
    assert.equal(back.columns[0].prompt, 'Code the combination; do not explain.');
    assert.equal(back.columns[1].maxOutputTokens, 4096);
    assert.deepEqual(back.warnings, []);
});

test('columns missing from the new sheet are dropped with a warning, prompts kept', () => {
    const back = deserializeConfig(savedConfig, ['Feedback', 'region']);
    assert.deepEqual(back.columns[0].columns, ['Feedback']);
    assert.equal(back.columns[0].prompt, 'Code the combination; do not explain.');
    assert.equal(back.warnings.length, 1);
    assert.match(back.warnings[0], /type_of_feedback/);
});

test('rejects files that are not configurations, and configs from a newer version', () => {
    assert.throws(() => deserializeConfig(null, cfgColumns), /not an Aidstack/);
    assert.throws(() => deserializeConfig({ columns: [] }, cfgColumns), /not an Aidstack/);
    assert.throws(() => deserializeConfig({ ...savedConfig, columns: [] }, cfgColumns), /no column analyses/);
    assert.throws(() => deserializeConfig({ ...savedConfig, version: 99 }, cfgColumns), /newer version/);
});

test('an out-of-range output length falls back to the default the select offers', () => {
    const odd = { ...savedConfig, columns: [{ ...savedConfig.columns[0], maxOutputTokens: 999999 }] };
    assert.equal(deserializeConfig(odd, cfgColumns).columns[0].maxOutputTokens, DEFAULT_OUTPUT_TOKENS);
    assert.equal(serializeConfig({ columns: [{ columns: ['Feedback'], maxOutputTokens: undefined }] })
        .columns[0].maxOutputTokens, DEFAULT_OUTPUT_TOKENS);
});

test('the single-column shorthand imports as a one-column list', () => {
    const legacy = { ...savedConfig, columns: [{ column: 'Feedback', prompt: 'Summarize.' }] };
    assert.deepEqual(deserializeConfig(legacy, cfgColumns).columns[0].columns, ['Feedback']);
});

test('the filename is dated so repeated exports do not overwrite each other', () => {
    assert.equal(configFilename(new Date('2026-09-19T10:00:00Z')), 'aidstack-config-2026-09-19.json');
});
