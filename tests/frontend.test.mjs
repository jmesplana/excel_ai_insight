import test from 'node:test';
import assert from 'node:assert/strict';
import { BatchRun } from '../static/js/batch-runner.js';
import { parseWorkbook } from '../static/js/spreadsheet.js';

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
