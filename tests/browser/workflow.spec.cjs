const {test, expect} = require('@playwright/test');
test('safe import, partial export, and resume without duplicate batches', async ({page}) => {
 const errors=[]; page.on('pageerror', e=>errors.push(e.message));
  await page.goto('/');
  await expect(page.locator('#start-analyzing-btn')).toBeVisible();

  await page.locator('#start-analyzing-btn').click();
  await page.locator('#config-next-btn').click();
  const csv='Name,Value\n'+Array.from({length:25},(_,i)=>`${i===0?'<img src=x onerror=window.injected=true>':'row'+i},${i}`).join('\n');
  await page.locator('#file').setInputFiles({name:'sample.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
  await page.locator('#upload-form button[type=submit]').click();


  await expect(page.locator('#preview-table')).toContainText('<img src=x onerror=window.injected=true>');
  if(await page.evaluate(()=>window.injected))throw Error('HTML executed');

  await page.locator('#preview-next-btn').click();
  await page.locator('select[name="column"]').first().selectOption('Name');
  await page.locator('input[name="prompt"]').first().fill('Classify sentiment');
  let fail = true; const calls=[];
  await page.route('**/analyze_batch', async route => {
   const payload=route.request().postDataJSON(); const start=payload.rows[0].rowIndex; calls.push(start);
   if(start===20 && fail) return route.fulfill({status:503,contentType:'application/json',body:JSON.stringify({error:'Temporary outage'})});
   return route.fulfill({status:200,contentType:'application/json',body:JSON.stringify({errors:0,results:payload.rows.map(row=>({rowIndex:row.rowIndex,values:{[payload.configs[0].outputColumnName]:'Positive'}}))})});
  });
  await page.locator('#analyze-btn').click();
  await expect(page.locator('#analysis-recovery')).toBeVisible();
  await expect(page.locator('#analyze-message')).toContainText('20 of 25');
  const download=page.waitForEvent('download'); await page.locator('#partial-download-btn').click();
  console.log('partial download',(await download).suggestedFilename());
  fail=false; await page.locator('#resume-analysis-btn').click();
  await expect(page.locator('#step-5')).toBeVisible();
  await expect(page.locator('#result-message')).toContainText('25 of 25');
  if(JSON.stringify(calls)!=='[0,20,20]')throw Error('Invalid recovery calls '+JSON.stringify(calls));
  console.log('recovery calls',calls);
  await expect(page.locator('.test-run-preview-panel')).toHaveCount(0);
  await page.setViewportSize({width:390,height:844});
  await expect(page.locator('#step-5')).toBeVisible();
  const safe=await page.evaluate(async()=>{const {renderMarkdown}=await import('/static/js/rendering.js'); return renderMarkdown('<img src=x onerror=alert(1)><script>alert(1)</script>**safe**');});
  if(safe.includes('onerror')||safe.includes('<script'))throw Error('Unsafe markdown');
  console.log('sanitized',safe);
  console.log('errors',errors);
  if(errors.length) throw Error(errors.join('; '));

});
