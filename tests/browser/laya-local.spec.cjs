const {test, expect} = require('@playwright/test');
const {spawn} = require('child_process');
const http = require('http');
const origin = 'https://aidstack-local-test.example';

// A real local connector and deterministic model server, reached by an HTTPS page.
test('hosted page connects directly to local Laya with CORS, retries and export', async ({page, context, request}) => {
  let modelCalls=0, fail=false;
  const model = http.createServer(async (req,res) => {
    let text=''; for await (const chunk of req) text+=chunk;
    const body=JSON.parse(text); modelCalls++;
    res.setHeader('Content-Type','application/json');
    if (fail) {fail=false;res.writeHead(503);res.end('{}');return;}
    const answers={};
    for (const [id,q] of Object.entries(body.questions)) {
      if(q.type==='noul') answers[id]={type:'noul',noul:.99};
      else {const keys=Object.keys(q.criteria);answers[id]={type:'choice',choice:keys[0],confidence:.99,probabilities:Object.fromEntries(keys.map((k,i)=>[k,i===0?1:0]))};}
    }
    res.end(JSON.stringify({model:'local-test',answers}));
  });
  await new Promise(resolve=>model.listen(0,'127.0.0.1',resolve));
  // Pick a free connector port without hardcoding a user's service port.
  const reservation=http.createServer();await new Promise(resolve=>reservation.listen(0,'127.0.0.1',resolve));
  const port=reservation.address().port;await new Promise(resolve=>reservation.close(resolve));
  const connector=`http://127.0.0.1:${port}`;
  const bridge=spawn('venv/bin/python',['local_laya_bridge.py'],{env:{...process.env,LAYA_ALLOWED_ORIGINS:origin,LAYA_BRIDGE_PORT:String(port),LAYA_BASE_URL:`http://127.0.0.1:${model.address().port}`},stdio:'ignore'});
  try {
    await expect.poll(async()=>{try{return (await request.get(connector+'/health')).status();}catch{return 0;}}).toBe(200);
    await context.grantPermissions(['local-network-access'],{origin});
    const remotePosts=[];
    await page.route(origin+'/**',async route=>{
      if(route.request().method()!=='GET'){remotePosts.push(route.request().url());return route.abort();}
      const url=new URL(route.request().url());
      const response=await request.get('http://127.0.0.1:8091'+url.pathname+url.search);
      await route.fulfill({response});
    });
    await page.goto(origin);
    await page.locator('#api-settings-btn').click();
    await page.locator('#settings-tab-classify').click();await page.locator('label[for=classifier-laya]').click();
    await expect(page.locator('#laya-start-command')).toContainText('--origin '+origin);
    await page.locator('#modal-laya-endpoint').fill(connector);
    await page.locator('#test-jev-connection-btn').click();
    await expect(page.locator('#test-jev-connection-result')).toContainText('Connected to Laya');
    await page.locator('#save-api-settings').click();
    await page.locator('#start-analyzing-btn').click();await page.locator('#mode-card-jev').click();
    await page.locator('#config-next-btn').click();
    await page.locator('#file').setInputFiles({name:'feedback.csv',mimeType:'text/csv',buffer:Buffer.from('Feedback\nBonjour\n')});
    await page.locator('#upload-form button[type=submit]').click();await page.locator('#preview-next-btn').click();
    await page.locator('[name=jev-output-column-name]').first().fill('Category');
    await page.locator('[name=jev-instructions]').first().fill('Choose');await page.locator('.jev-options').first().fill('A\nB');
    fail=true;
    await page.locator('#jev-run-btn').click();await expect(page.locator('#jev-coverage')).toContainText('1 failed');
    await page.locator('#jev-retry-results').click();await expect(page.locator('#jev-coverage')).toContainText('0 failed');
    const download=page.waitForEvent('download');await page.locator('#download-link').click();expect((await download).suggestedFilename()).toMatch(/\.xlsx$/);
    expect(modelCalls).toBe(3);expect(remotePosts).toEqual([]);
  } finally {
    bridge.kill('SIGTERM');await new Promise(resolve=>bridge.once('exit',resolve));
    await new Promise(resolve=>model.close(resolve));
  }
});
