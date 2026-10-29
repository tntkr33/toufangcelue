const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const {execFileSync}=require('node:child_process');
const assert=require('node:assert/strict');
const py=code=>execFileSync('py',['-3.14','-c',code],{encoding:'utf8',env:{...process.env,PYTHONIOENCODING:'utf-8'},maxBuffer:10000000});
(async()=>{
  const html=py('from strategy_excel_tool.server import INDEX_HTML; print(INDEX_HTML)');
  const result=JSON.parse(py("from strategy_excel_tool import server as s; import json; r=s.analyze_and_apply(s.build_demo_xlsx(['商品编码','商品名称','本期投放策略'],[['01614','Synthetic','1-30档2条']]),'district_type'); r['sessionId']='test'; print(json.dumps(r))"));
  let confirmed=false, choices=null, revision=0, token=result.zeroToken, fail=false;
  const calls=[], errors=[];
  const browser=await chromium.launch({channel:'msedge',headless:true});
  try {
    const page=await browser.newPage();page.on('pageerror',e=>errors.push(String(e)));
    await page.route('http://review.test/**',async route=>{
      const path=new URL(route.request().url()).pathname;
      if(path==='/')return route.fulfill({contentType:'text/html',body:html});
      if(path==='/api/templates')return route.fulfill({json:{templates:[{key:'district_type',exists:true,label:'Test',marker:'2'}],innerCodeMapping:{exists:true}}});
      if(path==='/api/page-events')return route.fulfill({status:204});
      if(path==='/api/analyze')return route.fulfill({json:result});
      const state=()=>({ok:true,revision,reviewToken:token,choicesConfirmed:confirmed,choices,exportProducts:result.exportProducts,customerSummary:null});
      if(path==='/api/customer/state')return route.fulfill({json:state()});
      if(path==='/api/customer/options')return route.fulfill({json:{...state(),options:{products:result.exportProducts,counties:[],markets:[]}}});
      if(path==='/api/customer/choices'){
        calls.push('confirm');
        if(fail)return route.fulfill({status:409,json:{ok:false,errors:['test conflict']}});
        choices=route.request().postDataJSON().choices;confirmed=true;revision++;
        return route.fulfill({json:state()});
      }
      if(path==='/api/save-location'){
        calls.push('save');assert.ok(confirmed,'Save dialog opened before confirmation');
        return route.fulfill({json:{ok:true,cancelled:true}});
      }
      throw new Error('Unexpected endpoint '+path);
    });
    await page.goto('http://review.test/');
    assert.deepEqual(await page.locator('#analyzeBtn, #reviewExportBtn, #exportBtn').evaluateAll(es=>es.map(e=>e.id)),['analyzeBtn','reviewExportBtn','exportBtn']);
    assert.ok(await page.locator('#reviewExportBtn').isDisabled());
    await page.locator('#strategyFile').setInputFiles({name:'test.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from('synthetic')});
    await page.locator('#analyzeBtn').click();await page.locator('#reviewPanel').waitFor({state:'visible'});
    await page.locator('#exportBtn').click();await page.locator('#zeroConfirm').waitFor();
    assert.deepEqual(calls,[]);
    await page.locator('#zeroCancel').click();await page.waitForFunction(()=>!reviewBusy);
    assert.deepEqual(calls,[]);
    await page.locator('#reviewExportBtn').click();await page.locator('#zeroConfirm').click();await page.waitForFunction(()=>!reviewBusy);
    assert.deepEqual(calls,['confirm']);
    await page.locator('#reviewExportBtn').click();
    await page.locator('#zeroExportDialog tbody select').first().selectOption('omit');
    assert.ok(await page.locator('#zeroConfirm').isDisabled());
    await page.locator('#zeroCancel').click();await page.waitForFunction(()=>!reviewBusy);
    assert.deepEqual(calls,['confirm']);
    assert.ok(confirmed);
    await page.locator('#exportBtn').click();await page.waitForFunction(()=>!reviewBusy);
    assert.deepEqual(calls,['confirm','save']);assert.equal(await page.locator('#zeroExportDialog').count(),0);
    confirmed=false;token='new-review';await page.evaluate(r=>reviewReceive(r),result);
    await page.locator('#exportBtn').click();await page.locator('#zeroConfirm').waitFor();
    assert.deepEqual(calls,['confirm','save']);
    await page.locator('#zeroConfirm').click();await page.waitForFunction(()=>!reviewBusy);
    assert.deepEqual(calls,['confirm','save','confirm','save']);
    confirmed=false;fail=true;
    await page.locator('#exportBtn').click();await page.locator('#zeroConfirm').click();await page.waitForFunction(()=>!reviewBusy);
    assert.deepEqual(calls,['confirm','save','confirm','save','confirm']);
    assert.match(await page.locator('#status').textContent(),/失败/);
    await page.locator('#reviewSearch').fill('01614');await page.locator('#reviewResults button').click();
    await page.locator('#reviewTable input').first().fill('7');
    assert.ok(await page.locator('#reviewExportBtn').isDisabled());
    assert.ok(await page.locator('#exportBtn').isDisabled());
    assert.deepEqual(errors,[]);
    console.log('Review-before-save passed: button order, cancel, standalone confirmation, reuse, invalidation, conflict, and draft guards.');
  }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
