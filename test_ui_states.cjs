const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const {execFileSync}=require('node:child_process');
const assert=require('node:assert/strict');
const py=code=>execFileSync('py',['-3.14','-c',code],{encoding:'utf8',env:{...process.env,PYTHONIOENCODING:'utf-8'},maxBuffer:10000000});
(async()=>{
  const html=py("from strategy_excel_tool.server import INDEX_HTML; print(INDEX_HTML.replace('__BUILD_VERSION__','v1.3'))");
  const result=JSON.parse(py("from strategy_excel_tool import server as s; import json; d=s.build_demo_xlsx(['商品编码','商品名称','本期投放策略'],[['01614','测试','1-30档2条']]); r=s.analyze_and_apply(d,'district_type'); r['sessionId']='test'; print(json.dumps(r))"));
  const browser=await chromium.launch({channel:'msedge',headless:true});
  try {
    const page=await browser.newPage();let release,cancelPicker,requests=0;
    await page.route('http://tool.test/**',async route=>{
      const path=new URL(route.request().url()).pathname;
      if(path==='/')return route.fulfill({contentType:'text/html',body:html});
      if(path==='/api/templates')return route.fulfill({json:{templates:[{key:'district_type',exists:true,label:'格式一',marker:'2'},{key:'county',exists:true,label:'格式二',marker:'1'}],innerCodeMapping:{exists:true},templateDirectory:'private'}});
      if(path==='/api/analyze'){await new Promise(r=>release=r);return route.fulfill({json:result});}
      if(path==='/api/shutdown')return route.fulfill({json:{ok:true}});
      if(path==='/api/page-events')return route.fulfill({status:204});
      if(path==='/api/customer/state')return route.fulfill({json:{ok:true,revision:0,reviewToken:result.zeroToken,filename:'',customerSummary:null,choicesConfirmed:true,choices:Object.fromEntries(result.exportProducts.map(p=>[String(p.sourceRow),'keep'])),exportProducts:result.exportProducts,busy:false}});
      if(path==='/api/save-location'){await new Promise(r=>cancelPicker=r);return route.fulfill({json:{ok:true,cancelled:true}});}
      requests++;return route.fulfill({status:500,json:{errors:['unexpected request']}});
    });
    await page.goto('http://tool.test/');
    await page.evaluate(()=>showMessages({warnings:Array.from({length:33},(_,i)=>'第'+(i+1)+'条校验提示，请复核当前策略。')}));
    assert.equal(await page.locator('#validationDetails').getAttribute('open')!==null,true);
    assert.equal(await page.locator('#messages .msg').count(),33);
    assert.ok((await page.locator('#messages').boundingBox()).height>1000);
    assert.equal(await page.locator('#messages').evaluate(e=>getComputedStyle(e).maxHeight),'none');
    assert.equal(await page.locator('#messages').evaluate(e=>e.scrollHeight<=e.clientHeight+1),true);
    await page.locator('#validationTitle').click();
    await page.evaluate(()=>showMessages({warnings:['新的校验提示']}));
    assert.equal(await page.locator('#validationDetails').getAttribute('open')!==null,true);
    await page.locator('#strategyFile').setInputFiles({name:'a.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from('test')});
    await page.locator('#analyzeBtn').click();await page.waitForTimeout(100);
    assert.equal(await page.locator('#strategyFile').isDisabled(),true);assert.equal(await page.locator('#templateKey').isDisabled(),true);
    assert.equal(await page.locator('#exitTool').isDisabled(),true);
    release();await page.locator('#reviewPanel').waitFor({state:'visible'});
    await page.locator('#reviewSearch').fill('01614');await page.locator('#reviewResults button').click();
    await page.locator('#exportBtn').click();
    assert.equal(await page.locator('#reviewTable input').first().isDisabled(),true);
    assert.equal(await page.locator('#analyzeBtn').isDisabled(),true);
    assert.equal(await page.locator('#exitTool').isDisabled(),true);
    // Simulate a draft arriving while the async picker is open; finally must honor it.
    await page.locator('#reviewTable input').first().evaluate(e=>{e.value='100';e.dispatchEvent(new Event('input',{bubbles:true}));});
    cancelPicker();
    await page.waitForFunction(()=>!document.getElementById('analyzeBtn').disabled);
    assert.equal(await page.locator('#exportBtn').isDisabled(),true);
    await page.locator('#exportBtn').evaluate(e=>{e.disabled=false;e.click();});assert.equal(requests,0);
    await page.locator('#discardReview').click();
    await page.locator('#editStrategy').click();await page.locator('#reviewText').fill('30档5条');
    await page.locator('#maximizeTable').click();await page.locator('#reviewTable input').first().fill('9');
    await page.locator('#applyNumbers').click();
    assert.match(await page.locator('#quantityDialog #quantityStatus').textContent(),/文字也有修改/);
    await page.locator('#maximizeTable').click();await page.locator('#quantityDialog').waitFor({state:'hidden'});
    await page.locator('#discardReview').click();
    await page.locator('#analyzeBtn').click();await page.waitForTimeout(100);
    await page.locator('#templateKey').evaluate(e=>{e.value='county';e.dispatchEvent(new Event('change'));});
    release();await page.waitForFunction(()=>!document.getElementById('analyzeBtn').disabled);
    assert.equal(await page.locator('#reviewPanel').isVisible(),false);assert.equal(await page.locator('#exportBtn').isDisabled(),true);
    page.once('dialog',d=>d.dismiss());await page.locator('#exitTool').click();
    assert.equal(await page.locator('#exitNotice').isVisible(),false);
    page.once('dialog',d=>d.accept());await page.locator('#exitTool').click();
    await page.locator('#exitNotice').waitFor({state:'visible'});
    assert.equal(await page.locator('main').isVisible(),false);
    assert.equal(await page.locator('#exitTool').isDisabled(),true);
    console.log('Async locks, stale responses, draft guard, maximized feedback, exit cancellation and confirmed exit UI passed.');
  } finally {await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
