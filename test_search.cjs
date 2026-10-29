const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const { execFileSync } = require('node:child_process');
const assert = require('node:assert/strict');
(async () => {
  const js = execFileSync('py', ['-3.14', '-c', 'from strategy_excel_tool.review_ui import REVIEW_JS; print(REVIEW_JS)'], {encoding:'utf8',env:{...process.env,PYTHONIOENCODING:'utf-8'}});
  const browser = await chromium.launch({channel:'msedge',headless:true});
  try {
    const page = await browser.newPage();
    await page.setContent('<div id="previewPanel"></div><button id="exportBtn"></button><button id="analyzeBtn"></button><input id="templateKey"><input id="strategyFile"><div id="statusNode"></div>');
    await page.addScriptTag({content:`let sessionId='test'; const escapeHtml=s=>String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('"','&quot;'); ${js}`});
    await page.evaluate(()=>reviewReceive({rules:[{sourceRow:1,productCode:'01361',productName:'海韵'},{sourceRow:2,productCode:'01686',productName:'海韵中支'}],reviewRows:[1,2].map(n=>({sourceRow:n,id:String(n),originalText:'original'+n,strategyText:'text'+n,warnings:[],county:'县',marketType:'城市',businessArea:'',assignments:Object.fromEntries(Array.from({length:30},(_,i)=>[i+1,10])),manualTiers:[]}))}));
    assert.equal(await page.locator('#reviewSearch').inputValue(),'');
    assert.equal(await page.locator('#reviewContent').isVisible(),false);
    await page.locator('#reviewSearch').fill('海韵');
    assert.equal(await page.locator('#reviewResults button').count(),2);
    assert.equal(await page.locator('#reviewContent').isVisible(),false);
    await page.locator('[data-product-row="2"]').click();
    assert.equal(await page.locator('#reviewSearch').inputValue(),'');
    assert.match(await page.locator('#selectedProduct').textContent(),/01686.*海韵中支/);
    assert.equal(await page.locator('#reviewOriginal').textContent(),'original2');
    await page.locator('#reviewSearch').fill('不存在');
    assert.equal(await page.locator('#reviewResults button').count(),0);
    await page.locator('#reviewSearch').press('Escape');
    assert.equal(await page.locator('#reviewSearch').inputValue(),'');
    await page.locator('#reviewTable input').first().fill('100');
    await page.locator('#reviewSearch').fill('01361');
    assert.equal(await page.locator('#reviewOriginal').textContent(),'original2');
    assert.equal(await page.locator('#reviewSearch').inputValue(),'');
    await page.evaluate(()=>reviewReceive({rules:[{sourceRow:1,productCode:'01361',productName:'海韵'}],reviewRows:[1,2].map(n=>({sourceRow:1,id:String(n),originalText:'原文',strategyText:'原文',warnings:[],county:'凤山县营销部',marketType:'',businessArea:'商圈'+n,assignments:Object.fromEntries(Array.from({length:30},(_,i)=>[i+1,n])),manualTiers:[]}))}));
    await page.locator('#reviewSearch').fill('01361');await page.locator('#reviewResults button').click();
    assert.equal(await page.locator('#chartMarkets button').count(),1);
    assert.equal(await page.locator('#chartMarkets button').textContent(),'全部业态');
    await page.locator('#chartBusiness').selectOption('商圈2');
    assert.equal(await page.locator('[data-chart-tier="30"] .chart-value').textContent(),'2');
    assert.equal(await page.locator('#chartCounty option').count(),1);
    console.log('Search, explicit selection, no matches, Escape, unsaved-edit protection and business-area chart passed');
  } finally { await browser.close(); }
})().catch(e=>{console.error(e);process.exitCode=1;});
