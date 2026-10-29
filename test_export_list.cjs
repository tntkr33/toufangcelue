const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const {execFileSync}=require('node:child_process');
const {mkdirSync}=require('node:fs');
const assert=require('node:assert/strict');
const py=code=>execFileSync('py',['-3.14','-c',code],{encoding:'utf8',env:{...process.env,PYTHONIOENCODING:'utf-8'}});
(async()=>{
  const js=py('from strategy_excel_tool.zero_export_ui import ZERO_EXPORT_JS; print(ZERO_EXPORT_JS)');
  const css=py('from strategy_excel_tool.ui_styles import APP_CSS; print(APP_CSS)');
  const browser=await chromium.launch({channel:'msedge',headless:true});
  try {
    const page=await browser.newPage({viewport:{width:1440,height:1000}});
    await page.setContent('<style>'+css+'</style><h1>导出清单测试</h1>');await page.addScriptTag({content:js});
    const items=[
      {sourceRow:2,productCode:'01',productName:'停投品规',strategyText:'停投',originalText:'停投',category:'stop',defaultChoice:'omit',reason:'整品规停投，建议不导出',warnings:[]},
      {sourceRow:3,productCode:'02',productName:'凌云',strategyText:'按价位段投放',originalText:'按价位段投放',category:'price',defaultChoice:'omit',reason:'整品规按价位段投放，建议不导出',warnings:['无法解析策略行：按价位段投放']},
      {sourceRow:4,productCode:'03',productName:'混合品规',strategyText:'南丹：按价位段投放。凤山：1-30档3条',originalText:'南丹：按价位段投放。凤山：1-30档3条',category:'review',defaultChoice:'keep',reason:'局部或混合策略，默认保留，请复核',warnings:['县区南丹后没有识别到档位规则']},
      ...Array.from({length:178},(_,i)=>({sourceRow:5+i,productCode:String(70000+i),productName:'正常品规'+i,strategyText:'1-30档5条',originalText:'1-30档5条',category:'normal',defaultChoice:'keep',reason:'正常导出',warnings:[]}))
    ];
    await page.evaluate(items=>{window.result='pending';confirmZeroExport({exportProducts:items,sourceProductCount:items.length}).then(r=>window.result=r);},items);
    assert.equal(await page.locator('tbody tr').count(),181);
    assert.equal(await page.locator('[data-export-row="3"] select').inputValue(),'omit');
    assert.equal(await page.locator('[data-export-row="4"] select').inputValue(),'keep');
    await page.locator('#exportProductSearch').fill('凌云');assert.equal(await page.locator('tbody tr').count(),1);
    assert.match(await page.locator('tbody').textContent(),/无法解析策略行/);
    await page.locator('#zeroSelectAll').check();await page.locator('#zeroBatchKeep').click();
    await page.locator('#exportProductSearch').fill('70177');await page.locator('tbody select').selectOption('omit');
    await page.locator('#exportProductSearch').fill('');
    assert.equal(await page.locator('[data-export-row="3"] select').inputValue(),'keep');
    await page.locator('#exportProductFilter').selectOption('review');
    assert.equal(await page.locator('#zeroBatchOmit').isDisabled(),true);
    await page.locator('#zeroSelectAll').check();await page.locator('#zeroBatchOmit').click();
    await page.locator('#exportProductFilter').selectOption('all');
    assert.equal(await page.locator('[data-export-row="3"] select').inputValue(),'keep');
    assert.equal(await page.locator('[data-export-row="4"] select').inputValue(),'omit');
    await page.locator('#exportProductSearch').fill('不存在');assert.equal(await page.locator('#exportListEmpty').isVisible(),true);
    await page.locator('#exportProductSearch').fill('');
    const shots='strategy_excel_tool/outputs/ui-v1.9-update';mkdirSync(shots,{recursive:true});
    await page.screenshot({path:shots+'/export-desktop.png'});
    await page.setViewportSize({width:390,height:844});await page.screenshot({path:shots+'/export-mobile.png'});
    assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));
    await page.locator('#zeroConfirm').click();
    const result=await page.evaluate(()=>window.result);
    assert.equal(Object.keys(result).length,181);assert.equal(result['3'],'keep');assert.equal(result['4'],'omit');assert.equal(result['182'],'omit');
    await page.evaluate(items=>{window.result='pending';confirmZeroExport({exportProducts:items.slice(0,1)}).then(r=>window.result=r);},items);
    assert.equal(await page.locator('#zeroConfirm').isDisabled(),true);
    await page.locator('#zeroCancel').click();assert.equal(await page.evaluate(()=>window.result),null);
    console.log('181-product export list: search all, exact warnings, mixed defaults, visible-only batch, retained decisions, no-match, cancel and all-omit guard passed.');
  } finally {await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
