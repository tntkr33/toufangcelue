const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const {execFileSync,spawn}=require('node:child_process');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const os=require('node:os');
const path=require('node:path');
const python=execFileSync('py',['-3.14','-c','import sys; print(sys.executable)'],{encoding:'utf8'}).trim();
const env={...process.env,PYTHONIOENCODING:'utf-8'};
(async()=>{
  const directory=fs.mkdtempSync(path.join(os.tmpdir(),'strategy-zero-'));
  const target=path.join(directory,'202610W1(10.01-10.04).xls');
  const server=spawn(python,['-u','-c',"from strategy_excel_tool import server as s, desktop_actions as d; from pathlib import Path; import os; root=Path(os.environ['TEST_EXPORT_DIR']); d.choose_save_path=lambda name:root/name; s.OUTPUT_DIR=root/'backup'; h,_=s.create_http_server('127.0.0.1',0,s.Handler); print(h.server_address[1],flush=True); h.serve_forever()"],{env:{...env,TEST_EXPORT_DIR:directory},stdio:['ignore','pipe','pipe']});
  let browser;
  try {
    const port=await new Promise((resolve,reject)=>{server.stdout.once('data',d=>resolve(Number(d.toString().trim())));server.once('error',reject);});
    browser=await chromium.launch({channel:'msedge',headless:true}); const page=await browser.newPage();
    await page.goto(`http://127.0.0.1:${port}`);
    const buffer=execFileSync(python,['-c',"from strategy_excel_tool import server as s; import sys; sys.stdout.buffer.write(s.build_demo_xlsx(['商品编码','商品名称','本期投放策略'],[['01614','停投品规','停投'],['01686','无库存品规','无库存'],['01361','正常品规','1-30档3条']]))"],{env});
    await page.locator('#strategyFile').setInputFiles({name:'test.xlsx',mimeType:'application/octet-stream',buffer});
    await page.locator('#analyzeBtn').click(); await page.locator('#reviewPanel').waitFor({state:'visible'});
    await page.locator('#exportName').fill('202610W1(10.01-10.04)');
    await page.locator('#exportBtn').click(); await page.locator('#zeroExportDialog').waitFor({state:'visible'});
    assert.equal(await page.locator('#zeroExportDialog tbody tr').count(),3);
    assert.equal(await page.locator('#zeroExportDialog tbody select').first().inputValue(),'omit');
    await page.locator('#zeroCancel').click(); assert.equal(fs.existsSync(target),false);
    await page.locator('#exportBtn').click(); await page.locator('#zeroSelectAll').check(); await page.locator('#zeroBatchKeep').click();
    assert.equal(await page.locator('#zeroExportDialog tbody select').first().inputValue(),'keep');
    await page.locator('#zeroExportDialog tbody select').first().selectOption('omit');
    await page.locator('#zeroConfirm').click(); await page.locator('#savedFileActions').waitFor({state:'visible'});
    const bytes=fs.readFileSync(target);
    execFileSync(python,['-c',"import sys,xlrd; s=xlrd.open_workbook(file_contents=sys.stdin.buffer.read(),formatting_info=True).sheet_by_index(0); assert s.nrows==92; assert s.rowinfo_map[0].hidden; assert {s.cell_value(r,2) for r in range(4,s.nrows)}=={'01686','01361'}; assert all(s.cell_value(r,c)==0 for r in range(4,48) for c in range(8,38))"],{env,input:bytes});
    console.log('Stop list defaults, cancel, bulk keep, individual omit and actual XLS output passed.');
  } finally {if(browser)await browser.close();server.kill();await new Promise(r=>server.exitCode!==null?r():server.once('exit',r));}
})().catch(e=>{console.error(e);process.exitCode=1;});
