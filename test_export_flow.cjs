const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const {execFileSync,spawn}=require('node:child_process');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const os=require('node:os');
const path=require('node:path');
const python=execFileSync('py',['-3.14','-c','import sys; print(sys.executable)'],{encoding:'utf8'}).trim();
const env={...process.env,PYTHONIOENCODING:'utf-8'};
(async()=>{
  const directory=fs.mkdtempSync(path.join(os.tmpdir(),'strategy-desktop-'));
  const script=`from strategy_excel_tool import server as s, desktop_actions as d
from pathlib import Path
import os,json
root=Path(os.environ['TEST_EXPORT_DIR'])
s.OUTPUT_DIR=root/'backup'
d.choose_save_path=lambda name: None if name=='cancel.xls' else root/('202610W2(10.12-10.18).xls' if name=='rename-me.xls' else name)
d.open_saved_path=lambda p,a: (root/'opened.json').write_text(json.dumps([str(p),a]),encoding='utf-8')
h,_=s.create_http_server('127.0.0.1',0,s.Handler)
print(h.server_address[1],flush=True)
h.serve_forever()
h.server_close()
`;
  const server=spawn(python,['-u','-c',script],{env:{...env,TEST_EXPORT_DIR:directory},stdio:['ignore','pipe','pipe']});
  let browser;
  server.stderr.on('data',d=>process.stderr.write(d));
  try {
    const port=await new Promise((resolve,reject)=>{server.stdout.once('data',d=>resolve(Number(d.toString().trim())));server.once('error',reject);});
    browser=await chromium.launch({channel:'msedge',headless:true}); const page=await browser.newPage();
    const errors=[]; page.on('pageerror',e=>errors.push(String(e)));
    const url=`http://127.0.0.1:${port}`;
    await page.goto(url);
    await page.waitForFunction(()=>document.getElementById('templateKey').options.length>0);
    assert.match(await page.locator('#templateKey').textContent(),/档位\+区县\+市场类型/);
    assert.ok(!(await page.locator('#templateKey').textContent()).includes('区倒'));
    const buffer=execFileSync(python,['-c',"from strategy_excel_tool import server as s; import sys; sys.stdout.buffer.write(s.build_demo_xlsx(['商品编码','商品名称','本期投放策略'],[['01686','海韵中支','1-30档2条']]))"],{env});
    await page.locator('#strategyFile').setInputFiles({name:'test.xlsx',mimeType:'application/octet-stream',buffer});
    await page.locator('#analyzeBtn').click(); await page.locator('#reviewPanel').waitFor({state:'visible'});
    function checkTitle(file,period) {
      const titles=JSON.parse(execFileSync(python,['-c',"import sys,xlrd,json; s=xlrd.open_workbook(file_contents=sys.stdin.buffer.read(),formatting_info=True).sheet_by_index(0); assert s.rowinfo_map[0].hidden; print(json.dumps([c for r in range(4) for c in s.row_values(r) if isinstance(c,str) and c.startswith('2026')]))"],{env,input:fs.readFileSync(file),encoding:'utf8'}));
      assert.ok(titles.includes(period),JSON.stringify(titles));
    }
    for(const [input,period] of [['202610W1(10.01-10.04)','202610W1(10.01-10.04)'],['rename-me','202610W2(10.12-10.18)']]) {
      await page.locator('#exportName').fill(input);
      await page.locator('#exportBtn').click();
      if(input==='202610W1(10.01-10.04)')await page.locator('#zeroConfirm').click();
      await page.locator('#savedFileActions').waitFor({state:'visible'});
      await page.waitForFunction(()=>!document.getElementById('exportBtn').disabled);
      assert.equal(await page.locator('#exportName').inputValue(),period);
      const saved=path.join(directory,period+'.xls');
      assert.ok((await page.locator('#status').textContent()).includes(saved));
      checkTitle(saved,period);
      for (const [button,action] of [['openSavedFile','open'],['revealSavedFile','reveal']]) {
        const response=page.waitForResponse(r=>r.url().endsWith('/api/open-export'));
        await page.locator('#'+button).click(); assert.equal((await response).status(),200);
        assert.deepEqual(JSON.parse(fs.readFileSync(path.join(directory,'opened.json'),'utf8')),[saved,action]);
      }
    }
    await page.screenshot({path:path.join(directory,'desktop.png'),fullPage:false});
    await page.setViewportSize({width:390,height:844});
    await page.screenshot({path:path.join(directory,'mobile.png'),fullPage:false});
    assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));
    await page.setViewportSize({width:1280,height:900});
    assert.equal((await page.request.post(url+'/api/open-export',{data:{saveId:'fake',action:'open'}})).status(),403);
    const token=await page.evaluate(()=>shutdownToken);
    assert.equal((await page.request.post(url+'/api/open-export',{headers:{'X-Tool-Token':token},data:{saveId:'fake',action:'open'}})).status(),400);
    await page.locator('#exportName').fill('cancel'); await page.locator('#exportBtn').click();
    await page.waitForFunction(()=>document.getElementById('status').textContent.includes('已取消保存'));
    assert.ok(await page.locator('#savedFileActions').isHidden());
    const invalid='202610W1(10.15-10.111)';
    await page.locator('#exportName').fill(invalid);
    await page.locator('#exportBtn').click(); await page.locator('#periodDialog').waitFor({state:'visible'});
    assert.equal(await page.locator(':focus').getAttribute('id'),'periodBack');
    await page.locator('#periodBack').click();
    await page.waitForFunction(()=>!document.getElementById('exportBtn').disabled);
    assert.ok(!fs.existsSync(path.join(directory,invalid+'.xls')));
    await page.locator('#exportBtn').click(); await page.locator('#periodDialog').waitFor({state:'visible'});
    const expected=await page.locator('#periodProposed').textContent();
    await page.locator('#periodAccept').click(); await page.locator('#savedFileActions').waitFor({state:'visible'});
    checkTitle(path.join(directory,invalid+'.xls'),expected);
    fs.unlinkSync(path.join(directory,invalid+'.xls'));
    await page.locator('#openSavedFile').click();
    await page.waitForFunction(()=>document.getElementById('status').textContent.includes('移动、删除'));
    assert.deepEqual(errors,[]);
    const other=await browser.newPage(); await other.goto(url);
    await page.reload(); await page.waitForTimeout(12000);
    assert.equal(server.exitCode,null);
    await other.close(); await page.waitForTimeout(12000);
    assert.equal(server.exitCode,null);
    await page.close();
    await new Promise((resolve,reject)=>{
      const timeout=setTimeout(()=>reject(new Error('Last tab did not stop the server')),16000);
      server.once('exit',code=>{clearTimeout(timeout);code===0?resolve():reject(new Error('Exit '+code));});
    });
    console.log('Native save target, open/reveal, rename/title, cancel, missing file, authorization, refresh/multi-tab/auto-exit passed. Screenshots: '+directory);
  } finally {
    if(browser)await browser.close();
    if(server.exitCode===null){server.kill();await new Promise(r=>server.once('exit',r));}
  }
})().catch(e=>{console.error(e);process.exitCode=1;});
