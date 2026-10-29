// Synthetic browser contract tests; no customer workbook or local service required.
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const {execFileSync, spawn} = require('node:child_process');
const {createInterface} = require('node:readline');
const {mkdtempSync} = require('node:fs');
const {tmpdir} = require('node:os');
const {join} = require('node:path');
const assert = require('node:assert/strict');
const py = code => execFileSync('py', ['-3.14', '-c', code], {
  cwd:__dirname, encoding:'utf8', maxBuffer:8 * 1024 * 1024,
  env:{...process.env, PYTHONIOENCODING:'utf-8'},
});

(async () => {
  const html = py(`from server import INDEX_HTML
from customer_ui import CUSTOMER_JS, CUSTOMER_CSS
page = INDEX_HTML
if "strategy.id = 'strategyWorkspace'" not in page:
    page = page.replace('    loadTemplates();', CUSTOMER_JS + '\\n    loadTemplates();')
    page = page.replace('</style>', CUSTOMER_CSS + '</style>')
print(page.replace('__SHUTDOWN_TOKEN__', 'test-token'))`);
  const products = [2, 3].map((sourceRow, i) => ({sourceRow, productCode:'00' + i,
    productName:'合成品规' + i, defaultChoice:'keep', category:'normal', reason:'正常',
    strategyText:'1-30档2条', originalText:'1-30档2条', warnings:[]}));
  let state = {ok:true, revision:1, reviewToken:'review-1', filename:'', customerSummary:null,
    choicesConfirmed:false, choices:null, exportProducts:products, busy:false};
  const options = {products, counties:['合成甲县','合成乙县'], markets:['城市','农村']};
  const requests = [], errors = [];
  let queryNumber = 0, queryDelay = 0, cancelSave = false, failExport = false;
  let exportRunning = false, progressCalls = 0, failQuery = false;
  const plans = new Map();
  const browser = await chromium.launch({channel:'msedge', headless:true});
  try {
    const page = await browser.newPage({viewport:{width:1440, height:1000}});
    page.on('pageerror', error => errors.push(error.message));
    await page.addInitScript(() => { window.EventSource = class {close() {}}; });
    await page.route('http://customer.test/**', async route => {
      const request = route.request(), url = new URL(request.url());
      const reply = (data, status = 200) => route.fulfill({status, contentType:'application/json', body:JSON.stringify(data)});
      if (url.pathname === '/') return route.fulfill({contentType:'text/html', body:html});
      if (url.pathname === '/api/templates') return reply({ok:true, templates:[], innerCodeMapping:{exists:true}, templateDirectory:'synthetic'});
      if (url.pathname === '/api/open-export') {
        requests.push({path:'open-export', payload:request.postDataJSON()}); return reply({ok:true});
      }
      if (!url.pathname.startsWith('/api/customer/')) return reply({ok:true});
      assert.equal(request.headers()['x-tool-token'], 'test-token');
      const path = url.pathname.split('/').pop();
      const payload = request.method() === 'POST' && path !== 'upload' ? request.postDataJSON() : null;
      requests.push({path, payload});
      if (path === 'state') return reply(state);
      if (path === 'options') return reply({...state, options});
      if (path === 'upload') {
        assert.match(request.postData(), /name="customerWorkbook"/);
        assert.match(request.postData(), /name="sessionId"/);
        assert.match(request.postData(), /review-1/);
        state = {...state, revision:state.revision + 1, filename:'synthetic.xlsx', customerSummary:{validCustomerCount:2, issueCount:1}};
        return reply(state);
      }
      if (path === 'choices') {
        assert.equal(payload.zeroToken, state.reviewToken);
        state = {...state, revision:state.revision + 1, choicesConfirmed:true, choices:payload.choices};
        return reply(state);
      }
      if (path === 'query') {
        if (failQuery) return reply({ok:false, errors:['合成查询错误']}, 409);
        const revision = state.revision;
        const id = payload.queryId || 'query-' + ++queryNumber;
        if (!payload.queryId) plans.set(id, {view:payload.view, filters:payload.filters});
        const plan = plans.get(id);
        const totalRows = Object.values(plan.filters).some(value => Array.isArray(value) && value.length === 0) ? 0 : 101;
        const rows = Array.from({length:Math.min(payload.pageSize, Math.max(0, totalRows - (payload.page - 1) * payload.pageSize))}, (_, i) => ({
          productCode:'0001', productName:i === 0 ? '<img src=x onerror=alert(1)>' : '合成品规', quantityTiao:2, quantityXiang:0.008,
        }));
        if (queryDelay) await new Promise(resolve => setTimeout(resolve, queryDelay));
        return reply({ok:true, queryId:id, revision, view:plan.view,
          columns:[{key:'productCode', label:'商品编码'}, {key:'productName', label:'商品名称'}, {key:'quantityTiao', label:'条'}, {key:'quantityXiang', label:'箱'}],
          rows, totalRows, page:payload.page, pageSize:payload.pageSize,
          totals:{productCount:2, customerCount:2, matchedCustomerCount:1, unmatchedCustomerCount:1, quantityTiao:2, quantityXiang:0.008, incomplete:true, detailRowCount:101}});
      }
      if (path === 'save-location') return reply(cancelSave ? {ok:true, cancelled:true} : {ok:true, saveId:'save-1'});
      if (path === 'export-start') { progressCalls = 0; return reply({ok:true, jobId:'job-1'}); }
      if (path === 'export-status') {
        progressCalls++;
        if (exportRunning) return reply({ok:true, status:'running', rowCount:50, totalRows:101, sheetCount:1});
        if (failExport) return reply({ok:true, status:'failed', error:'合成导出错误'});
        return reply({ok:true, status:'complete', rowCount:101, totalRows:101, sheetCount:1,
          result:{selectedPath:'C:/synthetic.xlsx', filename:'synthetic.xlsx', saveId:'save-1'}});
      }
      throw new Error('Unexpected endpoint: ' + path);
    });
    await page.goto('http://customer.test/');
    const wait = condition => page.waitForFunction(condition);
    const idle = () => wait(() => document.querySelector('#customerWorkspace').getAttribute('aria-busy') === 'false');
    const runQuery = async () => { await page.waitForTimeout(450); await idle(); };
    await page.locator('#customerTab').click();
    assert.equal(await page.locator('#customerFile').isDisabled(), true);
    assert.equal(await page.locator('#strategyWorkspace').isHidden(), true);
    assert.equal(await page.locator('#customerFile').getAttribute('accept'), '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await page.evaluate(() => { sessionId = 'session-1'; onCustomerStrategyChanged({}); });
    await wait(() => !document.querySelector('#customerFile').disabled && !document.querySelector('#customerControls').disabled);
    await page.locator('#customerProductsToggle').click();
    await page.locator('#customerProductSearch').fill('001');
    await page.waitForFunction(() => document.querySelectorAll('#customerProductsOptions input').length === 1);
    assert.equal(await page.locator('#customerQuery').count(), 0);
    await page.locator('#customerProductSearch').fill('');
    await page.locator('#customerProductsToggle').click();
    await page.locator('#customerFile').setInputFiles({name:'synthetic.xlsx', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer:Buffer.from('synthetic mock upload')});
    await idle();
    assert.match(await page.locator('#customerFileInfo').textContent(), /有效客户 2.*异常 1/);
    assert.match(await page.locator('#customerFilename').textContent(), /synthetic.xlsx/);
    assert.equal(await page.locator('#customerFile').isHidden(), true);
    assert.match(await page.locator('#customerPickFile').textContent(), /更换/);
    await page.waitForTimeout(350);
    assert.equal(requests.filter(r => r.path === 'query').length, 0);
    await page.locator('[data-parity="odd"]').click();
    await page.locator('[data-monthly="no"]').click();
    await runQuery();
    assert.equal(requests.filter(r => r.path === 'choices').length, 0);
    assert.match(await page.locator('#customerScopeStatus').textContent(), /品规范围待复核/);
    assert.equal(await page.locator('#customerTable tbody tr').count(), 50);
    assert.equal(await page.locator('#customerTable img').count(), 0);
    assert.match(await page.locator('#customerTable tbody tr').first().textContent(), /0001.*0\.008/);
    assert.equal(await page.locator('#customerIncomplete').isVisible(), true);
    assert.match(await page.locator('#customerMetrics').textContent(), /2 条 \/ 0\.008 箱/);
    assert.equal(await page.locator('#customerExport').isEnabled(), true);
    assert.equal(requests.filter(r => r.path === 'query').at(-1).payload.filters.weekParity, 'odd');
    await page.locator('[data-parity="even"]').click();
    assert.equal(await page.locator('#customerExport').isDisabled(), true);
    await runQuery();
    assert.equal(requests.filter(r => r.path === 'query').at(-1).payload.filters.weekParity, 'even');
    await page.locator('#customerNext').click(); await idle();
    const paging = requests.filter(r => r.path === 'query').at(-1).payload;
    assert.ok(paging.queryId); assert.equal(paging.page, 2); assert.equal(paging.filters, undefined);
    await page.locator('#customerPageSize').selectOption('100'); await idle();
    assert.equal(await page.locator('#customerTable tbody tr').count(), 100);
    await page.locator('#customerChoices').click();
    assert.equal(await page.locator('#customerWorkspace').isHidden(), true);
    await page.locator('#customerTab').click(); await idle();
    assert.equal(await page.locator('#customerExport').isEnabled(), true);
    await page.locator('#customerProductsToggle').click();
    await page.locator('#customerProductSearch').fill('001');
    assert.equal(await page.locator('#customerProductsOptions input').count(), 1);
    await page.locator('#customerProductsOptions input').uncheck();
    assert.equal(await page.locator('#customerExport').isDisabled(), true);
    await page.locator('#customerMinTier').fill('5');
    await page.locator('#strategyTab').click(); await page.locator('#customerTab').click();
    assert.equal(await page.locator('#customerMinTier').inputValue(), '5');
    assert.equal(await page.locator('#customerProductSearch').inputValue(), '001');
    assert.equal(await page.locator('#customerProductsToggle').textContent(), '已选 1 项');
    await runQuery();
    assert.deepEqual(requests.filter(r => r.path === 'query').at(-1).payload.filters.productRows, [2]);
    if (await page.locator('#customerProductsMenu').isHidden()) await page.locator('#customerProductsToggle').click();
    await page.locator('#customerProductsNone').click();
    await runQuery();
    assert.deepEqual(requests.filter(r => r.path === 'query').at(-1).payload.filters.productRows, []);
    assert.match(await page.locator('#customerEmpty').textContent(), /没有结果/);
    assert.equal(await page.locator('#customerExport').isDisabled(), true);
    await page.locator('#customerResetFilters').click();
    await page.locator('[data-view="issues"]').click(); await runQuery();
    assert.equal(requests.filter(r => r.path === 'query').at(-1).payload.view, 'issues');
    await page.setViewportSize({width:390, height:844});
    assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await page.locator('#customerProductsToggle').click();
    assert.ok(await page.evaluate(() => document.querySelector('#customerProductsMenu').getBoundingClientRect().right <= innerWidth));
    await page.locator('#customerProductsToggle').press('Escape');
    await page.setViewportSize({width:1440, height:1000});
    cancelSave = true;
    await page.locator('#customerExport').click(); await idle();
    assert.equal(requests.filter(r => r.path === 'export-start').length, 0);
    cancelSave = false; exportRunning = true;
    await page.locator('#customerExport').click();
    await wait(() => document.querySelector('#customerStatus').textContent.includes('50 / 101'));
    assert.equal(await page.locator('#analyzeBtn').isDisabled(), true);
    assert.equal(await page.locator('#exitTool').isDisabled(), true);
    assert.equal(await page.locator('#customerMinTier').isDisabled(), true);
    const exportRequest = requests.filter(r => r.path === 'export-start').at(-1).payload;
    assert.equal(exportRequest.queryId, requests.filter(r => r.path === 'query').at(-1).payload.queryId || 'query-' + queryNumber);
    assert.equal(exportRequest.page, undefined);
    exportRunning = false; await idle();
    assert.equal(await page.locator('#customerSaved').isVisible(), true);
    await page.locator('#customerOpen').click(); await idle();
    assert.deepEqual(requests.filter(r => r.path === 'open-export').at(-1).payload, {saveId:'save-1', action:'open'});
    await page.locator('#customerReveal').click(); await idle();
    assert.equal(requests.filter(r => r.path === 'open-export').at(-1).payload.action, 'reveal');
    failExport = true;
    await page.locator('#customerExport').click(); await idle();
    assert.match(await page.locator('#customerStatus').textContent(), /合成导出错误/);
    assert.equal(await page.locator('#analyzeBtn').isEnabled(), true);
    assert.equal(await page.locator('#customerSaved').isHidden(), true);
    failExport = false; queryDelay = 500;
    await page.locator('#customerMinTier').fill('6');
    await wait(() => document.querySelector('#customerStatus').textContent === '正在统计投放量…');
    state = {...state, revision:state.revision + 1};
    await page.evaluate(() => onCustomerChoicesChanged());
    await runQuery();
    assert.equal(await page.locator('#customerTableWrap').isVisible(), true);
    queryDelay = 0; failQuery = true;
    await page.locator('#customerMinTier').fill('8'); await runQuery();
    assert.match(await page.locator('#customerStatus').textContent(), /合成查询错误/);
    assert.equal(await page.locator('#customerExport').isDisabled(), true);
    failQuery = false;
    await page.locator('#customerMinTier').fill('31');
    const count = requests.filter(r => r.path === 'query').length;
    await runQuery(); assert.equal(requests.filter(r => r.path === 'query').length, count);
    await page.locator('#customerMinTier').fill('7');
    state = {...state, reviewToken:'review-2', revision:state.revision + 1, choicesConfirmed:false};
    await page.evaluate(() => onCustomerStrategyChanged({}));
    assert.equal(await page.locator('#customerMinTier').inputValue(), '7');
    assert.match(await page.locator('#customerFileInfo').textContent(), /synthetic.xlsx/);
    await page.evaluate(() => { sessionId = null; resetCustomerQuery(); });
    assert.equal(await page.locator('#customerMinTier').inputValue(), '1');
    assert.equal(await page.locator('#customerFile').isDisabled(), true);
    assert.match(await page.locator('#customerFileInfo').textContent(), /尚未上传/);
    assert.deepEqual(errors, []);
    console.log('Customer UI passed: catalog before upload, automatic queries without confirmation, pending review status, safe rendering, units, paging, filters, mobile, export locks, file actions, failures and stale responses.');
    await realServerSmoke(browser);
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });

async function realServerSmoke(browser) {
  const script = `import base64, json, sys, tempfile, threading
from pathlib import Path
import server as s
import desktop_actions as d
from test_customer_allocation import HEADERS, workbook, customer, reviewed
with tempfile.TemporaryDirectory() as directory:
    d.choose_save_path = lambda *args, **kwargs: Path(directory) / 'synthetic.xlsx'
    http, _ = s.create_http_server('127.0.0.1', 0, s.Handler)
    product = dict(sourceRow=3, productCode='00001', productName='Synthetic product', defaultChoice='keep', category='normal', reason='Synthetic', strategyText='Synthetic', originalText='Synthetic', warnings=[])
    s.SESSIONS['browser-synthetic'] = dict(templateKey='county', review=dict(zeroToken='real-review', reviewRows=[reviewed()], exportProducts=[product]))
    thread = threading.Thread(target=http.serve_forever)
    thread.start()
    content = workbook([r + ['一周一次'] for r in [customer(), customer('002', 30), customer('003', 0)]], HEADERS + ['订货周期'])
    print(json.dumps(dict(url='http://127.0.0.1:%s' % http.server_address[1], workbook=base64.b64encode(content).decode())), flush=True)
    try:
        sys.stdin.readline()
    finally:
        http.shutdown()
        http.server_close()
        thread.join()
`;
  const child = spawn('py', ['-3.14', '-u', '-c', script], {cwd:__dirname,
    env:{...process.env, PYTHONIOENCODING:'utf-8'}, stdio:['pipe','pipe','pipe']});
  let stderr = '';
  child.stderr.on('data', data => { stderr += data; });
  const exited = new Promise(resolve => child.once('exit', resolve));
  const lines = createInterface({input:child.stdout});
  let page;
  try {
    const info = await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('Server start timed out: ' + stderr)), 20000);
      lines.once('line', line => { clearTimeout(timer); try { resolve(JSON.parse(line)); } catch (error) { reject(error); } });
      child.once('error', error => { clearTimeout(timer); reject(error); });
      child.once('exit', code => { clearTimeout(timer); reject(new Error('Server exited ' + code + ': ' + stderr)); });
    });
    page = await browser.newPage({viewport:{width:1440, height:1000}});
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(info.url);
    await page.evaluate(() => { sessionId = 'browser-synthetic'; onCustomerStrategyChanged({}); });
    await page.locator('#customerTab').click();
    await page.locator('#customerFile').setInputFiles({name:'synthetic.xlsx', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer:Buffer.from(info.workbook, 'base64')});
    const idle = () => page.waitForFunction(() => document.querySelector('#customerWorkspace').getAttribute('aria-busy') === 'false');
    await idle();
    assert.match(await page.locator('#customerFileInfo').textContent(), /有效客户 2.*异常 1/);
    await page.locator('[data-parity="odd"]').click();
    await page.locator('[data-monthly="no"]').click();
    await page.waitForFunction(() => document.querySelectorAll('#customerTable tbody tr').length === 1);
    await idle();
    assert.equal(await page.locator('#customerTable tbody tr').count(), 1);
    assert.match(await page.locator('#customerMetrics').textContent(), /7 条 \/ 0\.028 箱/);
    assert.equal(await page.locator('#customerMetrics div').filter({hasText:'本周应投客户数'}).locator('dd').textContent(), '2');
    assert.equal(await page.locator('#customerIncomplete').isVisible(), true);
    await page.locator('#customerExport').click(); await idle();
    assert.equal(await page.locator('#customerSaved').isVisible(), true);
    assert.match(await page.locator('#customerSavedPath').textContent(), /synthetic\.xlsx/);
    await page.locator('[data-view="detail"]').click();
    await page.waitForTimeout(450); await idle();
    assert.equal(await page.locator('#customerTable tbody tr').count(), 2);
    assert.match(await page.locator('#customerTable tbody').textContent(), /001/);
    const screenshotDir = process.env.CUSTOMER_UI_SCREENSHOTS || mkdtempSync(join(tmpdir(), 'customer-ui-'));
    await page.screenshot({path:join(screenshotDir, 'customer-desktop.png'), fullPage:true});
    await page.setViewportSize({width:390, height:844});
    await page.screenshot({path:join(screenshotDir, 'customer-mobile.png'), fullPage:true});
    assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await page.locator('[data-view="issues"]').click();
    await page.waitForTimeout(450); await idle();
    assert.equal(await page.locator('#customerTable tbody tr').count(), 1);
    assert.deepEqual(errors, []);
    console.log('Real server passed: XLSX upload, summary keys, confirmation, summary/detail/issues matching, XLSX save, leading-zero codes and mobile layout. Screenshots: ' + screenshotDir);
  } finally {
    if (page) await page.close();
    child.stdin.end('\n');
    const timer = setTimeout(() => child.kill(), 5000);
    await exited; clearTimeout(timer); lines.close();
  }
}
