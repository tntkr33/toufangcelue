"""Customer query workspace, appended after REVIEW_JS and ZERO_EXPORT_JS."""

CUSTOMER_CSS = r'''
    .customer-tabs { display:flex; gap:8px; border-bottom:1px solid var(--line); padding-bottom:12px; margin-bottom:20px; }
    .customer-tabs button[aria-selected="false"] { background:white; color:var(--ink); border-color:var(--line); }
    .customer-tabs button[aria-selected="true"], #customerViews button[aria-pressed="true"] { background:var(--brand); color:white; }
    #customerWorkspace { min-width:0; }
    #customerWorkspace [hidden], #strategyWorkspace[hidden] { display:none !important; }
    #customerWorkspace h2 { margin:0; }
    .customer-heading,.customer-toolbar,.customer-pagination { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }
    .customer-heading { margin-bottom:16px; }
    #customerControls { border:0; padding:0; margin:0; min-width:0; }
    .customer-upload { display:grid; grid-template-columns:minmax(0,1fr) auto; align-items:end; gap:12px; padding-bottom:18px; border-bottom:1px solid var(--line); }
    #customerFile { width:100%; }
    #customerFileInfo { overflow-wrap:anywhere; margin:8px 0 0; }
    .customer-upload-box { padding:24px; background:#f2f8f5; border:1px solid #b8d5c8; border-radius:6px; min-width:0; }
    #customerUploadTitle { display:block; font-size:18px; font-weight:700; margin-bottom:8px; }
    #customerFilename { margin:12px 0; font-size:15px; overflow-wrap:anywhere; }
    #customerPickFile { min-height:46px; min-width:180px; }
    .customer-cycle-controls { display:flex; flex-wrap:wrap; gap:20px 40px; padding:20px 0; border-bottom:1px solid var(--line); }
    .customer-cycle-controls button[aria-pressed="true"] { background:var(--brand); color:white; }
    .customer-cycle-controls button { min-width:88px; white-space:nowrap; }
    #customerCycleWarning { padding:12px; background:#fff8e8; border-left:3px solid #ab730c; }
    .customer-filters { display:grid; grid-template-columns:minmax(200px,2fr) minmax(170px,1fr) minmax(150px,1fr); gap:16px; padding:18px 0; align-items:start; }
    .customer-picker { position:relative; min-width:0; }
    .customer-picker > button { width:100%; text-align:left; overflow-wrap:anywhere; }
    .customer-menu { position:absolute; z-index:20; top:100%; left:0; width:100%; min-width:0; padding:12px; background:white; border:1px solid var(--line); border-radius:4px; box-shadow:0 6px 18px #163b2418; }
    .customer-options { max-height:240px; overflow:auto; margin-top:8px; }
    #customerWorkspace .customer-check { display:flex; align-items:center; gap:8px; margin:0; padding:7px 0; font-weight:400; overflow-wrap:anywhere; }
    #customerWorkspace input[type="checkbox"] { width:16px; height:16px; flex:none; margin:0; accent-color:var(--brand); }
    .customer-check span { min-width:0; }
    .customer-tier { display:flex; align-items:center; gap:8px; }
    .customer-tier input { width:100%; min-width:0; min-height:40px; border:1px solid var(--line); border-radius:4px; padding:8px; }
    .customer-markets { display:flex; flex-wrap:wrap; gap:8px 20px; margin-bottom:16px; }
    #customerStatus { margin:14px 0; overflow-wrap:anywhere; }
    #customerIncomplete { border-left:3px solid #ab730c; background:#fff8e8; padding:12px; margin:16px 0; color:#725008; }
    #customerMetrics { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; border-block:1px solid var(--line); padding:16px 0; margin:16px 0; }
    #customerMetrics dt { color:var(--muted); font-size:12px; }
    #customerMetrics dd { margin:5px 0 0; font-size:20px; font-weight:700; overflow-wrap:anywhere; }
    #customerMetrics .customer-quantity { font-size:17px; }
    #customerWorkspace .table-wrap { max-height:520px; margin-top:16px; }
    #customerTable { width:100%; }
    #customerTable th,#customerTable td { white-space:nowrap; }
    #customerTable td:last-child { white-space:normal; min-width:140px; overflow-wrap:anywhere; }
    .customer-pagination { padding:12px 0; }
    .customer-pagination select { width:auto; }
    #customerSaved { padding-top:12px; overflow-wrap:anywhere; }
    #customerSavedPath { flex-basis:100%; }
    @media(max-width:800px) { .customer-filters { grid-template-columns:repeat(2,minmax(0,1fr)); } #customerMetrics { grid-template-columns:repeat(2,minmax(0,1fr)); } }
    @media(max-width:480px) { .customer-filters,.customer-upload { grid-template-columns:minmax(0,1fr); } .customer-tabs button { flex:1; padding:10px 8px; } #customerWorkspace .panel { padding:14px; } #customerViews { width:100%; display:flex; } #customerViews button { flex:1; padding:8px 5px; } .customer-pagination .actions { flex-wrap:wrap; } }
'''

CUSTOMER_JS = r'''
    (() => {
      const main = document.querySelector('main');
      const strategy = document.createElement('div');
      strategy.id = 'strategyWorkspace';
      strategy.setAttribute('role', 'tabpanel');
      strategy.setAttribute('aria-labelledby', 'strategyTab');
      // Move existing nodes, preserving their listeners and element references.
      while (main.firstChild) strategy.append(main.firstChild);
      main.append(strategy);
      const tabs = document.createElement('div');
      tabs.className = 'customer-tabs'; tabs.setAttribute('role', 'tablist');
      tabs.setAttribute('aria-label', '工作区');
      tabs.innerHTML = '<button id="strategyTab" role="tab" aria-controls="strategyWorkspace" aria-selected="true" type="button">策略解析与复核</button><button id="customerTab" role="tab" aria-controls="customerWorkspace" aria-selected="false" tabindex="-1" class="secondary" type="button">客户投放查询</button>';
      main.prepend(tabs);
      const root = document.createElement('section');
      root.id = 'customerWorkspace'; root.hidden = true;
      root.setAttribute('role', 'tabpanel'); root.setAttribute('aria-labelledby', 'customerTab');
      root.innerHTML = `
        <div class="panel">
          <div class="customer-heading"><h2>客户投放查询</h2><span class="muted">策略测算投放量 · 非实际订货量</span></div>
          <fieldset id="customerControls">
            <div class="customer-upload"><div class="customer-upload-box"><span id="customerUploadTitle">1. 上传客户信息表</span><p id="customerFilename">上传客户信息表，开始统计</p><input id="customerFile" type="file" hidden accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"><button id="customerPickFile" type="button">选择客户信息表</button><p id="customerFileInfo" class="muted">支持 .xlsx，上传后显示文件名和客户数量</p></div><button id="customerChoices" type="button" class="secondary">前往策略复核</button></div>
            <div class="customer-cycle-controls"><div><label>2. 选择策略周</label><div class="segments" role="group" aria-label="策略周"><button type="button" data-parity="odd" aria-pressed="false" class="secondary">单周</button><button type="button" data-parity="even" aria-pressed="false" class="secondary">双周</button></div></div><div><label>本周是否为每月一订投放周</label><div class="segments" role="group" aria-label="每月一订投放周"><button type="button" data-monthly="yes" aria-pressed="false" class="secondary">是，计入投放</button><button type="button" data-monthly="no" aria-pressed="false" class="secondary">否，投放为 0</button></div></div></div>
            <p id="customerScopeStatus" role="status"></p>
            <div class="customer-filters">
              <div class="customer-picker"><label for="customerProductsToggle">参与测算品规</label><button id="customerProductsToggle" type="button" class="secondary" aria-expanded="false" aria-controls="customerProductsMenu">全部品规</button>
                <div id="customerProductsMenu" class="customer-menu" hidden><label for="customerProductSearch">搜索品规或编码</label><input id="customerProductSearch" type="search" autocomplete="off"><div class="actions"><button type="button" id="customerProductsAll" class="quiet">全选</button><button type="button" id="customerProductsNone" class="quiet">取消全选</button></div><div id="customerProductsOptions" class="customer-options"></div></div></div>
              <div class="customer-picker"><label for="customerCountiesToggle">营销部 / 区县</label><button id="customerCountiesToggle" type="button" class="secondary" aria-expanded="false" aria-controls="customerCountiesMenu">全部营销部</button>
                <div id="customerCountiesMenu" class="customer-menu" hidden><div class="actions"><button type="button" id="customerCountiesAll" class="quiet">全选</button><button type="button" id="customerCountiesNone" class="quiet">取消全选</button></div><div id="customerCountiesOptions" class="customer-options"></div></div></div>
              <div><label for="customerMinTier">档位范围（含端点）</label><div class="customer-tier"><input id="customerMinTier" aria-label="起始档位" type="number" min="1" max="30" step="1" value="1"><span>至</span><input id="customerMaxTier" aria-label="结束档位" type="number" min="1" max="30" step="1" value="30"></div></div>
            </div>
            <label>市场类型（扩展）</label><div class="customer-markets" id="customerMarkets"></div>
            <div class="customer-toolbar"><div id="customerViews" class="segments" role="group" aria-label="查询视图"><button type="button" data-view="summary" aria-pressed="true">品规汇总</button><button type="button" data-view="detail" aria-pressed="false">客户明细</button><button type="button" data-view="issues" aria-pressed="false">未匹配清单</button></div><div class="actions"><button id="customerResetFilters" type="button" class="secondary">重置筛选</button><button id="customerRetry" type="button" hidden>重试</button></div></div>
          </fieldset>
          <p id="customerStatus" role="status" aria-live="polite">请先在策略页解析周策略表。</p>
          <p id="customerIncomplete" hidden role="status">存在缺项：未匹配记录未按 0 条计入，以下数量不是完整投放总量。请查看未匹配清单。</p>
          <p id="customerCycleWarning" hidden role="status"></p>
          <dl id="customerMetrics" hidden></dl>
          <div class="customer-toolbar"><span id="customerResultTitle">品规汇总</span><button type="button" id="customerExport" class="secondary" disabled>导出当前筛选全部结果 XLSX</button></div>
          <div class="table-wrap" id="customerTableWrap" hidden><table id="customerTable"><thead></thead><tbody></tbody></table></div>
          <p id="customerEmpty">尚未测算</p>
          <div class="customer-pagination"><span id="customerPageInfo">0 条结果</span><div class="actions"><label for="customerPageSize">每页</label><select id="customerPageSize"><option>50</option><option>100</option><option>200</option></select><button type="button" id="customerPrev" class="secondary" disabled>上一页</button><button type="button" id="customerNext" class="secondary" disabled>下一页</button></div></div>
          <div id="customerSaved" class="actions" hidden><span id="customerSavedPath"></span><button type="button" id="customerOpen" class="secondary">打开文件</button><button type="button" id="customerReveal" class="secondary">打开所在位置</button></div>
        </div>`;
      main.append(root);
      const find = id => document.getElementById(id);
      const labels = {summary:'品规汇总', detail:'客户明细', issues:'未匹配清单'};
      let boundSession = null, state = null, result = null, saved = null;
      let generation = 0, refreshId = 0, working = false, draftDirty = true, view = 'summary';
      let queryTimer = null, querying = 0;
      let schedule = {weekParity:null, monthlyWeek:null};
      const scheduleReady = () => schedule.weekParity !== null && schedule.monthlyWeek !== null;
      let options = {products:[], counties:[], markets:[]};
      let selection = {productRows:null, counties:null, markets:null};
      const status = message => { find('customerStatus').textContent = message; };
      const current = (version, sid) => generation === version && sessionId === sid;
      const busy = () => working || reviewBusy || !!state?.busy;
      const hasCustomers = () => state?.customerSummary != null;
      function controls() {
        find('customerControls').disabled = busy() || !sessionId;
        find('customerRetry').disabled = busy() || !hasCustomers();
        find('customerPageSize').disabled = busy() || !!querying || !result || draftDirty;
        find('customerExport').disabled = busy() || !!querying || draftDirty || !result?.totalRows;
        find('customerPrev').disabled = busy() || !!querying || draftDirty || !result || result.page <= 1;
        find('customerNext').disabled = busy() || !!querying || draftDirty || !result || result.page * result.pageSize >= result.totalRows;
        find('customerOpen').disabled = find('customerReveal').disabled = busy() || !saved;
        root.setAttribute('aria-busy', String(working || !!querying));
      }
      function clearResult(message) {
        generation++; result = null; draftDirty = true; saved = null;
        find('customerSaved').hidden = true;
        find('customerMetrics').hidden = find('customerIncomplete').hidden = find('customerTableWrap').hidden = true;
        find('customerCycleWarning').hidden = true;
        find('customerTable').querySelector('tbody').replaceChildren();
        find('customerPageInfo').textContent = '0 条结果';
        find('customerEmpty').hidden = false; find('customerEmpty').textContent = '等待更新';
        if (message) status(message);
        controls();
      }
      function dirty() {
        generation++; draftDirty = true; saved = null;
        find('customerSaved').hidden = true;
        find('customerMetrics').hidden = find('customerIncomplete').hidden = find('customerTableWrap').hidden = true;
        find('customerCycleWarning').hidden = true;
        find('customerEmpty').hidden = false; find('customerEmpty').textContent = hasCustomers() ? '正在更新筛选结果…' : '上传客户表后自动统计';
        find('customerPageInfo').textContent = '待更新';
        status(hasCustomers() ? '正在更新筛选结果…' : '请上传客户基本信息表。'); controls(); scheduleQuery();
      }
      function scheduleQuery() {
        clearTimeout(queryTimer);
        if (!hasCustomers() || !draftDirty) return;
        if (!scheduleReady()) {
          status('请选择单周 / 双周，以及本周是否为每月一订投放周。');
          find('customerEmpty').textContent = '完成上方周次选择后自动统计';
          return;
        }
        queryTimer = setTimeout(() => {
          if (busy()) { scheduleQuery(); return; }
          void query();
        }, 250);
      }
      async function api(path, payload, sid = sessionId) {
        const request = {headers:{'X-Tool-Token':shutdownToken}};
        if (payload !== undefined) {
          request.method = 'POST';
          if (payload instanceof FormData) request.body = payload;
          else { request.headers['Content-Type'] = 'application/json'; request.body = JSON.stringify({sessionId:sid, ...payload}); }
        }
        const response = await fetch('/api/customer/' + path, request);
        const data = await response.json();
        if (!response.ok || data.ok === false) {
          const message = (data.errors || [data.error || '请求失败，请重试。']).join('；');
          throw new Error(message);
        }
        return data;
      }
      function renderState() {
        const summary = state?.customerSummary;
        const count = keys => keys.map(key => summary?.[key]).find(value => typeof value === 'number');
        const valid = count(['validCustomerCount','validCount','customerCount','validCustomers']);
        const issues = count(['issueCount','invalidCount','invalidCustomerCount','issues']);
        find('customerUploadTitle').textContent = state?.filename ? '1. 客户信息表已导入' : '1. 上传客户信息表';
        find('customerFilename').textContent = state?.filename || '上传客户信息表，开始统计';
        find('customerPickFile').textContent = state?.filename ? '更换客户信息表' : '选择客户信息表';
        find('customerFileInfo').textContent = state?.filename
          ? state.filename + (valid === undefined ? '' : ` · 有效客户 ${valid}`) + (issues === undefined ? '' : ` · 异常 ${issues}`) + (summary.footerRowCount ? ` · 已跳过 ${summary.footerRowCount} 行制表附注` : '')
          : '尚未上传客户表 · 支持 .xlsx 文件';
        find('customerScopeStatus').textContent = !state ? '' : state.choicesConfirmed
          ? '品规范围已复核 · 已排除投放的品规按 0 条统计'
          : '品规范围待复核 · 当前按默认建议统计，已排除投放的品规按 0 条统计';
        controls();
      }
      async function refresh() {
        const sid = sessionId, version = generation, requestId = ++refreshId;
        if (!sid) { state = null; renderState(); return; }
        const data = await api('state?' + new URLSearchParams({sessionId:sid}), undefined, sid);
        if (!current(version, sid) || requestId !== refreshId) return;
        if (state && (data.revision !== state.revision || data.reviewToken !== state.reviewToken)) clearResult('数据已更新，请重新测算。');
        state = data; renderState();
        if (state.busy) status('统计报表正在导出，请等待完成后再操作。');
        else if (!hasCustomers()) status('请上传客户基本信息表。');
        if (state) {
          const optionsVersion = generation;
          const data = await api('options?' + new URLSearchParams({sessionId:sid}), undefined, sid);
          if (!current(optionsVersion, sid) || requestId !== refreshId) return;
          if (data.revision !== state.revision) { clearResult('数据已更新，请重新测算。'); state = data; renderState(); }
          options = data.options; renderOptions();
        }
        scheduleQuery();
      }
      function checkLabel(text, checked, onChange) {
        const label = document.createElement('label'); label.className = 'customer-check';
        const input = document.createElement('input'); input.type = 'checkbox'; input.checked = checked;
        const span = document.createElement('span'); span.textContent = text;
        input.onchange = () => onChange(input.checked); label.append(input, span); return label;
      }
      const values = key => key === 'productRows' ? options.products.map(p => p.sourceRow) : options[key];
      function choose(key, value, checked) {
        const selected = new Set(selection[key] === null ? values(key) : selection[key]);
        if (checked) selected.add(value); else selected.delete(value);
        selection[key] = [...selected]; dirty(); renderOptions();
      }
      function renderOptions() {
        const search = find('customerProductSearch').value.trim().toLowerCase();
        for (const [key, prefix, allLabel] of [['productRows','Products','全部品规'],['counties','Counties','全部营销部']]) {
          const target = find('customer' + prefix + 'Options'); target.replaceChildren();
          const items = key === 'productRows' ? options.products.map(p => ({value:p.sourceRow, label:`${p.productCode} ${p.productName}（第 ${p.sourceRow} 行）${p.allocationStatus === '已排除投放' ? ' · 已排除投放（0条）' : ''}`})) : options.counties.map(value => ({value, label:value}));
          for (const item of items) {
            if (key === 'productRows' && !item.label.toLowerCase().includes(search)) continue;
            target.append(checkLabel(item.label, selection[key] === null || selection[key].includes(item.value), checked => choose(key, item.value, checked)));
          }
          if (!target.children.length) target.textContent = '没有匹配选项';
          find('customer' + prefix + 'Toggle').textContent = selection[key] === null ? allLabel : `已选 ${selection[key].length} 项`;
        }
        const markets = find('customerMarkets'); markets.replaceChildren();
        const allMarkets = checkLabel('全部市场类型', selection.markets === null, checked => { selection.markets = checked ? null : []; dirty(); renderOptions(); });
        markets.append(allMarkets);
        for (const market of options.markets) markets.append(checkLabel(market, selection.markets === null || selection.markets.includes(market), checked => choose('markets', market, checked)));
      }
      function closeMenus() {
        for (const prefix of ['Products','Counties']) { find('customer' + prefix + 'Menu').hidden = true; find('customer' + prefix + 'Toggle').setAttribute('aria-expanded', 'false'); }
      }
      async function run(action) {
        if (busy() || !sessionId) return;
        working = true; setReviewBusy(true); closeMenus(); controls();
        try { await action(); }
        catch (error) { status(error.message); }
        finally { working = false; setReviewBusy(false); controls(); scheduleQuery(); }
      }
      function filters() {
        if (!scheduleReady()) throw new Error('请先选择策略周和每月一订投放周。');
        const minTier = Number(find('customerMinTier').value), maxTier = Number(find('customerMaxTier').value);
        if (!Number.isInteger(minTier) || !Number.isInteger(maxTier) || minTier < 1 || maxTier > 30 || minTier > maxTier) throw new Error('档位范围须为 1 至 30 的整数，且起始档位不能大于结束档位。');
        return {...selection, ...schedule, minTier, maxTier};
      }
      function renderResult() {
        const totals = result.totals || {}, metrics = find('customerMetrics'); metrics.replaceChildren();
        const metric = (label, value, quantity = false) => {
          const item = document.createElement('div'), dt = document.createElement('dt'), dd = document.createElement('dd');
          dt.textContent = label; dd.textContent = value; if (quantity) dd.className = 'customer-quantity'; item.append(dt, dd); metrics.append(item);
        };
        metric('参与品规', totals.productCount ?? '—'); metric('本周应投客户数（去重）', totals.eligibleCustomerCount ?? '—');
        metric('匹配 / 未匹配客户', `${totals.matchedCustomerCount ?? '—'} / ${totals.unmatchedCustomerCount ?? '—'}`);
        metric('测算投放量', `${totals.quantityTiao ?? '—'} 条 / ${Number(totals.quantityXiang || 0).toFixed(3)} 箱`, true);
        metrics.hidden = false; find('customerIncomplete').hidden = !totals.incomplete;
        find('customerCycleWarning').hidden = !totals.cycleReviewCustomerCount;
        find('customerCycleWarning').textContent = `有 ${totals.cycleReviewCustomerCount || 0} 位客户的订货周期或周次异常，投放量已按 0 条计算。请在“未匹配清单”中复核客户编码、名称和异常原因。`;
        find('customerResultTitle').textContent = labels[result.view];
        const head = find('customerTable').querySelector('thead'), body = find('customerTable').querySelector('tbody');
        head.replaceChildren(); body.replaceChildren(); const header = document.createElement('tr');
        for (const column of result.columns) { const th = document.createElement('th'); th.scope = 'col'; th.textContent = column.label; header.append(th); }
        head.append(header);
        for (const row of result.rows) {
          const tr = document.createElement('tr');
          result.columns.forEach((column, index) => { const td = document.createElement('td'); let value = Array.isArray(row) ? row[index] : row[column.key]; if (value != null && /xiang/i.test(column.key)) value = Number(value).toFixed(3); td.textContent = value ?? ''; tr.append(td); });
          body.append(tr);
        }
        find('customerTableWrap').hidden = result.rows.length === 0;
        find('customerEmpty').hidden = result.rows.length > 0; find('customerEmpty').textContent = '当前筛选条件下没有结果';
        find('customerPageInfo').textContent = `共 ${result.totalRows} 条 · 第 ${result.page} / ${Math.max(1, Math.ceil(result.totalRows / result.pageSize))} 页`;
        find('customerPageSize').value = String(result.pageSize);
        status('测算完成。' + (totals.incomplete ? '存在缺项，请查看未匹配清单。' : ''));
        controls();
      }
      async function query(page = 1, paging = false) {
        if (busy() || !hasCustomers()) return;
        const sid = sessionId, version = generation;
        querying++; find('customerRetry').hidden = true; controls();
        try {
          let payload;
          if (paging) {
            if (!result || draftDirty) return;
            payload = {queryId:result.queryId};
          } else {
            const requestedFilters = filters();
            payload = {revision:state.revision, view, filters:requestedFilters};
          }
          status('正在统计投放量…');
          const data = await api('query', {...payload, page, pageSize:Number(find('customerPageSize').value)}, sid);
          if (!current(version, sid)) return;
          if (data.revision !== state?.revision) { clearResult('数据版本已变化，请重新测算。'); return; }
          result = data; draftDirty = false; renderResult();
        } catch (error) {
          if (current(version, sid)) {
            clearResult(error.message); find('customerRetry').hidden = false;
          }
        } finally { querying--; controls(); }
      }
      function resetFilters() {
        selection = {productRows:null, counties:null, markets:null};
        find('customerMinTier').value = '1'; find('customerMaxTier').value = '30'; find('customerProductSearch').value = '';
        renderOptions();
      }
      function resetCustomerQuery() {
        boundSession = sessionId; state = null; options = {products:[], counties:[], markets:[]};
        clearResult('请先在策略页解析周策略表。'); resetFilters();
        view = 'summary'; updateView(); renderState();
        schedule = {weekParity:null, monthlyWeek:null}; updateSchedule();
        find('customerFile').value = ''; find('customerPageSize').value = '50';
      }
      function onCustomerStrategyChanged(data) {
        if (boundSession !== sessionId) resetCustomerQuery();
        boundSession = sessionId;
        clearResult('策略已更新，正在刷新查询数据…');
        // A same-session strategy mutation retains the uploaded workbook and filter drafts.
        void refresh().catch(error => status(error.message));
      }
      function onCustomerChoicesChanged() {
        clearResult('品规范围已更新，正在刷新查询数据…');
        void refresh().catch(error => status(error.message));
      }
      Object.assign(window, {onCustomerStrategyChanged, onCustomerChoicesChanged, resetCustomerQuery});
      function selectTab(customer) {
        strategy.hidden = customer; root.hidden = !customer;
        for (const [id, selected] of [['strategyTab', !customer], ['customerTab', customer]]) { find(id).setAttribute('aria-selected', String(selected)); find(id).tabIndex = selected ? 0 : -1; }
        if (customer && !working) void refresh().catch(error => { clearResult(error.message); });
        controls();
      }
      find('strategyTab').onclick = () => selectTab(false); find('customerTab').onclick = () => selectTab(true);
      tabs.onkeydown = event => { if (['ArrowLeft','ArrowRight','Home','End'].includes(event.key)) { event.preventDefault(); const customer = event.key === 'End' || (event.key !== 'Home' && root.hidden); selectTab(customer); find(customer ? 'customerTab' : 'strategyTab').focus(); } };
      function updateView() {
        root.querySelectorAll('[data-view]').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.view === view)));
        find('customerResultTitle').textContent = labels[view];
      }
      root.querySelectorAll('[data-view]').forEach(button => { button.onclick = () => { if (view === button.dataset.view) return; view = button.dataset.view; updateView(); dirty(); }; });
      for (const [prefix, key] of [['Products','productRows'], ['Counties','counties']]) {
        find('customer' + prefix + 'Toggle').onclick = () => { const menu = find('customer' + prefix + 'Menu'), open = menu.hidden; closeMenus(); menu.hidden = !open; find('customer' + prefix + 'Toggle').setAttribute('aria-expanded', String(open)); };
        find('customer' + prefix + 'All').onclick = () => { selection[key] = null; dirty(); renderOptions(); };
        find('customer' + prefix + 'None').onclick = () => { selection[key] = []; dirty(); renderOptions(); };
      }
      document.addEventListener('click', event => { if (!event.target.closest('.customer-picker')) closeMenus(); });
      root.addEventListener('keydown', event => { if (event.key === 'Escape') { const menu = event.target.closest('.customer-menu'); closeMenus(); if (menu) menu.parentElement.querySelector('button').focus(); } });
      find('customerProductSearch').oninput = renderOptions;
      function updateSchedule() {
        root.querySelectorAll('[data-parity]').forEach(button => button.setAttribute('aria-pressed', String(schedule.weekParity === button.dataset.parity)));
        root.querySelectorAll('[data-monthly]').forEach(button => button.setAttribute('aria-pressed', String(schedule.monthlyWeek === (button.dataset.monthly === 'yes'))));
      }
      root.querySelectorAll('[data-parity]').forEach(button => { button.onclick = () => { schedule.weekParity = button.dataset.parity; updateSchedule(); dirty(); }; });
      root.querySelectorAll('[data-monthly]').forEach(button => { button.onclick = () => { schedule.monthlyWeek = button.dataset.monthly === 'yes'; updateSchedule(); dirty(); }; });
      find('customerPickFile').onclick = () => find('customerFile').click();
      find('customerMinTier').oninput = find('customerMaxTier').oninput = dirty;
      find('customerResetFilters').onclick = () => { resetFilters(); dirty(); };
      find('customerRetry').onclick = () => { void refresh().catch(error => status(error.message)); };
      find('customerChoices').onclick = () => { selectTab(false); find('reviewExportBtn').scrollIntoView({block:'center'}); find('reviewExportBtn').focus(); };
      find('customerPrev').onclick = () => query(result.page - 1, true);
      find('customerNext').onclick = () => query(result.page + 1, true);
      find('customerPageSize').onchange = () => query(1, true);
      find('customerFile').onchange = () => run(async () => {
        const file = find('customerFile').files[0]; if (!file) return;
        clearResult('正在读取客户表…');
        try {
          await refresh(); const sid = sessionId, version = generation;
          const form = new FormData(); form.append('customerWorkbook', file); form.append('sessionId', sid); form.append('reviewToken', state.reviewToken);
          const data = await api('upload', form, sid);
          if (!current(version, sid)) return;
          state = data; renderState(); await refresh(); status('客户表已上传，正在自动统计投放量…');
        } finally { find('customerFile').value = ''; }
      });
      find('customerExport').onclick = () => run(async () => {
        if (!result || draftDirty || !result.totalRows) return;
        const snapshot = result, sid = sessionId, version = generation;
        saved = null; find('customerSaved').hidden = true; status('请选择统计报表保存位置…');
        const chosen = await api('save-location', {queryId:snapshot.queryId, filename:labels[snapshot.view] + '.xlsx'}, sid);
        if (!current(version, sid)) return;
        if (chosen.cancelled) { status('已取消保存，查询结果保留。'); return; }
        const job = await api('export-start', {queryId:snapshot.queryId, saveId:chosen.saveId}, sid);
        // Once started, keep the strategy/exit lock until a terminal server status.
        for (;;) {
          let progress;
          try { progress = await api('export-status?' + new URLSearchParams({sessionId:sid, jobId:job.jobId}), undefined, sid); }
          catch (error) { status('暂时无法获取导出进度，正在重试。请勿关闭工具。' + error.message); await new Promise(resolve => setTimeout(resolve, 2000)); continue; }
          if (progress.status === 'failed') throw new Error(progress.error || '统计导出失败，请重试。');
          if (progress.status === 'complete') {
            if (current(version, sid)) { saved = progress.result; find('customerSavedPath').textContent = saved.selectedPath || saved.filename; find('customerSaved').hidden = false; status(`已保存 ${progress.rowCount ?? snapshot.totalRows} 条结果` + (progress.sheetCount ? ` · ${progress.sheetCount} 个工作表` : '') + '。'); }
            return;
          }
          status(`正在导出 ${progress.rowCount ?? 0} / ${progress.totalRows ?? snapshot.totalRows} 条` + (progress.sheetCount ? ` · ${progress.sheetCount} 个工作表` : '') + '…');
          await new Promise(resolve => setTimeout(resolve, 700));
        }
      });
      for (const [id, action] of [['customerOpen','open'], ['customerReveal','reveal']]) find(id).onclick = () => run(async () => { if (saved) await fileAction('/api/open-export', {saveId:saved.saveId, action}); });
      new MutationObserver(controls).observe(analyzeBtn, {attributes:true, attributeFilter:['disabled']});
      renderOptions(); controls();
    })();
'''
