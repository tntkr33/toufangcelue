REVIEW_JS = r'''
    const expandIcon = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 3H3v5m13-5h5v5M3 16v5h5m13-5v5h-5"/></svg>';
    const collapseIcon = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 8h5V3m8 0v5h5M8 21v-5H3m18 0h-5v5"/></svg>';
    document.getElementById('previewPanel').insertAdjacentHTML('beforebegin', `
      <section class="panel" id="reviewPanel" hidden>
        <div class="section-heading"><div><div class="eyebrow">REVIEW WORKSPACE</div><h2>品规复核与修改</h2></div><p id="reviewState" role="status"></p></div>
        <div class="search-wrap">
          <label for="reviewSearch">搜索并选择品规</label>
          <input type="search" id="reviewSearch" placeholder="搜索品规名称或商品编码…" autocomplete="off" aria-controls="reviewResults" aria-expanded="false" />
          <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="10.5" cy="10.5" r="6.5"/><path d="m16 16 5 5"/></svg>
          <div id="reviewResults" hidden></div>
        </div>
        <div id="reviewEmpty">搜索并选择一个品规，开始复核</div>
        <div id="reviewContent" hidden>
          <div class="product-heading"><h3 id="selectedProduct"></h3><span id="selectedSource" class="muted"></span></div>
          <div class="strategy-section">
            <div class="strategy-heading"><h3>上传表中的原文</h3><button id="editStrategy" type="button" class="secondary" aria-expanded="false" aria-controls="strategyEditor">修改文字策略</button></div>
            <div class="strategy-preview" id="strategyPreview"><div id="reviewOriginal"></div><div id="appliedStrategy" hidden><label>已应用的文字策略</label><div id="reviewApplied"></div></div></div>
            <div id="strategyEditor" hidden>
              <label for="reviewText">修改当前文字策略</label><textarea id="reviewText" rows="4"></textarea>
              <div class="actions"><button id="applyText" type="button">重新解析并应用文字策略</button><button id="cancelText" type="button" class="quiet">取消文字修改</button></div>
            </div>
          </div>
          <div id="reviewWarnings" class="messages"></div>
          <div id="quantityAnchor"></div>
          <section id="quantityWorkspace" class="quantity-workspace" aria-label="档位数量表">
            <div class="quantity-heading"><div><h3>档位投放数量 <span class="muted"> / 条</span></h3><div id="quantityProduct"></div></div><button id="maximizeTable" class="secondary icon-button" title="最大化数量表" aria-label="最大化数量表" aria-expanded="false">${expandIcon}</button></div>
            <div class="quantity-toolbar">
              <div class="county-filter"><label for="reviewCounty">营销部 / 区县</label>
                <button id="reviewCounty" type="button" class="secondary" aria-expanded="false" aria-controls="countyMenu">全部营销部</button>
                <div id="countyMenu" hidden>
                  <div class="county-actions"><button id="countyAll" type="button" class="quiet">全选</button><button id="countyNone" type="button" class="quiet">取消全选</button><button id="countyInvert" type="button" class="quiet">反选</button></div>
                  <div id="countyOptions"></div>
                </div>
              </div>
              <div><label for="reviewMarket">业态</label><select id="reviewMarket"></select></div>
              <div class="actions"><button id="applyNumbers" type="button">应用数量修改</button><button id="discardReview" type="button" class="secondary">放弃未应用修改</button></div>
            </div>
            <div class="table-wrap"><table id="reviewTable"></table></div>
            <div class="table-footer"><span id="tableCount"></span><span id="quantityStatus" role="status"></span><span id="draftCount" role="status"></span></div>
          </section>
          <section class="chart-section">
            <div class="section-heading"><h3>档位投放图 <span class="muted"> / 条·档</span></h3><div class="legend"><span>解析数量</span><span>已人工修改</span></div></div>
            <div class="chart-toolbar">
              <div class="chart-dimensions"><label for="chartCounty">营销部</label><select id="chartCounty"></select><label for="chartBusiness" id="chartBusinessLabel" hidden>商圈</label><select id="chartBusiness" hidden></select></div>
              <div id="chartMarkets" class="segments" role="group" aria-label="图表业态"></div>
            </div>
            <div id="reviewChart" aria-label="30档至1档的投放条数"></div>
          </section>
        </div>
      </section>
      <dialog id="quantityDialog" aria-label="最大化档位数量表"></dialog>
      <dialog id="changeDialog" aria-labelledby="changeTitle">
        <div class="change-header"><div><div class="eyebrow">CONFIRM CHANGES</div><h2 id="changeTitle">确认应用修改</h2></div><button id="closeChanges" type="button" class="quiet icon-button" title="返回修改" aria-label="返回修改">×</button></div>
        <div class="change-scroll"><div id="changeProduct"></div><p id="changeSummary"></p><div id="changeText" hidden></div><div id="changeWarnings" class="messages"></div>
          <div class="table-wrap"><table id="changeTable"><thead><tr><th>营销部 / 区县</th><th>业态</th><th>商圈</th><th>档位</th><th>修改前</th><th>修改后</th></tr></thead><tbody></tbody></table></div>
          <p id="noQuantityChanges" class="muted" hidden>本次没有档位数量变化。</p>
        </div>
        <div class="change-footer"><span class="muted">确认后，后续导出使用修改后的数据。</span><div class="actions"><button id="cancelChanges" type="button" class="secondary">返回修改</button><button id="confirmChanges" type="button">确认并应用</button></div></div>
      </dialog>`);
    let reviewData=null, reviewSelected=null, reviewBusy=false;
    let numberDrafts=new Map(), textDirty=false, chartMarket='', pendingReview=null;
    let maximized=false, previousOverflow='';
    let countySelection=null, countyOptions=[];
    const el=id=>document.getElementById(id);
    const dirtyReview=()=>numberDrafts.size>0||textDirty;
    const productRules=()=>reviewData?.rules||[];
    function selectedLabel() {
      const p=productRules().find(p=>p.sourceRow===reviewSelected);
      return p?p.productCode+' '+p.productName:'';
    }
    function currentRows(){return (reviewData?.reviewRows||[]).filter(r=>r.sourceRow===reviewSelected);}
    function reviewStatus(message) {
      el('reviewState').textContent=message||(dirtyReview()?'存在未应用修改':(reviewSelected===null?'全部品规可搜索':'当前数据已应用'));
      el('reviewState').dataset.dirty=String(dirtyReview());
      el('quantityStatus').textContent=message||'';
      el('draftCount').textContent=numberDrafts.size?`${numberDrafts.size} 个档位待应用`:'无待应用数量修改';
      exportBtn.disabled=reviewBusy||dirtyReview()||!sessionId;
      el('reviewExportBtn').disabled=exportBtn.disabled;
    }
    function setReviewBusy(value) {
      reviewBusy=value;
      if(value)closeCountyMenu();
      for(const root of [el('reviewPanel'),el('quantityWorkspace')])root.querySelectorAll('input,select,textarea,button').forEach(c=>c.disabled=value);
      for(const c of [analyzeBtn,templateKey,strategyFile])c.disabled=value;
      if(el('exitTool'))el('exitTool').disabled=value;
      exportBtn.disabled=value||dirtyReview()||!sessionId;
      el('reviewExportBtn').disabled=exportBtn.disabled;
    }
    function reviewReceive(data) {
      reviewData=data;numberDrafts.clear();textDirty=false;el('reviewPanel').hidden=false;reviewOptions();
      if(typeof onCustomerStrategyChanged==='function')onCustomerStrategyChanged(data);
    }
    function reviewOptions() {
      if(!productRules().some(p=>p.sourceRow===reviewSelected))reviewSelected=null;
      closeReviewResults();el('reviewContent').hidden=reviewSelected===null;el('reviewEmpty').hidden=reviewSelected!==null;
      if(reviewSelected!==null)reviewProduct();else reviewStatus();
    }
    function closeReviewResults(clear=true) {
      el('reviewResults').hidden=true;el('reviewSearch').setAttribute('aria-expanded','false');
      if(clear)el('reviewSearch').value='';
    }
    function searchReviewProducts() {
      const q=el('reviewSearch').value.trim().toLowerCase();
      const products=productRules().filter(p=>(p.productCode+' '+p.productName).toLowerCase().includes(q));
      el('reviewResults').innerHTML=products.length?products.map(p=>`<button type="button" data-product-row="${p.sourceRow}"><span class="product-code">${escapeHtml(p.productCode)}</span><span>${escapeHtml(p.productName)}</span><span class="result-source">第 ${p.sourceRow} 行</span></button>`).join(''):'<p role="status" style="padding:12px 16px;margin:0">没有匹配的品规</p>';
      el('reviewResults').hidden=false;el('reviewSearch').setAttribute('aria-expanded','true');
    }
    function setEditor(open) {
      el('strategyEditor').hidden=!open;el('editStrategy').setAttribute('aria-expanded',String(open));
      el('editStrategy').textContent=open?'收起文字编辑':'修改文字策略';
    }
    function reviewProduct() {
      const rows=currentRows();if(!rows.length)return;
      const row=rows[0],p=productRules().find(p=>p.sourceRow===reviewSelected);
      el('selectedProduct').innerHTML=`<span class="product-code">${escapeHtml(p.productCode)}</span> <span>${escapeHtml(p.productName)}</span>`;
      el('selectedSource').textContent=`策略表第 ${reviewSelected} 行 · ${rows.length} 个维度`;
      el('quantityProduct').textContent=selectedLabel();el('reviewOriginal').textContent=row.originalText||'（空白策略）';
      el('reviewApplied').textContent=row.strategyText;el('appliedStrategy').hidden=row.originalText===row.strategyText;
      el('strategyPreview').classList.toggle('has-edit',row.originalText!==row.strategyText);
      el('reviewText').value=row.strategyText;setEditor(false);
      el('reviewWarnings').innerHTML=row.warnings.map(w=>`<div class="msg warning">${escapeHtml(w)}</div>`).join('');
      renderCountyPicker();
      for(const [id,key,label] of [['reviewMarket','marketType','全部业态']]) {
        const previous=el(id).value;
        el(id).innerHTML=`<option value="">${label}</option>`+[...new Set(rows.map(r=>r[key]).filter(Boolean))].map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
        if([...el(id).options].some(o=>o.value===previous))el(id).value=previous;
      }
      const county=el('chartCounty').value;
      el('chartCounty').innerHTML=[...new Set(rows.map(r=>r.county))].map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v||'全部地区')}</option>`).join('');
      if(rows.some(r=>r.county===county))el('chartCounty').value=county;
      renderReviewRows();renderChartSelectors();reviewStatus();
    }
    function closeCountyMenu() {el('countyMenu').hidden=true;el('reviewCounty').setAttribute('aria-expanded','false');}
    function updateCountyCaption() {
      const selected=countyOptions.filter(v=>countySelection.has(v));
      el('reviewCounty').textContent=selected.length===countyOptions.length?'全部营销部':
        selected.length===0?'未选择营销部':selected.length===1?(selected[0]||'全部地区'):`已选 ${selected.length} 个营销部`;
      el('reviewCounty').title=selected.map(v=>v||'全部地区').join('、');
      el('countyOptions').querySelectorAll('input').forEach(c=>c.checked=countySelection.has(c.value));
    }
    function renderCountyPicker() {
      const all=countySelection===null||countyOptions.every(v=>countySelection.has(v));
      countyOptions=[...new Set(currentRows().map(r=>r.county))];
      countySelection=all?new Set(countyOptions):new Set(countyOptions.filter(v=>countySelection.has(v)));
      el('countyOptions').innerHTML=countyOptions.map((v,i)=>`<div class="county-option"><label><input type="checkbox" value="${escapeHtml(v)}">${escapeHtml(v||'全部地区')}</label><button type="button" class="quiet" data-county-only="${i}">仅筛选此项</button></div>`).join('');
      updateCountyCaption();
    }
    function renderReviewRows() {
      const rows=currentRows().filter(r=>countySelection.has(r.county)&&(!el('reviewMarket').value||r.marketType===el('reviewMarket').value));
      const tiers=Array.from({length:30},(_,i)=>30-i);
      el('reviewTable').innerHTML='<thead><tr><th>营销部 / 业态</th>'+tiers.map(t=>`<th>${t}档</th>`).join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr><td>'+escapeHtml(r.county||'全部地区')+'<span class="row-dimension">'+escapeHtml([r.marketType,r.businessArea].filter(Boolean).join(' / ')||'全部业态')+'</span></td>'+tiers.map(t=>{
        const draft=numberDrafts.get(r.id+'|'+t),manual=r.manualTiers.map(String).includes(String(t));
        const value=draft?(draft.value??''):r.assignments[t];
        return `<td><input class="tier-input ${draft?'draft':manual?'manual':''}" aria-label="${escapeHtml([r.county,r.marketType,r.businessArea,t+'档'].filter(Boolean).join(' '))}" title="${draft?'待应用修改':manual?'已人工修改':'解析数量'}" type="number" min="0" max="1000000" step="1" data-id="${r.id}" data-tier="${t}" value="${value}" /></td>`;
      }).join('')+'</tr>').join('')+'</tbody>';
      el('tableCount').textContent=`${rows.length} 个维度 · 30 档`;
      if(!rows.length)el('reviewTable').querySelector('tbody').innerHTML='<tr><td colspan="31">'+(countySelection.size?'没有符合当前筛选条件的记录':'未选择营销部 / 区县')+'</td></tr>';
    }
    function renderChartSelectors() {
      const rows=currentRows().filter(r=>r.county===el('chartCounty').value),markets=[...new Set(rows.map(r=>r.marketType))];
      if(!markets.includes(chartMarket))chartMarket=markets[0]??'';
      el('chartMarkets').innerHTML=markets.map(v=>`<button type="button" data-market="${escapeHtml(v)}" aria-pressed="${v===chartMarket}">${escapeHtml(v||'全部业态')}</button>`).join('');
      const business=[...new Set(rows.filter(r=>r.marketType===chartMarket).map(r=>r.businessArea))],previous=el('chartBusiness').value;
      el('chartBusiness').innerHTML=business.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v||'未指定商圈')}</option>`).join('');
      if(business.includes(previous))el('chartBusiness').value=previous;
      el('chartBusiness').hidden=el('chartBusinessLabel').hidden=!business.some(Boolean);renderChart();
    }
    function currentChartRow() {return currentRows().find(r=>r.county===el('chartCounty').value&&r.marketType===chartMarket&&r.businessArea===el('chartBusiness').value);}
    function renderChart() {
      const row=currentChartRow();if(!row){el('reviewChart').textContent='没有匹配维度';return;}
      const values=Array.from({length:30},(_,i)=>{const t=30-i,d=numberDrafts.get(row.id+'|'+t);return d?(d.value===null?NaN:d.value):Number(row.assignments[t]);});
      const max=Math.max(1,...values.filter(Number.isFinite));
      el('reviewChart').innerHTML=values.map((v,i)=>`<button type="button" class="chart-bar ${numberDrafts.has(row.id+'|'+(30-i))?'draft':row.manualTiers.map(String).includes(String(30-i))?'manual':''}" data-chart-tier="${30-i}" title="${escapeHtml([row.county,row.marketType,row.businessArea].filter(Boolean).join(' / '))} · ${30-i}档：${Number.isFinite(v)?v:'未填写'}条" aria-label="定位${30-i}档，${Number.isFinite(v)?v:'未填写'}条" style="height:${Math.max(2,(Number.isFinite(v)?v:0)/max*170)}px"><span class="chart-value">${Number.isFinite(v)?v:'?'}</span><span class="chart-tier">${30-i}</span></button>`).join('');
    }
    el('reviewTable').addEventListener('input',e=>{
      if(!e.target.dataset.id)return;
      const input=e.target,tier=Number(input.dataset.tier),value=input.value===''?null:Number(input.value),key=input.dataset.id+'|'+tier;
      const row=currentRows().find(r=>r.id===input.dataset.id);
      if(value===Number(row.assignments[tier]))numberDrafts.delete(key);else numberDrafts.set(key,{id:row.id,tier,value});
      input.classList.toggle('draft',numberDrafts.has(key));reviewStatus();renderChart();
    });
    el('reviewChart').addEventListener('click',e=>{
      const bar=e.target.closest('[data-chart-tier]'),row=currentChartRow();if(!bar||!row)return;
      let input=el('reviewTable').querySelector(`input[data-id="${row.id}"][data-tier="${bar.dataset.chartTier}"]`);
      if(!input){countySelection=new Set([row.county]);updateCountyCaption();el('reviewMarket').value=row.marketType;renderReviewRows();input=el('reviewTable').querySelector(`input[data-id="${row.id}"][data-tier="${bar.dataset.chartTier}"]`);}
      if(input){input.focus();input.scrollIntoView({block:'nearest',inline:'center'});}
    });
    el('editStrategy').addEventListener('click',()=>{setEditor(el('strategyEditor').hidden);if(!el('strategyEditor').hidden)el('reviewText').focus();});
    el('cancelText').addEventListener('click',()=>{el('reviewText').value=currentRows()[0]?.strategyText||'';textDirty=false;setEditor(false);reviewStatus();});
    el('reviewText').addEventListener('input',()=>{textDirty=el('reviewText').value!==currentRows()[0]?.strategyText;reviewStatus();});
    el('reviewSearch').addEventListener('input',()=>{if(dirtyReview()){closeReviewResults();reviewStatus('请先应用或放弃当前修改，再搜索。');return;}searchReviewProducts();});
    el('reviewSearch').addEventListener('focus',()=>{if(!dirtyReview()&&!reviewBusy)searchReviewProducts();});
    el('reviewResults').addEventListener('click',e=>{
      const button=e.target.closest('[data-product-row]');if(!button||reviewBusy)return;
      if(dirtyReview()){closeReviewResults();reviewStatus('请先应用或放弃当前修改，再切换品规。');return;}
      reviewSelected=Number(button.dataset.productRow);closeReviewResults();reviewOptions();
    });
    el('reviewSearch').addEventListener('keydown',e=>{
      if(e.key==='Escape'){e.preventDefault();closeReviewResults();return;}
      if(e.key==='ArrowDown'||e.key==='Enter'){e.preventDefault();if(dirtyReview()||reviewBusy)return;if(el('reviewResults').hidden)searchReviewProducts();el('reviewResults').querySelector('button')?.focus();}
    });
    el('reviewResults').addEventListener('keydown',e=>{
      const buttons=[...el('reviewResults').querySelectorAll('button')],i=buttons.indexOf(document.activeElement);
      if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();buttons[Math.max(0,Math.min(buttons.length-1,i+(e.key==='ArrowDown'?1:-1)))]?.focus();}
      if(e.key==='Escape'){e.preventDefault();el('reviewSearch').focus();closeReviewResults();}
    });
    document.addEventListener('click',e=>{if(e.target!==el('reviewSearch')&&!el('reviewResults').contains(e.target))closeReviewResults();});
    el('reviewCounty').addEventListener('click',()=>{const open=el('countyMenu').hidden;el('countyMenu').hidden=!open;el('reviewCounty').setAttribute('aria-expanded',String(open));});
    el('countyOptions').addEventListener('change',e=>{if(e.target.type!=='checkbox')return;if(e.target.checked)countySelection.add(e.target.value);else countySelection.delete(e.target.value);updateCountyCaption();renderReviewRows();});
    el('countyOptions').addEventListener('click',e=>{const b=e.target.closest('[data-county-only]');if(!b)return;countySelection=new Set([countyOptions[Number(b.dataset.countyOnly)]]);updateCountyCaption();renderReviewRows();closeCountyMenu();el('reviewCounty').focus();});
    for(const id of ['countyAll','countyNone','countyInvert'])el(id).addEventListener('click',()=>{countySelection=new Set(id==='countyAll'?countyOptions:id==='countyNone'?[]:countyOptions.filter(v=>!countySelection.has(v)));updateCountyCaption();renderReviewRows();});
    document.addEventListener('click',e=>{if(!e.target.closest('.county-filter'))closeCountyMenu();});
    el('countyMenu').parentElement.addEventListener('keydown',e=>{if(e.key==='Escape'&&!el('countyMenu').hidden){e.preventDefault();e.stopPropagation();closeCountyMenu();el('reviewCounty').focus();}});
    el('reviewMarket').addEventListener('change',renderReviewRows);
    el('chartCounty').addEventListener('change',renderChartSelectors);
    el('chartMarkets').addEventListener('click',e=>{const b=e.target.closest('[data-market]');if(b){chartMarket=b.dataset.market;renderChartSelectors();}});
    el('chartBusiness').addEventListener('change',renderChart);
    el('discardReview').addEventListener('click',()=>{numberDrafts.clear();textDirty=false;reviewOptions();});
    function restoreQuantity() {
      if(!maximized)return;
      el('quantityAnchor').after(el('quantityWorkspace'));maximized=false;document.body.style.overflow=previousOverflow;
      el('maximizeTable').innerHTML=expandIcon;el('maximizeTable').title='最大化数量表';el('maximizeTable').setAttribute('aria-label','最大化数量表');el('maximizeTable').setAttribute('aria-expanded','false');el('maximizeTable').focus();
    }
    el('maximizeTable').addEventListener('click',()=>{
      if(maximized){el('quantityDialog').close();restoreQuantity();return;}
      previousOverflow=document.body.style.overflow;document.body.style.overflow='hidden';el('quantityDialog').append(el('quantityWorkspace'));maximized=true;
      el('maximizeTable').innerHTML=collapseIcon;el('maximizeTable').title='退出最大化';el('maximizeTable').setAttribute('aria-label','退出最大化');el('maximizeTable').setAttribute('aria-expanded','true');el('quantityDialog').showModal();el('maximizeTable').focus();
    });
    el('quantityDialog').addEventListener('close',restoreQuantity);
    function openConfirmation(result,payload) {
      pendingReview=payload;
      const previous=new Map(currentRows().map(r=>[r.id,r])),changes=[];
      for(const row of result.reviewRows.filter(r=>r.sourceRow===reviewSelected)) {
        const old=previous.get(row.id);if(!old)continue;
        for(let tier=30;tier>=1;tier--)if(Number(old.assignments[tier])!==Number(row.assignments[tier]))changes.push({row,tier,before:old.assignments[tier],after:row.assignments[tier]});
      }
      el('changeProduct').textContent=selectedLabel();
      el('changeSummary').textContent=`${payload.action==='text'?'文字策略重新解析':'手工调整数量'} · ${new Set(changes.map(c=>c.row.id)).size} 个维度 · ${changes.length} 个档位数量变更`;
      el('changeText').hidden=payload.action!=='text';
      el('changeText').innerHTML=payload.action==='text'?`<div><label>修改前文字</label><pre>${escapeHtml(currentRows()[0].strategyText)}</pre></div><div><label>修改后文字</label><pre>${escapeHtml(payload.text)}</pre></div>`:'';
      const warnings=[...new Set(result.reviewRows.filter(r=>r.sourceRow===reviewSelected).flatMap(r=>r.warnings))];
      if(payload.action==='text') {
        const count=currentRows().reduce((n,r)=>n+r.manualTiers.length,0);
        if(count||numberDrafts.size)warnings.unshift(`重新解析将清除该品规 ${count} 个已应用的手工档位覆盖、${numberDrafts.size} 个未应用数量修改；下方为重新解析后的最终数量差异。`);
      }
      el('changeWarnings').innerHTML=warnings.map(w=>`<div class="msg warning">${escapeHtml(w)}</div>`).join('');
      el('changeTable').querySelector('tbody').innerHTML=changes.map(({row,tier,before,after})=>`<tr><td>${escapeHtml(row.county||'全部地区')}</td><td>${escapeHtml(row.marketType||'全部业态')}</td><td>${escapeHtml(row.businessArea||'—')}</td><td>${tier}档</td><td>${before}</td><td class="new-value">${after}</td></tr>`).join('');
      el('changeTable').parentElement.hidden=!changes.length;el('noQuantityChanges').hidden=changes.length>0;
      el('changeDialog').showModal();el('cancelChanges').focus();
    }
    function cancelConfirmation(){if(el('confirmChanges').disabled)return;pendingReview=null;el('changeDialog').close();}
    el('cancelChanges').addEventListener('click',cancelConfirmation);el('closeChanges').addEventListener('click',cancelConfirmation);
    el('changeDialog').addEventListener('cancel',e=>{if(el('confirmChanges').disabled)e.preventDefault();});
    el('changeDialog').addEventListener('close',()=>{pendingReview=null;setReviewBusy(false);reviewStatus();});
    async function applyReview(action) {
      if(reviewBusy||reviewSelected===null)return;
      if(action==='numbers'&&textDirty){reviewStatus('文字也有修改，请先应用文字或取消文字修改。');return;}
      const changes=[...numberDrafts.values()];
      if(action==='numbers'&&!changes.length){reviewStatus('没有需要应用的数量修改。');return;}
      if(action==='text'&&!textDirty&&!numberDrafts.size&&!currentRows().some(r=>r.manualTiers.length)){reviewStatus('文字策略未变化。');return;}
      if(action==='numbers'&&changes.some(c=>!Number.isInteger(c.value)||c.value<0||c.value>1000000)){reviewStatus('条数必须为0到1000000的整数，不能留空。');return;}
      const payload={sessionId,action,sourceRow:reviewSelected,text:el('reviewText').value,changes};
      setReviewBusy(true);reviewStatus('正在生成修改确认清单…');
      try {
        const response=await fetch('/api/review-preview',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
        const data=await response.json();if(!response.ok||!data.ok)throw new Error((data.errors||['预览修改失败']).join('；'));
        openConfirmation(data,payload);
      }catch(error){setReviewBusy(false);reviewStatus('预览修改失败：'+error.message);}
    }
    el('confirmChanges').addEventListener('click',async()=>{
      if(!pendingReview)return;
      const payload=pendingReview;for(const id of ['confirmChanges','cancelChanges','closeChanges'])el(id).disabled=true;el('confirmChanges').textContent='正在应用…';
      try {
        const response=await fetch('/api/review',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
        const data=await response.json();if(!response.ok||!data.ok)throw new Error((data.errors||['应用失败']).join('；'));
        showMessages(data);showSummary(data.summary);showPreview(data.previewRows);el('rulesJson').textContent=JSON.stringify(data.rules,null,2);
        reviewReceive(data);el('changeDialog').close();setReviewBusy(false);reviewStatus('修改已应用，导出将使用这些数量。');statusNode.textContent='人工修改已应用，可导出 XLS。';
      }catch(error){el('changeWarnings').insertAdjacentHTML('beforeend',`<div class="msg error">应用失败：${escapeHtml(error.message)}</div>`);}
      finally{for(const id of ['confirmChanges','cancelChanges','closeChanges'])el(id).disabled=false;el('confirmChanges').textContent='确认并应用';}
    });
    el('applyText').addEventListener('click',()=>applyReview('text'));el('applyNumbers').addEventListener('click',()=>applyReview('numbers'));
    function invalidateReview() {
      countySelection=null;countyOptions=[];closeCountyMenu();
      if(maximized){el('quantityDialog').close();restoreQuantity();}
      sessionId=null;reviewData=null;reviewSelected=null;numberDrafts.clear();textDirty=false;el('reviewPanel').hidden=true;exportBtn.disabled=true;closeReviewResults();
      el('reviewExportBtn').disabled=true;
      for(const id of ['summaryPanel','previewPanel','rulesPanel'])el(id).hidden=true;statusNode.textContent='文件或格式已变化，请重新解析。';
      if(typeof resetCustomerQuery==='function')resetCustomerQuery();
    }
    strategyFile.addEventListener('change',invalidateReview);templateKey.addEventListener('change',invalidateReview);
    analyzeBtn.addEventListener('click',e=>{
      if((dirtyReview()||reviewData?.reviewRows?.some(r=>r.manualTiers.length||r.strategyText!==r.originalText))&&!confirm('重新解析上传表会清除本次人工修改，是否继续？')){e.stopImmediatePropagation();return;}invalidateReview();
    },true);
    window.addEventListener('beforeunload',e=>{if(dirtyReview()){e.preventDefault();e.returnValue='';}});
'''
