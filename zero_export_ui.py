ZERO_EXPORT_JS = r'''
    async function confirmZeroExport(info) {
      const dialog = document.createElement('dialog');
      dialog.id = 'zeroExportDialog';
      dialog.setAttribute('aria-labelledby', 'zeroExportTitle');
      dialog.innerHTML = `<div class="change-header"><h2 id="zeroExportTitle">确认品规导出清单</h2></div>
        <div class="export-list-body">
          <p class="muted">保留导出使用当前解析和已应用的人工修改数量；不导出将移除该品规全部维度行。</p>
          <div class="export-list-filters"><div><label for="exportProductSearch">搜索品规、编码或策略</label><input id="exportProductSearch" type="search" placeholder="搜索全部品规…" autocomplete="off"></div>
          <div><label for="exportProductFilter">筛选清单</label><select id="exportProductFilter"><option value="all">全部品规</option><option value="suggested">建议不导出</option><option value="review">需人工复核</option><option value="omitted">已选不导出</option></select></div></div>
          <div class="actions export-batch"><label><input type="checkbox" id="zeroSelectAll"> 全选当前筛选结果</label>
          <button type="button" id="zeroBatchOmit">批量不导出</button><button type="button" id="zeroBatchKeep" class="secondary">批量保留导出</button><span id="exportSelectedCount" class="muted"></span></div>
          <div class="table-wrap export-list-table"><table><thead><tr><th>选择</th><th>策略行 / 编码</th><th>品规及建议</th><th>文字策略及校验提示</th><th>导出处理</th></tr></thead><tbody></tbody></table><p id="exportListEmpty" hidden>没有匹配的品规</p></div>
        </div>
        <div class="change-footer"><p id="zeroExportSummary" aria-live="polite"></p><div class="actions"><button type="button" id="zeroCancel" class="secondary">返回复核</button><button type="button" id="zeroConfirm">确认并导出</button></div></div>`;
      const items = (info.exportProducts || []).slice().sort((a,b) => {
        const rank = p => p.defaultChoice === 'omit' ? 0 : p.category === 'review' ? 1 : 2;
        return rank(a)-rank(b) || a.sourceRow-b.sourceRow;
      });
      if (info.confirmOnly) {
        dialog.querySelector('#zeroExportTitle').textContent = '确认参与测算的品规';
        dialog.querySelector('#zeroConfirm').textContent = '确认品规范围';
      }
      if (info.reviewOnly) {
        dialog.querySelector('#zeroExportTitle').textContent = '复核品规导出清单';
        dialog.querySelector('#zeroConfirm').textContent = '确认并应用清单';
      }
      const choices = new Map(items.map(p=>[String(p.sourceRow),info.confirmedChoices?.[String(p.sourceRow)] || p.defaultChoice]));
      const checked = new Set();
      const find = selector => dialog.querySelector(selector);
      const visibleItems = () => {
        const query=find('#exportProductSearch').value.trim().toLowerCase(), filter=find('#exportProductFilter').value;
        return items.filter(p => (filter==='all'||filter==='suggested'&&p.defaultChoice==='omit'||filter==='review'&&p.category==='review'||filter==='omitted'&&choices.get(String(p.sourceRow))==='omit') &&
          [p.productCode,p.productName,p.strategyText,p.originalText,p.reason,...(p.warnings||[])].join(' ').toLowerCase().includes(query));
      };
      function refreshSummary() {
        const omitted=[...choices.values()].filter(v=>v==='omit').length, count=items.length-omitted;
        find('#zeroExportSummary').textContent=`不导出 ${omitted} 个，保留导出 ${count} 个；共 ${items.length} 个品规。`+(count===0?' 请至少保留一个品规，或返回复核。':'');
        find('#zeroConfirm').disabled=count===0;
        const visible=visibleItems(), selected=visible.filter(p=>checked.has(String(p.sourceRow))).length;
        const all=find('#zeroSelectAll');all.checked=visible.length>0&&selected===visible.length;all.indeterminate=selected>0&&selected<visible.length;all.disabled=!visible.length;
        find('#exportSelectedCount').textContent=`当前显示 ${visible.length} 个，勾选 ${selected} 个`;
        find('#zeroBatchOmit').disabled=find('#zeroBatchKeep').disabled=selected===0;
      }
      function render() {
        const visible=visibleItems(), body=find('tbody');body.replaceChildren();
        for(const item of visible) {
          const id=String(item.sourceRow), tr=document.createElement('tr');tr.dataset.exportRow=id;
          const td=document.createElement('td'), check=document.createElement('input');check.type='checkbox';check.checked=checked.has(id);check.setAttribute('aria-label','批量选择 '+item.productName);
          check.onchange=()=>{if(check.checked)checked.add(id);else checked.delete(id);refreshSummary();};td.append(check);tr.append(td);
          const code=document.createElement('td');code.textContent=`第 ${item.sourceRow} 行\n${item.productCode}`;tr.append(code);
          const name=document.createElement('td'), title=document.createElement('strong'), reason=document.createElement('div');title.textContent=item.productName;reason.textContent=item.reason;reason.className='export-reason';name.append(title,reason);tr.append(name);
          const policy=document.createElement('td'), text=document.createElement('div');text.textContent=item.originalText===item.strategyText?item.strategyText:'原文：'+item.originalText+'\n当前：'+item.strategyText;policy.append(text);
          if(item.warnings?.length){const warnings=document.createElement('div');warnings.className='export-policy-warnings';warnings.textContent=item.warnings.join('\n');policy.append(warnings);}
          tr.append(policy);
          const action=document.createElement('td'), select=document.createElement('select');select.setAttribute('aria-label',item.productName+' 导出处理');
          select.innerHTML='<option value="omit">不导出</option><option value="keep">保留导出</option>';select.value=choices.get(id);
          select.onchange=()=>{choices.set(id,select.value);if(find('#exportProductFilter').value==='omitted')render();else refreshSummary();};action.append(select);tr.append(action);body.append(tr);
        }
        find('#exportListEmpty').hidden=visible.length>0;refreshSummary();
      }
      find('#exportProductSearch').oninput=render;find('#exportProductFilter').onchange=render;
      find('#zeroSelectAll').onchange=e=>{for(const p of visibleItems()){const id=String(p.sourceRow);if(e.target.checked)checked.add(id);else checked.delete(id);}render();};
      for(const [id,value] of [['zeroBatchOmit','omit'],['zeroBatchKeep','keep']])find('#'+id).onclick=()=>{for(const p of visibleItems())if(checked.has(String(p.sourceRow)))choices.set(String(p.sourceRow),value);render();};
      document.body.append(dialog);render();
      return new Promise(resolve=>{
        const finish=value=>{dialog.close();dialog.remove();resolve(value);};
        find('#zeroCancel').onclick=()=>finish(null);
        dialog.oncancel=e=>{e.preventDefault();finish(null);};
        find('#zeroConfirm').onclick=()=>finish(Object.fromEntries(choices));
        dialog.showModal();find('#zeroCancel').focus();
      });
    }
'''
