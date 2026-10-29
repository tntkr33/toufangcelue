APP_CSS = r'''
    :root { color-scheme:light; --ink:#202724; --muted:#747c77; --line:#dfe5e1; --bg:#f4f6f5; --brand:#117653; --bad:#b42318; --warn:#936014; }
    * { box-sizing:border-box; }
    [hidden] { display:none !important; }
    body { margin:0; font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif; color:var(--ink); background:var(--bg); letter-spacing:0; }
    button,input,select,textarea { font:inherit; }
    button,input,select,textarea,summary { outline-offset:3px; }
    :focus-visible { outline:2px solid var(--brand); }
    button { border:1px solid transparent; border-radius:6px; padding:9px 16px; background:var(--brand); color:white; cursor:pointer; font-weight:600; min-height:40px; }
    button:hover:not(:disabled) { filter:brightness(.96); }
    button:disabled { opacity:.45; cursor:not-allowed; }
    button.secondary { background:white; border-color:var(--line); color:var(--ink); }
    button.quiet { background:transparent; color:var(--muted); border-color:transparent; }
    .icon-button { min-width:40px; padding:8px; display:inline-flex; align-items:center; justify-content:center; }
    .icon-button svg { width:19px; height:19px; fill:none; stroke:currentColor; stroke-width:1.7; }
    header { display:flex; align-items:center; justify-content:space-between; gap:16px; padding:18px 32px; background:white; border-bottom:1px solid var(--line); }
    .brand { display:flex; align-items:center; gap:12px; min-width:0; flex-wrap:wrap; }
    .brand-mark { display:grid; grid-template-columns:repeat(2,5px); gap:3px; background:var(--brand); padding:10px; border-radius:7px; }
    .brand-mark i { width:5px; height:5px; background:white; border-radius:1px; }
    h1 { font-size:20px; line-height:1.4; margin:0; font-weight:650; }
    .version { font-size:12px; font-weight:600; padding:3px 9px; color:var(--brand); background:#eaf4ee; border-radius:4px; white-space:nowrap; }
    .local-badge { color:var(--muted); font-size:12px; white-space:nowrap; }
    .local-badge::before { content:""; display:inline-block; width:6px; height:6px; background:#30966a; border-radius:50%; margin-right:7px; }
    #exitNotice { margin:24px 32px; }
    #exitNotice p { margin:12px 0 0; color:var(--muted); }
    @media(max-width:640px) { header { flex-wrap:wrap; } #exitTool { min-height:34px; padding:6px 12px; } }
    main { width:100%; margin:0; padding:24px 32px 48px; }
    .panel { min-width:0; background:white; padding:24px; margin-bottom:16px; border-bottom:1px solid var(--line); }
    .section-heading { display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap; margin-bottom:16px; }
    h2,h3,p { margin-top:0; }
    h2 { font-size:18px; margin-bottom:0; font-weight:650; }
    h3 { font-size:15px; margin:0; font-weight:600; }
    .eyebrow { color:var(--muted); font-size:12px; margin-bottom:4px; }
    .muted { color:var(--muted); font-size:12px; }
    .setup-grid { display:grid; grid-template-columns:minmax(0,1.15fr) minmax(0,1fr) minmax(0,1fr); gap:24px; }
    .grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:24px; }
    .grid > div,.setup-grid > div { min-width:0; }
    label { display:block; font-size:13px; font-weight:600; margin-bottom:8px; }
    input[type=file],input[type=text],input[type=search],select,textarea { width:100%; min-width:0; border:1px solid var(--line); border-radius:6px; background:white; color:var(--ink); padding:10px 12px; }
    input[type=file] { padding:6px; height:44px; font-size:13px; }
    input::file-selector-button { background:#eff3f0; color:var(--ink); border:0; border-radius:4px; padding:6px 12px; margin-right:12px; cursor:pointer; }
    input::placeholder { color:#959c97; }
    .actions { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
    .setup-actions { margin-top:20px; padding-top:16px; border-top:1px solid #edf0ee; }
    .setup-actions .status { flex:1; }
    .hint { color:var(--muted); font-size:12px; line-height:1.6; margin-top:7px; overflow-wrap:anywhere; }
    .status { font-size:13px; color:var(--muted); overflow-wrap:anywhere; min-width:0; }
    .messages { display:grid; gap:6px; }
    .msg { padding:10px 12px; border-radius:4px; font-size:13px; line-height:1.6; overflow-wrap:anywhere; }
    .warning { background:#fff9ee; color:var(--warn); border-left:3px solid #d9ae57; }
    .error { background:#fff1ef; color:var(--bad); border-left:3px solid #df7469; }
    .ok { background:#edf7f0; color:var(--brand); border-left:3px solid #51a879; }
    summary { cursor:pointer; font-size:13px; font-weight:600; }
    details[open] > summary { margin-bottom:14px; }
    #messagePanel { padding:14px 24px; }
    #messages { max-height:none; overflow:visible; }
    #summaryPanel { padding:16px 24px; }
    .summary { display:grid; grid-template-columns:repeat(6,minmax(0,1fr)); gap:16px; }
    .metric { border-right:1px solid var(--line); padding-right:16px; }
    .metric:last-child { border-right:0; }
    .metric b { font-size:24px; font-weight:650; display:block; font-variant-numeric:tabular-nums; }
    .metric span { font-size:12px; color:var(--muted); }
    .extra-metrics { margin-top:12px; color:var(--muted); }
    .extra-metrics > div { display:flex; gap:20px; flex-wrap:wrap; font-size:12px; }
    .search-wrap { position:relative; width:100%; margin:18px 0; }
    #reviewSearch { height:54px; font-size:16px; padding:12px 48px 12px 18px; background:#f8faf9; border-color:#cfdad3; }
    .search-wrap label { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
    .search-wrap > svg { position:absolute; right:18px; top:17px; width:20px; height:20px; stroke:var(--muted); fill:none; stroke-width:1.8; pointer-events:none; }
    #reviewResults { position:absolute; z-index:12; top:calc(100% + 6px); left:0; right:0; max-height:300px; overflow:auto; background:white; border:1px solid var(--line); border-radius:6px; box-shadow:0 12px 28px #1b32251a; }
    #reviewResults button { display:flex; align-items:center; gap:14px; width:100%; text-align:left; white-space:normal; background:white; color:var(--ink); border-bottom:1px solid #edf0ee; border-radius:0; padding:12px 16px; font-weight:500; }
    #reviewResults button:hover,#reviewResults button:focus { background:#edf6f0; }
    .product-code { font-variant-numeric:tabular-nums; color:var(--brand); background:#edf5ef; border-radius:4px; padding:3px 8px; font-size:12px; }
    .result-source { margin-left:auto; font-size:12px; color:var(--muted); white-space:nowrap; }
    #reviewState { color:var(--muted); font-size:12px; margin:0; }
    #reviewState[data-dirty=true] { color:var(--warn); }
    #reviewEmpty { padding:40px 16px; text-align:center; color:var(--muted); border-top:1px solid var(--line); }
    #selectedProduct { display:flex; align-items:center; flex-wrap:wrap; gap:10px; font-size:19px; margin:0; }
    .product-heading { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; margin:22px 0 16px; }
    .strategy-section { border-top:1px solid var(--line); padding:16px 0; }
    .strategy-heading { display:flex; justify-content:space-between; align-items:center; gap:12px; margin-bottom:8px; }
    #reviewOriginal,#reviewApplied { white-space:pre-wrap; line-height:1.8; max-height:128px; overflow:auto; font-size:14px; }
    #appliedStrategy { border-left:2px solid #bdd9c7; padding-left:20px; }
    .strategy-preview { display:grid; grid-template-columns:minmax(0,1fr); gap:24px; }
    .strategy-preview.has-edit { grid-template-columns:repeat(2,minmax(0,1fr)); }
    #strategyEditor { background:#f7f9f7; padding:18px; margin-top:14px; border-radius:6px; }
    #reviewText { display:block; min-height:120px; max-height:400px; resize:vertical; line-height:1.7; margin-bottom:12px; }
    #reviewWarnings { margin-bottom:12px; max-height:180px; overflow:auto; }
    .quantity-workspace { min-width:0; padding-top:18px; border-top:1px solid var(--line); }
    .quantity-heading { display:flex; justify-content:space-between; align-items:center; gap:12px; margin-bottom:12px; }
    #quantityProduct { color:var(--muted); font-size:12px; margin-top:3px; }
    .quantity-toolbar { display:flex; align-items:end; gap:12px; flex-wrap:wrap; margin-bottom:14px; }
    .quantity-toolbar label { margin-bottom:5px; font-size:12px; color:var(--muted); }
    .quantity-toolbar select { width:180px; }
    .quantity-toolbar .actions { margin-left:auto; }
    .county-filter { position:relative; min-width:180px; }
    #reviewCounty { width:100%; min-width:180px; text-align:left; font-weight:400; }
    #countyMenu { position:absolute; left:0; top:100%; margin-top:4px; width:340px; max-width:calc(100vw - 48px); z-index:10; background:white; border:1px solid var(--line); border-radius:6px; box-shadow:0 6px 18px #20352b20; }
    .county-actions { display:flex; gap:2px; padding:6px; border-bottom:1px solid var(--line); }
    .county-actions button { flex:1; padding:6px; min-height:34px; }
    #countyOptions { overflow:auto; max-height:min(360px,45dvh); padding:4px; }
    .county-option { display:flex; align-items:center; gap:8px; padding:2px 4px; }
    .county-option:hover { background:#f2f7f4; }
    .county-option label { flex:1; display:flex; align-items:center; gap:8px; margin:0; color:var(--ink); cursor:pointer; }
    .county-option input { width:16px; height:16px; accent-color:var(--brand); }
    .county-option button { min-height:32px; padding:4px 8px; font-size:12px; color:var(--brand); white-space:nowrap; }
    @media(max-width:640px) { #reviewCounty { min-width:0; } }
    table { width:100%; border-collapse:separate; border-spacing:0; font-size:13px; background:white; }
    th,td { border-bottom:1px solid #e9eeeb; padding:10px 12px; text-align:left; vertical-align:middle; }
    th { position:sticky; top:0; z-index:2; background:#f2f5f3; font-weight:600; color:#56635b; white-space:nowrap; }
    .table-wrap { max-height:400px; overflow:auto; border:1px solid var(--line); border-radius:6px; }
    #reviewTable { width:max-content; min-width:100%; }
    #reviewTable th:not(:first-child) { min-width:72px; text-align:center; }
    #reviewTable th:first-child,#reviewTable td:first-child { position:sticky; left:0; z-index:1; min-width:200px; max-width:260px; background:white; border-right:1px solid var(--line); }
    #reviewTable th:first-child { z-index:3; background:#f2f5f3; }
    #reviewTable tr:hover td { background:#f7faf8; }
    #reviewTable td:not(:first-child) { padding:4px; }
    .row-dimension { display:block; color:var(--muted); font-size:11px; margin-top:2px; }
    .tier-input { width:64px; padding:6px 3px; height:34px; text-align:center; border:1px solid transparent; border-radius:4px; background:transparent; font-variant-numeric:tabular-nums; appearance:textfield; }
    .tier-input::-webkit-inner-spin-button { -webkit-appearance:none; }
    .tier-input:hover,.tier-input:focus { border-color:#a9c5b4; background:white; }
    .tier-input.manual { color:#92590d; background:#fff5df; border-color:#e4c38a; }
    .tier-input.draft { color:#206b4b; background:#e7f4eb; border-color:#80bb97; }
    .table-footer { display:flex; justify-content:space-between; flex-wrap:wrap; gap:12px; color:var(--muted); font-size:12px; padding:10px 0; }
    #quantityDialog { max-width:none; max-height:none; width:calc(100vw - 24px); height:calc(100dvh - 24px); margin:auto; padding:24px; border:1px solid var(--line); border-radius:8px; background:white; color:var(--ink); }
    #quantityDialog::backdrop { background:#192c2480; }
    #quantityDialog .quantity-workspace { display:flex; flex-direction:column; height:100%; padding:0; border:0; }
    #quantityDialog .table-wrap { flex:1; min-height:0; max-height:none; }
    .chart-section { margin-top:20px; border-top:1px solid var(--line); padding-top:20px; }
    .chart-toolbar { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px; margin-top:14px; }
    .chart-dimensions { display:flex; align-items:center; flex-wrap:wrap; gap:10px; }
    .chart-dimensions select { width:220px; }
    .chart-dimensions label { margin:0; font-size:12px; color:var(--muted); }
    .segments { display:flex; flex-wrap:wrap; gap:3px; background:#eff3f0; padding:4px; border-radius:6px; }
    .segments button { min-height:32px; padding:5px 16px; background:transparent; color:var(--muted); border-radius:4px; font-size:13px; }
    .segments button[aria-pressed=true] { background:white; color:var(--brand); box-shadow:0 1px 4px #213c2412; }
    #reviewChart { display:flex; align-items:flex-end; gap:8px; height:250px; overflow-x:auto; padding:26px 4px 30px; margin-top:20px; border-bottom:1px solid var(--line); background:repeating-linear-gradient(to top,transparent 0,transparent 47px,#edf1ee 48px,transparent 49px); }
    .chart-bar { position:relative; flex:1; min-width:27px; min-height:2px; border:0; border-radius:3px 3px 0 0; padding:0; background:#278760; }
    .chart-bar.manual { background:#b88b39; }
    .chart-bar.draft { background:#69aa84; }
    .chart-bar span { position:absolute; left:0; right:0; color:var(--ink); font-size:11px; font-weight:500; }
    .chart-value { bottom:calc(100% + 5px); }
    .chart-tier { top:calc(100% + 8px); color:var(--muted) !important; }
    .legend { display:flex; gap:18px; font-size:12px; color:var(--muted); }
    .legend span::before { content:""; display:inline-block; width:8px; height:8px; border-radius:2px; background:#278760; margin-right:6px; }
    .legend span:last-child::before { background:#b88b39; }
    #changeDialog { width:min(1000px,calc(100vw - 32px)); max-width:none; max-height:calc(100dvh - 40px); padding:0; border:1px solid var(--line); border-radius:8px; color:var(--ink); }
    #changeDialog[open] { display:flex; flex-direction:column; }
    #changeDialog::backdrop { background:#192c2480; }
    #zeroExportDialog { width:min(1280px,calc(100vw - 32px)); max-width:none; max-height:calc(100dvh - 32px); padding:0; border:1px solid var(--line); border-radius:8px; color:var(--ink); }
    #zeroExportDialog[open] { display:flex; flex-direction:column; }
    #zeroExportDialog::backdrop { background:#192c2480; }
    .export-list-body { min-height:0; overflow:auto; padding:0 24px 16px; }
    .export-list-filters { display:grid; grid-template-columns:minmax(0,1fr) 200px; gap:16px; }
    .export-batch { margin:16px 0; }
    .export-batch label { display:flex; align-items:center; gap:8px; margin:0; }
    #zeroExportDialog input[type=checkbox] { width:16px; height:16px; accent-color:var(--brand); }
    .export-list-table { max-height:44dvh; }
    .export-list-table table { min-width:820px; table-layout:fixed; }
    .export-list-table th:nth-child(1) { width:55px; }
    .export-list-table th:nth-child(2) { width:110px; }
    .export-list-table th:nth-child(3) { width:210px; }
    .export-list-table th:last-child { width:150px; }
    .export-list-table td { white-space:pre-wrap; overflow-wrap:anywhere; vertical-align:top; }
    .export-reason { font-size:12px; color:var(--muted); margin-top:6px; }
    .export-policy-warnings { margin-top:8px; padding-left:8px; border-left:2px solid #d1a44b; color:var(--warn); font-size:12px; }
    #exportListEmpty { padding:16px; text-align:center; color:var(--muted); }
    #zeroExportSummary { margin:0; }
    @media(max-width:640px) { .export-list-filters { grid-template-columns:minmax(0,1fr); gap:8px; } .export-list-body { padding:0 16px 12px; } .export-batch { gap:8px; } .export-batch label { width:100%; } }
    #periodDialog { width:min(680px,calc(100vw - 32px)); max-height:calc(100dvh - 32px); padding:0; border:1px solid var(--line); border-radius:8px; color:var(--ink); }
    #periodDialog::backdrop { background:#192c2480; }
    #periodInputName { overflow-wrap:anywhere; }
    #periodProposed { font-size:18px; font-weight:650; color:var(--brand); }
    .change-header { padding:20px 24px; border-bottom:1px solid var(--line); display:flex; align-items:center; justify-content:space-between; gap:12px; }
    .change-scroll { padding:20px 24px; overflow:auto; min-height:0; }
    #changeProduct { font-weight:600; font-size:17px; margin-bottom:6px; }
    #changeSummary { font-size:13px; color:var(--muted); }
    #changeText { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:16px; margin-bottom:16px; }
    #changeText pre { white-space:pre-wrap; overflow-wrap:anywhere; font-family:inherit; font-size:13px; line-height:1.7; margin:0; max-height:200px; }
    #changeWarnings { margin-bottom:12px; }
    #changeDialog .table-wrap { max-height:none; }
    #changeTable { min-width:600px; }
    .new-value { color:var(--brand); font-weight:700; background:#eff8f1; }
    .change-footer { padding:16px 24px; border-top:1px solid var(--line); display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }
    @media(max-width:640px) { #changeText { grid-template-columns:minmax(0,1fr); } .change-header,.change-scroll,.change-footer { padding:16px; } .change-footer .actions { width:100%; } .change-footer button { flex:1; } }
    #previewPanel,#rulesPanel { padding:16px 24px; }
    pre { margin:12px 0 0; max-height:360px; overflow:auto; border-radius:4px; padding:16px; font:12px/1.6 Consolas,monospace; background:#f4f7f5; color:#385245; }
    @media(min-width:1800px) { .table-wrap { max-height:480px; } #reviewTable th:not(:first-child) { min-width:64px; } .tier-input { width:56px; } }
    @media(max-width:1000px) { main { padding:20px; } header { padding:16px 20px; } .setup-grid { grid-template-columns:repeat(2,minmax(0,1fr)); } .setup-grid > div:last-child { grid-column:1/-1; } .summary { grid-template-columns:repeat(3,minmax(0,1fr)); } }
    @media(max-width:640px) { main { padding:12px; } header { padding:14px 16px; } h1 { font-size:16px; } .brand { gap:8px; } .brand-mark,.local-badge { display:none; } .panel { padding:16px; } .setup-grid,.grid,.strategy-preview.has-edit { grid-template-columns:minmax(0,1fr); gap:16px; } #reviewSearch { font-size:14px; height:50px; } .setup-actions .status { flex-basis:100%; } .quantity-toolbar > div { flex:1; min-width:0; } .quantity-toolbar select { width:100%; } .quantity-toolbar .actions { flex-basis:100%; margin-left:0; } .quantity-toolbar .actions button { flex:1; padding:9px; } .metric b { font-size:21px; } #summaryPanel { padding:16px; } .chart-dimensions { width:100%; } .chart-dimensions select { flex:1; width:150px; } .segments { width:100%; } .segments button { flex:1; padding:5px 9px; } #quantityDialog { width:100vw; height:100dvh; padding:12px; border:0; border-radius:0; } #reviewTable th:first-child,#reviewTable td:first-child { min-width:145px; max-width:145px; } #selectedProduct { font-size:17px; } #appliedStrategy { padding:0; border:0; } }
'''
