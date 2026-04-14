<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no"/>
<title>AML Checker</title>

<!-- PWA META TAGS -->
<meta name="application-name" content="AML Checker"/>
<meta name="apple-mobile-web-app-capable" content="yes"/>
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent"/>
<meta name="apple-mobile-web-app-title" content="AML Checker"/>
<meta name="mobile-web-app-capable" content="yes"/>
<meta name="theme-color" content="#080c10"/>
<meta name="description" content="AML Transaction Screening — PMLA 2002, FATF Standards, RBI Guidelines"/>

<!-- PWA MANIFEST (inline via data URI) -->
<link rel="manifest" href="data:application/json;charset=utf-8,%7B%22name%22%3A%22AML+Transaction+Checker%22%2C%22short_name%22%3A%22AML+Checker%22%2C%22description%22%3A%22Anti-Money+Laundering+Transaction+Screening%22%2C%22start_url%22%3A%22.%2F%22%2C%22display%22%3A%22standalone%22%2C%22background_color%22%3A%22%23080c10%22%2C%22theme_color%22%3A%22%23080c10%22%2C%22orientation%22%3A%22portrait%22%2C%22icons%22%3A%5B%7B%22src%22%3A%22data%3Aimage%2Fpng%3Bbase64%2CiVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAACXBIWXMAAAsTAAALEwEAmpwYAAAF8WlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNy4yLWMwMDAgNzkuMWI2NWE3OWI0LCAyMDIyLzA2LzEzLTIyOjAxOjAxICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdFJlZj0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlUmVmIyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgMjMuNSAoV2luZG93cykiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6RTY5NUU3RjkwOUYyMTFFREI4RjhBNTc2MUM3NTlBNkIiIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6RTY5NUU3RkEwOUYyMTFFREI4RjhBNTc2MUM3NTlBNkIiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDpFNjk1RTdGNzA5RjIxMUVEQjhGOEE1NzYxQzc1OUE2QiIgc3RSZWY6ZG9jdW1lbnRJRD0ieG1wLmRpZDpFNjk1RTdGODA5RjIxMUVEQjhGOEE1NzYxQzc1OUE2QiIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gPD94cGFja2V0IGVuZD0iciI%2FPsVPGp0AAAOUSURBVHhe7d07axRRGIDhM7vZbBKjBi8QFSxUsLGwEiystLGwsrOwFCystbSyULCwULCwUMHCQsHCQsHCi4ggKoqKt3jBe7xfmHEd92QyO7OzM%2Bfs%2Bz7wFbKzk%2Bx5v5wzs7Mhz3cAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA%22%2C%22sizes%22%3A%22192x192%22%2C%22type%22%3A%22image%2Fpng%22%7D%5D%7D"/>

<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Syne:wght@700;800&display=swap" rel="stylesheet"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>

<style>
*{box-sizing:border-box;margin:0;padding:0}
::-webkit-scrollbar{width:4px}
::-webkit-scrollbar-track{background:#0d1117}
::-webkit-scrollbar-thumb{background:#1f3a5f;border-radius:3px}
body{min-height:100vh;background:#080c10;font-family:'DM Mono','Courier New',monospace;color:#c9d1d9;overscroll-behavior:none}

/* INSTALL BANNER */
#install-banner{
  display:none;position:fixed;bottom:0;left:0;right:0;z-index:999;
  background:linear-gradient(135deg,#1f6feb,#2f81f7);
  padding:14px 20px;display:none;align-items:center;justify-content:space-between;gap:12px;
  box-shadow:0 -4px 20px rgba(47,129,247,.3);
}
#install-banner.show{display:flex}
.install-text{font-size:13px;color:#fff;line-height:1.4}
.install-text strong{display:block;font-family:'Syne',sans-serif;font-size:15px}
.install-btn{background:#fff;color:#1f6feb;border:none;border-radius:6px;padding:9px 18px;font-family:'DM Mono',monospace;font-size:12px;font-weight:500;cursor:pointer;white-space:nowrap}
.install-close{background:none;border:none;color:rgba(255,255,255,.7);font-size:20px;cursor:pointer;padding:0 4px}

/* OFFLINE BADGE */
#offline-badge{
  display:none;position:fixed;top:10px;right:10px;z-index:999;
  background:#ff3b3b;color:#fff;font-size:10px;padding:4px 10px;
  border-radius:20px;letter-spacing:.08em;font-weight:500;
}

/* STATUS BAR SPACER (for standalone mode) */
@media all and (display-mode: standalone){
  body{padding-top:env(safe-area-inset-top)}
  #install-banner{padding-bottom:calc(14px + env(safe-area-inset-bottom))}
}

/* HEADER */
.header{background:linear-gradient(180deg,#0d1117,#080c10);border-bottom:1px solid #1f3a5f;padding:16px 16px 0;position:sticky;top:0;z-index:100}
.header-inner{max-width:900px;margin:0 auto}
.live-row{display:flex;align-items:center;gap:8px;margin-bottom:3px}
.live-dot{width:7px;height:7px;border-radius:50%;background:#2f81f7;box-shadow:0 0 8px #2f81f7;animation:pulse 2s infinite;flex-shrink:0}
.live-label{font-size:10px;color:#7d8590;letter-spacing:.12em}
h1{font-family:'Syne',sans-serif;font-size:22px;font-weight:800;color:#e6edf3;letter-spacing:-.02em;margin-bottom:3px}
.subtitle{font-size:11px;color:#7d8590;margin-bottom:14px}
.tabs{display:flex;gap:0;overflow-x:auto;-webkit-overflow-scrolling:touch}
.tabs::-webkit-scrollbar{display:none}
.tab-btn{cursor:pointer;border:none;outline:none;font-family:'DM Mono',monospace;font-size:11px;text-transform:uppercase;letter-spacing:.06em;padding:9px 16px;background:transparent;color:#7d8590;border-bottom:2px solid transparent;transition:all .2s;white-space:nowrap;flex-shrink:0}
.tab-btn.active{background:#161b22;color:#e6edf3;border-bottom:2px solid #2f81f7;border-radius:6px 6px 0 0}

/* MAIN */
.main{max-width:900px;margin:0 auto;padding:16px 16px 120px}

/* PANEL */
.panel{background:#0d1117;border:1px solid #1f3a5f;border-radius:10px;padding:18px;margin-bottom:16px}
.panel-title{font-size:10px;color:#7d8590;letter-spacing:.1em;margin-bottom:16px;text-transform:uppercase}

/* FORM */
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:12px}
.field{margin-bottom:12px}
label{font-size:10px;letter-spacing:.07em;color:#7d8590;display:block;margin-bottom:5px;text-transform:uppercase}
.req{color:#ff3b3b;margin-left:2px}
input,select,textarea{
  background:#080c10;border:1px solid #1f3a5f;color:#c9d1d9;
  border-radius:6px;padding:11px 12px;width:100%;
  font-family:'DM Mono',monospace;font-size:14px;outline:none;
  transition:border-color .2s;-webkit-appearance:none;
  -webkit-tap-highlight-color:transparent;
}
input:focus,select:focus,textarea:focus{border-color:#2f81f7;box-shadow:0 0 0 3px rgba(47,129,247,.1)}
select{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' fill='%237d8590' viewBox='0 0 16 16'%3E%3Cpath d='M7.247 11.14 2.451 5.658C1.885 5.013 2.345 4 3.204 4h9.592a1 1 0 0 1 .753 1.659l-4.796 5.48a1 1 0 0 1-1.506 0z'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:32px}
select option{background:#0d1117}
textarea{resize:vertical;min-height:60px}
input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none}

/* BUTTONS */
.btn{cursor:pointer;border:none;border-radius:8px;font-family:'DM Mono',monospace;font-size:13px;font-weight:500;padding:13px 20px;transition:all .2s;letter-spacing:.04em;-webkit-tap-highlight-color:transparent}
.btn:active{transform:scale(.97);filter:brightness(.9)}
.btn-primary{background:linear-gradient(135deg,#1f6feb,#2f81f7);color:#fff}
.btn-reset{background:#161b22;color:#7d8590;border:1px solid #1f3a5f}
.btn-export{background:#161b22;color:#34d399;border:1px solid #10b981;width:100%;margin-top:10px}
.btn-row{display:flex;gap:10px;margin-top:6px}

/* UPLOAD */
.upload-zone{border:2px dashed #1f3a5f;border-radius:10px;padding:36px 20px;text-align:center;position:relative;transition:all .25s}
.upload-zone.drag{border-color:#2f81f7;background:rgba(47,129,247,.05)}
.upload-zone input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.upload-icon{font-size:32px;margin-bottom:10px}
.upload-title{font-family:'Syne',sans-serif;font-size:15px;font-weight:700;color:#e6edf3;margin-bottom:5px}
.upload-sub{font-size:11px;color:#7d8590}
.upload-badges{display:flex;justify-content:center;gap:8px;margin-top:10px}
.badge{font-size:10px;padding:3px 10px;border-radius:4px;letter-spacing:.07em;font-weight:500}
.badge-excel{background:rgba(16,185,129,.15);color:#34d399;border:1px solid #10b98155}
.badge-pdf{background:rgba(239,68,68,.12);color:#f87171;border:1px solid #ef444455}

/* PROGRESS */
.progress-wrap{margin-top:14px;display:none}
.progress-bar{height:4px;background:#1f3a5f;border-radius:2px;overflow:hidden}
.progress-fill{height:100%;background:linear-gradient(90deg,#1f6feb,#2f81f7);border-radius:2px;width:0;transition:width .3s}
.progress-label{font-size:11px;color:#7d8590;margin-top:6px;text-align:center}

/* RESULT */
.result-card{border-radius:10px;padding:18px;margin-bottom:14px;animation:slidein .3s ease}
.result-header{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:16px}
.result-name{font-family:'Syne',sans-serif;font-size:20px;font-weight:800;color:#e6edf3}
.result-sub{font-size:11px;color:#7d8590;margin-top:3px;line-height:1.6}
.risk-badge-lg{font-family:'Syne',sans-serif;font-size:13px;font-weight:800;padding:9px 18px;border-radius:6px;letter-spacing:.1em;white-space:nowrap}
.clean-box{background:rgba(16,185,129,.08);border:1px solid #10b981;border-radius:8px;padding:13px 14px;display:flex;align-items:center;gap:10px}
.flag-card{border-radius:0 7px 7px 0;padding:12px 13px;margin-bottom:9px}
.flag-top{display:flex;align-items:center;gap:7px;margin-bottom:4px;flex-wrap:wrap}
.flag-level-pill{font-size:10px;padding:2px 8px;border-radius:4px;letter-spacing:.07em;font-weight:500}
.flag-rule{font-size:13px;font-weight:500}
.flag-detail{font-size:12px;color:#8b949e;line-height:1.6}
.disclaimer{margin-top:12px;padding:9px 12px;background:rgba(255,255,255,.03);border-radius:6px;font-size:10px;color:#7d8590}

/* TABLE */
.table-wrap{overflow-x:auto;border-radius:8px;border:1px solid #1f3a5f;-webkit-overflow-scrolling:touch}
table{width:100%;border-collapse:collapse;font-size:12px;min-width:650px}
th{padding:10px 12px;text-align:left;font-size:10px;letter-spacing:.09em;color:#7d8590;text-transform:uppercase;border-bottom:1px solid #1f3a5f;background:#0d1117;white-space:nowrap}
td{padding:10px 12px;border-bottom:1px solid #1f3a5f11;color:#c9d1d9;vertical-align:middle}
tr:last-child td{border-bottom:none}
.risk-pill{font-size:10px;font-weight:700;font-family:'Syne',sans-serif;padding:3px 10px;border-radius:20px;letter-spacing:.06em}

/* HISTORY */
.hist-item{background:#0d1117;border-radius:8px;padding:13px 14px;display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:9px}
.hist-name{font-weight:500;color:#e6edf3;font-size:14px;margin-bottom:2px}
.hist-sub{font-size:11px;color:#7d8590}
.hist-meta{font-size:10px;color:#484f58;margin-top:2px}
.stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-top:10px;background:#0d1117;border:1px solid #1f3a5f;border-radius:8px;padding:14px}
.stat{text-align:center}
.stat-num{font-size:22px;font-weight:800;font-family:'Syne',sans-serif}
.stat-lbl{font-size:10px;color:#7d8590;letter-spacing:.07em}
.sum-bar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}
.sum-chip{padding:5px 12px;border-radius:6px;font-size:11px;font-weight:500;border:1px solid}
.alert{padding:12px 14px;border-radius:8px;font-size:12px;margin-bottom:14px;display:flex;gap:9px;align-items:flex-start}
.alert-error{background:rgba(255,59,59,.08);border:1px solid #ff3b3b55;color:#ff6b6b}
.alert-info{background:rgba(47,129,247,.08);border:1px solid #2f81f755;color:#79c0ff}
.empty{text-align:center;padding:50px 20px;color:#7d8590;font-size:13px}
.empty span{display:block;font-size:11px;opacity:.6;margin-top:6px}

@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
@keyframes slidein{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}

@media(max-width:500px){
  .grid-2,.grid-3{grid-template-columns:1fr}
  .result-header{flex-direction:column}
  .risk-badge-lg{align-self:flex-start}
  .stats-grid{grid-template-columns:repeat(2,1fr)}
  h1{font-size:19px}
}
</style>
</head>
<body>

<!-- OFFLINE BADGE -->
<div id="offline-badge">✕ OFFLINE</div>

<!-- INSTALL BANNER -->
<div id="install-banner">
  <div class="install-text">
    <strong>Install AML Checker</strong>
    Add to home screen — works offline
  </div>
  <button class="install-btn" id="install-btn">Install</button>
  <button class="install-close" onclick="dismissInstall()">✕</button>
</div>

<!-- HEADER -->
<div class="header">
  <div class="header-inner">
    <div class="live-row">
      <span class="live-dot"></span>
      <span class="live-label">AML SCREENING — PMLA 2002 · FATF · RBI</span>
    </div>
    <h1>AML Transaction Checker</h1>
    <p class="subtitle">Single check or upload Excel / PDF for batch screening</p>
    <div class="tabs">
      <button class="tab-btn active" id="tab-single"  onclick="switchTab('single')">✎ Single Check</button>
      <button class="tab-btn"        id="tab-upload"  onclick="switchTab('upload')">⬆ Upload File</button>
      <button class="tab-btn"        id="tab-history" onclick="switchTab('history')">◈ History (<span id="hist-count">0</span>)</button>
    </div>
  </div>
</div>

<!-- MAIN -->
<div class="main">

<!-- TAB 1: SINGLE CHECK -->
<div id="page-single">
  <div class="panel">
    <div class="panel-title">Fill Transaction Details (<span class="req">*</span> required)</div>
    <div class="grid-2">
      <div><label>Transaction ID</label><input id="txn_id" placeholder="TXN001" autocomplete="off"/></div>
      <div><label>Account ID</label><input id="account_id" placeholder="ACC101" autocomplete="off"/></div>
    </div>
    <div class="field">
      <label>Customer Name <span class="req">*</span></label>
      <input id="customer_name" placeholder="Full name of account holder" autocomplete="off"/>
    </div>
    <div class="grid-2">
      <div><label>Amount (INR) <span class="req">*</span></label><input id="amount" type="number" placeholder="e.g. 1200000"/></div>
      <div><label>Date</label><input id="date" type="date"/></div>
    </div>
    <div class="grid-2">
      <div>
        <label>Transaction Type <span class="req">*</span></label>
        <select id="txn_type">
          <option value="">-- Select --</option>
          <optgroup label="── Cash ──">
            <option>CASH_DEPOSIT</option><option>CASH_WITHDRAWAL</option>
          </optgroup>
          <optgroup label="── Bank Transfers ──">
            <option>NEFT</option><option>RTGS</option><option>IMPS</option>
            <option>UPI</option><option>WIRE_TRANSFER</option><option>SWIFT</option>
            <option>CHEQUE</option><option>DD</option>
          </optgroup>
          <optgroup label="── High Risk ──">
            <option>CRYPTO</option><option>HAWALA</option><option>SHELL_COMPANY</option>
            <option>OFFSHORE</option><option>BEARER_BOND</option><option>CASINO</option>
            <option>VIRTUAL_ASSET</option>
          </optgroup>
          <optgroup label="── Other ──">
            <option>INSURANCE_PREMIUM</option><option>LOAN_REPAYMENT</option>
            <option>TRADE_FINANCE</option><option>INVESTMENT</option>
          </optgroup>
        </select>
      </div>
      <div>
        <label>Channel</label>
        <select id="channel">
          <option>BRANCH</option><option>ATM</option><option>ONLINE</option>
          <option>MOBILE</option><option>AGENT</option><option>SWIFT</option>
        </select>
      </div>
    </div>
    <div class="grid-2">
      <div>
        <label>Country <span class="req">*</span></label>
        <select id="country"><option value="">-- Select Country --</option></select>
      </div>
      <div><label>Currency</label>
        <select id="currency">
          <option>INR</option><option>USD</option><option>EUR</option><option>GBP</option>
          <option>AED</option><option>SGD</option><option>JPY</option><option>CNY</option>
        </select>
      </div>
    </div>
    <div class="field">
      <label>Counterparty / Beneficiary</label>
      <input id="counterparty" placeholder="Sender or receiver name" autocomplete="off"/>
    </div>
    <div class="field">
      <label>Notes / Remarks</label>
      <textarea id="notes" placeholder="Purpose of transaction..."></textarea>
    </div>
    <div class="btn-row">
      <button class="btn btn-primary" style="flex:1" onclick="runSingleCheck()">▶ &nbsp;RUN AML CHECK</button>
      <button class="btn btn-reset" onclick="resetSingle()">✕</button>
    </div>
  </div>
  <div id="single-result"></div>
</div>

<!-- TAB 2: UPLOAD -->
<div id="page-upload" style="display:none">
  <div class="panel">
    <div class="panel-title">Required columns in your file</div>
    <div style="font-size:12px;color:#8b949e;line-height:1.9">
      Your file must have these column headers (order doesn't matter):
      <div style="margin:10px 0;padding:10px 12px;background:#080c10;border-radius:6px;border:1px solid #1f3a5f;font-size:11px;color:#79c0ff;line-height:2">
        customer_name &nbsp;·&nbsp; amount &nbsp;·&nbsp; country &nbsp;·&nbsp; txn_type<br/>
        <span style="color:#484f58">(optional: txn_id · account_id · date · channel · counterparty · notes)</span>
      </div>
      For PDF — must be a text-based PDF, not a scanned image.
    </div>
    <button class="btn btn-export" onclick="downloadTemplate()">⬇ &nbsp;Download Excel Template</button>
  </div>
  <div class="panel">
    <div class="panel-title">Upload Transaction File</div>
    <div class="upload-zone" id="drop-zone" ondragover="dragOver(event)" ondragleave="dragLeave()" ondrop="dropFile(event)">
      <input type="file" id="file-input" accept=".xlsx,.xls,.pdf" onchange="handleFile(this.files[0])"/>
      <div class="upload-icon">📂</div>
      <div class="upload-title">Tap to browse or drop file</div>
      <div class="upload-sub">Supports Excel and PDF</div>
      <div class="upload-badges">
        <span class="badge badge-excel">XLSX / XLS</span>
        <span class="badge badge-pdf">PDF</span>
      </div>
    </div>
    <div class="progress-wrap" id="progress-wrap">
      <div class="progress-bar"><div class="progress-fill" id="progress-fill"></div></div>
      <div class="progress-label" id="progress-label">Processing...</div>
    </div>
  </div>
  <div id="batch-result"></div>
</div>

<!-- TAB 3: HISTORY -->
<div id="page-history" style="display:none">
  <div id="history-list">
    <div class="empty">No transactions yet.<span>Check a transaction or upload a file to begin.</span></div>
  </div>
  <div id="history-stats"></div>
</div>

</div>

<!-- SERVICE WORKER REGISTRATION + JS -->
<script>
// ─── SERVICE WORKER (offline support) ───────────────────────
if('serviceWorker' in navigator){
  const swCode=`
    const CACHE='aml-v1';
    const ASSETS=[self.location.href];
    self.addEventListener('install',e=>{
      e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS)));
      self.skipWaiting();
    });
    self.addEventListener('activate',e=>{
      e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));
      self.clients.claim();
    });
    self.addEventListener('fetch',e=>{
      e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(res=>{
        const clone=res.clone();
        caches.open(CACHE).then(c=>c.put(e.request,clone));
        return res;
      }).catch(()=>caches.match(e.request))));
    });
  `;
  const blob=new Blob([swCode],{type:'application/javascript'});
  const url=URL.createObjectURL(blob);
  navigator.serviceWorker.register(url).catch(()=>{});
}

// ─── OFFLINE INDICATOR ───────────────────────────────────────
function updateOnline(){
  document.getElementById('offline-badge').style.display=navigator.onLine?'none':'block';
}
window.addEventListener('online',updateOnline);
window.addEventListener('offline',updateOnline);
updateOnline();

// ─── PWA INSTALL PROMPT ──────────────────────────────────────
let deferredPrompt=null;
window.addEventListener('beforeinstallprompt',e=>{
  e.preventDefault();
  deferredPrompt=e;
  if(!localStorage.getItem('installDismissed')){
    document.getElementById('install-banner').classList.add('show');
  }
});
document.getElementById('install-btn').addEventListener('click',async()=>{
  if(!deferredPrompt) return;
  deferredPrompt.prompt();
  const {outcome}=await deferredPrompt.userChoice;
  deferredPrompt=null;
  document.getElementById('install-banner').classList.remove('show');
});
function dismissInstall(){
  document.getElementById('install-banner').classList.remove('show');
  localStorage.setItem('installDismissed','1');
}

// ─── COUNTRIES ───────────────────────────────────────────────
const ALL_COUNTRIES=["Afghanistan","Albania","Algeria","Andorra","Angola","Antigua and Barbuda","Argentina","Armenia","Australia","Austria","Azerbaijan","Bahamas","Bahrain","Bangladesh","Barbados","Belarus","Belgium","Belize","Benin","Bhutan","Bolivia","Bosnia and Herzegovina","Botswana","Brazil","Brunei","Bulgaria","Burkina Faso","Burundi","Cabo Verde","Cambodia","Cameroon","Canada","Central African Republic","Chad","Chile","China","Colombia","Comoros","Congo (Brazzaville)","Congo (Kinshasa)","Costa Rica","Croatia","Cuba","Cyprus","Czech Republic","Denmark","Djibouti","Dominica","Dominican Republic","Ecuador","Egypt","El Salvador","Equatorial Guinea","Eritrea","Estonia","Eswatini","Ethiopia","Fiji","Finland","France","Gabon","Gambia","Georgia","Germany","Ghana","Greece","Grenada","Guatemala","Guinea","Guinea-Bissau","Guyana","Haiti","Honduras","Hungary","Iceland","India","Indonesia","Iran","Iraq","Ireland","Israel","Italy","Jamaica","Japan","Jordan","Kazakhstan","Kenya","Kiribati","Kuwait","Kyrgyzstan","Laos","Latvia","Lebanon","Lesotho","Liberia","Libya","Liechtenstein","Lithuania","Luxembourg","Madagascar","Malawi","Malaysia","Maldives","Mali","Malta","Marshall Islands","Mauritania","Mauritius","Mexico","Micronesia","Moldova","Monaco","Mongolia","Montenegro","Morocco","Mozambique","Myanmar","Namibia","Nauru","Nepal","Netherlands","New Zealand","Nicaragua","Niger","Nigeria","North Korea","North Macedonia","Norway","Oman","Pakistan","Palau","Palestine","Panama","Papua New Guinea","Paraguay","Peru","Philippines","Poland","Portugal","Qatar","Romania","Russia","Rwanda","Saint Kitts and Nevis","Saint Lucia","Saint Vincent and the Grenadines","Samoa","San Marino","Sao Tome and Principe","Saudi Arabia","Senegal","Serbia","Seychelles","Sierra Leone","Singapore","Slovakia","Slovenia","Solomon Islands","Somalia","South Africa","South Korea","South Sudan","Spain","Sri Lanka","Sudan","Suriname","Sweden","Switzerland","Syria","Taiwan","Tajikistan","Tanzania","Thailand","Timor-Leste","Togo","Tonga","Trinidad and Tobago","Tunisia","Turkey","Turkmenistan","Tuvalu","Uganda","Ukraine","United Arab Emirates","United Kingdom","United States","Uruguay","Uzbekistan","Vanuatu","Vatican City","Venezuela","Vietnam","Yemen","Zambia","Zimbabwe"];
(function(){const s=document.getElementById('country');ALL_COUNTRIES.forEach(c=>{const o=document.createElement('option');o.value=c;o.textContent=c;s.appendChild(o);});})();

// ─── AML ENGINE ──────────────────────────────────────────────
const CTR=1000000,WIRE=500000;
const HRC=new Set(["iran","north korea","myanmar","syria","yemen","pakistan","haiti","panama","philippines","south sudan","sudan","venezuela","russia","afghanistan","libya"]);
const HRT=new Set(["CRYPTO","HAWALA","SHELL_COMPANY","OFFSHORE","BEARER_BOND","CASINO","VIRTUAL_ASSET"]);
const RS={HIGH:{bg:"#2a0a0a",border:"#ff3b3b",badge:"#ff3b3b",text:"#ff6b6b",bt:"#fff"},MEDIUM:{bg:"#1f1500",border:"#f59e0b",badge:"#f59e0b",text:"#fbbf24",bt:"#fff"},LOW:{bg:"#001a1a",border:"#06b6d4",badge:"#06b6d4",text:"#22d3ee",bt:"#fff"},CLEAN:{bg:"#001a0a",border:"#10b981",badge:"#10b981",text:"#34d399",bt:"#001a0a"}};

function runAML(t){
  const flags=[],amt=parseFloat(t.amount)||0,type=(t.txn_type||'').toUpperCase().trim(),cntry=(t.country||'').trim().toLowerCase(),cash=["CASH_DEPOSIT","CASH_WITHDRAWAL"].includes(type);
  if(cash&&amt>=CTR) flags.push({rule:"CTR — PMLA Rule 3",level:"HIGH",detail:`Cash Rs.${fmt(amt)} ≥ Rs.10 lakh. Mandatory CTR to FIU-IND within 15 days of month end.`});
  if(!cash&&amt>=CTR) flags.push({rule:"Large Transaction — PMLA §12",level:"HIGH",detail:`Non-cash Rs.${fmt(amt)} ≥ Rs.10 lakh. STR must be filed with FIU-IND within 7 days.`});
  if(HRC.has(cntry)) flags.push({rule:`High-Risk Country — FATF [${t.country}]`,level:"HIGH",detail:`${t.country} is FATF grey/blacklisted. Enhanced Due Diligence (EDD) mandatory before processing.`});
  if(["WIRE_TRANSFER","SWIFT"].includes(type)&&amt>=WIRE) flags.push({rule:"Cross-Border Wire — FEMA/RBI",level:"HIGH",detail:`Wire Rs.${fmt(amt)} ≥ Rs.5 lakh. Cross Border Wire Transfer Report (CWTR) required.`});
  if(HRT.has(type)) flags.push({rule:`High-Risk Type [${t.txn_type}]`,level:"MEDIUM",detail:`${t.txn_type} carries elevated ML/TF risk. Enhanced customer due diligence required.`});
  if(amt>=100000&&amt%100000===0) flags.push({rule:"Round Amount Anomaly",level:"MEDIUM",detail:`Rs.${fmt(amt)} is perfectly round — common indicator of layering in AML typologies.`});
  if(cash&&amt>0&&amt<CTR*0.95) flags.push({rule:"Potential Structuring — PMLA §3",level:"MEDIUM",detail:`Cash Rs.${fmt(amt)} is near but below Rs.10L CTR threshold. Watch for repeated transactions — possible smurfing.`});
  const risk=flags.some(f=>f.level==="HIGH")?"HIGH":flags.some(f=>f.level==="MEDIUM")?"MEDIUM":flags.length>0?"LOW":"CLEAN";
  return{flags,risk};
}

// ─── SINGLE CHECK ────────────────────────────────────────────
let history=[];
function runSingleCheck(){
  const name=v('customer_name'),amount=v('amount'),cntry=v('country'),ttype=v('txn_type');
  if(!name||!amount||!cntry||!ttype){
    document.getElementById('single-result').innerHTML=`<div class="alert alert-error">⚠ Fill in Customer Name, Amount, Country and Transaction Type.</div>`;
    return;
  }
  const txn={txn_id:v('txn_id')||'—',account_id:v('account_id')||'—',customer_name:name,date:v('date')||today(),amount,currency:v('currency')||'INR',txn_type:ttype,channel:v('channel'),country:cntry,counterparty:v('counterparty')||'—',notes:v('notes')||''};
  const{flags,risk}=runAML(txn);
  const entry={...txn,flags,risk,timestamp:new Date().toLocaleTimeString()};
  history.unshift(entry);
  document.getElementById('single-result').innerHTML=renderCard(entry);
  document.getElementById('hist-count').textContent=history.length;
  renderHistory();
  window.scrollTo({top:document.getElementById('single-result').offsetTop-20,behavior:'smooth'});
}

function resetSingle(){
  ['txn_id','account_id','customer_name','amount','counterparty','notes'].forEach(id=>document.getElementById(id).value='');
  ['country','txn_type'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('channel').value='BRANCH';
  document.getElementById('currency').value='INR';
  document.getElementById('date').value='';
  document.getElementById('single-result').innerHTML='';
}

// ─── FILE UPLOAD ─────────────────────────────────────────────
function handleFile(file){
  if(!file)return;
  const ext=file.name.split('.').pop().toLowerCase();
  if(!['xlsx','xls','pdf'].includes(ext)){showBatchAlert('error','Only .xlsx, .xls and .pdf files accepted.');return;}
  showProg(true,'Reading file...');
  if(ext==='pdf')readPDF(file);else readExcel(file);
}
function readExcel(file){
  const r=new FileReader();
  r.onload=function(e){
    try{
      setProg(40,'Parsing Excel...');
      const data=new Uint8Array(e.target.result);
      const wb=XLSX.read(data,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
      setProg(80,'Running AML checks...');
      if(!rows||!rows.length){showBatchAlert('error','No data rows found in file.');showProg(false);return;}
      processBatch(rows,file.name);
    }catch(err){showBatchAlert('error','Excel parse error: '+err.message);showProg(false);}
  };
  r.readAsArrayBuffer(file);
}
async function readPDF(file){
  try{
    pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    const ab=await file.arrayBuffer();
    setProg(20,'Loading PDF...');
    const pdf=await pdfjsLib.getDocument({data:ab}).promise;
    let text='';
    for(let i=1;i<=pdf.numPages;i++){
      setProg(20+Math.floor((i/pdf.numPages)*55),'Reading page '+i+'/'+pdf.numPages+'...');
      const page=await pdf.getPage(i);
      const content=await page.getTextContent();
      text+=content.items.map(it=>it.str).join(' ')+'\n';
    }
    setProg(80,'Parsing data...');
    const rows=parsePDF(text);
    if(!rows||!rows.length){showBatchAlert('error','Could not extract table data from this PDF. Try saving as Excel instead.');showProg(false);return;}
    processBatch(rows,file.name);
  }catch(err){showBatchAlert('error','PDF error: '+err.message);showProg(false);}
}
function parsePDF(text){
  const known=["customer_name","amount","country","txn_type","txn_id","account_id","date","channel","counterparty","notes","customer","name","type"];
  const lines=text.split('\n').map(l=>l.trim()).filter(l=>l.length>3);
  let hi=-1,headers=[];
  for(let i=0;i<lines.length;i++){
    const words=lines[i].toLowerCase().split(/\s{2,}|\t|,/);
    if(words.filter(w=>known.some(h=>w.includes(h))).length>=2){hi=i;headers=words.map(w=>w.trim());break;}
  }
  if(hi<0||headers.length<2)return[];
  const rows=[];
  for(let i=hi+1;i<lines.length;i++){
    const cells=lines[i].split(/\s{2,}|\t|,/);
    if(cells.length<2)continue;
    const row={};
    headers.forEach((h,idx)=>{row[normKey(h)]=cells[idx]||'';});
    if(row.customer_name||row.amount)rows.push(row);
  }
  return rows;
}

// ─── COLUMN MAPPING ──────────────────────────────────────────
const CMAP={customer_name:["customer_name","customer","name","account holder"],amount:["amount","amt","transaction amount","txn amount","value"],country:["country","nation","location"],txn_type:["txn_type","type","transaction type","txn type","payment type"],txn_id:["txn_id","txn id","transaction id","id","ref"],account_id:["account_id","account id","account","acc"],date:["date","txn date","value date"],channel:["channel","mode"],counterparty:["counterparty","beneficiary","sender","payee"],notes:["notes","remarks","narration","description"]};
function normKey(raw){const r=(raw||'').toString().toLowerCase().trim().replace(/\s+/g,'_');for(const[k,a]of Object.entries(CMAP)){if(a.some(x=>r===x||r.includes(x.replace(/\s/g,'_'))))return k;}return r;}
function mapRow(raw){const out={};for(const k of Object.keys(raw)){out[normKey(k)]=raw[k];}return out;}

// ─── BATCH PROCESSOR ─────────────────────────────────────────
function processBatch(rows,fname){
  const results=[];
  rows.forEach((raw,i)=>{
    const row=mapRow(raw);
    if(!row.customer_name&&!row.amount)return;
    const{flags,risk}=runAML(row);
    const entry={txn_id:row.txn_id||'R'+(i+1),account_id:row.account_id||'—',customer_name:row.customer_name||'Unknown',date:row.date||today(),amount:row.amount||0,currency:row.currency||'INR',txn_type:row.txn_type||'—',channel:row.channel||'—',country:row.country||'—',counterparty:row.counterparty||'—',notes:row.notes||'',flags,risk,timestamp:new Date().toLocaleTimeString()};
    results.push(entry);
    history.unshift(entry);
  });
  setProg(100,'Done!');
  setTimeout(()=>showProg(false),500);
  document.getElementById('hist-count').textContent=history.length;
  renderHistory();
  renderBatch(results,fname);
}

function renderBatch(results,fname){
  if(!results.length){showBatchAlert('error','No valid transaction rows found.');return;}
  const c={HIGH:0,MEDIUM:0,LOW:0,CLEAN:0};
  results.forEach(r=>c[r.risk]++);
  let sum=`<div class="sum-bar">`;
  ['HIGH','MEDIUM','LOW','CLEAN'].forEach(r=>{const s=RS[r];sum+=`<span class="sum-chip" style="background:${s.bg};border-color:${s.border};color:${s.text}">${r}: ${c[r]}</span>`;});
  sum+=`</div>`;
  let rows='';
  results.forEach(r=>{
    const s=RS[r.risk];
    const ft=r.flags.length===0?'—':r.flags.map(f=>`<span style="color:${RS[f.level].text}">[${f.level}] ${f.rule}</span>`).join('<br/>');
    rows+=`<tr><td style="color:#7d8590;font-size:11px">${r.txn_id}</td><td><strong style="color:#e6edf3">${r.customer_name}</strong></td><td>Rs.${fmt(r.amount)}</td><td style="font-size:11px">${r.txn_type}</td><td style="font-size:11px">${r.country}</td><td><span class="risk-pill" style="background:${s.badge};color:${s.bt}">${r.risk}</span></td><td style="font-size:11px;line-height:1.8">${ft}</td></tr>`;
  });
  document.getElementById('batch-result').innerHTML=`
    <div class="panel">
      <div class="panel-title">Batch Results — ${results.length} transactions from "${fname}"</div>
      ${sum}
      <div class="table-wrap"><table>
        <thead><tr><th>ID</th><th>Customer</th><th>Amount</th><th>Type</th><th>Country</th><th>Risk</th><th>Flags</th></tr></thead>
        <tbody>${rows}</tbody>
      </table></div>
      <button class="btn btn-export" onclick="exportBatch()">⬇ &nbsp;Export Batch Report as CSV</button>
    </div>`;
  window._batch=results;
}

// ─── RENDER CARD ─────────────────────────────────────────────
function renderCard(e){
  const s=RS[e.risk];
  let fh='';
  if(!e.flags.length){
    fh=`<div class="clean-box"><span style="font-size:20px">✅</span><div><div style="color:#34d399;font-weight:500;font-size:13px">No AML Triggers Detected</div><div style="color:#7d8590;font-size:12px;margin-top:2px">Standard monitoring applies.</div></div></div>`;
  }else{
    e.flags.forEach(f=>{const fs=RS[f.level];fh+=`<div class="flag-card" style="background:rgba(0,0,0,.3);border-left:3px solid ${fs.border};border-top:1px solid ${fs.border}22;border-right:1px solid ${fs.border}22;border-bottom:1px solid ${fs.border}22"><div class="flag-top"><span class="flag-level-pill" style="background:${fs.badge}22;color:${fs.text}">${f.level}</span><span class="flag-rule" style="color:${fs.text}">${f.rule}</span></div><div class="flag-detail">${f.detail}</div></div>`;});
  }
  return `<div class="result-card" style="background:${s.bg};border:1px solid ${s.border}"><div class="result-header"><div><div style="font-size:10px;color:#7d8590;letter-spacing:.1em;margin-bottom:4px">SCREENING RESULT</div><div class="result-name">${e.customer_name}</div><div class="result-sub">${e.txn_type} · Rs.${fmt(e.amount)} · ${e.country} · ${e.date}</div></div><div class="risk-badge-lg" style="background:${s.badge};color:${s.bt}">${e.risk}</div></div>${fh}<div class="disclaimer">⚠ Compliance screening tool only. Final STR/CTR decisions must be made by a qualified Principal Officer — PMLA 2002 §12.</div><button class="btn btn-export" onclick="exportSingle()">⬇ Export as CSV</button></div>`;
}

// ─── HISTORY ─────────────────────────────────────────────────
function renderHistory(){
  const hl=document.getElementById('history-list'),sa=document.getElementById('history-stats');
  if(!history.length){hl.innerHTML=`<div class="empty">No transactions yet.<span>Check or upload to begin.</span></div>`;sa.innerHTML='';return;}
  hl.innerHTML=history.map(h=>{const s=RS[h.risk];return`<div class="hist-item" style="border:1px solid ${s.border}55"><div style="flex:1;min-width:0"><div class="hist-name">${h.customer_name}</div><div class="hist-sub">${h.txn_type} · Rs.${fmt(h.amount)} · ${h.country}</div><div class="hist-meta">${h.flags.length} flag${h.flags.length!==1?'s':''} · ${h.timestamp}</div></div><div class="risk-pill" style="background:${s.badge};color:${s.bt};padding:5px 13px;border-radius:20px">${h.risk}</div></div>`;}).join('');
  const c={HIGH:0,MEDIUM:0,LOW:0,CLEAN:0};history.forEach(h=>c[h.risk]++);
  sa.innerHTML=`<div class="stats-grid">${['HIGH','MEDIUM','LOW','CLEAN'].map(r=>`<div class="stat"><div class="stat-num" style="color:${RS[r].text}">${c[r]}</div><div class="stat-lbl">${r}</div></div>`).join('')}</div><button class="btn btn-export" onclick="exportAll()">⬇ Export Full History as CSV</button>`;
}

// ─── CSV EXPORT ───────────────────────────────────────────────
function dl(rows,name){
  const cols=['txn_id','account_id','customer_name','date','amount','currency','txn_type','channel','country','counterparty','notes','risk','flags'];
  const lines=[cols.join(',')];
  rows.forEach(r=>{const ft=r.flags.map(f=>`[${f.level}] ${f.rule}`).join(' | ');lines.push([r.txn_id,r.account_id,`"${r.customer_name}"`,r.date,r.amount,r.currency||'INR',r.txn_type,r.channel,`"${r.country}"`,`"${r.counterparty||''}"`,`"${(r.notes||'').replace(/"/g,"'")}"`,r.risk,`"${ft}"`].join(','));});
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([lines.join('\n')],{type:'text/csv'}));a.download=`${name}_${today()}.csv`;a.click();
}
function exportSingle(){if(history.length)dl([history[0]],'aml_result');}
function exportBatch(){if(window._batch)dl(window._batch,'aml_batch');}
function exportAll(){if(history.length)dl(history,'aml_history');}

function downloadTemplate(){
  const cols=['txn_id','account_id','customer_name','date','amount','currency','txn_type','channel','country','counterparty','notes'];
  const sample=[['TXN001','ACC101','Ramesh Kumar','2026-04-04','1200000','INR','CASH_DEPOSIT','BRANCH','India','Self','Large deposit'],['TXN002','ACC102','Priya Sharma','2026-04-04','600000','INR','WIRE_TRANSFER','ONLINE','Iran','Tehran Traders','Import'],['TXN003','ACC103','Kavita Joshi','2026-04-04','25000','INR','NEFT','ONLINE','India','Amazon','Monthly vendor']];
  const ws=XLSX.utils.aoa_to_sheet([cols,...sample]);
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Transactions');
  XLSX.writeFile(wb,'AML_Template.xlsx');
}

// ─── DRAG DROP ────────────────────────────────────────────────
function dragOver(e){e.preventDefault();document.getElementById('drop-zone').classList.add('drag')}
function dragLeave(){document.getElementById('drop-zone').classList.remove('drag')}
function dropFile(e){e.preventDefault();dragLeave();const f=e.dataTransfer.files[0];if(f)handleFile(f);}

// ─── PROGRESS ────────────────────────────────────────────────
function showProg(show,lbl=''){
  const w=document.getElementById('progress-wrap');w.style.display=show?'block':'none';
  if(show){document.getElementById('progress-fill').style.width='0';document.getElementById('progress-label').textContent=lbl;}
}
function setProg(p,lbl){document.getElementById('progress-fill').style.width=p+'%';document.getElementById('progress-label').textContent=lbl;}

// ─── TABS ─────────────────────────────────────────────────────
function switchTab(tab){
  ['single','upload','history'].forEach(t=>{
    document.getElementById('page-'+t).style.display=t===tab?'block':'none';
    document.getElementById('tab-'+t).className='tab-btn'+(t===tab?' active':'');
  });
}

// ─── ALERTS ───────────────────────────────────────────────────
function showBatchAlert(type,msg){document.getElementById('batch-result').innerHTML=`<div class="alert alert-${type}">${type==='error'?'❌':'✅'} &nbsp;${msg}</div>`;}

// ─── HELPERS ──────────────────────────────────────────────────
function v(id){return document.getElementById(id).value.trim();}
function fmt(n){return parseFloat(n).toLocaleString('en-IN');}
function today(){return new Date().toISOString().slice(0,10);}
</script>
</body>
</html>
