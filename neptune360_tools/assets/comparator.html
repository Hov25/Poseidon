<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Neptune 360 · Import File Comparison</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2/dist/tabler-icons.min.css">
<style>
:root {
  --font-sans: 'Segoe UI', system-ui, -apple-system, sans-serif;
  --border-radius-lg: 10px;
  --border-radius-md: 6px;
}
body { background: #0f1117; margin: 0; }
* { box-sizing: border-box; margin: 0; padding: 0; }
.wrap { padding: 1.5rem; font-family: var(--font-sans); background: #0f1117; border-radius: var(--border-radius-lg); min-height: 200px; color: #e2e4ea; }
.header { display: flex; align-items: center; gap: 10px; margin-bottom: 1.5rem; }
.header-icon { width: 36px; height: 36px; background: #1e2a3a; border-radius: var(--border-radius-md); display: flex; align-items: center; justify-content: center; }
.header-icon i { font-size: 18px; color: #7eb8ff; }
.title { font-size: 17px; font-weight: 500; color: #f0f2f8; }
.subtitle { font-size: 13px; color: #9ba3b5; margin-top: 2px; }
.drop-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 1rem; }
.drop-zone { border: 1.5px dashed #3a4160; border-radius: var(--border-radius-lg); padding: 1.5rem 1rem; text-align: center; cursor: pointer; transition: background 0.15s, border-color 0.15s; position: relative; background: #161b27; }
.drop-zone:hover, .drop-zone.drag-over { background: #1a2236; border-color: #7eb8ff; }
.drop-zone.loaded { border-style: solid; border-color: #2a6b4a; background: #0f1f18; }
.drop-zone input[type=file] { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }
.drop-zone i { font-size: 24px; color: #7b85a0; display: block; margin-bottom: 8px; }
.drop-zone.loaded i { color: #4ade9a; }
.drop-label { font-size: 13px; color: #9ba3b5; }
.drop-zone.loaded .drop-label { color: #4ade9a; font-weight: 500; }
.drop-tag { font-size: 11px; color: #6b7590; margin-top: 4px; }
.compare-btn { width: 100%; padding: 10px; font-size: 14px; font-weight: 500; border-radius: var(--border-radius-md); border: 0.5px solid #3a4160; background: #1a1f2e; color: #d8dce8; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px; margin-bottom: 1.5rem; transition: background 0.15s; }
.compare-btn:hover { background: #22293d; }
.compare-btn:disabled { opacity: 0.35; cursor: default; }
.results { display: none; }
.results.show { display: block; }
.summary-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 1.5rem; }
.scard { background: #161b27; border: 0.5px solid #2d3348; border-radius: var(--border-radius-md); padding: 0.75rem 1rem; }
.scard-label { font-size: 11px; color: #9ba3b5; text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 4px; }
.scard-val { font-size: 22px; font-weight: 500; color: #f0f2f8; }
.scard-val.ok { color: #4ade9a; }
.scard-val.warn { color: #fcd34d; }
.scard-val.danger { color: #fc8181; }
.legend { display: flex; gap: 16px; margin-bottom: 1rem; flex-wrap: wrap; }
.legend-item { display: flex; align-items: center; gap: 6px; font-size: 12px; color: #9ba3b5; }
.legend-swatch { width: 12px; height: 12px; border-radius: 2px; flex-shrink: 0; }
.sw-warn   { background: #2d2000; border: 1px solid #b45309; }
.sw-danger { background: #2d0f0f; border: 1px solid #991b1b; }
.sw-info   { background: #0f1f36; border: 1px solid #1e4080; }
.sw-ok     { background: #0a1f14; border: 1px solid #166534; }
.filter-row { display: flex; gap: 8px; margin-bottom: 1rem; flex-wrap: wrap; align-items: center; }
.filter-row label { font-size: 13px; color: #9ba3b5; }
.pill { font-size: 12px; padding: 4px 12px; border-radius: 999px; border: 0.5px solid #3a4160; background: #161b27; color: #9ba3b5; cursor: pointer; transition: background 0.15s; }
.pill:hover { background: #1e2535; color: #c8cfe0; }
.pill.active { background: #1a2a42; border-color: #3d6cb5; color: #7eb8ff; }
.search-box { margin-left: auto; }
.search-box input { font-size: 13px; padding: 5px 10px; border-radius: var(--border-radius-md); border: 0.5px solid #3a4160; background: #161b27; color: #d8dce8; width: 180px; outline: none; }
.search-box input:focus { border-color: #7eb8ff; }
.search-box input::placeholder { color: #6b7590; }
.table-wrap { border: 0.5px solid #2d3348; border-radius: var(--border-radius-lg); overflow: hidden; }
.diff-table { width: 100%; border-collapse: collapse; font-size: 13px; table-layout: fixed; }
.diff-table th { background: #131822; font-weight: 500; font-size: 11px; color: #8a93a8; text-transform: uppercase; letter-spacing: 0.04em; padding: 8px 10px; border-bottom: 0.5px solid #2d3348; text-align: left; }
.diff-table td { padding: 7px 10px; border-bottom: 0.5px solid #1e2535; vertical-align: top; word-break: break-word; color: #d8dce8; }
.diff-table tr:last-child td { border-bottom: none; }
.diff-table tr.acct-header td { background: #13182b; border-top: 1px solid #2d3348; padding: 5px 10px; }
.diff-table tr.acct-header.ah-warn td { background: #1f1700; border-top-color: #92400e; }
.diff-table tr.acct-header.ah-danger td { background: #1f0a0a; border-top-color: #991b1b; }
.diff-table tr.acct-header.ah-info td { background: #0a1525; border-top-color: #1e4080; }
.acct-header-inner { display: flex; align-items: center; gap: 8px; }
.acct-header-key { font-weight: 500; font-size: 12px; color: #e2e6f0; }
.acct-header-acct { font-size: 11px; color: #9ba3b5; }
.diff-table tr.row-match td { background: #0d1117; color: #c8cfe0; }
.diff-table tr.row-mismatch td { background: #1f1700; color: #e8dcc0; }
.diff-table tr.row-missing-a td { background: #1f0a0a; color: #e8c0c0; }
.diff-table tr.row-missing-b td { background: #0a1525; color: #b8cce8; }
.badge { display: inline-block; font-size: 10px; padding: 2px 7px; border-radius: 999px; font-weight: 500; white-space: nowrap; }
.badge-ok     { background: #0a1f14; color: #4ade9a; border: 0.5px solid #166534; }
.badge-warn   { background: #2d2000; color: #fcd34d; border: 0.5px solid #b45309; }
.badge-danger { background: #2d0f0f; color: #fc8181; border: 0.5px solid #991b1b; }
.badge-info   { background: #0f1f36; color: #7eb8ff; border: 0.5px solid #1e4080; }
.col-key { width: 17%; } .col-acct { width: 15%; } .col-field { width: 14%; } .col-fa { width: 20%; } .col-fb { width: 20%; } .col-status { width: 14%; }
.export-row { display: flex; gap: 8px; margin-top: 1rem; }
.export-btn { font-size: 12px; padding: 6px 14px; border-radius: var(--border-radius-md); border: 0.5px solid #3a4160; background: #161b27; color: #9ba3b5; cursor: pointer; display: flex; align-items: center; gap: 6px; transition: background 0.15s; }
.export-btn:hover { background: #1e2535; color: #d8dce8; }
.empty { padding: 2rem; text-align: center; color: #6b7590; font-size: 13px; }
</style>
</head>
<body>

<div class="wrap">
  <h2 class="sr-only" style="position:absolute;width:1px;height:1px;overflow:hidden">Neptune 360 Import File Comparison</h2>
  <div class="header">
    <div class="header-icon"><i class="ti ti-file-diff" aria-hidden="true"></i></div>
    <div>
      <div class="title">Neptune 360 Import File Comparison</div>
      <div class="subtitle">Compare two import files — customer and account fields side by side</div>
    </div>
  </div>

  <div class="drop-row">
    <div class="drop-zone" id="zone-a" ondragover="onDrag(event,'a')" ondragleave="offDrag('a')" ondrop="onDrop(event,'a')">
      <input type="file" accept=".txt,.dat,.imp,.csv" onchange="loadFile(event,'a')">
      <i class="ti ti-file-upload" id="icon-a" aria-hidden="true"></i>
      <div class="drop-label" id="label-a">Drop file A here or click to browse</div>
      <div class="drop-tag" id="tag-a">Reference file</div>
    </div>
    <div class="drop-zone" id="zone-b" ondragover="onDrag(event,'b')" ondragleave="offDrag('b')" ondrop="onDrop(event,'b')">
      <input type="file" accept=".txt,.dat,.imp,.csv" onchange="loadFile(event,'b')">
      <i class="ti ti-file-upload" id="icon-b" aria-hidden="true"></i>
      <div class="drop-label" id="label-b">Drop file B here or click to browse</div>
      <div class="drop-tag" id="tag-b">Comparison file</div>
    </div>
  </div>

  <button class="compare-btn" id="compare-btn" disabled onclick="runCompare()">
    <i class="ti ti-arrows-diff" aria-hidden="true"></i> Compare Files
  </button>

  <div class="results" id="results">
    <div class="summary-cards" id="summary-cards"></div>
    <div class="legend">
      <div class="legend-item"><div class="legend-swatch sw-danger"></div> Only in File A / Missing from B</div>
      <div class="legend-item"><div class="legend-swatch sw-info"></div> Only in File B / Missing from A</div>
      <div class="legend-item"><div class="legend-swatch sw-warn"></div> Field value differs</div>
      <div class="legend-item"><div class="legend-swatch sw-ok"></div> All fields match</div>
    </div>
    <div class="filter-row">
      <label>Show:</label>
      <span class="pill active" onclick="setFilter('all',this)">All accounts</span>
      <span class="pill" onclick="setFilter('problems',this)">Problems only</span>
      <span class="pill" onclick="setFilter('mismatch',this)">Mismatches</span>
      <span class="pill" onclick="setFilter('missing',this)">Missing records</span>
      <span class="pill" onclick="setFilter('match',this)">Full matches</span>
      <div class="search-box">
        <input type="text" placeholder="Search key or account…" oninput="setSearch(this.value)" aria-label="Search records">
      </div>
    </div>
    <div class="table-wrap">
      <table class="diff-table" role="table">
        <thead>
          <tr>
            <th class="col-key">Premises Key</th>
            <th class="col-acct">Account #</th>
            <th class="col-field">Field</th>
            <th class="col-fa">File A Value</th>
            <th class="col-fb">File B Value</th>
            <th class="col-status">Status</th>
          </tr>
        </thead>
        <tbody id="diff-body"></tbody>
      </table>
    </div>
    <div class="export-row">
      <button class="export-btn" onclick="exportCSV()"><i class="ti ti-download" aria-hidden="true"></i> Export CSV</button>
      <button class="export-btn" onclick="copyText()"><i class="ti ti-copy" aria-hidden="true"></i> Copy summary</button>
    </div>
  </div>
</div>

<script>
const files = { a: null, b: null };
let accountGroups = [];
let activeFilter = 'all';
let searchTerm = '';

const FIELDS = [
  { name: 'Customer name',  ex: r => (r.customerName||'').trim() },
  { name: 'Account number', ex: r => (r.accountNumber||'').trim() },
  { name: 'Account status', ex: r => (r.accountStatus||'').trim() },
  { name: 'Address 1',      ex: r => (r.address1||'').trim() },
  { name: 'Address 2',      ex: r => (r.address2||'').trim() },
  { name: 'City',           ex: r => (r.city||'').trim() },
  { name: 'State',          ex: r => (r.state||'').trim() },
  { name: 'Zip',            ex: r => (r.zip||'').trim() },
  { name: 'Email',          ex: r => (r.email||'').trim() },
  { name: 'Phone 1',        ex: r => (r.phone1||'').trim() },
  { name: 'Phone 2',        ex: r => (r.phone2||'').trim() },
  { name: 'Custom 1',       ex: r => (r.custom1||'').trim() },
  { name: 'Custom 2',       ex: r => (r.custom2||'').trim() },
];

function parseFile(text) {
  const lines = text.split(/\r?\n/);
  const premises = {};
  lines.forEach(line => {
    if (!line || line.length < 5) return;
    const rt = line.substring(0,5).trim().toUpperCase();
    if (rt === 'PRMDT') {
      const key = line.substring(83,103).trim();
      if (!key) return;
      premises[key] = { recordType:'PRMDT', premisesKey:key,
        address1:line.substring(5,31), address2:line.substring(31,57),
        customerName:line.substring(57,83), accountNumber:line.substring(103,123),
        accountStatus:line.substring(123,127),
        custom1:line.length>127?line.substring(127,153):'',
        custom2:line.length>153?line.substring(153,179):'',
        city:'', state:'', zip:'', email:'', phone1:'', phone2:'' };
    } else if (rt === 'PRMD2') {
      const key = line.substring(5,25).trim();
      if (!key) return;
      premises[key] = { recordType:'PRMD2', premisesKey:key,
        customerName:line.substring(25,51), address1:line.substring(25,51), address2:'',
        accountNumber:line.substring(103,123), accountStatus:line.substring(123,127),
        city:   line.length>189?line.substring(189,215):'',
        state:  line.length>215?line.substring(215,217):'',
        zip:    line.length>217?line.substring(217,228):'',
        phone1: line.length>228?line.substring(228,238):'',
        phone2: line.length>238?line.substring(238,248):'',
        email:  line.length>618?line.substring(618,668):'',
        custom1:line.length>438?line.substring(438,464):'',
        custom2:line.length>464?line.substring(464,490):'' };
    }
  });
  return premises;
}

function loadFile(e, which) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    files[which] = { text: ev.target.result, name: file.name, size: file.size };
    document.getElementById('zone-'+which).classList.add('loaded');
    document.getElementById('icon-'+which).className = 'ti ti-circle-check';
    document.getElementById('label-'+which).textContent = file.name;
    document.getElementById('tag-'+which).textContent = (file.size/1024).toFixed(1)+' KB';
    checkReady();
  };
  reader.readAsText(file);
}

function onDrag(e,w){e.preventDefault();document.getElementById('zone-'+w).classList.add('drag-over');}
function offDrag(w){document.getElementById('zone-'+w).classList.remove('drag-over');}
function onDrop(e,w){e.preventDefault();offDrag(w);const f=e.dataTransfer.files[0];if(!f)return;loadFile({target:{files:[f]}},w);}
function checkReady(){document.getElementById('compare-btn').disabled=!(files.a&&files.b);}

function runCompare() {
  const a = parseFile(files.a.text);
  const b = parseFile(files.b.text);
  const allKeys = [...new Set([...Object.keys(a),...Object.keys(b)])].sort();
  accountGroups = [];
  allKeys.forEach(key => {
    const ra=a[key], rb=b[key];
    const acct=(ra||rb).accountNumber.trim();
    const rows=[];
    if (!ra) {
      rows.push({field:'(record)',valA:'—',valB:'present in B',status:'missing-a'});
    } else if (!rb) {
      rows.push({field:'(record)',valA:'present in A',valB:'—',status:'missing-b'});
    } else {
      FIELDS.forEach(f=>{
        const va=f.ex(ra),vb=f.ex(rb);
        if(va||vb) rows.push({field:f.name,valA:va||'(blank)',valB:vb||'(blank)',status:va===vb?'match':'mismatch'});
      });
    }
    const hasMismatch=rows.some(r=>r.status==='mismatch');
    const hasMissingA=rows.some(r=>r.status==='missing-a');
    const hasMissingB=rows.some(r=>r.status==='missing-b');
    const accountStatus=hasMissingA?'missing-a':hasMissingB?'missing-b':hasMismatch?'mismatch':'match';
    accountGroups.push({key,acct,rows,accountStatus});
  });
  renderSummary();
  renderTable();
  document.getElementById('results').classList.add('show');
}

function renderSummary() {
  const total=accountGroups.length;
  const mismatches=accountGroups.filter(g=>g.accountStatus==='mismatch').length;
  const missing=accountGroups.filter(g=>g.accountStatus==='missing-a'||g.accountStatus==='missing-b').length;
  const matches=accountGroups.filter(g=>g.accountStatus==='match').length;
  document.getElementById('summary-cards').innerHTML=`
    <div class="scard"><div class="scard-label">Total Accounts</div><div class="scard-val">${total}</div></div>
    <div class="scard"><div class="scard-label">Clean Matches</div><div class="scard-val ok">${matches}</div></div>
    <div class="scard"><div class="scard-label">Field Mismatches</div><div class="scard-val ${mismatches?'warn':'ok'}">${mismatches}</div></div>
    <div class="scard"><div class="scard-label">Missing Records</div><div class="scard-val ${missing?'danger':'ok'}">${missing}</div></div>`;
}

function renderTable() {
  let groups=accountGroups;
  if(activeFilter==='problems') groups=groups.filter(g=>g.accountStatus!=='match');
  else if(activeFilter==='mismatch') groups=groups.filter(g=>g.accountStatus==='mismatch');
  else if(activeFilter==='missing') groups=groups.filter(g=>g.accountStatus==='missing-a'||g.accountStatus==='missing-b');
  else if(activeFilter==='match') groups=groups.filter(g=>g.accountStatus==='match');
  if(searchTerm){const s=searchTerm.toLowerCase();groups=groups.filter(g=>g.key.toLowerCase().includes(s)||g.acct.toLowerCase().includes(s));}
  const body=document.getElementById('diff-body');
  if(!groups.length){body.innerHTML='<tr><td colspan="6" class="empty">No records match this filter.</td></tr>';return;}
  const html=[];
  groups.forEach(g=>{
    const isProblem=g.accountStatus!=='match';
    const ahClass=g.accountStatus==='missing-a'?'ah-danger':g.accountStatus==='missing-b'?'ah-info':g.accountStatus==='mismatch'?'ah-warn':'';
    const statusBadge=g.accountStatus==='match'?'<span class="badge badge-ok">All Match</span>'
      :g.accountStatus==='mismatch'?'<span class="badge badge-warn">Has Differences</span>'
      :g.accountStatus==='missing-a'?'<span class="badge badge-danger">Only in B</span>'
      :'<span class="badge badge-info">Only in A</span>';
    const iconColor=g.accountStatus==='mismatch'?'#fcd34d':isProblem?'#fc8181':'#4ade9a';
    const iconName=isProblem?'ti-alert-triangle':'ti-circle-check';
    html.push(`<tr class="acct-header ${ahClass}"><td colspan="6"><div class="acct-header-inner">
      <i class="ti ${iconName}" aria-hidden="true" style="font-size:14px;color:${iconColor}"></i>
      <span class="acct-header-key">${esc(g.key)}</span>
      <span class="acct-header-acct">Acct: ${esc(g.acct)}</span>
      <span style="margin-left:auto">${statusBadge}</span>
    </div></td></tr>`);
    g.rows.forEach(r=>{
      const rc=r.status==='mismatch'?'row-mismatch':r.status==='missing-a'?'row-missing-a':r.status==='missing-b'?'row-missing-b':'row-match';
      const badge=r.status==='match'?'<span class="badge badge-ok">Match</span>'
        :r.status==='mismatch'?'<span class="badge badge-warn">Mismatch</span>'
        :r.status==='missing-a'?'<span class="badge badge-danger">Only in B</span>'
        :'<span class="badge badge-info">Only in A</span>';
      html.push(`<tr class="${rc}"><td>${esc(g.key)}</td><td>${esc(g.acct)}</td><td>${esc(r.field)}</td><td>${esc(r.valA)}</td><td>${esc(r.valB)}</td><td>${badge}</td></tr>`);
    });
  });
  body.innerHTML=html.join('');
}

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function setFilter(f,el){activeFilter=f;document.querySelectorAll('.pill').forEach(p=>p.classList.remove('active'));el.classList.add('active');renderTable();}
function setSearch(v){searchTerm=v;renderTable();}

function exportCSV(){
  const rows=[['Premises Key','Account Number','Field','File A Value','File B Value','Status','Account Status']];
  accountGroups.forEach(g=>g.rows.forEach(r=>rows.push([g.key,g.acct,r.field,r.valA,r.valB,r.status,g.accountStatus])));
  const csv=rows.map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(',')).join('\n');
  const blob=new Blob([csv],{type:'text/csv'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');a.href=url;a.download='neptune360_comparison.csv';a.click();URL.revokeObjectURL(url);
}

function copyText(){
  const problems=accountGroups.filter(g=>g.accountStatus!=='match').length;
  const matches=accountGroups.filter(g=>g.accountStatus==='match').length;
  const text=`Neptune 360 Import File Comparison — ${files.a.name} vs ${files.b.name}\nTotal accounts: ${accountGroups.length}\nClean matches: ${matches}\nAccounts with issues: ${problems}`;
  navigator.clipboard.writeText(text).catch(()=>{});
}
</script>
</body>
</html>
