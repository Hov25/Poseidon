<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Neptune 360 · EXP → CSV Converter</title>
<style>
  :root {
    --bg:        #0f1117;
    --surface:   #1a1d27;
    --surface2:  #22263a;
    --border:    #2e3350;
    --accent:    #4f8ef7;
    --accent2:   #6ee7b7;
    --text:      #e8eaf6;
    --muted:     #8b92b8;
    --danger:    #f87171;
    --warn:      #fbbf24;
    --success:   #34d399;
    --radius:    10px;
    --font:      'Segoe UI', system-ui, sans-serif;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    background: var(--bg);
    color: var(--text);
    font-family: var(--font);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 32px 16px 48px;
  }

  /* ── Header ── */
  .header {
    display: flex;
    align-items: center;
    gap: 14px;
    margin-bottom: 32px;
  }
  .logo {
    width: 44px; height: 44px;
    background: linear-gradient(135deg, var(--accent), var(--accent2));
    border-radius: 12px;
    display: flex; align-items: center; justify-content: center;
    font-size: 22px;
  }
  .header h1 { font-size: 22px; font-weight: 700; letter-spacing: .3px; }
  .header p  { font-size: 13px; color: var(--muted); margin-top: 2px; }

  /* ── Main card ── */
  .card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    width: 100%;
    max-width: 820px;
    padding: 28px;
    margin-bottom: 20px;
  }
  .card-title {
    font-size: 13px;
    font-weight: 600;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: .8px;
    margin-bottom: 16px;
  }

  /* ── Drop zone ── */
  #dropzone {
    border: 2px dashed var(--border);
    border-radius: var(--radius);
    padding: 48px 24px;
    text-align: center;
    cursor: pointer;
    transition: border-color .2s, background .2s;
    position: relative;
  }
  #dropzone:hover, #dropzone.drag-over {
    border-color: var(--accent);
    background: rgba(79,142,247,.06);
  }
  #dropzone.has-file {
    border-color: var(--accent2);
    background: rgba(110,231,183,.05);
  }
  #dropzone svg { margin-bottom: 12px; opacity: .6; }
  #dropzone .drop-title { font-size: 16px; font-weight: 600; margin-bottom: 6px; }
  #dropzone .drop-sub   { font-size: 13px; color: var(--muted); }
  #dropzone .file-badge {
    display: inline-flex; align-items: center; gap: 8px;
    background: var(--surface2); border: 1px solid var(--border);
    border-radius: 6px; padding: 6px 14px; font-size: 13px; margin-top: 12px;
    color: var(--accent2);
  }
  #fileInput { display: none; }

  /* ── Options ── */
  .options-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
  }
  @media (max-width: 540px) { .options-grid { grid-template-columns: 1fr; } }

  .opt-group label {
    display: block;
    font-size: 12px;
    color: var(--muted);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: .6px;
    margin-bottom: 8px;
  }
  .toggle-row {
    display: flex; align-items: center; gap: 10px;
    font-size: 13px; color: var(--text); margin-bottom: 8px;
  }
  .toggle {
    position: relative; width: 36px; height: 20px; cursor: pointer;
  }
  .toggle input { opacity: 0; width: 0; height: 0; }
  .slider {
    position: absolute; inset: 0;
    background: var(--surface2); border-radius: 20px;
    transition: background .2s;
  }
  .slider::before {
    content: '';
    position: absolute;
    width: 14px; height: 14px;
    left: 3px; top: 3px;
    background: var(--muted);
    border-radius: 50%;
    transition: transform .2s, background .2s;
  }
  .toggle input:checked + .slider { background: var(--accent); }
  .toggle input:checked + .slider::before {
    transform: translateX(16px);
    background: #fff;
  }

  /* ── Buttons ── */
  .btn-row { display: flex; gap: 12px; flex-wrap: wrap; margin-top: 4px; }
  .btn {
    display: inline-flex; align-items: center; gap: 7px;
    padding: 10px 20px; border-radius: 8px;
    font-size: 14px; font-weight: 600; cursor: pointer;
    border: none; transition: opacity .15s, transform .1s;
  }
  .btn:active { transform: scale(.97); }
  .btn:disabled { opacity: .4; cursor: not-allowed; }
  .btn-primary {
    background: linear-gradient(135deg, var(--accent), #3b78e7);
    color: #fff;
  }
  .btn-secondary {
    background: var(--surface2);
    border: 1px solid var(--border);
    color: var(--text);
  }
  .btn-success {
    background: linear-gradient(135deg, var(--success), #059669);
    color: #fff;
  }

  /* ── Progress ── */
  #progressWrap { display: none; margin-top: 20px; }
  .progress-bar-bg {
    background: var(--surface2);
    border-radius: 999px;
    height: 8px;
    overflow: hidden;
    margin-top: 8px;
  }
  .progress-bar-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    border-radius: 999px;
    width: 0%;
    transition: width .2s;
  }
  #progressLabel { font-size: 13px; color: var(--muted); }

  /* ── Stats ── */
  #stats { display: none; }
  .stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
    gap: 12px;
    margin-bottom: 20px;
  }
  .stat-box {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 14px 16px;
  }
  .stat-box .val { font-size: 24px; font-weight: 700; color: var(--accent); }
  .stat-box .lbl { font-size: 12px; color: var(--muted); margin-top: 2px; }

  /* ── Preview table ── */
  #previewWrap { display: none; }
  .table-scroll {
    overflow-x: auto;
    border: 1px solid var(--border);
    border-radius: 8px;
    margin-top: 12px;
    max-height: 340px;
    overflow-y: auto;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
    white-space: nowrap;
  }
  thead th {
    background: var(--surface2);
    padding: 9px 12px;
    text-align: left;
    font-size: 11px;
    font-weight: 600;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: .5px;
    position: sticky; top: 0;
    border-bottom: 1px solid var(--border);
  }
  tbody tr { border-bottom: 1px solid var(--border); }
  tbody tr:last-child { border-bottom: none; }
  tbody tr:hover { background: var(--surface2); }
  tbody td { padding: 7px 12px; color: var(--text); }
  .badge {
    display: inline-block;
    padding: 2px 7px; border-radius: 4px;
    font-size: 10px; font-weight: 700; letter-spacing: .4px;
  }
  .badge-green  { background: rgba(52,211,153,.15); color: var(--success); }
  .badge-yellow { background: rgba(251,191,36,.15);  color: var(--warn); }
  .badge-red    { background: rgba(248,113,113,.15); color: var(--danger); }
  .badge-blue   { background: rgba(79,142,247,.15);  color: var(--accent); }

  /* ── Log ── */
  #logWrap { display: none; }
  #log {
    background: #0a0c14;
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 14px;
    font-family: 'Consolas', monospace;
    font-size: 12px;
    line-height: 1.7;
    max-height: 180px;
    overflow-y: auto;
    color: var(--muted);
    margin-top: 12px;
  }
  .log-ok   { color: var(--success); }
  .log-warn { color: var(--warn); }
  .log-err  { color: var(--danger); }
  .log-info { color: var(--accent); }
</style>
</head>
<body>

<div class="header">
  <div class="logo">🌊</div>
  <div>
    <h1>Neptune 360 · EXP → CSV</h1>
    <p>Meter reading export file converter</p>
  </div>
</div>

<!-- Step 1: File -->
<div class="card">
  <div class="card-title">1 · Select Export File</div>
  <div id="dropzone" onclick="document.getElementById('fileInput').click()">
    <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
      <path d="M4 14v4a2 2 0 002 2h12a2 2 0 002-2v-4M12 3v11M8 7l4-4 4 4"/>
    </svg>
    <div class="drop-title">Drop your .exp file here</div>
    <div class="drop-sub">or click to browse</div>
    <div id="fileBadge" style="display:none" class="file-badge">
      📄 <span id="fileName"></span>
    </div>
  </div>
  <input type="file" id="fileInput" accept=".exp,.EXP,text/*">
</div>

<!-- Step 2: Options -->
<div class="card">
  <div class="card-title">2 · Output Options</div>
  <div class="options-grid">
    <div class="opt-group">
      <label>Include in CSV</label>
      <div class="toggle-row">
        <label class="toggle"><input type="checkbox" id="optConsumption" checked><span class="slider"></span></label>
        Calculated Consumption
      </div>
      <div class="toggle-row">
        <label class="toggle"><input type="checkbox" id="optReadCode" checked><span class="slider"></span></label>
        Read Code Description
      </div>
      <div class="toggle-row">
        <label class="toggle"><input type="checkbox" id="optLimits" checked><span class="slider"></span></label>
        Hi / Low Limits
      </div>
      <div class="toggle-row">
        <label class="toggle"><input type="checkbox" id="optCollectionID" checked><span class="slider"></span></label>
        Collection ID (MIU/Endpoint)
      </div>
    </div>
    <div class="opt-group">
      <label>Formatting</label>
      <div class="toggle-row">
        <label class="toggle"><input type="checkbox" id="optFmtDate" checked><span class="slider"></span></label>
        Format Dates (MM/DD/YYYY)
      </div>
      <div class="toggle-row">
        <label class="toggle"><input type="checkbox" id="optStripZeros" checked><span class="slider"></span></label>
        Strip Leading Zeros from Readings
      </div>
      <div class="toggle-row">
        <label class="toggle"><input type="checkbox" id="optFlagZeroCons" checked><span class="slider"></span></label>
        Flag Zero-Consumption Rows
      </div>
      <div class="toggle-row">
        <label class="toggle"><input type="checkbox" id="optFlagCodes" checked><span class="slider"></span></label>
        Flag Non-RR Read Codes
      </div>
    </div>
  </div>
</div>

<!-- Step 3: Convert -->
<div class="card">
  <div class="card-title">3 · Convert</div>
  <div class="btn-row">
    <button class="btn btn-primary" id="btnConvert" onclick="convert()" disabled>
      ⚡ Convert to CSV
    </button>
    <button class="btn btn-secondary" id="btnReset" onclick="reset()">
      ↺ Reset
    </button>
    <button class="btn btn-success" id="btnDownload" onclick="download()" style="display:none">
      ⬇ Download CSV
    </button>
  </div>

  <div id="progressWrap">
    <div id="progressLabel">Parsing file…</div>
    <div class="progress-bar-bg"><div class="progress-bar-fill" id="progressFill"></div></div>
  </div>

  <!-- Log -->
  <div id="logWrap">
    <div class="card-title" style="margin-top:20px">Processing Log</div>
    <div id="log"></div>
  </div>
</div>

<!-- Stats -->
<div class="card" id="stats">
  <div class="card-title">Summary</div>
  <div class="stats-grid" id="statsGrid"></div>
</div>

<!-- Preview -->
<div class="card" id="previewWrap">
  <div class="card-title">Preview (first 50 rows)</div>
  <div class="table-scroll">
    <table id="previewTable"></table>
  </div>
</div>

<script>
// ── State ──────────────────────────────────────────
let fileContent = null;
let csvContent  = null;
let csvFilename = 'readings.csv';

// ── Read Code Map ──────────────────────────────────
const READ_CODES = {
  RR:'Radio Read', RH:'Radio High', RL:'Radio Low',
  RZ:'Radio Zero Consumption', RN:'Radio Negative Consumption',
  RT:'Radio Tamper', RV:'Radio Verified', RI:'Radio Inactive',
  RA:'Radio Alpha', KR:'Keyed Read', KH:'Keyed High', KL:'Keyed Low',
  KZ:'Keyed Zero', KN:'Keyed Negative', KV:'Keyed Verified',
  KI:'Keyed Inactive', KA:'Keyed Alpha', ER:'External Read (Probe)',
  EF:'External Failure', FC:'Failure Read Compare',
  AR:'Admin Read', AH:'Admin High Fail', AL:'Admin Low Fail',
  AZ:'Admin Zero Use', AU:'Admin Inactive',
};

const STATUS_MAP = { CO:'Complete', IN:'Incomplete', SK:'Skipped' };

// ── File drop / select ─────────────────────────────
const dropzone = document.getElementById('dropzone');
dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('drag-over'); });
dropzone.addEventListener('dragleave',()=> dropzone.classList.remove('drag-over'));
dropzone.addEventListener('drop', e => {
  e.preventDefault();
  dropzone.classList.remove('drag-over');
  handleFile(e.dataTransfer.files[0]);
});
document.getElementById('fileInput').addEventListener('change', e => handleFile(e.target.files[0]));

function handleFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    fileContent = e.target.result;
    csvFilename = file.name.replace(/\.exp$/i, '') + '.csv';
    document.getElementById('fileName').textContent = file.name + ' (' + (file.size/1024).toFixed(1) + ' KB)';
    document.getElementById('fileBadge').style.display = 'inline-flex';
    dropzone.classList.add('has-file');
    document.getElementById('btnConvert').disabled = false;
    addLog('info', `Loaded: ${file.name} (${file.size.toLocaleString()} bytes)`);
    showLog();
  };
  reader.readAsText(file, 'latin1');
}

// ── Core Parser ────────────────────────────────────
function parseExp(text) {
  const opts = {
    consumption: document.getElementById('optConsumption').checked,
    readCode:    document.getElementById('optReadCode').checked,
    limits:      document.getElementById('optLimits').checked,
    collID:      document.getElementById('optCollectionID').checked,
    fmtDate:     document.getElementById('optFmtDate').checked,
    stripZeros:  document.getElementById('optStripZeros').checked,
    flagZero:    document.getElementById('optFlagZeroCons').checked,
    flagCodes:   document.getElementById('optFlagCodes').checked,
  };

  const lines = text.split(/\r?\n/);
  const rows = [];
  let premises = {}, meter = {}, ordst = {};
  let counts = { premises:0, meters:0, readings:0, skipped:0, warnings:0, zeroCons:0 };

  function fmtDate(d) {
    d = (d||'').trim();
    if (opts.fmtDate && d.length === 8 && /^\d{8}$/.test(d))
      return `${d.slice(4,6)}/${d.slice(6,8)}/${d.slice(0,4)}`;
    return d;
  }
  function cleanNum(s) {
    s = (s||'').trim();
    if (opts.stripZeros && /^\d+$/.test(s)) return String(parseInt(s,10));
    return s;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.length < 5) continue;
    const rt = line.slice(0,5).trimEnd();

    if (rt === 'PRMDT') {
      premises = {
        address:      line.slice(5,31).trim(),
        customerName: line.slice(57,83).trim(),
        premisesKey:  line.slice(83,103).trim(),
        accountNum:   line.slice(103,123).trim().replace(/^0+/,'') || '0',
        acctStatus:   line.slice(123,127).trim(),
      };
      counts.premises++;
      meter = {}; ordst = {};
    }
    else if (rt === 'MTRDT') {
      meter = {
        meterNumber:  line.slice(37,57).trim(),
        meterSize:    line.slice(85,93).trim(),
        meterUOM:     line.slice(107,110).trim(),
        prevReadDate: line.slice(244,252).trim(),
      };
      counts.meters++;
      ordst = {};
    }
    else if (rt === 'ORDST') {
      ordst = {
        readDate:    line.slice(5,13).trim(),
        readTime:    line.slice(13,19).trim(),
        readerID:    line.slice(24,44).trim(),
        orderStatus: line.slice(44,46).trim(),
      };
    }
    else if (rt === 'RDGDT') {
      counts.readings++;
      const collID   = line.slice(9,22).trim();
      const dials    = line.slice(49,51).trim();
      const decimals = line.slice(53,55).trim();
      const hiLim    = cleanNum(line.slice(58,68));
      const loLim    = cleanNum(line.slice(68,78));
      const prevRead = cleanNum(line.slice(78,88));
      const curRead  = cleanNum(line.slice(88,98));
      const readCode = line.slice(108,110).trim();

      let consumption = '';
      if (opts.consumption) {
        const p = parseInt(prevRead,10), c = parseInt(curRead,10);
        if (!isNaN(p) && !isNaN(c)) consumption = String(c - p);
      }

      if (opts.skipZero && consumption === '0') { counts.skipped++; continue; }

      const isZero = consumption === '0';
      if (isZero) counts.zeroCons++;

      const status = STATUS_MAP[ordst.orderStatus] || ordst.orderStatus || '';
      const rcDesc = READ_CODES[readCode] || readCode;
      const isFlag = opts.flagCodes && readCode !== 'RR' && readCode !== '';
      if (isFlag) counts.warnings++;

      const row = {
        'Address':          premises.address      || '',
        'Customer Name':    premises.customerName || '',
        'Account Number':   premises.accountNum   || '',
        'Account Status':   premises.acctStatus   || '',
        'Meter Number':     meter.meterNumber     || '',
        'Meter Size':       meter.meterSize       || '',
        'Dials':            dials                 || '',
      };
      if (opts.collID)    row['Collection ID']    = collID;
      row['Prev Read Date']   = fmtDate(meter.prevReadDate || '');
      row['Previous Reading'] = prevRead;
      row['Read Date']        = fmtDate(ordst.readDate || '');
      row['Current Reading']  = curRead;
      if (opts.consumption) row['Consumption']   = consumption;
      row['Order Status']     = status;
      if (opts.readCode) row['Read Code']        = readCode ? `${readCode} - ${rcDesc}` : '';
      else               row['Read Code']        = readCode;
      // Zero consumption flag — always included, shown as text for CSV compatibility
      if (opts.flagZero) row['Zero Consumption'] = isZero ? 'YES' : '';
      if (opts.flagCodes) row['Read Code Flag']  = isFlag ? 'REVIEW' : '';
      if (opts.limits) {
        row['Hi Limit'] = hiLim;
        row['Low Limit'] = loLim;
      }
      rows.push(row);
    }
  }
  return { rows, counts };
}

// ── Convert ────────────────────────────────────────
function convert() {
  if (!fileContent) return;
  document.getElementById('progressWrap').style.display = 'block';
  setProgress(10, 'Parsing records…');
  addLog('info', 'Starting conversion…');

  setTimeout(() => {
    try {
      setProgress(40, 'Building CSV…');
      const { rows, counts } = parseExp(fileContent);
      setProgress(80, 'Rendering preview…');

      if (rows.length === 0) {
        addLog('err', 'No RDGDT records found. Is this a valid Neptune 360 export file?');
        setProgress(0, '');
        return;
      }

      // Build CSV string
      const headers = Object.keys(rows[0]);
      const csvLines = [headers.join(',')];
      for (const row of rows) {
        csvLines.push(headers.map(h => {
          const v = String(row[h] ?? '').replace(/"/g, '""');
          return v.includes(',') || v.includes('"') || v.includes('\n') ? `"${v}"` : v;
        }).join(','));
      }
      csvContent = csvLines.join('\r\n');

      setProgress(100, 'Done!');
      addLog('ok', `Parsed ${counts.premises} premises, ${counts.meters} meters, ${counts.readings} readings.`);
      if (counts.zeroCons) addLog('warn', `${counts.zeroCons} rows flagged for zero consumption.`);
      if (counts.skipped)  addLog('warn', `Skipped ${counts.skipped} rows.`);
      if (counts.warnings) addLog('warn', `${counts.warnings} rows have non-RR read codes — flagged for review.`);
      addLog('ok', `Output: ${rows.length} CSV rows, ${headers.length} columns.`);

      renderStats(counts, rows);
      renderPreview(rows, headers);

      document.getElementById('btnDownload').style.display = 'inline-flex';
    } catch(e) {
      addLog('err', 'Error: ' + e.message);
      console.error(e);
    }
  }, 50);
}

// ── Download ───────────────────────────────────────
function download() {
  if (!csvContent) return;
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = csvFilename;
  a.click();
  addLog('ok', `Downloaded: ${csvFilename}`);
}

// ── Reset ──────────────────────────────────────────
function reset() {
  fileContent = null; csvContent = null;
  document.getElementById('fileInput').value = '';
  document.getElementById('fileBadge').style.display = 'none';
  document.getElementById('fileName').textContent = '';
  dropzone.classList.remove('has-file', 'drag-over');
  document.getElementById('btnConvert').disabled = true;
  document.getElementById('btnDownload').style.display = 'none';
  document.getElementById('progressWrap').style.display = 'none';
  document.getElementById('logWrap').style.display = 'none';
  document.getElementById('stats').style.display = 'none';
  document.getElementById('previewWrap').style.display = 'none';
  document.getElementById('log').innerHTML = '';
  document.getElementById('statsGrid').innerHTML = '';
  document.getElementById('previewTable').innerHTML = '';
  setProgress(0,'');
}

// ── Progress ───────────────────────────────────────
function setProgress(pct, label) {
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('progressLabel').textContent = label;
}

// ── Log ────────────────────────────────────────────
function addLog(type, msg) {
  const el = document.getElementById('log');
  const cls = { ok:'log-ok', warn:'log-warn', err:'log-err', info:'log-info' }[type] || '';
  el.innerHTML += `<div class="${cls}">${new Date().toLocaleTimeString()} › ${msg}</div>`;
  el.scrollTop = el.scrollHeight;
}
function showLog() {
  document.getElementById('logWrap').style.display = 'block';
}

// ── Stats ──────────────────────────────────────────
function renderStats(counts, rows) {
  const flagged = rows.filter(r => r['Read Code Flag'] === 'REVIEW').length;

  const items = [
    { val: counts.premises,         lbl: 'Premises' },
    { val: counts.meters,           lbl: 'Meters' },
    { val: rows.length,             lbl: 'CSV Rows' },
    { val: counts.zeroCons || '—',  lbl: 'Zero Consumption' },
    { val: flagged || '—',          lbl: 'Flagged Read Codes' },
    { val: counts.skipped || '—',   lbl: 'Skipped' },
  ];
  document.getElementById('statsGrid').innerHTML = items.map(i =>
    `<div class="stat-box"><div class="val">${i.val}</div><div class="lbl">${i.lbl}</div></div>`
  ).join('');
  document.getElementById('stats').style.display = 'block';
}

// ── Preview Table ──────────────────────────────────
function renderPreview(rows, headers) {
  const preview = rows.slice(0, 50);
  const statusColor = { Complete:'badge-green', Incomplete:'badge-yellow', Skipped:'badge-red' };

  let html = '<thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
  for (const row of preview) {
    const isZero = row['Zero Consumption'] === 'YES';
    const rowStyle = isZero ? ' style="background:rgba(251,191,36,.07)"' : '';
    html += `<tr${rowStyle}>` + headers.map(h => {
      let v = row[h] ?? '';
      if (h === 'Order Status') {
        const cls = statusColor[v] || 'badge-blue';
        v = `<span class="badge ${cls}">${v}</span>`;
      } else if (h === 'Zero Consumption' && v === 'YES') {
        v = `<span class="badge badge-yellow">ZERO</span>`;
      } else if (h === 'Read Code Flag' && v === 'REVIEW') {
        v = `<span class="badge badge-yellow">⚠ REVIEW</span>`;
      } else if (h === 'Consumption') {
        const n = parseInt(v, 10);
        if (!isNaN(n) && n < 0) v = `<span style="color:var(--danger);font-weight:600">${v}</span>`;
        else if (!isNaN(n) && n === 0) v = `<span style="color:var(--warn);font-weight:600">${v}</span>`;
      }
      return `<td>${v}</td>`;
    }).join('') + '</tr>';
  }
  html += '</tbody>';
  document.getElementById('previewTable').innerHTML = html;
  document.getElementById('previewWrap').style.display = 'block';
}
</script>
</body>
</html>
