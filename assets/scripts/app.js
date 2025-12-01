// assets/scripts/app.js (final updated)
// Features:
// - Auto-load .xlsx in /data via optional files.json manifest or fallback list
// - Robust column header matching & alias mapping
// - Editable table with LocalStorage autosave
// - Multi-pack & Normal cards + clickable filters
// - All controls (container/status/search) filter summary, charts, multipack/normal counts and table
// - Export per-file: Data + Analytics sheets, downloaded with timestamped filename
// - Auto-export all loaded files on unload (browser download)

// CONFIG: manifest path and fallback file URLs (adjust if needed)
const MANIFEST_PATH = '/data/files.json';
const FALLBACK_FILES = ['/data/NOV_Shippment_Boxes_R4.xlsx'];

// expected canonical columns (we'll normalize/match to these)
const EXPECTED_COLS = ['NO', 'ContainerNum', 'BoxNum', 'Container', 'BoxName', 'ItemCount', 'Kits', 'Factory', 'REMARKS'];

// Alias map: common alternative header names -> canonical column
const COLUMN_ALIASES = {
  // BoxNum variations
  'basebox': 'BoxNum',
  'box no': 'BoxNum',
  'boxno': 'BoxNum',
  'box number': 'BoxNum',
  'box_number': 'BoxNum',
  // REMARKS variations
  'remark': 'REMARKS',
  'remarks': 'REMARKS',
  'status': 'REMARKS',
  'inspection status': 'REMARKS'
  // add more aliases if you find new variants
};

// App state
const files = {}; // key -> { name, workbook, sheetName, rows, columns }
let activeKey = null;

// DOM elements
const fileInput = document.getElementById('fileInput');
const exportBtn = document.getElementById('exportCsv');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const rowsCount = document.getElementById('rowsCount');
const containerFilter = document.getElementById('containerFilter');
const statusFilter = document.getElementById('statusFilter');
const searchInput = document.getElementById('searchInput');
const summaryWrap = document.getElementById('summaryWrap');
const multipackCard = document.getElementById('multipackCard');
const normalPackCard = document.getElementById('normalPackCard');
const multipackCountEl = document.getElementById('multipackCount');
const normalCountEl = document.getElementById('normalCount');

const factoryFilter = document.getElementById('factoryFilter');

// files select (insert at left of actions if exists)
const filesSelect = document.createElement('select');
filesSelect.id = 'filesSelect';
const actionsDiv = document.querySelector('.actions');
if (actionsDiv) actionsDiv.insertBefore(filesSelect, actionsDiv.firstChild);

// charts
let progressChart = null;
let containerBarChart = null;

// ---------------------- Utilities ----------------------

function makeKey(name) {
  return String(name).replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-\.]/g, '').toLowerCase();
}

// Normalizes header strings to a canonical comparator form
function normalizeKey(str) {
  return String(str || '')
    .trim()
    .replace(/\u00A0/g, '')        // non-breaking spaces
    .replace(/[^\w]/g, '')         // remove all non-word chars (keeps letters+digits+underscore)
    .toLowerCase();
}

// Return canonical column name for a given header key (or null)
function canonicalColumnFor(header) {
  const nk = normalizeKey(header);
  if (!nk) return null;
  // direct match against expected columns
  for (const c of EXPECTED_COLS) {
    if (normalizeKey(c) === nk) return c;
  }
  // check alias map
  if (COLUMN_ALIASES[nk]) return COLUMN_ALIASES[nk];
  // also allow plural variant (e.g., remarks -> REMARKS)
  if (nk.endsWith('s')) {
    const singular = nk.slice(0, -1);
    if (COLUMN_ALIASES[singular]) return COLUMN_ALIASES[singular];
  }
  return null;
}

function isCompleted(remarks) {
  if (remarks === null || remarks === undefined) return false;
  return /done/i.test(String(remarks));
}

function nowTimestampForName() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}-${pad(d.getMinutes())}`;
}

// LocalStorage helpers
function saveFileToLocalStorage(key) {
  try {
    localStorage.setItem(`inspection_${key}`, JSON.stringify(files[key].rows || []));
  } catch (e) {
    console.warn('localStorage save failed', e);
  }
}
function loadFileFromLocalStorage(key) {
  try {
    const raw = localStorage.getItem(`inspection_${key}`);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      files[key].rows = parsed;
      return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

// ---------------------- File loading ----------------------

// read uploaded File object (from input)
function readFileAndAdd(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const key = makeKey(file.name);
      addWorkbook(key, file.name, wb);
    } catch (err) {
      console.error('Failed to parse uploaded file', err);
      alert('Failed to parse uploaded file: ' + file.name + '\n' + err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// fetch a URL and add workbook
function fetchXlsxAndAdd(url, nameHint) {
  return fetch(url)
    .then(r => {
      if (!r.ok) throw new Error(`fetch ${url} failed (${r.status})`);
      return r.arrayBuffer();
    })
    .then(ab => {
      const data = new Uint8Array(ab);
      const wb = XLSX.read(data, { type: 'array' });
      const key = makeKey(nameHint || url.split('/').pop());
      addWorkbook(key, nameHint || url.split('/').pop(), wb);
    })
    .catch(err => {
      console.info('Could not fetch', url, err);
    });
}

// Add workbook to files map and normalize rows with robust matching
function addWorkbook(key, name, workbook) {
  if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) return;

  const targetSheet = workbook.SheetNames.find(s => s.toLowerCase() === 'f200_boxes') || workbook.SheetNames[0];
  const ws = workbook.Sheets[targetSheet];
  const rowsRaw = XLSX.utils.sheet_to_json(ws, { defval: '' });

  // If sheet empty, still create empty rows array
  const normalized = rowsRaw.map(r => {
    const nr = {};
    // First, map any headers to canonical columns
    const headerKeys = Object.keys(r);
    const headerMap = {}; // normalized -> original header
    headerKeys.forEach(h => headerMap[normalizeKey(h)] = h);

    // For each expected column, find best matching header
    EXPECTED_COLS.forEach(col => {
      // try exact normalized match
      const desiredNorm = normalizeKey(col);
      // check direct
      let foundOriginal = null;
      if (headerMap[desiredNorm]) foundOriginal = headerMap[desiredNorm];
      else {
        // try aliases
        const aliasCandidate = Object.entries(COLUMN_ALIASES).find(([k, v]) => v === col);
        if (aliasCandidate) {
          const aliasNorm = aliasCandidate[0];
          if (headerMap[aliasNorm]) foundOriginal = headerMap[aliasNorm];
        }
      }
      // if not found yet, try any header that normalizes equal to desiredNorm
      if (!foundOriginal) {
        const matchHeader = headerKeys.find(h => normalizeKey(h) === desiredNorm);
        if (matchHeader) foundOriginal = matchHeader;
      }
      // if still not found, try other heuristics: check if header normalized contains desiredNorm or vice versa
      if (!foundOriginal) {
        const matchHeader = headerKeys.find(h => {
          const nh = normalizeKey(h);
          return nh.includes(desiredNorm) || desiredNorm.includes(nh);
        });
        if (matchHeader) foundOriginal = matchHeader;
      }

      // Lastly check aliases map explicitly
      if (!foundOriginal) {
        // iterate aliases keys
        for (const ak of Object.keys(COLUMN_ALIASES)) {
          if (COLUMN_ALIASES[ak] === col && headerMap[ak]) {
            foundOriginal = headerMap[ak];
            break;
          }
        }
      }

      nr[col] = foundOriginal ? r[foundOriginal] : '';
    });

    // keep extras unchanged (so we don't lose extra columns)
    Object.keys(r).forEach(k => {
      const can = canonicalColumnFor(k);
      // if this header didn't map to an EXPECTED_COLS canonical, keep it as an extra
      if (!can) {
        nr[k] = r[k];
      }
    });

    return nr;
  });

  // If there were no rows (empty sheet), create an empty rows array but keep columns from EXPECTED_COLS
  const columns = (normalized[0] ? Object.keys(normalized[0]) : EXPECTED_COLS.slice());

  files[key] = {
    name,
    workbook,
    sheetName: targetSheet,
    rows: normalized,
    columns
  };

  // add to UI select
  const opt = document.createElement('option');
  opt.value = key;
  opt.textContent = name;
  filesSelect.appendChild(opt);

  // load autosaved edits (if present)
  // loadFileFromLocalStorage(key);

  // if first file, activate it
  if (!activeKey) setActiveFile(key);

  // update multipack/normal counts for active file if it's the one we just added
  if (activeKey === key) {
    // trigger a filtered render to update UI
    renderFilteredAndLive();

    // --- NEW: force initial multipack/normal counts ---
    // const allRows = files[activeKey].rows || [];
    // updateMultipackNormalCounts(allRows);
    // updateStatusCounts(allRows);

  }
}

// ---------------------- UI & Rendering ----------------------

function setActiveFile(key) {
  if (!files[key]) return;
  activeKey = key;
  filesSelect.value = key;
  buildFactoryFilter();
  buildContainerFilter();
  renderFilteredAndLive();
}

function buildFactoryFilter() {
  if (!activeKey) return;

  const factories = new Set();
  files[activeKey].rows.forEach(r => {
    const f = String(r.Factory ?? '').trim();
    if (f) factories.add(f);
  });

  factoryFilter.innerHTML =
    `<option value="all">All</option>` +
    [...factories].sort().map(f => `<option value="${f}">${f}</option>`).join('');
}

function buildContainerFilter() {
  if (!activeKey) return;
  const set = new Set();
  files[activeKey].rows.forEach(r => {
    const v = String(r.ContainerNum ?? '').trim();
    if (v) set.add(v);
  });

  const containerArray = Array.from(set).map(c => {
    // Try to parse number, fallback to string
    const n = parseInt(c, 10);
    return { original: c, number: isNaN(n) ? Infinity : n };
  });

  // Sort numerically
  containerArray.sort((a, b) => a.number - b.number);

  const opts = ['<option value="all">All</option>']
    .concat(containerArray.map(c => `<option value="${c.original}">${c.original}</option>`))
    .join('');

  containerFilter.innerHTML = opts;
}

// Main filtering function: returns filtered rows and renders everything based on filtered rows
function renderFilteredAndLive() {
  if (!activeKey) return;
  const allRows = files[activeKey].rows || [];
  const fFactory = factoryFilter.value || 'all';
  const fContainer = (containerFilter.value || 'all');
  const fStatus = (statusFilter.value || 'all');
  const q = (searchInput.value || '').trim().toLowerCase();

  const filtered = allRows.filter(r => {
    // Factory filter
    if (fFactory !== 'all' && String(r.Factory ?? '') !== fFactory)
      return false;

    // container filter
    if (fContainer !== 'all') {
      if (String(r.ContainerNum ?? '') !== fContainer) return false;
    }

    // status filter
    if (fStatus !== 'all') {
      const rem = String(r.REMARKS ?? '').toLowerCase();
      const isDone = /done/i.test(rem);
      const isInProg = /(in progress|inprogress)/i.test(rem);
      const isNotStarted = rem.trim() === '' || /(not started|n\/a|na)/i.test(rem);
      if (fStatus === 'Finished' && !isDone) return false;
      if (fStatus === 'In Progress' && !isInProg) return false;
      if (fStatus === 'Not Started' && !isNotStarted) return false;
      if (fStatus === 'Remaining' && isDone) return false;
    }

    // search filter
    if (q) {
      const hay = Object.values(r).join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }

    return true;
  });

  // render UI using filtered rows
  renderTable(filtered);
  renderSummary(filtered);
  renderCharts(filtered);
  updateMultipackNormalCounts(filtered);

  // rows count = filtered count
  rowsCount.textContent = filtered.length;
}


function renderTable(rows) {
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';
  if (!activeKey) return;
  const cols = files[activeKey].columns || EXPECTED_COLS;

  // --- Build table header ---
  const trh = document.createElement('tr');
  cols.forEach(c => {
    const th = document.createElement('th');
    th.textContent = c;
    trh.appendChild(th);
  });
  tableHead.appendChild(trh);

  // --- Build table body ---
  rows.forEach(r => {
    const tr = document.createElement('tr');

    cols.forEach(c => {
      const td = document.createElement('td');

      if (c === 'REMARKS') {
        // --- Create dropdown for REMARKS ---
        const select = document.createElement('select');
        const options = ['', 'Done', 'In Progress'];
        options.forEach(opt => {
          const el = document.createElement('option');
          el.value = opt;
          el.textContent = opt;
          if ((r[c] ?? '').toLowerCase() === opt.toLowerCase()) el.selected = true;
          select.appendChild(el);
        });

        // --- Apply initial color ---
        if ((r[c] ?? '').toLowerCase() === 'done') select.style.backgroundColor = '#66FF66'; // green
        else if ((r[c] ?? '').toLowerCase() === 'in progress') select.style.backgroundColor = '#FFF867'; // yellow
        else select.style.backgroundColor = '';

        select.addEventListener('change', () => {
          const allRows = files[activeKey].rows;
          const keyProps = ['NO', 'BoxNum', 'ContainerNum'];
          const identifier = keyProps.map(k => r[k] ?? '').join('||');
          let idx = allRows.findIndex(rr => keyProps.map(k => rr[k] ?? '').join('||') === identifier);
          if (idx === -1) {
            idx = allRows.findIndex(rr => String(rr.BoxNum ?? '') === String(r.BoxNum ?? '') &&
              String(rr.ContainerNum ?? '') === String(r.ContainerNum ?? ''));
          }
          if (idx >= 0) {
            files[activeKey].rows[idx][c] = select.value;
            saveFileToLocalStorage(activeKey);

            // --- Update color immediately ---
            if (select.value.toLowerCase() === 'done') select.style.backgroundColor = '#66FF66';
            else if (select.value.toLowerCase() === 'in progress') select.style.backgroundColor = '#FFF867';
            else select.style.backgroundColor = '';

            renderFilteredAndLive();
          }
        });

        td.innerHTML = '';
        td.appendChild(select);
      } else {
        // --- Regular editable cells ---
        td.contentEditable = true;
        td.spellcheck = false;
        td.textContent = r[c] ?? '';

        td.addEventListener('input', () => {
          const allRows = files[activeKey].rows;
          const keyProps = ['NO', 'BoxNum', 'ContainerNum'];
          const identifier = keyProps.map(k => r[k] ?? '').join('||');
          let idx = allRows.findIndex(rr => keyProps.map(k => rr[k] ?? '').join('||') === identifier);
          if (idx === -1) {
            idx = allRows.findIndex(rr => String(rr.BoxNum ?? '') === String(r.BoxNum ?? '') &&
              String(rr.ContainerNum ?? '') === String(r.ContainerNum ?? ''));
          }
          if (idx >= 0) {
            files[activeKey].rows[idx][c] = td.textContent;
            saveFileToLocalStorage(activeKey);
            renderFilteredAndLive();
          }
        });
      }

      tr.appendChild(td);
    });

    tableBody.appendChild(tr);
  });
}


const applyRemarkSelect = document.getElementById('applyRemark');
const applyAllBtn = document.getElementById('applyAllBtn');

applyAllBtn.addEventListener('click', () => {
  if (!activeKey) return;
  const val = applyRemarkSelect.value;
  const allRows = files[activeKey].rows || [];

  const fFactory = factoryFilter.value || 'all';
  const fContainer = containerFilter.value || 'all';
  const fStatus = statusFilter.value || 'all';
  const q = (searchInput.value || '').trim().toLowerCase();

  allRows.forEach(r => {
    let visible = true;

    // Factory
    if (fFactory !== 'all' && String(r.Factory ?? '') !== fFactory)
      visible = false;

    // Container
    if (fContainer !== 'all' && String(r.ContainerNum ?? '') !== fContainer)
      visible = false;

    // Status: SAME LOGIC as renderFilteredAndLive()
    const rem = String(r.REMARKS ?? '').toLowerCase();
    const isDone = /done/i.test(rem);
    const isInProg = /(in progress|inprogress)/i.test(rem);
    const isNotStarted = rem.trim() === '' || /(not started|n\/a|na)/i.test(rem);

    if (fStatus === 'Finished' && !isDone) visible = false;
    if (fStatus === 'In Progress' && !isInProg) visible = false;
    if (fStatus === 'Not Started' && !isNotStarted) visible = false;
    if (fStatus === 'Remaining' && isDone) visible = false;

    // Search filter
    if (q && !Object.values(r).join(' ').toLowerCase().includes(q))
      visible = false;

    // Apply update only to visible rows
    if (visible) r.REMARKS = val;
  });

  saveFileToLocalStorage(activeKey);
  renderFilteredAndLive();
});


function classifyStatus(remarks) {
  const rem = String(remarks ?? '').trim().toLowerCase();

  // Completed
  if (/done/i.test(rem)) return 'Completed';

  // In Progress
  if (/in\s*progress/i.test(rem)) return 'In Progress';

  // Not Started
  if (rem === '' || /^(n\/a|na|not started)$/i.test(rem)) return 'Not Started';

  return 'Not Started';
}

function renderSummary(rows) {
  const total = rows.length;

  let completed = 0;
  let inProgress = 0;
  let notStarted = 0;

  rows.forEach(r => {
    const s = classifyStatus(r.REMARKS);
    if (s === 'Completed') completed++;
    else if (s === 'In Progress') inProgress++;
    else if (s === 'Not Started') notStarted++;
  });

  // Remaining = In Progress + Not Started
  const remaining = inProgress + notStarted;

  const percent = total === 0 ? 0 : Math.round((completed / total) * 100);


  summaryWrap.innerHTML = `
  <div class="card">
    <strong>Total Boxes</strong>
    <div class="big">${total}</div>
    <div class="muted">All rows</div>
  </div>

  <div class="card">
    <strong>Completed</strong>
    <div class="big">${completed}</div>
    <div class="muted">Finished (${percent}%)</div>
  </div>

  <div class="card">
    <strong>In Progress</strong>
    <div class="big">${inProgress}</div>
    <div class="muted">Under inspection</div>
  </div>

  <div class="card">
    <strong>Not Started</strong>
    <div class="big">${notStarted}</div>
    <div class="muted">Not yet inspected</div>
  </div>

  <div class="card">
    <strong>Remaining</strong>
    <div class="big">${remaining}</div>
    <div class="muted">In Progress + Not Started</div>
  </div>
`;
}

// rows parameter is filtered rows for counting multipack/normal
function updateMultipackNormalCounts(rows) {
  if (!rows) { multipackCountEl.textContent = '0'; normalCountEl.textContent = '0'; return; }
  let multi = 0, normal = 0;
  rows.forEach(r => {
    const ic = Number(r.ItemCount ?? r.Itemcount ?? 0) || 0;
    if (ic > 1) multi++;
    else normal++;
  });
  multipackCountEl.textContent = multi;
  normalCountEl.textContent = normal;
}

// clicking cards: apply ItemCount filters on top of current container/status/search filters
if (multipackCard) multipackCard.addEventListener('click', () => {
  if (!activeKey) return;
  const allRows = files[activeKey].rows || [];
  const fContainer = (containerFilter.value || 'all');
  const fStatus = (statusFilter.value || 'all');
  const q = (searchInput.value || '').trim().toLowerCase();
  const filtered = allRows.filter(r => {
    if (fContainer !== 'all' && String(r.ContainerNum ?? '') !== fContainer) return false;
    if (fStatus !== 'all') {
      const rem = String(r.REMARKS ?? '').toLowerCase();
      const isDone = /done/i.test(rem);
      const isInProg = /(in progress|inprogress)/i.test(rem);
      const isNotStarted = rem.trim() === '' || /(not started|n\/a|na)/i.test(rem);
      if (fStatus === 'Finished' && !isDone) return false;
      if (fStatus === 'In Progress' && !isInProg) return false;
      if (fStatus === 'Not Started' && !isNotStarted) return false;
      if (fStatus === 'Remaining' && isDone) return false;
    }
    if (q) {
      const hay = Object.values(r).join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    // multipack only
    const ic = Number(r.ItemCount ?? r.Itemcount ?? 0) || 0;
    return ic > 1;
  });
  renderTable(filtered);
  renderSummary(filtered);
  renderCharts(filtered);
  updateMultipackNormalCounts(filtered);
  rowsCount.textContent = filtered.length;
});

if (normalPackCard) normalPackCard.addEventListener('click', () => {
  if (!activeKey) return;
  const allRows = files[activeKey].rows || [];
  const fContainer = (containerFilter.value || 'all');
  const fStatus = (statusFilter.value || 'all');
  const q = (searchInput.value || '').trim().toLowerCase();
  const filtered = allRows.filter(r => {
    if (fContainer !== 'all' && String(r.ContainerNum ?? '') !== fContainer) return false;
    if (fStatus !== 'all') {
      const rem = String(r.REMARKS ?? '').toLowerCase();
      const isDone = /done/i.test(rem);
      const isInProg = /(in progress|inprogress)/i.test(rem);
      const isNotStarted = rem.trim() === '' || /(not started|n\/a|na)/i.test(rem);
      if (fStatus === 'Finished' && !isDone) return false;
      if (fStatus === 'In Progress' && !isInProg) return false;
      if (fStatus === 'Not Started' && !isNotStarted) return false;
      if (fStatus === 'Remaining' && isDone) return false;
    }
    if (q) {
      const hay = Object.values(r).join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    // normal only
    const ic = Number(r.ItemCount ?? r.Itemcount ?? 0) || 0;
    return ic === 1;
  });
  renderTable(filtered);
  renderSummary(filtered);
  renderCharts(filtered);
  updateMultipackNormalCounts(filtered);
  rowsCount.textContent = filtered.length;
});




function renderCharts(rows) {
  // compute per container
  const byContainer = {};
  rows.forEach(r => {
    const cont = String(r.ContainerNum ?? 'NA');
    if (!byContainer[cont]) byContainer[cont] = { total: 0, finished: 0 };
    byContainer[cont].total++;
    if (isCompleted(r.REMARKS)) byContainer[cont].finished++;
  });

  const labels = Object.keys(byContainer).sort();
  const finishedData = labels.map(l => byContainer[l].finished);
  const remainingData = labels.map(l => byContainer[l].total - byContainer[l].finished);

  // BAR CHART
  const ctxBar = document.getElementById('boxesByContainerChart').getContext('2d');
  if (containerBarChart) containerBarChart.destroy();
  containerBarChart = new Chart(ctxBar, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Finished', data: finishedData },
        { label: 'Remaining', data: remainingData }
      ]
    },
    options: {
      plugins: { legend: { position: 'bottom' } },
      responsive: true,
      scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } }
    }
  });

  // FORCE SAME HEIGHTS
  const MAX_HEIGHT = 100;

  const progressCtx = document.getElementById('progressChart').getContext('2d');
  const barCtx = document.getElementById('boxesByContainerChart').getContext('2d');

  progressCtx.canvas.parentElement.style.height = MAX_HEIGHT + '%';
  progressCtx.canvas.style.maxHeight = MAX_HEIGHT + '%';

  barCtx.canvas.parentElement.style.height = MAX_HEIGHT + '%';
  barCtx.canvas.style.maxHeight = MAX_HEIGHT + '%';

  // DONUT CHART
  const totals = labels.reduce((acc, l) => {
    acc.total += byContainer[l].total;
    acc.finished += byContainer[l].finished;
    return acc;
  }, { total: 0, finished: 0 });

  const totalRemaining = Math.max(0, totals.total - totals.finished);

  const ctxDonut = document.getElementById('progressChart').getContext('2d');
  if (progressChart) progressChart.destroy();

  progressChart = new Chart(ctxDonut, {
    type: 'doughnut',
    data: {
      labels: ['Finished', 'Remaining'],
      datasets: [
        { data: [totals.finished, totalRemaining] }
      ]
    },
    options: {
      plugins: {
        legend: { position: 'bottom' },
        datalabels: {
          color: '#fff',
          font: { weight: 'bold', size: 16 },
          formatter: v => v
        }
      },
      cutout: '60%',
      responsive: true,
      maintainAspectRatio: false
    },
    plugins: [ChartDataLabels]
  });
}


// ---------------------- Export / Download ----------------------
function exportWorkbookWithAnalytics(key) {
  const entry = files[key];
  if (!entry) return;

  // -----------------------------
  // 1) DATA SHEET
  // -----------------------------
  const dataRows = entry.rows.map(r => {
    const out = {};
    EXPECTED_COLS.forEach(c => out[c] = r[c] ?? "");
    Object.keys(r).forEach(k => { if (!EXPECTED_COLS.includes(k)) out[k] = r[k]; });
    return out;
  });
  const wsData = XLSX.utils.json_to_sheet(dataRows);

  // -----------------------------
  // ANALYTICS BUILDER
  // -----------------------------
  function buildAnalytics(rows) {
    const byContainer = {};

    rows.forEach(r => {
      const cont = String(r.ContainerNum ?? "NA");
      if (!byContainer[cont]) byContainer[cont] = { total: 0, finished: 0 };
      byContainer[cont].total++;
      if (isCompleted(r.REMARKS)) byContainer[cont].finished++;
    });

    const out = [];
    let total = 0, finished = 0;

    Object.keys(byContainer).sort().forEach(cont => {
      const v = byContainer[cont];
      const remaining = v.total - v.finished;
      const pct = v.total === 0 ? 0 : Math.round((v.finished / v.total) * 100);

      out.push({
        Container: cont,
        TotalBoxes: v.total,
        Finished: v.finished,
        Remaining: remaining,
        CompletionPercent: pct + "%"
      });

      total += v.total;
      finished += v.finished;
    });

    // ALL ROW
    const pctAll = total === 0 ? 0 : Math.round((finished / total) * 100);
    out.push({
      Container: "ALL",
      TotalBoxes: total,
      Finished: finished,
      Remaining: total - finished,
      CompletionPercent: pctAll + "%"
    });

    return XLSX.utils.json_to_sheet(out);
  }

  const wsAnalytics = buildAnalytics(entry.rows);

  // -----------------------------
  // FACTORY ORDER
  // -----------------------------
  const factoryOrder = ["F200", "F100", "AIO"];
  const factories = [...new Set(entry.rows.map(r => r.Factory || "UNKNOWN"))];

  // -----------------------------
  // SUMMARY SHEET (UPDATED)
  // -----------------------------
  function buildSummarySheet() {
    const summaryData = [];

    let allTotal = 0,
      allFinished = 0;

    factoryOrder.forEach(fac => {
      const filtered = entry.rows.filter(r => r.Factory === fac);
      if (filtered.length === 0) return;

      const total = filtered.length;
      const finished = filtered.filter(r => isCompleted(r.REMARKS)).length;
      const pct = total === 0 ? 0 : Math.round((finished / total) * 100);

      summaryData.push({
        Factory: fac,
        TotalBoxes: total,
        Completed: finished,
        Remaining: total - finished,
        CompletionPercent: pct + "%"
      });

      allTotal += total;
      allFinished += finished;
    });

    // ---- ALL FACTORIES ROW ----
    const pctAll = allTotal === 0 ? 0 : Math.round((allFinished / allTotal) * 100);

    summaryData.push({
      Factory: "ALL",
      TotalBoxes: allTotal,
      Completed: allFinished,
      Remaining: allTotal - allFinished,
      CompletionPercent: pctAll + "%"
    });

    // Build sheet
    const ws = XLSX.utils.json_to_sheet(summaryData, { origin: "A2" });

    // ---- Merged Title Cell ----
    ws["A1"] = { t: "s", v: "Shipment Inspection Summary" };
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }];

    // ---- Column Widths ----
    ws["!cols"] = [
      { wch: 12 },  // Factory
      { wch: 12 },  // TotalBoxes
      { wch: 12 },  // Completed
      { wch: 12 },  // Remaining
      { wch: 18 }   // CompletionPercent
    ];

    return ws;
  }

  const wsSummary = buildSummarySheet();

  // -----------------------------
  // BUILD WORKBOOK
  // -----------------------------
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsData, "Data");
  XLSX.utils.book_append_sheet(wb, wsAnalytics, "Analytics");

  factories.forEach(fac => {
    const filtered = entry.rows.filter(r => r.Factory === fac);
    const ws = buildAnalytics(filtered);
    const safe = `Analytics_${fac}`.replace(/[^A-Za-z0-9_]/g, "");
    XLSX.utils.book_append_sheet(wb, ws, safe);
  });

  XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");

  // -----------------------------
  // SAVE FILE
  // -----------------------------
  const outName =
    `${entry.name.replace(/\s+/g, '_')}_${nowTimestampForName()}.xlsx`;
  XLSX.writeFile(wb, outName);
}


// export active
exportBtn.addEventListener('click', () => {
  if (!activeKey) { alert('No active file'); return; }
  exportWorkbookWithAnalytics(activeKey);
});


// ---------------------- Autosave / Autoexport on unload ----------------------

function autoExportAllAndPersist() {
  Object.keys(files).forEach(k => {
    try { saveFileToLocalStorage(k); } catch (e) { }
    try { exportWorkbookWithAnalytics(k); } catch (e) { console.warn('export failed for', k, e); }
  });
}
window.addEventListener('beforeunload', () => {
  autoExportAllAndPersist();
});

// ---------------------- Event wiring ----------------------

fileInput.addEventListener('change', e => {
  const list = Array.from(e.target.files || []);
  list.forEach(f => readFileAndAdd(f));
  e.target.value = '';
});
filesSelect.addEventListener('change', e => {
  const k = e.target.value;
  if (k) setActiveFile(k);
});
factoryFilter.addEventListener('change', renderFilteredAndLive);
containerFilter.addEventListener('change', renderFilteredAndLive);
statusFilter.addEventListener('change', renderFilteredAndLive);
searchInput.addEventListener('input', renderFilteredAndLive);

// ---------------------- Auto-load on init ----------------------

(function init() {
  // placeholder for filesSelect
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = '-- Select/Open File --';
  placeholder.disabled = true;
  placeholder.selected = true;
  filesSelect.appendChild(placeholder);

  // try manifest first
  fetch(MANIFEST_PATH).then(r => {
    if (!r.ok) throw new Error('no manifest');
    return r.json();
  }).then(list => {
    if (!Array.isArray(list)) throw new Error('manifest invalid');
    list.forEach(fname => fetchXlsxAndAdd(`/data/${fname}`, fname));
  }).catch(() => {
    // fallback to known list
    FALLBACK_FILES.forEach(p => fetchXlsxAndAdd(p));
  });
})();
