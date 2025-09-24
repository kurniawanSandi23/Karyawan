// --- Utility helpers ---
function tryParseNumber(v) {
  const n = parseFloat(v);
  return isNaN(n) ? v : n;
}

// --- DOM references ---
const fileInput = document.getElementById('file-input');
const btnBrowse = document.getElementById('btn-browse');
const dropArea = document.getElementById('drop-area');
const tableHead = document.getElementById('table-head');
const tableBody = document.getElementById('table-body');
const statTotal = document.getElementById('stat-total');
const statNames = document.getElementById('stat-names');
const statWfh = document.getElementById('stat-wfh');
const statWfo = document.getElementById('stat-wfo');
const colFilter = document.getElementById('col-filter');
const exportCsvBtn = document.getElementById('export-csv');

let rawData = [];
let dataTable = null;
let pieChart = null;
let barChart = null;

// --- Events ---
btnBrowse.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

dropArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropArea.classList.add('border-primary');
});
dropArea.addEventListener('dragleave', () => {
  dropArea.classList.remove('border-primary');
});
dropArea.addEventListener('drop', (e) => {
  e.preventDefault();
  dropArea.classList.remove('border-primary');
  handleFiles(e.dataTransfer.files);
});

colFilter.addEventListener('change', () => refreshTableWithFilter());
exportCsvBtn.addEventListener('click', () => exportCurrentTableCSV());

// --- File handling using SheetJS ---
function handleFiles(files) {
  if (!files || files.length === 0)
    return alert('Pilih file Excel (.xlsx) terlebih dahulu.');

  const f = files[0];
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (json.length === 0)
      return alert('Sheet kosong atau format tidak dikenal.');

    rawData = json.map(r => normalizeRow(r));
    renderAll();
  };
  reader.readAsArrayBuffer(f);
}

function normalizeRow(r) {
  const obj = {};
  Object.keys(r).forEach(k => (obj[k.trim()] = r[k]));
  const keys = Object.keys(obj);

  const mapKey = (candidates) => {
    const lower = keys.map(k => k.toLowerCase());
    for (const c of candidates) {
      const idx = lower.findIndex(x =>
        x === c.toLowerCase() || x.includes(c.toLowerCase())
      );
      if (idx !== -1) return keys[idx];
    }
    return null;
  };

  return {
    raw: r,
    nama: obj[mapKey(['nama','name'])] || '',
    tanggal: obj[mapKey(['tanggal','date'])] || '',
    status: obj[mapKey(['attendance status','status'])] || '',
    lokasi: obj[mapKey(['clock in location','location'])] || '',
    jam: tryParseNumber(obj[mapKey(['total jam kerja','total hours'])])
  };
}

// --- Render UI ---
function renderAll() {
  buildTable();
  updateQuickStats();
  renderCharts();
}

function buildTable() {
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  const displayHeaders = ['nama', 'tanggal', 'status', 'lokasi', 'jam'];
  displayHeaders.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h.toUpperCase();
    tableHead.appendChild(th);
  });

  rawData.forEach(row => {
    const tr = document.createElement('tr');
    displayHeaders.forEach(h => {
      const td = document.createElement('td');
      td.textContent = row[h] ?? '';
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });

  if (dataTable) dataTable.destroy();
  $(document).ready(function () {
    dataTable = $('#data-table').DataTable({
      paging: true,
      searching: true,
      info: true,
      lengthMenu: [5, 10, 25, 50],
      pageLength: 10,
      order: [[1, 'asc']],
      destroy: true
    });
  });
}

function refreshTableWithFilter() {
  const val = colFilter.value;
  if (!dataTable) return;
  if (val === 'all') dataTable.search('').draw();
  else dataTable.column(2).search(val, true, false, true).draw();
}

function updateQuickStats() {
  statTotal.textContent = rawData.length;
  const uniqueNames = new Set(rawData.map(r => r.nama || '').filter(x => x));
  statNames.textContent = uniqueNames.size;
  const countWfh = rawData.filter(r => /wfh/i.test(r.status)).length;
  const countWfo = rawData.filter(r => /wfo/i.test(r.status)).length;
  statWfh.textContent = countWfh;
  statWfo.textContent = countWfo;
}

// --- Charts ---
function renderCharts() {
  const statuses = {};
  const perName = {};
  rawData.forEach(r => {
    const s = r.status || 'Other';
    statuses[s] = (statuses[s] || 0) + 1;
    const n = r.nama || '(unknown)';
    perName[n] = (perName[n] || 0) + 1;
  });

  // pie chart
  const pctx = document.getElementById('pieChart').getContext('2d');
  if (pieChart) pieChart.destroy();
  pieChart = new Chart(pctx, {
    type: 'pie',
    data: { labels: Object.keys(statuses), datasets: [{ data: Object.values(statuses) }] },
    options: { plugins: { legend: { position: 'bottom' } } }
  });

  // bar chart
  const sortedNames = Object.entries(perName).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const bctx = document.getElementById('barChart').getContext('2d');
  if (barChart) barChart.destroy();
  barChart = new Chart(bctx, {
    type: 'bar',
    data: {
      labels: sortedNames.map(x => x[0]),
      datasets: [{ label: 'Jumlah', data: sortedNames.map(x => x[1]) }]
    },
    options: {
      indexAxis: 'y',
      plugins: { legend: { display: false } },
      scales: { x: { beginAtZero: true } }
    }
  });
}

// --- CSV Export ---
function exportCurrentTableCSV() {
  if (!dataTable) return alert('Tidak ada data untuk diexport');
  const filtered = dataTable.rows({ search: 'applied' }).data().toArray();
  if (filtered.length === 0) return alert('Tidak ada baris setelah filter');

  const headers = ['NAMA','TANGGAL','STATUS','LOKASI','JAM'];
  const rows = filtered.map(r =>
    r.map(cell => '"' + String(cell).replace(/"/g, '""') + '"').join(',')
  );
  const csv = [headers.join(',')].concat(rows).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'attendance_export.csv';
  a.click();
  URL.revokeObjectURL(url);
}
