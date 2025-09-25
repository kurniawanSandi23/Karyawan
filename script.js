let rawData = [];
let dataTable = null;
let pieChart = null;
let barChart = null;
let summaryTable = null;

// elements
const fileInput = document.getElementById('file-input');
const btnBrowse = document.getElementById('btn-browse');
const btnClear = document.getElementById('btn-clear');
const dropArea = document.getElementById('drop-area');
const toggleBtn = document.getElementById('toggle-dark');

// helper: normalize column name tolerant to case/space
function findColumnKey(obj, variants = []) {
  const keys = Object.keys(obj || {});
  for (const v of variants) {
    const low = v.toLowerCase();
    const found = keys.find(k => (k || '').toLowerCase().trim() === low);
    if (found) return found;
  }
  // fuzzy fallback: includes
  for (const key of keys) {
    for (const v of variants) {
      if ((key || '').toLowerCase().includes(v.toLowerCase())) return key;
    }
  }
  return null;
}

// File handlers
btnBrowse.addEventListener('click', () => fileInput.click());
btnClear.addEventListener('click', () => {
  rawData = [];
  destroyAllTables();
  document.getElementById('table-head').innerHTML = '';
  document.getElementById('table-body').innerHTML = '';
  updateQuickStats();
});

fileInput.addEventListener('change', (e) => handleFiles(e.target.files));
dropArea.addEventListener('dragover', (e) => { e.preventDefault(); dropArea.classList.add('dragover'); });
dropArea.addEventListener('dragleave', () => dropArea.classList.remove('dragover'));
dropArea.addEventListener('drop', (e) => { e.preventDefault(); dropArea.classList.remove('dragover'); handleFiles(e.dataTransfer.files); });

// filters & export
document.getElementById('col-filter').addEventListener('change', () => refreshTableWithFilter());
document.getElementById('export-csv').addEventListener('click', () => exportCurrentTableCSV());
document.getElementById('download-template').addEventListener('click', () => downloadTemplateCSV());

// File reading
function handleFiles(files) {
  if (!files || files.length === 0) return alert('Pilih file Excel (.xlsx) terlebih dahulu.');
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (json.length === 0) return alert('Sheet kosong atau format tidak dikenal.');
      rawData = json;
      renderAll();
      runIconPulse(); // start pulse after load
    } catch (err) {
      console.error(err);
      alert('Gagal memproses file. Pastikan file Excel valid.');
    }
  };
  reader.readAsArrayBuffer(files[0]);
}

function renderAll() {
  buildTable();
  updateQuickStats();
  renderCharts();
  renderSummaryTable();
  renderGroupSummary('project_name', 'project-table', 'project-body');
  renderGroupSummary('tl', 'tl-table', 'tl-body');
  renderGroupSummary('pm', 'pm-table', 'pm-body');
  renderGroupSummary('gm', 'gm-table', 'gm-body');
}

// Build main table
function buildTable() {
  const tableHead = document.getElementById('table-head');
  const tableBody = document.getElementById('table-body');
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  const headers = Object.keys(rawData[0] || {});
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = prettifyHeader(h);
    tableHead.appendChild(th);
  });

  rawData.forEach(row => {
    const tr = document.createElement('tr');
    headers.forEach(h => {
      const td = document.createElement('td');
      td.textContent = row[h] ?? '';
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });

  if (dataTable) dataTable.destroy();
  dataTable = $('#data-table').DataTable({
    scrollX: true,
    paging: true,
    searching: true,
    info: true,
    pageLength: 10,
    order: [[0, 'asc']],
    columnDefs: [{ targets: '_all', defaultContent: '' }]
  });
}

// prettify header labels
function prettifyHeader(h) {
  if (!h) return '';
  return h.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
}

// Filter
function refreshTableWithFilter() {
  const val = document.getElementById('col-filter').value;
  if (!dataTable) return;

  const headers = Object.keys(rawData[0] || {});
  const statusColKey = findColumnKey(rawData[0], ['attendance status', 'status', 'attendance']);
  const statusColIndex = headers.findIndex(h => h === statusColKey);

  if (statusColIndex === -1) {
    alert("Kolom 'Attendance Status' tidak ditemukan.");
    return;
  }

  if (val === 'all') {
    dataTable.column(statusColIndex).search('').draw();
  } else {
    // use regex to match, case-insensitive
    dataTable.column(statusColIndex).search(val, true, false, true).draw();
  }
}

// Quick stats
function updateQuickStats() {
  document.getElementById('stat-total').textContent = rawData.length || 0;

  const nameKeyCandidate = rawData[0] ? findColumnKey(rawData[0], ['nama', 'name']) : null;
  const namesArr = (rawData || []).map(r => ((nameKeyCandidate && r[nameKeyCandidate]) || '').toString().trim().toUpperCase()).filter(x => x);
  const uniqueNames = new Set(namesArr);
  document.getElementById('stat-names').textContent = uniqueNames.size;

  const statusKey = rawData[0] ? findColumnKey(rawData[0], ['attendance status', 'status', 'attendance']) : null;
  const countWfh = (rawData || []).filter(r => statusKey && /wfh/i.test(r[statusKey] || '')).length;
  const countWfo = (rawData || []).filter(r => statusKey && /wfo/i.test(r[statusKey] || '')).length;
  document.getElementById('stat-wfh').textContent = countWfh;
  document.getElementById('stat-wfo').textContent = countWfo;

  animateCountUp('#stat-total');
  animateCountUp('#stat-names');
  animateCountUp('#stat-wfh');
  animateCountUp('#stat-wfo');
}

// simple count-up animation
function animateCountUp(selector, duration = 700) {
  const el = document.querySelector(selector);
  if (!el) return;
  const end = parseInt(el.textContent || '0', 10) || 0;
  let start = 0;
  const stepTime = Math.max(20, Math.floor(duration / (end || 1)));
  el.textContent = '0';
  const timer = setInterval(() => {
    start += Math.ceil(end / (duration / stepTime));
    if (start >= end) {
      el.textContent = end;
      clearInterval(timer);
    } else {
      el.textContent = start;
    }
  }, stepTime);
}

// Charts
function renderCharts() {
  const statuses = {};
  const perName = {};
  const nameKey = rawData[0] ? findColumnKey(rawData[0], ['nama', 'name']) : null;
  const statusKey = rawData[0] ? findColumnKey(rawData[0], ['attendance status', 'status', 'attendance']) : null;

  rawData.forEach(r => {
    const s = (statusKey && r[statusKey]) ? r[statusKey] : 'Lainnya';
    statuses[s] = (statuses[s] || 0) + 1;
    const n = (nameKey && r[nameKey] ? r[nameKey] : '(unknown)').toString().trim().toUpperCase();
    perName[n] = (perName[n] || 0) + 1;
  });

  // pie
  if (pieChart) pieChart.destroy();
  pieChart = new Chart(document.getElementById('pieChart'), {
    type: 'pie',
    data: { labels: Object.keys(statuses), datasets: [{ data: Object.values(statuses) }] },
    options: { plugins: { legend: { position: 'bottom' } }, animation: { duration: 700 } }
  });

  // bar
  if (barChart) barChart.destroy();
  const sortedNames = Object.entries(perName).sort((a, b) => b[1] - a[1]).slice(0, 10);
  barChart = new Chart(document.getElementById('barChart'), {
    type: 'bar',
    data: { labels: sortedNames.map(x => x[0]), datasets: [{ label: 'Jumlah Absensi', data: sortedNames.map(x => x[1]) }] },
    options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true } }, animation: { duration: 700 } }
  });
}

// Summary per name
function renderSummaryTable() {
  const summary = {};
  const nameKey = rawData[0] ? findColumnKey(rawData[0], ['nama', 'name']) : null;
  const statusKey = rawData[0] ? findColumnKey(rawData[0], ['attendance status', 'status', 'attendance']) : null;

  rawData.forEach(r => {
    const nama = (nameKey && r[nameKey]) ? r[nameKey].toString().trim().toUpperCase() : '(unknown)';
    const status = (statusKey && r[statusKey]) ? r[statusKey].toString().toUpperCase() : 'LAINNYA';
    if (!summary[nama]) summary[nama] = { WFH: 0, WFO: 0, HOLIDAY: 0, LAINNYA: 0 };
    if (status.includes('WFH')) summary[nama].WFH++;
    else if (status.includes('WFO')) summary[nama].WFO++;
    else if (status.includes('HOLIDAY')) summary[nama].HOLIDAY++;
    else summary[nama].LAINNYA++;
  });

  const tbody = document.getElementById('summary-body');
  tbody.innerHTML = '';
  Object.keys(summary).forEach(nama => {
    const row = summary[nama];
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${nama}</td><td>${row.WFH}</td><td>${row.WFO}</td><td>${row.HOLIDAY}</td><td>${row.LAINNYA}</td>`;
    tbody.appendChild(tr);
  });

  if (summaryTable) summaryTable.destroy();
  summaryTable = $('#summary-table').DataTable({
    paging: true,
    searching: true,
    info: false,
    pageLength: 10
  });
}

// Generic group summary
function renderGroupSummary(field, tableId, bodyId) {
  const summary = {};
  const statusKey = rawData[0] ? findColumnKey(rawData[0], ['attendance status', 'status', 'attendance']) : null;

  rawData.forEach(r => {
    const key = (r[field] || '(unknown)').toString().trim() || '(unknown)';
    const status = (statusKey && r[statusKey]) ? r[statusKey].toString().toUpperCase() : 'LAINNYA';
    if (!summary[key]) summary[key] = { WFH: 0, WFO: 0, HOLIDAY: 0, LAINNYA: 0 };
    if (status.includes('WFH')) summary[key].WFH++;
    else if (status.includes('WFO')) summary[key].WFO++;
    else if (status.includes('HOLIDAY')) summary[key].HOLIDAY++;
    else summary[key].LAINNYA++;
  });

  const tbody = document.getElementById(bodyId);
  tbody.innerHTML = '';
  Object.keys(summary).forEach(k => {
    const row = summary[k];
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${k}</td><td>${row.WFH}</td><td>${row.WFO}</td><td>${row.HOLIDAY}</td><td>${row.LAINNYA}</td>`;
    tbody.appendChild(tr);
  });

  if ($.fn.DataTable.isDataTable(`#${tableId}`)) $(`#${tableId}`).DataTable().destroy();
  $(`#${tableId}`).DataTable({
    paging: true,
    searching: true,
    info: false,
    pageLength: 10
  });
}

// Export CSV
function exportCurrentTableCSV() {
  if (!dataTable) return alert('Tidak ada data untuk diexport');
  const filtered = dataTable.rows({ search: 'applied' }).data().toArray();
  if (filtered.length === 0) return alert('Tidak ada baris setelah filter');

  const headers = Object.keys(rawData[0] || {});
  const rows = filtered.map(r => r.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(','));
  const csv = [headers.join(',')].concat(rows).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'attendance_export.csv';
  a.click();
  URL.revokeObjectURL(url);
}

// utility: download a small template CSV file
function downloadTemplateCSV() {
  const sample = ['nama,Attendance Status,project_name,tl,pm,gm', 'BUDI,WFO,Project A,TL1,PM1,GM1', 'SITI,WFH,Project B,TL2,PM2,GM2'].join('\n');
  const blob = new Blob([sample], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'template_attendance.csv';
  a.click();
  URL.revokeObjectURL(url);
}

// dark/light toggle
toggleBtn.addEventListener('click', () => {
  document.body.classList.toggle('dark');
  const label = toggleBtn.querySelector('.toggle-label');
  const icon = toggleBtn.querySelector('i');
  if (document.body.classList.contains('dark')) {
    label.textContent = 'Mode Terang';
    icon.className = 'fa fa-sun me-1';
  } else {
    label.textContent = 'Mode Gelap';
    icon.className = 'fa fa-moon me-1';
  }
});

// cleanup
function destroyAllTables() {
  if (dataTable) { dataTable.destroy(); dataTable = null; }
  if (summaryTable) { summaryTable.destroy(); summaryTable = null; }
  // destroy other DataTables if exist
  ['project-table','tl-table','pm-table','gm-table'].forEach(id => {
    if ($.fn.DataTable.isDataTable(`#${id}`)) $(`#${id}`).DataTable().destroy();
  });
  if (pieChart) { pieChart.destroy(); pieChart = null; }
  if (barChart) { barChart.destroy(); barChart = null; }
}

// icon pulse periodic
let beaconInterval = null;
function runIconPulse() {
  clearInterval(beaconInterval);
  document.querySelectorAll('.icon-beacon').forEach(el => el.classList.add('pulse'));
  beaconInterval = setInterval(() => {
    document.querySelectorAll('.icon-beacon').forEach(el => {
      el.classList.remove('pulse');
      void el.offsetWidth; // reflow
      el.classList.add('pulse');
    });
  }, 4800);
}

// smooth scrolling for nav links
document.querySelectorAll('.nav-link').forEach(a => {
  a.addEventListener('click', (e) => {
    const href = a.getAttribute('href');
    if (!href || !href.startsWith('#')) return;
    e.preventDefault();
    const target = document.querySelector(href);
    if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
});

// auto-run small UI effects on page load
window.addEventListener('load', () => {
  // set initial pulses (even without data)
  document.querySelectorAll('.icon-beacon').forEach(el => el.classList.add('pulse'));
  setTimeout(() => runIconPulse(), 1000);
});

// === Animasi Futuristic Data Stream ala Vecteezy ===
(function() {
  const canvas = document.getElementById('bg-futuristic');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  let width, height;
  let lines = [];
  let particles = [];

  const lineCount = 30;
  const particleCount = 80;
  const maxLineLength = 200;

  function resize() {
    width = canvas.width = window.innerWidth;
    height = canvas.height = window.innerHeight;
  }
  window.addEventListener('resize', resize);
  resize();

  // Line object
  function Line() {
    this.x = Math.random() * width;
    this.y = Math.random() * height;
    this.length = 50 + Math.random() * (maxLineLength - 50);
    this.angle = Math.random() * Math.PI * 2;
    this.speed = 0.2 + Math.random() * 0.5;
    this.opacity = 0.1 + Math.random() * 0.2;
  }
  Line.prototype.update = function() {
    this.x += Math.cos(this.angle) * this.speed;
    this.y += Math.sin(this.angle) * this.speed;
    // jika keluar layar, reset
    if (this.x < -this.length || this.x > width + this.length || this.y < -this.length || this.y > height + this.length) {
      this.x = Math.random() * width;
      this.y = Math.random() * height;
    }
  };
  Line.prototype.draw = function() {
    ctx.strokeStyle = `rgba(0, 150, 255, ${this.opacity})`; // biru-cyber
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(this.x, this.y);
    ctx.lineTo(this.x + Math.cos(this.angle) * this.length, this.y + Math.sin(this.angle) * this.length);
    ctx.stroke();
  };

  // Particle object
  function Particle() {
    this.x = Math.random() * width;
    this.y = Math.random() * height;
    this.radius = 1 + Math.random() * 2;
    this.speedX = (Math.random() - 0.5) * 0.6;
    this.speedY = (Math.random() - 0.5) * 0.6;
    this.opacity = 0.3 + Math.random() * 0.3;
  }
  Particle.prototype.update = function() {
    this.x += this.speedX;
    this.y += this.speedY;
    if (this.x < 0 || this.x > width) this.speedX *= -1;
    if (this.y < 0 || this.y > height) this.speedY *= -1;
  };
  Particle.prototype.draw = function() {
    ctx.fillStyle = `rgba(0, 150, 255, ${this.opacity})`;
    ctx.beginPath();
    ctx.arc(this.x, this.y, this.radius, 0, Math.PI * 2);
    ctx.fill();
  };

  // inisialisasi
  for (let i = 0; i < lineCount; i++) {
    lines.push(new Line());
  }
  for (let j = 0; j < particleCount; j++) {
    particles.push(new Particle());
  }

  function animate() {
    ctx.clearRect(0, 0, width, height);

    // draw lines
    lines.forEach(line => {
      line.update();
      line.draw();
    });

    // draw particles
    particles.forEach(p => {
      p.update();
      p.draw();
    });

    requestAnimationFrame(animate);
  }

  animate();
})();
