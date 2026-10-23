// Student Attendance Intelligence Dashboard - Robust Production Script
let uploadedRows = []; // Array of row objects from the uploaded Excel
let tableColumns = []; // Array of column header names
let currentFileName = null;
let currentWorkbook = null; // Store loaded workbook for sheet switcher
let detectedCols = { student: null, status: null, batch: null, faculty: null, date: null };
let statusChartInstance = null;
let batchChartInstance = null;

// Course code mapping
const COURSE_MAPPING = {
  'HA': 'Hospital Administration',
  'DM': 'Digital Marketing',
  'FS': 'Full Stack Web Development',
  'AI': 'Artificial Intelligence & Data Science'
};

function getCourseName(batch) {
  if (!batch) return 'General Course';
  const clean = String(batch).trim().toUpperCase();
  const p2 = clean.substring(0, 2);
  return COURSE_MAPPING[p2] || `${p2} Course`;
}

// Global trigger to open file dialog reliably
function triggerFileInput() {
  const fi = document.getElementById('excelFileInput');
  if (fi) {
    fi.value = null;
    fi.click();
  }
}

// Native input change handler
function handleFileInputChange(input) {
  if (input && input.files && input.files[0]) {
    handleFileUpload(input.files[0]);
  }
}

// Intelligent Header Row Detection
function findBestHeaderRow(aoa) {
  let bestIdx = 0;
  let maxScore = -1;

  for (let i = 0; i < Math.min(25, aoa.length); i++) {
    const row = aoa[i];
    if (!row || !Array.isArray(row)) continue;
    const cells = row.map(c => String(c !== undefined && c !== null ? c : '').toLowerCase().trim());
    let score = 0;

    cells.forEach(c => {
      if (!c) return;
      if (c === 'student name' || c === 'candidate name' || c === 'student' || c === 'candidate') score += 15;
      else if (c.includes('student') || c.includes('candidate')) score += 10;
      else if (c === 'name' && !c.includes('holiday') && !c.includes('course') && !c.includes('day')) score += 8;

      if (c === 'status' || c === 'attendance status' || c === 'p/a' || c === 'present/absent') score += 15;
      else if (c === 'attendance') score += 10;

      if (c === 'batch' || c === 'batch code' || c === 'class' || c === 'section') score += 10;
      else if (c.includes('batch') || c.includes('class')) score += 8;

      if (c === 'date' || c === 'attendance date' || c === 'day' || c === 'month') score += 8;
      else if (c.includes('date')) score += 6;

      if (c === 'taken by' || c === 'recorded by' || c === 'faculty' || c === 'faculty name' || c === 'teacher') score += 8;
      else if (c.includes('faculty') || c.includes('taken by') || c.includes('teacher')) score += 6;

      if (c === 'sl no' || c === 'sl.no' || c === 's.no' || c === 'id' || c === 'roll no') score += 6;
    });

    const filledCount = cells.filter(Boolean).length;
    if (filledCount >= 3) score += filledCount * 2;

    if (score > maxScore) {
      maxScore = score;
      bestIdx = i;
    }
  }
  return bestIdx;
}

// Intelligent Sheet Selection for Multi-sheet Workbooks
function findBestSheet(workbook) {
  let bestSheetName = workbook.SheetNames[0];
  let bestRows = [];
  let bestScore = -999999;

  for (let s = 0; s < workbook.SheetNames.length; s++) {
    const sName = workbook.SheetNames[s];
    const sheet = workbook.Sheets[sName];
    if (!sheet) continue;

    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (!aoa || aoa.length === 0) continue;

    const headerIdx = findBestHeaderRow(aoa);
    const rawJson = XLSX.utils.sheet_to_json(sheet, { range: headerIdx, defval: '' });

    // Filter out completely blank rows
    const validRows = rawJson.filter(r => {
      return Object.values(r).some(v => String(v !== undefined && v !== null ? v : '').trim().length > 0);
    });

    if (validRows.length === 0) continue;

    const headers = Object.keys(validRows[0] || {}).map(h => h.toLowerCase().trim());
    let score = validRows.length; // 1 point per row

    // Bonus points for key attendance columns
    const hasStudent = headers.some(h => h.includes('student') || (h.includes('name') && !h.includes('holiday') && !h.includes('course')));
    const hasStatus = headers.some(h => h === 'status' || h.includes('attendance') || h === 'p/a');
    const hasBatch = headers.some(h => h.includes('batch') || h.includes('class'));
    const hasFaculty = headers.some(h => h.includes('taken by') || h.includes('faculty') || h.includes('teacher'));
    const hasDate = headers.some(h => h.includes('date'));

    if (hasStudent) score += 600;
    if (hasStatus) score += 600;
    if (hasBatch) score += 400;
    if (hasFaculty) score += 300;
    if (hasDate) score += 300;

    // Sheet Name heuristics
    const sLower = sName.toLowerCase();
    if (sLower.includes('attendance') || sLower.includes('log') || sLower.includes('daily')) score += 1500;
    if (sLower.includes('roster') || sLower.includes('data')) score += 600;
    if (sLower.includes('dashboard') || sLower.includes('summary')) score -= 200;
    if (sLower.includes('instruction') || sLower.includes('mapping') || sLower.includes('list') || sLower.includes('read me')) score -= 5000;

    if (score > bestScore) {
      bestScore = score;
      bestSheetName = sName;
      bestRows = validRows;
    }
  }

  // Fallback if scoring found nothing
  if (bestRows.length === 0 && workbook.SheetNames.length > 0) {
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
    const headerIdx = findBestHeaderRow(aoa);
    bestRows = XLSX.utils.sheet_to_json(firstSheet, { range: headerIdx, defval: '' });
    bestSheetName = workbook.SheetNames[0];
  }

  return { bestSheetName, bestRows };
}

// Master File Upload Handler
function handleFileUpload(file) {
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      if (typeof XLSX === 'undefined') {
        alert('Excel library (SheetJS) is loading. Please try again in 1 second.');
        return;
      }

      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });

      if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
        alert('No readable sheets found in this Excel file.');
        return;
      }

      currentWorkbook = workbook;

      // Smart sheet detection
      const { bestSheetName, bestRows } = findBestSheet(workbook);

      if (bestRows.length === 0) {
        alert(`Could not extract any data rows from "${file.name}". Please ensure the file has a header row and data rows.`);
        return;
      }

      // Setup multi-sheet switcher if more than 1 sheet
      const sheetSwitcher = document.getElementById('sheetSwitcher');
      const sheetSelect = document.getElementById('sheetSelect');
      if (workbook.SheetNames.length > 1 && sheetSwitcher && sheetSelect) {
        sheetSelect.innerHTML = '';
        workbook.SheetNames.forEach(name => {
          const opt = document.createElement('option');
          opt.value = name;
          opt.textContent = name;
          if (name === bestSheetName) opt.selected = true;
          sheetSelect.appendChild(opt);
        });
        sheetSwitcher.classList.remove('hidden');
        sheetSwitcher.classList.add('flex');

        sheetSelect.onchange = function() {
          loadSpecificSheet(currentWorkbook, sheetSelect.value);
        };
      } else if (sheetSwitcher) {
        sheetSwitcher.classList.add('hidden');
        sheetSwitcher.classList.remove('flex');
      }

      // Load the detected sheet
      loadDataset(bestRows, file.name);

    } catch (err) {
      console.error('Error processing file:', err);
      alert('Error reading Excel file: ' + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
}

function loadSpecificSheet(workbook, sheetName) {
  if (!workbook) workbook = currentWorkbook;
  if (!workbook) return;
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return;
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headerRowIdx = findBestHeaderRow(aoa);
  const json = XLSX.utils.sheet_to_json(sheet, { range: headerRowIdx, defval: '' });
  const validRows = json.filter(r => {
    return Object.values(r).some(v => String(v !== undefined && v !== null ? v : '').trim().length > 0);
  });
  loadDataset(validRows, currentFileName);
}

function resetFilterInputs() {
  if (document.getElementById('studentFilter')) document.getElementById('studentFilter').value = 'ALL';
  if (document.getElementById('batchFilter')) document.getElementById('batchFilter').value = 'ALL';
  if (document.getElementById('statusFilter')) document.getElementById('statusFilter').value = 'ALL';
  if (document.getElementById('facultyFilter')) document.getElementById('facultyFilter').value = 'ALL';
  const searchInput = document.getElementById('searchInput');
  if (searchInput) searchInput.value = '';
  const clearBtn = document.getElementById('clearSearchBtn');
  if (clearBtn) clearBtn.classList.add('hidden');
}

function loadDataset(rows, fileName) {
  uploadedRows = rows || [];
  currentFileName = fileName || 'Uploaded Spreadsheet';

  // Determine Columns from keys of the first row
  tableColumns = uploadedRows.length > 0 ? Object.keys(uploadedRows[0]) : [];

  // Reset filters to ALL so data is never hidden by old filter state
  resetFilterInputs();

  // Detect Column Semantics
  detectColumns();

  // Populate Dropdown Filters
  populateFilters();

  // Update Status Banner
  updateStatusBar();

  // Render KPIs & Table
  applyFilters();
}

function detectColumns() {
  detectedCols = { student: null, status: null, batch: null, faculty: null, date: null };

  // Pass 1: Look for exact / high-confidence matches
  tableColumns.forEach(col => {
    const cLower = col.toLowerCase().trim();

    if (!detectedCols.student) {
      if (cLower === 'student name' || cLower === 'candidate name' || cLower === 'student' || cLower === 'candidate') {
        detectedCols.student = col;
      }
    }

    if (!detectedCols.status) {
      if (cLower === 'status' || cLower === 'attendance status' || cLower === 'p/a' || cLower === 'present/absent') {
        detectedCols.status = col;
      }
    }

    if (!detectedCols.batch) {
      if (cLower === 'batch' || cLower === 'batch code' || cLower === 'class' || cLower === 'section') {
        detectedCols.batch = col;
      }
    }

    if (!detectedCols.faculty) {
      if (cLower === 'taken by' || cLower === 'recorded by' || cLower === 'faculty' || cLower === 'faculty name' || cLower === 'teacher') {
        detectedCols.faculty = col;
      }
    }

    if (!detectedCols.date) {
      if (cLower === 'date' || cLower === 'attendance date' || cLower === 'day') {
        detectedCols.date = col;
      }
    }
  });

  // Pass 2: Loose keyword matching for any unresolved columns
  tableColumns.forEach(col => {
    const cLower = col.toLowerCase().trim();

    if (!detectedCols.student) {
      if ((cLower.includes('student') || cLower.includes('name')) && !cLower.includes('course') && !cLower.includes('faculty') && !cLower.includes('holiday')) {
        detectedCols.student = col;
      }
    }

    if (!detectedCols.status) {
      if (cLower.includes('attendance') && !cLower.includes('taken') && !cLower.includes('by') && !cLower.includes('count')) {
        detectedCols.status = col;
      }
    }

    if (!detectedCols.batch) {
      if (cLower.includes('batch') || cLower.includes('class') || cLower.includes('section')) {
        detectedCols.batch = col;
      }
    }

    if (!detectedCols.faculty) {
      if (cLower.includes('faculty') || cLower.includes('teacher') || cLower.includes('taken by') || cLower.includes('mentor')) {
        detectedCols.faculty = col;
      }
    }

    if (!detectedCols.date) {
      if (cLower.includes('date') && !cLower.includes('birth')) {
        detectedCols.date = col;
      }
    }
  });

  // Pass 3: Inspect row values if status is still not detected
  if (!detectedCols.status && uploadedRows.length > 0) {
    for (let c of tableColumns) {
      let matchCount = 0;
      for (let i = 0; i < Math.min(25, uploadedRows.length); i++) {
        const val = String(uploadedRows[i][c] || '').toLowerCase().trim();
        if (val === 'present' || val === 'absent' || val === 'p' || val === 'a' || val === 'late' || val.includes('holiday') || val.includes('not marked')) {
          matchCount++;
        }
      }
      if (matchCount >= 2) {
        detectedCols.status = c;
        break;
      }
    }
  }
}

function updateStatusBar() {
  const fileBadge = document.getElementById('fileBadge');
  const fileNameText = document.getElementById('fileNameText');
  const fileStatsText = document.getElementById('fileStatsText');

  if (!currentFileName || uploadedRows.length === 0) {
    if (fileBadge) {
      fileBadge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-amber-100 text-amber-800 font-bold border border-amber-200';
      fileNameText.textContent = 'No file uploaded yet';
    }
    if (fileStatsText) fileStatsText.textContent = 'Please upload an attendance Excel file above to see your data.';
    return;
  }

  if (fileBadge) {
    fileBadge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-emerald-100 text-emerald-800 font-bold border border-emerald-300';
    fileNameText.textContent = `Active File: ${currentFileName}`;
  }

  const studentCount = detectedCols.student ? new Set(uploadedRows.map(r => r[detectedCols.student]).filter(Boolean)).size : '—';
  const batchCount = detectedCols.batch ? new Set(uploadedRows.map(r => r[detectedCols.batch]).filter(Boolean)).size : '—';
  if (fileStatsText) {
    fileStatsText.textContent = `${uploadedRows.length.toLocaleString()} total rows • ${tableColumns.length} columns • ${studentCount} students • ${batchCount} batches`;
  }

  if (window.lucide) lucide.createIcons();
}

function populateFilters() {
  const studentSelect = document.getElementById('studentFilter');
  const batchSelect = document.getElementById('batchFilter');
  const facultySelect = document.getElementById('facultyFilter');

  // 1. Student Filter
  if (studentSelect) {
    studentSelect.innerHTML = '<option value="ALL">All Students</option>';
    if (detectedCols.student) {
      const students = [...new Set(uploadedRows.map(r => String(r[detectedCols.student] || '').trim()).filter(Boolean))].sort();
      students.forEach(s => {
        const opt = document.createElement('option');
        opt.value = s;
        opt.textContent = s;
        studentSelect.appendChild(opt);
      });
    }
    studentSelect.value = 'ALL';
  }

  // 2. Batch Filter
  if (batchSelect) {
    batchSelect.innerHTML = '<option value="ALL">All Batches</option>';
    if (detectedCols.batch) {
      const batches = [...new Set(uploadedRows.map(r => String(r[detectedCols.batch] || '').trim()).filter(Boolean))].sort();
      batches.forEach(b => {
        const opt = document.createElement('option');
        opt.value = b;
        opt.textContent = b;
        batchSelect.appendChild(opt);
      });
    }
    batchSelect.value = 'ALL';
  }

  // 3. Faculty Filter
  if (facultySelect) {
    facultySelect.innerHTML = '<option value="ALL">All Faculty</option>';
    if (detectedCols.faculty) {
      const faculties = [...new Set(uploadedRows.map(r => String(r[detectedCols.faculty] || '').trim()).filter(Boolean))].sort();
      faculties.forEach(f => {
        const opt = document.createElement('option');
        opt.value = f;
        opt.textContent = f;
        facultySelect.appendChild(opt);
      });
    }
    facultySelect.value = 'ALL';
  }
}

function getFilteredRows() {
  if (uploadedRows.length === 0) return [];

  const selStudent = document.getElementById('studentFilter')?.value || 'ALL';
  const selBatch = document.getElementById('batchFilter')?.value || 'ALL';
  const selStatus = document.getElementById('statusFilter')?.value || 'ALL';
  const selFaculty = document.getElementById('facultyFilter')?.value || 'ALL';
  const search = (document.getElementById('searchInput')?.value || '').toLowerCase().trim();

  return uploadedRows.filter(r => {
    // Student Filter
    if (selStudent !== 'ALL' && detectedCols.student) {
      if (String(r[detectedCols.student] || '').trim() !== selStudent) return false;
    }
    // Batch Filter
    if (selBatch !== 'ALL' && detectedCols.batch) {
      if (String(r[detectedCols.batch] || '').trim() !== selBatch) return false;
    }
    // Faculty Filter
    if (selFaculty !== 'ALL' && detectedCols.faculty) {
      if (String(r[detectedCols.faculty] || '').trim() !== selFaculty) return false;
    }
    // Status Filter
    if (selStatus !== 'ALL' && detectedCols.status) {
      const val = String(r[detectedCols.status] || '').toLowerCase().trim();
      if (selStatus === 'Present' && !val.includes('present') && val !== 'p' && val !== '1' && val !== 'attended') return false;
      if (selStatus === 'Absent' && !val.includes('absent') && val !== 'a' && val !== '0' && val !== 'missed') return false;
      if (selStatus === 'Late' && !val.includes('late')) return false;
      if (selStatus === 'Holiday' && !val.includes('holiday') && !val.includes('leave')) return false;
    }
    // Search Box across all columns
    if (search) {
      const match = tableColumns.some(col => String(r[col] || '').toLowerCase().includes(search));
      if (!match) return false;
    }
    return true;
  });
}

function applyFilters() {
  const filtered = getFilteredRows();

  // 1. Update KPI Summary Cards
  updateKPIs(filtered);

  // 2. Render Main Data Table
  renderTable(filtered);

  // 3. Render Visual Analytics Charts
  renderCharts(filtered);

  // 4. Render Faculty Breakdown Table (if faculty detected)
  renderFacultyTable(filtered);

  // 5. Update row count badge
  const rowCountBadge = document.getElementById('rowCountBadge');
  const tableRecordCount = document.getElementById('tableRecordCount');
  if (rowCountBadge) rowCountBadge.textContent = `Showing ${filtered.length.toLocaleString()} of ${uploadedRows.length.toLocaleString()} rows`;
  if (tableRecordCount) tableRecordCount.textContent = `${filtered.length.toLocaleString()} records displayed`;

  if (window.lucide) lucide.createIcons();
}

function updateKPIs(filtered) {
  let present = 0;
  let absent = 0;
  let students = new Set();
  let batches = new Set();

  filtered.forEach(r => {
    if (detectedCols.student && r[detectedCols.student]) {
      students.add(String(r[detectedCols.student]).trim());
    }
    if (detectedCols.batch && r[detectedCols.batch]) {
      batches.add(String(r[detectedCols.batch]).trim());
    }
    if (detectedCols.status) {
      const s = String(r[detectedCols.status] || '').toLowerCase().trim();
      if (s === 'present' || s === 'p' || s === '1' || s === 'attended' || s.includes('present')) {
        present++;
      } else if (s === 'absent' || s === 'a' || s === '0' || s === 'missed' || s.includes('absent')) {
        absent++;
      }
    }
  });

  const totalMarked = present + absent;
  const pct = totalMarked > 0 ? ((present / totalMarked) * 100).toFixed(1) : '—';

  const kpiTotalEl = document.getElementById('kpiTotalRows');
  const kpiPresentEl = document.getElementById('kpiPresent');
  const kpiAbsentEl = document.getElementById('kpiAbsent');
  const kpiStudentsEl = document.getElementById('kpiStudents');
  const kpiBatchesEl = document.getElementById('kpiBatches');

  if (kpiTotalEl) kpiTotalEl.textContent = filtered.length.toLocaleString();
  if (kpiPresentEl) kpiPresentEl.textContent = present.toLocaleString();
  if (kpiAbsentEl) kpiAbsentEl.textContent = absent.toLocaleString();
  if (kpiStudentsEl) kpiStudentsEl.textContent = students.size.toLocaleString();
  if (kpiBatchesEl) kpiBatchesEl.textContent = batches.size.toLocaleString();

  // Percentage & Tier Badge
  const pctEl = document.getElementById('kpiPercentage');
  const tierBadge = document.getElementById('kpiTierBadge');
  const progressBar = document.getElementById('kpiProgressBar');

  if (pctEl && tierBadge && progressBar) {
    if (pct === '—') {
      pctEl.textContent = '—';
      pctEl.className = 'text-2xl font-black text-slate-400';
      tierBadge.textContent = 'No Data';
      tierBadge.className = 'text-[9px] font-bold px-1.5 py-0.5 rounded-full bg-slate-100 text-slate-600';
      progressBar.style.width = '0%';
      progressBar.className = 'h-full bg-slate-300';
    } else {
      const numPct = parseFloat(pct);
      pctEl.textContent = `${pct}%`;
      progressBar.style.width = `${Math.min(100, Math.max(0, numPct))}%`;

      if (numPct >= 90) {
        pctEl.className = 'text-2xl font-black text-emerald-700';
        tierBadge.textContent = 'Good (>=90%)';
        tierBadge.className = 'text-[9px] font-bold px-1.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-300';
        progressBar.className = 'h-full bg-emerald-600';
      } else if (numPct >= 75) {
        pctEl.className = 'text-2xl font-black text-amber-700';
        tierBadge.textContent = 'Warning (75-89%)';
        tierBadge.className = 'text-[9px] font-bold px-1.5 py-0.5 rounded-full bg-amber-100 text-amber-800 border border-amber-300';
        progressBar.className = 'h-full bg-amber-500';
      } else {
        pctEl.className = 'text-2xl font-black text-rose-700';
        tierBadge.textContent = 'Shortage (<75%)';
        tierBadge.className = 'text-[9px] font-bold px-1.5 py-0.5 rounded-full bg-rose-100 text-rose-800 border border-rose-300';
        progressBar.className = 'h-full bg-rose-600';
      }
    }
  }
}

function renderCharts(filtered) {
  const section = document.getElementById('chartsSection');
  if (!window.Chart || !section) return;

  if (filtered.length === 0 || !detectedCols.status) {
    section.classList.add('hidden');
    return;
  }

  section.classList.remove('hidden');

  // 1. Overall Status Donut Chart
  let present = 0;
  let absent = 0;
  let other = 0;

  filtered.forEach(r => {
    const s = String(r[detectedCols.status] || '').toLowerCase().trim();
    if (s === 'present' || s === 'p' || s === '1' || s === 'attended' || s.includes('present')) present++;
    else if (s === 'absent' || s === 'a' || s === '0' || s === 'missed' || s.includes('absent')) absent++;
    else other++;
  });

  const ctxStatus = document.getElementById('statusChart')?.getContext('2d');
  if (ctxStatus) {
    if (statusChartInstance) statusChartInstance.destroy();
    statusChartInstance = new Chart(ctxStatus, {
      type: 'doughnut',
      data: {
        labels: ['Present', 'Absent', ...(other > 0 ? ['Other / Excused'] : [])],
        datasets: [{
          data: [present, absent, ...(other > 0 ? [other] : [])],
          backgroundColor: ['#10b981', '#f43f5e', '#f59e0b'],
          borderWidth: 2,
          borderColor: '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { boxWidth: 12, font: { size: 11 } }
          }
        }
      }
    });
  }

  // 2. Batch Comparison Bar Chart
  const ctxBatch = document.getElementById('batchChart')?.getContext('2d');
  if (ctxBatch) {
    if (batchChartInstance) batchChartInstance.destroy();

    if (detectedCols.batch) {
      const batchStats = {};
      filtered.forEach(r => {
        const b = String(r[detectedCols.batch] || 'Unassigned').trim();
        if (!b) return;
        if (!batchStats[b]) batchStats[b] = { present: 0, total: 0 };
        const s = String(r[detectedCols.status] || '').toLowerCase().trim();
        if (s === 'present' || s === 'p' || s === '1' || s === 'attended' || s.includes('present')) {
          batchStats[b].present++;
          batchStats[b].total++;
        } else if (s === 'absent' || s === 'a' || s === '0' || s === 'missed' || s.includes('absent')) {
          batchStats[b].total++;
        }
      });

      const labels = Object.keys(batchStats).slice(0, 10);
      const data = labels.map(l => {
        const item = batchStats[l];
        return item.total > 0 ? Math.round((item.present / item.total) * 100) : 0;
      });

      batchChartInstance = new Chart(ctxBatch, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [{
            label: 'Attendance Rate',
            data: data,
            backgroundColor: '#3b82f6',
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: {
              beginAtZero: true,
              max: 100,
              ticks: {
                callback: function(v) { return v + '%'; },
                font: { size: 10 }
              }
            },
            x: {
              ticks: { font: { size: 10 } }
            }
          },
          plugins: {
            legend: { display: false }
          }
        }
      });
    }
  }
}

function renderTable(rows) {
  const thead = document.getElementById('dataTableHeader');
  const tbody = document.getElementById('dataTableBody');

  if (!thead || !tbody) return;

  if (tableColumns.length === 0 || rows.length === 0) {
    thead.innerHTML = `<tr><th class="py-3 px-4 text-center text-slate-400">${uploadedRows.length === 0 ? 'Please upload an Excel file above to see data' : 'No records match your active filters'}</th></tr>`;
    tbody.innerHTML = `<tr><td class="py-12 text-center text-slate-400">${uploadedRows.length === 0 ? 'Upload an Excel spreadsheet above to display records.' : 'No rows match the selected filters. Click Reset to clear filters.'}</td></tr>`;
    return;
  }

  // Header row
  let headerHtml = '<tr><th class="py-2.5 px-3 w-12 text-center bg-slate-100 text-slate-600 font-bold border-r border-slate-200">#</th>';
  tableColumns.forEach(col => {
    headerHtml += `<th class="py-2.5 px-3.5 whitespace-nowrap bg-slate-100 text-slate-700 font-bold border-r border-slate-200 uppercase text-[11px] tracking-wider">${col}</th>`;
  });
  headerHtml += '</tr>';
  thead.innerHTML = headerHtml;

  // Body rows
  let bodyHtml = '';
  rows.forEach((r, idx) => {
    bodyHtml += `<tr class="hover:bg-blue-50/50 transition">`;
    bodyHtml += `<td class="py-2 px-3 text-center text-slate-400 font-mono text-xs border-r border-slate-100">${idx + 1}</td>`;

    tableColumns.forEach(col => {
      let val = r[col];
      // Format Date if needed
      if (val instanceof Date) {
        val = val.toISOString().split('T')[0];
      }

      const strVal = String(val !== undefined && val !== null ? val : '').trim();
      let formattedVal = strVal;

      // Smart pill styling for Status column
      if (col === detectedCols.status) {
        const lower = strVal.toLowerCase();
        if (lower === 'present' || lower === 'p' || lower === '1' || lower === 'attended') {
          formattedVal = `<span class="inline-block px-2 py-0.5 rounded text-xs bg-emerald-100 text-emerald-800 font-bold">Present</span>`;
        } else if (lower === 'absent' || lower === 'a' || lower === '0' || lower === 'missed') {
          formattedVal = `<span class="inline-block px-2 py-0.5 rounded text-xs bg-rose-100 text-rose-800 font-bold">Absent</span>`;
        } else if (lower.includes('late')) {
          formattedVal = `<span class="inline-block px-2 py-0.5 rounded text-xs bg-amber-100 text-amber-800 font-bold">${strVal}</span>`;
        } else if (lower.includes('holiday') || lower.includes('leave')) {
          formattedVal = `<span class="inline-block px-2 py-0.5 rounded text-xs bg-purple-100 text-purple-800 font-bold">${strVal}</span>`;
        }
      }

      bodyHtml += `<td class="py-2 px-3.5 whitespace-nowrap text-slate-700 border-r border-slate-100">${formattedVal}</td>`;
    });

    bodyHtml += `</tr>`;
  });

  tbody.innerHTML = bodyHtml;
}

function renderFacultyTable(rows) {
  const section = document.getElementById('facultySection');
  const tbody = document.getElementById('facultyTableBody');

  if (!section || !tbody) return;

  if (!detectedCols.faculty || rows.length === 0) {
    section.classList.add('hidden');
    return;
  }

  const faculties = [...new Set(rows.map(r => String(r[detectedCols.faculty] || '').trim()).filter(Boolean))].sort();

  if (faculties.length === 0) {
    section.classList.add('hidden');
    return;
  }

  section.classList.remove('hidden');

  let html = '';
  faculties.forEach((fac, idx) => {
    const facRows = rows.filter(r => String(r[detectedCols.faculty] || '').trim() === fac);
    let present = 0;
    let absent = 0;
    let batches = new Set();

    facRows.forEach(r => {
      if (detectedCols.batch && r[detectedCols.batch]) batches.add(String(r[detectedCols.batch]).trim());
      if (detectedCols.status) {
        const s = String(r[detectedCols.status] || '').toLowerCase().trim();
        if (s === 'present' || s === 'p' || s === '1' || s === 'attended' || s.includes('present')) present++;
        else if (s === 'absent' || s === 'a' || s === '0' || s === 'missed' || s.includes('absent')) absent++;
      }
    });

    const total = present + absent;
    const pct = total > 0 ? ((present / total) * 100).toFixed(1) : '—';
    const badgeClass = pct !== '—' && parseFloat(pct) >= 90 ? 'bg-emerald-100 text-emerald-800' : (pct !== '—' && parseFloat(pct) >= 75 ? 'bg-amber-100 text-amber-800' : 'bg-rose-100 text-rose-800');

    html += `
      <tr class="hover:bg-slate-50 transition">
        <td class="py-2.5 px-3 text-center text-slate-400 font-medium">${idx + 1}</td>
        <td class="py-2.5 px-4 font-bold text-slate-800">${fac}</td>
        <td class="py-2.5 px-3 text-center text-slate-600">${batches.size > 0 ? [...batches].join(', ') : '—'}</td>
        <td class="py-2.5 px-3 text-center font-bold text-slate-700">${facRows.length}</td>
        <td class="py-2.5 px-3 text-center font-semibold text-emerald-700">${present}</td>
        <td class="py-2.5 px-3 text-center font-semibold text-rose-700">${absent}</td>
        <td class="py-2.5 px-4 text-center">
          <span class="inline-block px-2.5 py-1 rounded text-xs font-bold ${badgeClass}">
            ${pct !== '—' ? pct + '%' : '—'}
          </span>
        </td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

function handleSearchInput() {
  const input = document.getElementById('searchInput');
  const clearBtn = document.getElementById('clearSearchBtn');
  if (clearBtn) {
    if (input && input.value.trim().length > 0) {
      clearBtn.classList.remove('hidden');
    } else {
      clearBtn.classList.add('hidden');
    }
  }
  applyFilters();
}

function clearSearch() {
  const input = document.getElementById('searchInput');
  const clearBtn = document.getElementById('clearSearchBtn');
  if (input) {
    input.value = '';
    input.focus();
  }
  if (clearBtn) clearBtn.classList.add('hidden');
  applyFilters();
}

function resetFilters() {
  resetFilterInputs();
  applyFilters();
}

function clearData() {
  if (uploadedRows.length === 0) return;
  if (!confirm('Are you sure you want to clear the uploaded data?')) return;

  uploadedRows = [];
  tableColumns = [];
  currentFileName = null;
  currentWorkbook = null;
  detectedCols = { student: null, status: null, batch: null, faculty: null, date: null };

  document.getElementById('sheetSwitcher')?.classList.add('hidden');
  if (statusChartInstance) { statusChartInstance.destroy(); statusChartInstance = null; }
  if (batchChartInstance) { batchChartInstance.destroy(); batchChartInstance = null; }
  document.getElementById('chartsSection')?.classList.add('hidden');
  updateStatusBar();
  populateFilters();
  applyFilters();
}

function exportToExcel() {
  if (uploadedRows.length === 0) {
    alert('Please upload an Excel file first before exporting.');
    return;
  }

  const filtered = getFilteredRows();
  if (filtered.length === 0) {
    alert('No rows match your current filter selection to export.');
    return;
  }

  const ws = XLSX.utils.json_to_sheet(filtered);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Filtered Attendance Data');
  XLSX.writeFile(wb, `Export_${currentFileName || 'Attendance'}.xlsx`);
}

// Resilient App Initialization
function initApp() {
  const fileInput = document.getElementById('excelFileInput');
  const dropZone = document.getElementById('dropZone');
  const exportBtn = document.getElementById('exportExcelBtn');

  if (fileInput) {
    fileInput.addEventListener('change', function(e) {
      handleFileInputChange(fileInput);
    });
  }

  // Prevent browser window from opening dragged files
  window.addEventListener('dragover', function(e) { e.preventDefault(); }, false);
  window.addEventListener('drop', function(e) { e.preventDefault(); }, false);

  if (dropZone) {
    ['dragenter', 'dragover'].forEach(name => {
      dropZone.addEventListener(name, function(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('border-blue-600', 'bg-blue-100/70');
      }, false);
    });

    ['dragleave', 'drop'].forEach(name => {
      dropZone.addEventListener(name, function(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('border-blue-600', 'bg-blue-100/70');
      }, false);
    });

    dropZone.addEventListener('drop', function(e) {
      e.preventDefault();
      e.stopPropagation();
      const dt = e.dataTransfer;
      const file = dt && dt.files && dt.files[0];
      if (file) {
        handleFileUpload(file);
      }
    });
  }

  if (exportBtn) {
    exportBtn.addEventListener('click', exportToExcel);
  }

  updateStatusBar();
  applyFilters();
}

// Run immediately if DOM is ready, or on DOMContentLoaded
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}
