/* ============================================================
   NISSI Application Controller
   ─────────────────────────────────────────────────────────────
   Orchestrates the UI: step navigation, file upload, trailer
   selection, SORT configuration, processing, and download.
   ============================================================ */

import { parseExcelFile } from './parser.js';
import { expandRange, processAll } from './processor.js';
import { downloadExcel } from './writer.js';

/* ── State ────────────────────────────────────────────────── */
const state = {
  currentStep: 1,
  file: null,
  parsedData: null,        // { rows, headers, trailerColumns, sheetName, totalRows }
  selectedTrailers: [],
  processedResult: null,   // { success, data, stats, error }
};

/* ── DOM References ───────────────────────────────────────── */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const dom = {
  // Step containers
  steps: [null, $('#step-1'), $('#step-2'), $('#step-3'), $('#step-4')],
  wizardSteps: [null, $('#wizard-step-1'), $('#wizard-step-2'), $('#wizard-step-3'), $('#wizard-step-4')],
  connectors: [null, null, $('#connector-1-2'), $('#connector-2-3'), $('#connector-3-4')],

  // Step 1
  uploadZone:   $('#upload-zone'),
  fileInput:    $('#file-input'),
  browseBtn:    $('#browse-btn'),
  fileInfo:     $('#file-info'),
  fileName:     $('#file-name'),
  fileMeta:     $('#file-meta'),
  removeFileBtn:$('#remove-file-btn'),
  nextStep2Btn: $('#next-step-2-btn'),

  // Step 2
  trailerCheckboxes: $('#trailer-checkboxes'),
  selectAllBtn:      $('#select-all-btn'),
  deselectAllBtn:    $('#deselect-all-btn'),
  backStep1Btn:      $('#back-step-1-btn'),
  nextStep3Btn:      $('#next-step-3-btn'),

  // Step 3
  sort1Input: $('#sort1-input'),
  sort2Input: $('#sort2-input'),
  sort3Input: $('#sort3-input'),
  sort4Input: $('#sort4-input'),
  sort1Preview: $('#sort1-preview'),
  sort2Preview: $('#sort2-preview'),
  sort3Preview: $('#sort3-preview'),
  sort4Preview: $('#sort4-preview'),
  capacityInfo: $('#capacity-info'),
  backStep2Btn: $('#back-step-2-btn'),
  processBtn:   $('#process-btn'),

  // Step 4
  processingView: $('#processing-view'),
  resultsView:    $('#results-view'),
  statsGrid:      $('#stats-grid'),
  downloadBtn:    $('#download-btn'),
  startOverBtn:   $('#start-over-btn'),

  // Modal
  errorModal:       $('#error-modal'),
  errorModalMessage:$('#error-modal-message'),
  errorModalClose:  $('#error-modal-close'),

  // Toast
  toastContainer: $('#toast-container'),
};


/* ── Step Navigation ──────────────────────────────────────── */
function goToStep(step) {
  const prev = state.currentStep;
  state.currentStep = step;

  // Update step content visibility
  for (let i = 1; i <= 4; i++) {
    dom.steps[i].classList.toggle('active', i === step);
  }

  // Update wizard nav
  for (let i = 1; i <= 4; i++) {
    const ws = dom.wizardSteps[i];
    ws.classList.remove('active', 'completed');
    if (i < step) ws.classList.add('completed');
    else if (i === step) ws.classList.add('active');
  }

  // Update connectors
  for (let i = 2; i <= 4; i++) {
    const conn = dom.connectors[i];
    if (conn) conn.classList.toggle('filled', i <= step);
  }
}


/* ── Step 1: File Upload ──────────────────────────────────── */
function initUpload() {
  const zone = dom.uploadZone;
  const input = dom.fileInput;

  // Click to browse
  dom.browseBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    input.click();
  });
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
  });

  // File selected via input
  input.addEventListener('change', () => {
    if (input.files.length > 0) handleFile(input.files[0]);
  });

  // Drag & drop
  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('drag-over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
  });

  // Remove file
  dom.removeFileBtn.addEventListener('click', resetFile);

  // Next
  dom.nextStep2Btn.addEventListener('click', async () => {
    if (!state.file) return;
    try {
      dom.nextStep2Btn.disabled = true;
      dom.nextStep2Btn.textContent = 'Parsing…';
      state.parsedData = await parseExcelFile(state.file);
      renderTrailerCheckboxes();
      goToStep(2);
    } catch (err) {
      showError(err.message);
    } finally {
      dom.nextStep2Btn.disabled = false;
      dom.nextStep2Btn.innerHTML = 'Continue <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>';
    }
  });
}

function handleFile(file) {
  const validTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
  ];
  const isValid = validTypes.includes(file.type) || /\.xlsx?$/i.test(file.name);

  if (!isValid) {
    showToast('Please upload a valid .xlsx or .xls file.', 'error');
    return;
  }

  state.file = file;
  dom.fileName.textContent = file.name;
  dom.fileMeta.textContent = `${formatFileSize(file.size)} • ${new Date(file.lastModified).toLocaleDateString()}`;
  dom.uploadZone.classList.add('hidden');
  dom.fileInfo.classList.remove('hidden');
}

function resetFile() {
  state.file = null;
  state.parsedData = null;
  dom.fileInput.value = '';
  dom.uploadZone.classList.remove('hidden');
  dom.fileInfo.classList.add('hidden');
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}


/* ── Step 2: Trailer Selection ────────────────────────────── */
function renderTrailerCheckboxes() {
  const container = dom.trailerCheckboxes;
  container.innerHTML = '';
  state.selectedTrailers = [];

  for (const col of state.parsedData.trailerColumns) {
    const item = document.createElement('label');
    item.className = 'checkbox-item';
    item.innerHTML = `
      <input type="checkbox" value="${escapeHtml(col)}" id="chk-${escapeId(col)}">
      <div class="custom-check">
        <svg viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>
      </div>
      <span class="checkbox-label">${escapeHtml(col)}</span>
    `;
    const input = item.querySelector('input');

    input.addEventListener('change', () => {
      item.classList.toggle('checked', input.checked);
      updateSelectedTrailers();
    });

    container.appendChild(item);
  }

  // Select/Deselect all
  dom.selectAllBtn.onclick = () => toggleAllTrailers(true);
  dom.deselectAllBtn.onclick = () => toggleAllTrailers(false);

  updateSelectedTrailers();
}

function toggleAllTrailers(checked) {
  const items = dom.trailerCheckboxes.querySelectorAll('.checkbox-item');
  items.forEach(item => {
    const input = item.querySelector('input');
    input.checked = checked;
    item.classList.toggle('checked', checked);
  });
  updateSelectedTrailers();
}

function updateSelectedTrailers() {
  const inputs = dom.trailerCheckboxes.querySelectorAll('input:checked');
  state.selectedTrailers = Array.from(inputs).map(i => i.value);
  dom.nextStep3Btn.disabled = state.selectedTrailers.length === 0;
}

function initTrailerStep() {
  dom.backStep1Btn.addEventListener('click', () => goToStep(1));
  dom.nextStep3Btn.addEventListener('click', () => {
    if (state.selectedTrailers.length === 0) {
      showToast('Please select at least one trailer column.', 'error');
      return;
    }
    goToStep(3);
    updateCapacityInfo(); // Initialize capacity display
  });
}


/* ── Step 3: SORT Configuration ───────────────────────────── */
function initSortConfig() {
  const inputs = [
    { input: dom.sort1Input, preview: dom.sort1Preview },
    { input: dom.sort2Input, preview: dom.sort2Preview },
    { input: dom.sort3Input, preview: dom.sort3Preview },
    { input: dom.sort4Input, preview: dom.sort4Preview },
  ];

  for (const { input, preview } of inputs) {
    input.addEventListener('input', () => {
      renderSortPreview(input, preview);
      updateCapacityInfo();
      updateProcessButton();
    });
  }

  dom.backStep2Btn.addEventListener('click', () => goToStep(2));
  dom.processBtn.addEventListener('click', runProcessing);
}

function renderSortPreview(input, previewEl) {
  const values = expandRange(input.value);
  previewEl.innerHTML = '';

  if (values.length === 0 && input.value.trim() !== '') {
    // Invalid or blank
    const tag = document.createElement('span');
    tag.className = 'sort-tag empty-tag';
    tag.textContent = '(blank)';
    previewEl.appendChild(tag);
  } else if (values.length === 0) {
    // Empty input — show nothing
  } else {
    for (const val of values) {
      const tag = document.createElement('span');
      tag.className = 'sort-tag';
      tag.textContent = val;
      previewEl.appendChild(tag);
    }
  }
}

function updateCapacityInfo() {
  const s1 = expandRange(dom.sort1Input.value);
  const s2 = expandRange(dom.sort2Input.value);
  const s3 = expandRange(dom.sort3Input.value);
  const s4 = expandRange(dom.sort4Input.value);

  const info = dom.capacityInfo;

  if (s1.length === 0 || s2.length === 0 || s3.length === 0) {
    info.classList.add('hidden');
    return;
  }

  const c4 = s4.length > 0 ? s4.length : 1;
  const capacity = s1.length * s2.length * s3.length * c4;

  info.classList.remove('hidden');

  const formula = s4.length > 0
    ? `${s1.length} × ${s2.length} × ${s3.length} × ${s4.length}`
    : `${s1.length} × ${s2.length} × ${s3.length}`;

  info.className = 'capacity-info ok';
  info.innerHTML = `
    <span class="cap-icon">📊</span>
    <span class="cap-text">Total capacity: ${formula} =</span>
    <span class="cap-number">${capacity.toLocaleString()} combinations</span>
  `;
}

function updateProcessButton() {
  const s1 = expandRange(dom.sort1Input.value);
  const s2 = expandRange(dom.sort2Input.value);
  const s3 = expandRange(dom.sort3Input.value);

  dom.processBtn.disabled = s1.length === 0 || s2.length === 0 || s3.length === 0;
}


/* ── Step 4: Processing & Results ─────────────────────────── */
async function runProcessing() {
  goToStep(4);
  dom.processingView.classList.remove('hidden');
  dom.resultsView.classList.add('hidden');

  // Use setTimeout to let the UI update before heavy processing
  await new Promise(r => setTimeout(r, 100));

  try {
    const result = processAll(
      state.parsedData.rows,
      state.selectedTrailers,
      dom.sort1Input.value,
      dom.sort2Input.value,
      dom.sort3Input.value,
      dom.sort4Input.value,
    );

    if (!result.success) {
      showError(result.error);
      goToStep(3);
      return;
    }

    state.processedResult = result;
    renderResults(result);
  } catch (err) {
    showError(err.message || 'An unexpected error occurred during processing.');
    goToStep(3);
  }
}

function renderResults(result) {
  dom.processingView.classList.add('hidden');
  dom.resultsView.classList.remove('hidden');

  const { stats } = result;
  dom.statsGrid.innerHTML = '';

  const statItems = [
    { label: 'Input Rows', value: stats.totalInputRows.toLocaleString() },
    { label: 'Normal Rows Kept', value: stats.normalRowsKept.toLocaleString() },
    { label: 'Regular Scan Rows', value: stats.regularScanRows.toLocaleString() },
    { label: 'Rows Removed', value: stats.rowsRemoved.toLocaleString() },
    { label: 'Total Output Rows', value: stats.totalOutputRows.toLocaleString() },
    { label: 'Sort Capacity', value: stats.capacity.toLocaleString() },
  ];

  for (const item of statItems) {
    const card = document.createElement('div');
    card.className = 'stat-card';
    card.innerHTML = `
      <div class="stat-label">${item.label}</div>
      <div class="stat-value">${item.value}</div>
    `;
    dom.statsGrid.appendChild(card);
  }
}

function initResultsStep() {
  dom.downloadBtn.addEventListener('click', () => {
    if (!state.processedResult || !state.processedResult.data) return;

    try {
      // Generate filename with timestamp
      const timestamp = new Date().toISOString().slice(0, 10);
      const filename = `NISSI_Processed_${timestamp}.xlsx`;
      downloadExcel(state.processedResult.data, filename);
      showToast('File downloaded successfully!', 'success');
    } catch (err) {
      showError('Failed to generate Excel file: ' + err.message);
    }
  });

  dom.startOverBtn.addEventListener('click', () => {
    // Reset everything
    state.file = null;
    state.parsedData = null;
    state.selectedTrailers = [];
    state.processedResult = null;
    dom.fileInput.value = '';
    dom.uploadZone.classList.remove('hidden');
    dom.fileInfo.classList.add('hidden');
    dom.sort1Input.value = '';
    dom.sort2Input.value = '';
    dom.sort3Input.value = '';
    dom.sort4Input.value = '';
    dom.sort1Preview.innerHTML = '';
    dom.sort2Preview.innerHTML = '';
    dom.sort3Preview.innerHTML = '';
    dom.sort4Preview.innerHTML = '';
    dom.capacityInfo.classList.add('hidden');
    dom.trailerCheckboxes.innerHTML = '';
    goToStep(1);
  });
}


/* ── Error Modal ──────────────────────────────────────────── */
function showError(message) {
  dom.errorModalMessage.textContent = message;
  dom.errorModal.classList.remove('hidden');
}

function initErrorModal() {
  dom.errorModalClose.addEventListener('click', () => {
    dom.errorModal.classList.add('hidden');
  });

  dom.errorModal.addEventListener('click', (e) => {
    if (e.target === dom.errorModal) {
      dom.errorModal.classList.add('hidden');
    }
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !dom.errorModal.classList.contains('hidden')) {
      dom.errorModal.classList.add('hidden');
    }
  });
}


/* ── Toast Notifications ──────────────────────────────────── */
function showToast(message, type = 'success') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `
    <span>${type === 'error' ? '⚠️' : '✓'}</span>
    <span>${escapeHtml(message)}</span>
  `;
  dom.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.classList.add('toast-out');
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}


/* ── Utilities ────────────────────────────────────────────── */
function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function escapeId(str) {
  return str.replace(/[^a-zA-Z0-9-_]/g, '_');
}


/* ── Initialize ───────────────────────────────────────────── */
function init() {
  initUpload();
  initTrailerStep();
  initSortConfig();
  initResultsStep();
  initErrorModal();
}

// Wait for DOM
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
