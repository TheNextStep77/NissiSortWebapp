/* ============================================================
   NISSI Excel Writer
   ─────────────────────────────────────────────────────────────
   Generates the final .xlsx output as a Blob (no auto-download)
   so the caller can name the file before saving.
   Column order: Old_Barcode | Variant | BC | Units | SORT1-4 | ScanCount
   ============================================================ */

/**
 * Build a processed Excel workbook and return it as a Blob.
 * Caller is responsible for triggering the download with the chosen filename.
 *
 * @param {Object[]} outputRows - Array of output row objects from processor.js
 * @returns {Blob} .xlsx file blob
 */
export function buildExcelBlob(outputRows) {
  const XLSX = window.XLSX;
  if (!XLSX) {
    throw new Error('SheetJS library (XLSX) is not loaded.');
  }

  // Column order: Old_Barcode | Variant (moved to col 2) | BC | Units | SORT1-4 | ScanCount
  const columnOrder = [
    'Old_Barcode',
    'Variant',
    'BC',
    'Units',
    'SORT1',
    'SORT2',
    'SORT3',
    'SORT4',
    'ScanCount',
  ];

  const ws = XLSX.utils.json_to_sheet(outputRows, { header: columnOrder });

  // Column widths
  ws['!cols'] = [
    { wch: 18 },  // Old_Barcode
    { wch: 14 },  // Variant
    { wch: 24 },  // BC
    { wch: 8 },   // Units
    { wch: 8 },   // SORT1
    { wch: 8 },   // SORT2
    { wch: 8 },   // SORT3
    { wch: 8 },   // SORT4
    { wch: 12 },  // ScanCount
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Processed Output');

  const wbOut = XLSX.write(wb, {
    bookType: 'xlsx',
    type: 'array',
    bookSST: false,
  });

  return new Blob([wbOut], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/**
 * Trigger a browser download from a Blob.
 *
 * @param {Blob} blob
 * @param {string} filename
 */
export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename.endsWith('.xlsx') ? filename : filename + '.xlsx';
  document.body.appendChild(link);
  link.click();
  setTimeout(() => {
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }, 200);
}
