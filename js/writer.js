/* ============================================================
   NISSI Excel Writer
   ─────────────────────────────────────────────────────────────
   Generates and downloads the final .xlsx output file
   using SheetJS with values only (no formulas).
   ============================================================ */

/**
 * Generate and trigger download of the processed Excel file.
 *
 * @param {Object[]} outputRows - Array of cleaned output row objects with final column names.
 * @param {string} [filename='NISSI_Processed_Output.xlsx'] - The download filename.
 */
export function downloadExcel(outputRows, filename = 'NISSI_Processed_Output.xlsx') {
  const XLSX = window.XLSX;
  if (!XLSX) {
    throw new Error('SheetJS library (XLSX) is not loaded.');
  }

  // Define the exact column order for the output
  const columnOrder = [
    'Old_Barcode',
    'BC',
    'Article',
    'Units',
    'SORT1',
    'SORT2',
    'SORT3',
    'SORT4',
    'ScanCount',
  ];

  // Build worksheet from JSON with explicit column ordering
  const ws = XLSX.utils.json_to_sheet(outputRows, {
    header: columnOrder,
  });

  // Set column widths for readability
  ws['!cols'] = [
    { wch: 18 },  // Old_Barcode
    { wch: 24 },  // BC
    { wch: 14 },  // Article
    { wch: 8 },   // Units
    { wch: 8 },   // SORT1
    { wch: 8 },   // SORT2
    { wch: 8 },   // SORT3
    { wch: 8 },   // SORT4
    { wch: 12 },  // ScanCount
  ];

  // Create workbook
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Processed Output');

  // Write and download — bookType: 'xlsx' ensures .xlsx format
  // type: 'array' generates an ArrayBuffer for clean download
  const wbOut = XLSX.write(wb, {
    bookType: 'xlsx',
    type: 'array',
    bookSST: false, // Don't use shared string table — keeps values simple
  });

  // Create blob and trigger download
  const blob = new Blob([wbOut], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();

  // Cleanup
  setTimeout(() => {
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }, 200);
}
