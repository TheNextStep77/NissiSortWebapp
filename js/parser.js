/* ============================================================
   NISSI Excel Parser
   ─────────────────────────────────────────────────────────────
   Reads an uploaded .xlsx file using SheetJS, validates
   required columns, and returns structured row data.
   ============================================================ */

/** Required source columns that must exist in the sheet. */
const REQUIRED_COLUMNS = ['SWXref', 'BC', 'SKU NUM'];

/**
 * Parse an uploaded Excel file and extract data from the first sheet
 * (or the "Article Detail" sheet if found).
 *
 * @param {File} file - The uploaded .xlsx File object.
 * @returns {Promise<{
 *   rows: Object[],
 *   headers: string[],
 *   trailerColumns: string[],
 *   sheetName: string,
 *   totalRows: number
 * }>}
 * @throws {Error} If the file cannot be parsed or required columns are missing.
 */
export async function parseExcelFile(file) {
  const XLSX = window.XLSX;
  if (!XLSX) {
    throw new Error('SheetJS library (XLSX) is not loaded. Please check your internet connection and refresh.');
  }

  // Read file as ArrayBuffer
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });

  // Find the right sheet — prefer "Article Detail", fall back to the first sheet
  let sheetName = workbook.SheetNames.find(
    name => name.toLowerCase().replace(/\s+/g, '') === 'articledetail'
  );
  if (!sheetName) {
    sheetName = workbook.SheetNames[0];
  }

  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error(`Could not find any sheet in the uploaded workbook.`);
  }

  // Convert sheet to array of objects (first row = headers)
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (rawRows.length === 0) {
    throw new Error(`The sheet "${sheetName}" appears to be empty.`);
  }

  // Extract all headers from the first row
  const headers = Object.keys(rawRows[0]);

  // Validate required columns
  const missing = REQUIRED_COLUMNS.filter(col =>
    !headers.some(h => h.trim() === col)
  );

  if (missing.length > 0) {
    throw new Error(
      `Missing required column${missing.length > 1 ? 's' : ''}: ${missing.join(', ')}. ` +
      `Please ensure your Excel file contains these columns.`
    );
  }

  // Identify trailer columns (columns whose names start with "Trailer")
  const trailerColumns = headers
    .filter(h => /^trailer\s*/i.test(h.trim()))
    .sort((a, b) => {
      // Sort by the number after "Trailer" — e.g. "Trailer 1" < "Trailer 2"
      const numA = parseInt(a.replace(/\D+/g, ''), 10) || 0;
      const numB = parseInt(b.replace(/\D+/g, ''), 10) || 0;
      return numA - numB;
    });

  if (trailerColumns.length === 0) {
    throw new Error(
      'No trailer columns found. Expected columns named "Trailer 1", "Trailer 2", etc.'
    );
  }

  return {
    rows: rawRows,
    headers,
    trailerColumns,
    sheetName,
    totalRows: rawRows.length,
  };
}
