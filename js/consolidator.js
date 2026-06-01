/* ============================================================
   NISSI Scan Consolidator
   ─────────────────────────────────────────────────────────────
   Reads multiple text/notepad files, merges all scanned entries
   into one Excel workbook with two sheets:
     Sheet 1 ("Raw Data")    — every entry, one per row, with source file
     Sheet 2 ("Summary")     — unique entries + count of occurrences
   ============================================================ */

/**
 * Read a single text file and return an array of non-empty trimmed lines.
 * Handles Windows (\r\n), Unix (\n), and old Mac (\r) line endings.
 * Strips BOM if present.
 *
 * @param {File} file
 * @returns {Promise<string[]>}
 */
async function readTextFile(file) {
  let text = await file.text();

  // Strip BOM (byte order mark) that appears in some Notepad files
  if (text.charCodeAt(0) === 0xFEFF) {
    text = text.slice(1);
  }

  return text
    .split(/\r?\n|\r/)
    .map(line => line.trim())
    .filter(line => line.length > 0);
}

/**
 * Process multiple text files into a consolidated Excel workbook.
 *
 * @param {FileList|File[]} files — the uploaded .txt files
 * @returns {{
 *   workbookBlob: Blob,
 *   totalEntries: number,
 *   uniqueEntries: number,
 *   fileCount: number,
 *   topEntry: { value: string, count: number } | null
 * }}
 */
export async function consolidateTextFiles(files) {
  const XLSX = window.XLSX;
  if (!XLSX) {
    throw new Error('SheetJS library (XLSX) is not loaded. Please check your internet connection and refresh.');
  }

  if (!files || files.length === 0) {
    throw new Error('No files were provided. Please select at least one text file.');
  }

  // ── 1. Read all files ──────────────────────────────────
  const allEntries = [];          // { value, source }
  const fileNames = [];

  for (const file of files) {
    const lines = await readTextFile(file);
    const shortName = file.name;
    fileNames.push(shortName);

    for (const line of lines) {
      allEntries.push({ value: line, source: shortName });
    }
  }

  if (allEntries.length === 0) {
    throw new Error('All selected files appear to be empty. No entries found.');
  }

  // ── 2. Build frequency map ─────────────────────────────
  const freqMap = new Map();
  for (const entry of allEntries) {
    freqMap.set(entry.value, (freqMap.get(entry.value) || 0) + 1);
  }

  // Sort summary by count descending, then value ascending for ties
  const summary = Array.from(freqMap.entries())
    .map(([value, count]) => ({ value, count }))
    .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value));

  // ── 3. Build Sheet 1 — Raw Data ────────────────────────
  const rawRows = allEntries.map(e => ({
    'Entry': e.value,
    'Source File': e.source,
  }));

  const ws1 = XLSX.utils.json_to_sheet(rawRows);

  // Auto-size columns (rough estimate)
  ws1['!cols'] = [
    { wch: Math.max(14, ...rawRows.slice(0, 200).map(r => r['Entry'].length)) },
    { wch: Math.max(12, ...fileNames.map(n => n.length)) },
  ];

  // ── 4. Build Sheet 2 — Summary ─────────────────────────
  const summaryRows = summary.map(s => ({
    'Entry': s.value,
    'Count': s.count,
  }));

  const ws2 = XLSX.utils.json_to_sheet(summaryRows);

  ws2['!cols'] = [
    { wch: Math.max(14, ...summary.slice(0, 200).map(s => s.value.length)) },
    { wch: 8 },
  ];

  // ── 5. Assemble workbook ───────────────────────────────
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, 'Raw Data');
  XLSX.utils.book_append_sheet(wb, ws2, 'Summary');

  // Write to binary
  const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbOut], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  return {
    workbookBlob: blob,
    totalEntries: allEntries.length,
    uniqueEntries: summary.length,
    fileCount: files.length,
    topEntry: summary.length > 0 ? summary[0] : null,
  };
}
