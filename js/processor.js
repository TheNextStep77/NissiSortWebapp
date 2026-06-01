/* ============================================================
   NISSI Processing Engine
   ─────────────────────────────────────────────────────────────
   Handles: range expansion, units calculation, row filtering,
   capacity validation, and the 4-stage SORT allocation.
   ============================================================ */

/**
 * Expand a user-provided range string into an array of individual values.
 *
 * Accepted formats:
 *   "A-E"       → ["A","B","C","D","E"]
 *   "1-6"       → ["1","2","3","4","5","6"]
 *   "A,B,C"     → ["A","B","C"]
 *   "A-C,X,Y"   → ["A","B","C","X","Y"]
 *   "blank"     → []
 *   ""          → []
 *
 * @param {string} input - The raw range string from the user.
 * @returns {string[]} Array of individual allowed values (deduplicated).
 */
export function expandRange(input) {
  if (!input || input.trim() === '') return [];

  const trimmed = input.trim().toLowerCase();
  if (trimmed === 'blank' || trimmed === 'none' || trimmed === 'n/a') return [];

  const segments = input.split(',').map(s => s.trim()).filter(Boolean);
  const values = [];

  for (const segment of segments) {
    const rangeParts = segment.split('-').map(s => s.trim());

    if (rangeParts.length === 2 && rangeParts[0].length === 1 && rangeParts[1].length === 1) {
      const [startChar, endChar] = rangeParts;
      const isAlpha = /^[A-Za-z]$/.test(startChar) && /^[A-Za-z]$/.test(endChar);
      const isNum   = /^\d$/.test(startChar) && /^\d$/.test(endChar);

      if (isAlpha) {
        const s = startChar.toUpperCase().charCodeAt(0);
        const e = endChar.toUpperCase().charCodeAt(0);
        const step = s <= e ? 1 : -1;
        for (let c = s; step > 0 ? c <= e : c >= e; c += step) {
          values.push(String.fromCharCode(c));
        }
      } else if (isNum) {
        const s = parseInt(startChar, 10);
        const e = parseInt(endChar, 10);
        const step = s <= e ? 1 : -1;
        for (let n = s; step > 0 ? n <= e : n >= e; n += step) {
          values.push(String(n));
        }
      } else {
        // Not a valid range — treat the whole segment as a literal value
        values.push(segment.toUpperCase());
      }
    } else if (rangeParts.length === 1) {
      values.push(rangeParts[0].toUpperCase());
    } else {
      values.push(segment.toUpperCase());
    }
  }

  // Deduplicate while preserving order
  return [...new Map(values.map(v => [v, v])).values()];
}


/**
 * Calculate Units by summing only the user-selected trailer columns.
 *
 * @param {Object[]} rows - Array of row objects from the parsed Excel.
 * @param {string[]} selectedTrailers - Column names to include, e.g. ["Trailer 1", "Trailer 3"].
 * @returns {Object[]} Rows with an added `__units` property.
 */
export function calculateUnits(rows, selectedTrailers) {
  return rows.map(row => {
    let units = 0;
    for (const col of selectedTrailers) {
      const raw = row[col];
      const num = parseFloat(raw);
      if (!isNaN(num)) units += num;
    }
    return { ...row, __units: units };
  });
}


/**
 * Determine whether a row is a "Regular Scan" row.
 * Checks every value in the row for the substring "Regular Scan" (case-insensitive).
 *
 * @param {Object} row - A single row object.
 * @returns {boolean}
 */
export function isRegularScan(row) {
  for (const key of Object.keys(row)) {
    if (key.startsWith('__')) continue; // skip internal fields
    const val = row[key];
    if (typeof val === 'string' && val.toLowerCase().includes('regular scan')) {
      return true;
    }
  }
  return false;
}


/**
 * Filter rows into normal rows and Regular Scan rows.
 *
 * - Normal rows with Units < 10 are removed.
 * - Normal rows with Units >= 10 are kept.
 * - Regular Scan rows are always kept (regardless of Units).
 *
 * @param {Object[]} rows - Rows with `__units` already calculated.
 * @returns {{ normalRows: Object[], regularScanRows: Object[], removedCount: number }}
 */
export function filterRows(rows) {
  const normalRows = [];
  const regularScanRows = [];
  let removedCount = 0;

  for (const row of rows) {
    if (isRegularScan(row)) {
      regularScanRows.push({ ...row, __isRegularScan: true });
    } else if (row.__units >= 10) {
      normalRows.push({ ...row });
    } else {
      removedCount++;
    }
  }

  return { normalRows, regularScanRows, removedCount };
}


/**
 * Validate that the number of normal rows doesn't exceed available SORT combinations.
 *
 * @param {number} normalRowCount
 * @param {string[]} s1 - Expanded SORT1 values.
 * @param {string[]} s2 - Expanded SORT2 values.
 * @param {string[]} s3 - Expanded SORT3 values.
 * @param {string[]} s4 - Expanded SORT4 values.
 * @returns {{ valid: boolean, capacity: number, message: string }}
 */
export function validateCapacity(normalRowCount, s1, s2, s3, s4) {
  const c1 = s1.length || 1;
  const c2 = s2.length || 1;
  const c3 = s3.length || 1;
  const c4 = s4.length > 0 ? s4.length : 1;
  const capacity = c1 * c2 * c3 * c4;

  if (normalRowCount > capacity) {
    return {
      valid: false,
      capacity,
      message:
        `The selected sort pattern creates ${capacity.toLocaleString()} unique combinations, ` +
        `but the file requires ${normalRowCount.toLocaleString()} rows. ` +
        `Please expand one or more SORT ranges.`,
    };
  }

  return { valid: true, capacity, message: '' };
}


/**
 * Staged SORT Allocation — the heart of the engine.
 *
 * Assigns SORT1–SORT4 to normal rows in four sequential stages.
 * Each stage re-sorts the data by previously assigned fields + Units DESC,
 * then cycles through the allowed values within each group.
 *
 * This approach guarantees unique SORT combinations as long as
 * rowCount <= capacity (validated beforehand).
 *
 * @param {Object[]} normalRows - Filtered normal rows with `__units`.
 * @param {string[]} s1Vals - Allowed SORT1 values.
 * @param {string[]} s2Vals - Allowed SORT2 values.
 * @param {string[]} s3Vals - Allowed SORT3 values.
 * @param {string[]} s4Vals - Allowed SORT4 values (empty array if unused).
 * @returns {Object[]} Rows with `__sort1`, `__sort2`, `__sort3`, `__sort4` assigned.
 */
export function allocateSorts(normalRows, s1Vals, s2Vals, s3Vals, s4Vals) {
  const rows = normalRows.map(r => ({ ...r })); // shallow copy

  // ── Step 1: Assign SORT1 ──────────────────────────────────
  // Sort by Units descending (stable).
  stableSort(rows, (a, b) => b.__units - a.__units);

  for (let i = 0; i < rows.length; i++) {
    rows[i].__sort1 = s1Vals[i % s1Vals.length];
  }

  // ── Step 2: Assign SORT2 ──────────────────────────────────
  // Sort by SORT1 ascending, then Units descending.
  stableSort(rows, (a, b) => {
    const cmp = compareSortValues(a.__sort1, b.__sort1);
    return cmp !== 0 ? cmp : b.__units - a.__units;
  });

  cycleWithinGroups(rows, r => r.__sort1, s2Vals, '__sort2');

  // ── Step 3: Assign SORT3 ──────────────────────────────────
  // Sort by SORT1 asc, SORT2 asc, then Units desc.
  stableSort(rows, (a, b) => {
    let cmp = compareSortValues(a.__sort1, b.__sort1);
    if (cmp !== 0) return cmp;
    cmp = compareSortValues(a.__sort2, b.__sort2);
    return cmp !== 0 ? cmp : b.__units - a.__units;
  });

  cycleWithinGroups(rows, r => r.__sort1 + '\x00' + r.__sort2, s3Vals, '__sort3');

  // ── Step 4: Assign SORT4 ──────────────────────────────────
  if (s4Vals.length > 0) {
    stableSort(rows, (a, b) => {
      let cmp = compareSortValues(a.__sort1, b.__sort1);
      if (cmp !== 0) return cmp;
      cmp = compareSortValues(a.__sort2, b.__sort2);
      if (cmp !== 0) return cmp;
      cmp = compareSortValues(a.__sort3, b.__sort3);
      return cmp !== 0 ? cmp : b.__units - a.__units;
    });

    cycleWithinGroups(
      rows,
      r => r.__sort1 + '\x00' + r.__sort2 + '\x00' + r.__sort3,
      s4Vals,
      '__sort4'
    );
  } else {
    for (const row of rows) {
      row.__sort4 = '';
    }
  }

  return rows;
}


/**
 * Apply Regular Scan fixed SORT values: M-O-O-S.
 *
 * @param {Object[]} regularScanRows
 * @returns {Object[]} Rows with SORT values set.
 */
export function applyRegularScanSorts(regularScanRows) {
  return regularScanRows.map(row => ({
    ...row,
    __sort1: 'M',
    __sort2: 'O',
    __sort3: 'O',
    __sort4: 'S',
  }));
}


/**
 * Apply the final output sort order:
 *   1. SORT1 ascending
 *   2. SORT2 ascending
 *   3. SORT3 ascending
 *   4. SORT4 ascending
 *   5. Units descending
 *
 * @param {Object[]} rows - All output rows (normal + Regular Scan).
 * @returns {Object[]} Sorted rows.
 */
export function applyFinalSort(rows) {
  return stableSort([...rows], (a, b) => {
    let cmp = compareSortValues(a.__sort1, b.__sort1);
    if (cmp !== 0) return cmp;
    cmp = compareSortValues(a.__sort2, b.__sort2);
    if (cmp !== 0) return cmp;
    cmp = compareSortValues(a.__sort3, b.__sort3);
    if (cmp !== 0) return cmp;
    cmp = compareSortValues(a.__sort4, b.__sort4);
    if (cmp !== 0) return cmp;
    return b.__units - a.__units;
  });
}


/**
 * Build the final output row objects with only the required columns,
 * renamed and in the correct order.
 *
 * @param {Object[]} rows - Fully processed rows.
 * @returns {Object[]} Clean output rows.
 */
export function buildOutputRows(rows) {
  return rows.map(row => {
    const bc = row['BC'] != null ? String(row['BC']).trim() : '';
    const scanCount = row.__sort1 ? 0 : '';

    return {
      Old_Barcode: row['SWXref'] ?? '',
      BC: bc,
      Article: row['SKU NUM'] ?? '',
      Units: row.__units ?? 0,
      SORT1: row.__sort1 ?? '',
      SORT2: row.__sort2 ?? '',
      SORT3: row.__sort3 ?? '',
      SORT4: row.__sort4 ?? '',
      ScanCount: scanCount,
    };
  });
}


/**
 * Full processing pipeline — orchestrates all steps.
 *
 * @param {Object[]} rawRows - Parsed rows from the Excel file.
 * @param {string[]} selectedTrailers - User-selected trailer column names.
 * @param {string} sort1Input - Raw SORT1 range input from user.
 * @param {string} sort2Input - Raw SORT2 range input from user.
 * @param {string} sort3Input - Raw SORT3 range input from user.
 * @param {string} sort4Input - Raw SORT4 range input from user.
 * @returns {{ success: boolean, data?: Object[], stats?: Object, error?: string }}
 */
export function processAll(rawRows, selectedTrailers, sort1Input, sort2Input, sort3Input, sort4Input) {
  // 1. Expand SORT ranges
  const s1 = expandRange(sort1Input);
  const s2 = expandRange(sort2Input);
  const s3 = expandRange(sort3Input);
  const s4 = expandRange(sort4Input);

  // Validate SORT1–3 are not empty
  if (s1.length === 0) return { success: false, error: 'SORT1 range cannot be empty. Please provide at least one value.' };
  if (s2.length === 0) return { success: false, error: 'SORT2 range cannot be empty. Please provide at least one value.' };
  if (s3.length === 0) return { success: false, error: 'SORT3 range cannot be empty. Please provide at least one value.' };

  // 2. Calculate Units
  const rowsWithUnits = calculateUnits(rawRows, selectedTrailers);

  // 3. Filter rows
  const { normalRows, regularScanRows, removedCount } = filterRows(rowsWithUnits);

  // 4. Validate capacity
  const validation = validateCapacity(normalRows.length, s1, s2, s3, s4);
  if (!validation.valid) {
    return { success: false, error: validation.message };
  }

  // 5. Staged SORT allocation for normal rows
  const allocatedNormal = allocateSorts(normalRows, s1, s2, s3, s4);

  // 6. Apply Regular Scan SORT values
  const allocatedRegScan = applyRegularScanSorts(regularScanRows);

  // 7. Combine and apply final sort
  const combined = [...allocatedNormal, ...allocatedRegScan];
  const sorted = applyFinalSort(combined);

  // 8. Build output rows
  const output = buildOutputRows(sorted);

  return {
    success: true,
    data: output,
    stats: {
      totalInputRows: rawRows.length,
      normalRowsKept: allocatedNormal.length,
      regularScanRows: allocatedRegScan.length,
      rowsRemoved: removedCount,
      totalOutputRows: output.length,
      capacity: validation.capacity,
      sort1Values: s1.join(', '),
      sort2Values: s2.join(', '),
      sort3Values: s3.join(', '),
      sort4Values: s4.length > 0 ? s4.join(', ') : '(unused)',
    },
  };
}


/* ── Internal Helpers ─────────────────────────────────────── */

/**
 * Stable sort using the built-in Array.sort (stable in all modern engines since ES2019).
 * Falls back to a merge sort if needed for older environments.
 */
function stableSort(arr, compareFn) {
  // Tag each element with its original index for stability guarantee
  const tagged = arr.map((el, i) => ({ el, i }));
  tagged.sort((a, b) => {
    const cmp = compareFn(a.el, b.el);
    return cmp !== 0 ? cmp : a.i - b.i; // stable tie-break
  });
  for (let k = 0; k < arr.length; k++) {
    arr[k] = tagged[k].el;
  }
  return arr;
}


/**
 * Compare two SORT values for ascending order.
 * Handles both alphabetic and numeric single-character values correctly.
 */
function compareSortValues(a, b) {
  if (a === b) return 0;
  if (a === '' || a == null) return -1;
  if (b === '' || b == null) return 1;

  // If both are single digits, compare numerically
  const aNum = Number(a);
  const bNum = Number(b);
  if (!isNaN(aNum) && !isNaN(bNum)) return aNum - bNum;

  // Otherwise, string comparison
  return a < b ? -1 : a > b ? 1 : 0;
}


/**
 * Cycle through allowed values within each group, assigning to the target field.
 *
 * @param {Object[]} rows - Sorted rows.
 * @param {Function} groupKeyFn - Function that returns the group key for a row.
 * @param {string[]} values - Allowed values to cycle through.
 * @param {string} targetField - The internal field name to assign (e.g. '__sort2').
 */
function cycleWithinGroups(rows, groupKeyFn, values, targetField) {
  let currentGroup = null;
  let idx = 0;

  for (const row of rows) {
    const key = groupKeyFn(row);
    if (key !== currentGroup) {
      currentGroup = key;
      idx = 0;
    }
    row[targetField] = values[idx % values.length];
    idx++;
  }
}
