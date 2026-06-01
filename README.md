# 📦 Nissi Fulfillment and Distribution — Excel Processing Engine

A client-side web application that processes Excel workbooks for Nissi's fulfillment and distribution workflow. Upload trailer spreadsheets, configure sort parameters, and download cleaned output files with intelligent SORT code allocation — all running entirely in the browser.

---

## ✨ Features

- **Drag & drop Excel upload** with column auto-detection
- **Dynamic trailer column selection** for flexible Units calculation
- **Real-time SORT range expansion** with live capacity preview
- **4-stage hierarchical SORT allocation** algorithm guaranteeing unique combinations
- **Regular Scan row detection** with automatic M-O-O-S assignment
- **Capacity validation** preventing overflows before processing
- **Values-only .xlsx output** — no formulas, clean column mapping
- **100% client-side** — zero server, zero data leaves the browser

---

## 🏗️ Architecture

```
┌────────────────────────────────────────────────────┐
│                    Browser                         │
│                                                    │
│   ┌──────────┐    ┌──────────────┐    ┌────────┐  │
│   │  Parser  │───▶│  Processor   │───▶│ Writer │  │
│   │(SheetJS) │    │  (Engine)    │    │(SheetJS)│  │
│   └──────────┘    └──────────────┘    └────────┘  │
│        ▲                                   │       │
│        │           ┌──────────┐            ▼       │
│   .xlsx Upload ◀───│  App.js  │───▶ .xlsx Download │
│                    │(UI Logic)│                     │
│                    └──────────┘                     │
└────────────────────────────────────────────────────┘
```

### File Structure

```
nissi-app/
├── index.html          # Main page — wizard UI with animated background
├── css/
│   └── styles.css      # Light blue theme, card layout, animations
├── js/
│   ├── app.js          # Application controller — step navigation, events
│   ├── parser.js       # Excel file parser — SheetJS wrapper
│   ├── processor.js    # Core processing engine — all business logic
│   └── writer.js       # Excel output generator — .xlsx creation
└── README.md           # This file
```

---

## 🧠 Processing Engine — Technical Deep Dive

The processing engine (`processor.js`) implements all the business logic for the NISSI workflow. Here's how each stage works:

### 1. Excel Parsing (parser.js)

Uses **[SheetJS (xlsx)](https://sheetjs.com/)** to read uploaded Excel workbooks entirely in-browser:

- Reads the file as an `ArrayBuffer` via the File API
- Locates the "Article Detail" sheet (or falls back to the first sheet)
- Converts the sheet to an array of JavaScript objects using `XLSX.utils.sheet_to_json()`
- **Auto-detects trailer columns** by scanning headers for the pattern `/^Trailer\s+/i`
- **Validates required columns** exist: `SWXref`, `BC`, `SKU NUM`

### 2. Range Expansion Algorithm

User inputs like `A-E`, `1-6`, or `A-C,X,Y` are expanded into arrays of individual values:

```
"A-E"       → ["A", "B", "C", "D", "E"]
"1-6"       → ["1", "2", "3", "4", "5", "6"]
"A-C,X,Y"   → ["A", "B", "C", "X", "Y"]
"blank"     → []  (SORT4 unused)
```

The algorithm:
1. Split input by commas into segments
2. For each segment, check if it matches the range pattern `X-Y`
3. If alphabetic range → expand via char codes (`charCodeAt` / `fromCharCode`)
4. If numeric range → expand via integer iteration
5. Deduplicate while preserving order

### 3. Units Calculation

Units = sum of **only the user-selected trailer columns** for each row. Non-selected trailers are completely ignored.

```javascript
// For selected trailers ["Trailer 1", "Trailer 3"]:
// Row: { "Trailer 1": 5, "Trailer 2": 100, "Trailer 3": 12 }
// Units = 5 + 12 = 17  (Trailer 2 ignored)
```

### 4. Row Filtering & Pattern Recognition

Rows are classified into three categories:

| Category | Condition | Action |
|----------|-----------|--------|
| **Normal (keep)** | Units ≥ 10, not Regular Scan | Kept for SORT allocation |
| **Normal (remove)** | Units < 10, not Regular Scan | Removed from output |
| **Regular Scan** | Any cell contains "Regular Scan" | Always kept, assigned M-O-O-S |

**Regular Scan detection** scans every cell value in a row for the case-insensitive substring `"regular scan"`. This pattern-matching approach handles the text appearing in any column (description, notes, etc.).

### 5. Capacity Validation

Before allocating SORT values, the engine validates:

```
Capacity = |SORT1| × |SORT2| × |SORT3| × |SORT4|
```

If `normal_row_count > capacity`, processing stops with a descriptive error:

> *"The selected sort pattern creates 1,440 unique combinations, but the file requires 1,625 rows. Please expand one or more SORT ranges."*

### 6. 4-Stage SORT Allocation Algorithm ⭐

This is the core algorithm — a **hierarchical round-robin distribution** that assigns SORT codes in stages, ensuring even distribution of unit volume across sort groups.

**Why not simple sequential assignment?** Sequential codes (AAAA, AAAB, AAAC...) would cluster high-volume items together. The staged approach distributes them evenly.

#### Stage 1: Assign SORT1
```
1. Sort ALL normal rows by Units DESCENDING
2. Cycle through SORT1 values in order
   Row 1 (highest units) → A
   Row 2 → B
   Row 3 → C
   Row 4 → D
   Row 5 → E
   Row 6 → A  (cycle repeats)
   Row 7 → B
   ...
```

#### Stage 2: Assign SORT2
```
1. Re-sort by SORT1 ASC, then Units DESC
2. Within EACH SORT1 group, cycle SORT2 values independently
   Group A: highest→1, next→2, next→3, ..., 6, back to 1
   Group B: highest→1, next→2, ...
   (each group starts its own cycle)
```

#### Stage 3: Assign SORT3
```
1. Re-sort by SORT1 ASC, SORT2 ASC, then Units DESC
2. Within each (SORT1, SORT2) group, cycle SORT3 values
```

#### Stage 4: Assign SORT4
```
1. Re-sort by SORT1 ASC, SORT2 ASC, SORT3 ASC, then Units DESC
2. Within each (SORT1, SORT2, SORT3) group, cycle SORT4 values
```

**Uniqueness guarantee**: Since each stage subdivides groups from the previous stage, and values cycle within groups, every (SORT1, SORT2, SORT3, SORT4) combination is unique — as long as total rows ≤ capacity (validated in step 5).

**Stable sorting**: The algorithm uses a tagged stable sort (elements maintain their relative order when comparison values are equal) to ensure deterministic results.

### 7. Output Generation (writer.js)

Uses SheetJS to create the final `.xlsx` file:

- Builds an array of output objects with exactly 9 columns in order:
  `Old_Barcode`, `BC`, `Article`, `Units`, `SORT1`, `SORT2`, `SORT3`, `SORT4`, `ScanCount`
- Column mapping: `SWXref → Old_Barcode`, `SKU NUM → Article`, `BC → BC (trimmed)`
- `ScanCount` = 0 for every row where SORT1 has data
- Regular Scan rows: SORT1=M, SORT2=O, SORT3=O, SORT4=S
- Output is **values only** (no formulas, no shared string table)
- Triggers a browser download via `Blob` + `URL.createObjectURL`

---

## 🚀 Getting Started

### Prerequisites
- A modern web browser (Chrome, Firefox, Safari, Edge)
- No server, no Node.js, no build step required

### Run Locally

```bash
# Clone the repo
git clone https://github.com/saadhaniftaj/Nissi-Trailer-Sort.git
cd Nissi-Trailer-Sort

# Serve with any static file server
python3 -m http.server 8080
# or
npx serve .
```

Open **http://localhost:8080** in your browser.

### Usage

1. **Upload** your `.xlsx` workbook containing the Article Detail sheet
2. **Select** which Trailer columns to include in the Units sum
3. **Configure** SORT1–SORT4 character ranges (e.g., A-E, 1-6)
4. **Process** and download the cleaned output file

---

## ✅ Acceptance Tests

| Test | Expected Result |
|------|----------------|
| Select Trailer X and Y only | Units = sum of those two columns only |
| Normal row with Units = 9 | Row removed from output |
| Normal row with Units = 10 | Row kept in output |
| Regular Scan row with Units = 0 | Row kept; SORT = M-O-O-S |
| Normal rows exceed SORT capacity | Error shown before processing |
| BC column has extra spaces | Trimmed in output |
| Output columns | Old_Barcode, BC, Article, Units, SORT1–4, ScanCount |

---

## 🛠️ Tech Stack

| Component | Technology |
|-----------|-----------|
| UI | HTML5 + CSS3 + Vanilla JS |
| Excel I/O | [SheetJS (xlsx)](https://sheetjs.com/) via CDN |
| Processing | Pure JavaScript (ES Modules) |
| Hosting | Any static file server |
| Dependencies | Zero npm packages |

---

## 📄 License

© 2026 Nissi Fulfillment and Distribution. All rights reserved.
