# excel_itemizer_docx_tables

Generate clean, print-ready **Word (.docx)** reports from an **Excel (.xlsx/.xls)** sheet, automatically itemized into tidy tables (with optional grouping, sorting, and totals).

---

## What It Does

- Reads a source Excel workbook (single sheet or by name).
- Normalizes headers and data types (dates, numbers, text).
- Builds one or more **Word tables** with an optional title and metadata block.
- (Optionally) groups rows (e.g., by Category/Client/Month) and adds per‑group subtotals and a grand total.
- Applies a Word table style so the output looks consistent and professional.

> Ideal for statements, movement logs, purchase/item lists, or any “itemized” report your clients expect as a .docx.

---

## Libraries Used

This script relies on well-supported Python packages:

- **pandas** – fast, reliable data handling and column ops  
- **openpyxl** – Excel engine for reading `.xlsx` files (used by pandas)  
- **python-docx** – create and style the Word `.docx` document  

Optional (only if you need to read legacy `.xls` files):

- **xlrd** – engine for old Excel `.xls` files (not needed for `.xlsx`)

---

## Requirements

- **Python**: 3.9+ (3.10 or newer recommended)
- **Pip packages**:

```bash
pip install -U pandas openpyxl python-docx
# optional, only for legacy .xls:
pip install xlrd
```

Or use `requirements.txt`:

```
pandas>=2.0
openpyxl>=3.1
python-docx>=1.1.0
xlrd>=2.0 ; python_version >= "3.9"  # optional, .xls only
```

---

## Usage

Basic:

```bash
python excel_itemizer_docx_tables.py INPUT.xlsx OUTPUT.docx
```

Common options (if your version of the script exposes CLI flags):

```
--sheet SHEET_NAME              # read a specific sheet (default: first sheet)
--columns COL1 COL2 ...         # choose column order in the Word table
--group-by COL                  # group rows by this column; add per-group tables/totals
--sort-by COL[:asc|:desc]       # sort before building tables (e.g., "Date:asc", "Amount:desc")
--date-format "%Y-%m-%d"        # enforce a date display format
--table-style "Light Shading Accent 1"  # Word table style to apply
--title "My Report Title"       # title placed at the top of the document
--meta "Client=Acme, Month=Aug 2025"    # key=value pairs under the title
--currency-cols Amount,Total    # format these numeric columns as currency
--autofit                       # auto fit columns to content
--landscape                     # page orientation (default: portrait)
--margins "0.7,0.7,0.7,0.7"     # inches: top,right,bottom,left
```

> If your copy of the script doesn’t include some flags above, you can still run the **basic form**; those options are purely for convenience and formatting.

---

## Input Expectations

- **Headers**: first row in the sheet should be column names (e.g., `Date, Description, Category, Quantity, Unit Price, Amount`).
- **Dates**: stored as real Excel dates or ISO strings (the script attempts to parse).
- **Numbers**: amounts/quantities should be numeric (no currency symbols in the cells is best).
- **Empty rows**: are ignored.

---

## Output

- A `.docx` file containing:
  - An optional **title** and **metadata** block.
  - One or more **tables**:
    - If `--group-by` is used: a titled table per group, each with a **subtotal** row, plus a **grand total**.
    - If not grouped: a single table (optionally with a total row).
  - Consistent **table style** (editable in Word after export).

---

## Examples

Create a single table, use specific columns and a nice style:

```bash
python excel_itemizer_docx_tables.py data/movements.xlsx out/report.docx   --columns Date Description Category Amount   --sort-by Date:asc   --table-style "Light Shading Accent 1"   --title "Informe de Movimientos"   --meta "Cliente=ACME, Rango=Julio 2025"   --currency-cols Amount
```

Group by Category with subtotals and a grand total:

```bash
python excel_itemizer_docx_tables.py data/items.xlsx out/items_by_category.docx   --group-by Category   --columns Category Item Qty UnitPrice Amount   --sort-by Category:asc   --currency-cols Amount,UnitPrice
```

Read a specific sheet and print in landscape:

```bash
python excel_itemizer_docx_tables.py data/report_book.xlsx out/landscape.docx   --sheet "August"   --landscape   --autofit
```

---

## Table Styling Tips

- The default style (e.g., **“Light Shading Accent 1”**) is widely available in Word.
- If your Office install is localized, style names can differ. If a style name isn’t found, the script will fall back to the document’s default table style; you can change it later in Word.

---

## Troubleshooting

- **ModuleNotFoundError: pandas / python-docx / openpyxl**  
  Install the dependencies: `pip install pandas openpyxl python-docx`.

- **“Excel file is open / Permission denied” on Windows**  
  Close the Excel file (and the output Word file) before running.

- **“ValueError: Excel file format cannot be determined”**  
  Ensure the file extension matches the content (`.xlsx` vs `.xls`). For `.xls`, install `xlrd`.

- **Dates show as numbers**  
  Set `--date-format` (e.g., `--date-format "%d/%m/%Y"`) or ensure the column is parsed as a date in Excel.

- **Currency symbols duplicated**  
  Keep numeric cells **numeric** in Excel (no “$”), then tell the script which columns are currency via `--currency-cols`.

- **Columns out of order / missing**  
  Use `--columns` to explicitly control column order. Ensure headers in Excel match what you pass.

---

## Suggested Project Structure

```
.
├─ excel_itemizer_docx_tables.py
├─ requirements.txt
├─ data/
│  └─ sample.xlsx
└─ out/
   └─ report.docx
```

---

## Development Notes

- The script builds Word tables via `python-docx` APIs (`Document`, `add_table`, cell paragraphs/runs) and applies styles and alignments after data insertion for consistent layout.
- Numeric formatting (currency, thousands separators) is handled in Python to keep the Word file simple and portable.
- Grouping and totals are computed with `pandas` (e.g., `groupby`, `sum`, `sort_values`).

---

## License

MIT (or your preferred license). Add a `LICENSE` file if you plan to share.

---

## Changelog

- **v1.0** – Initial release: Excel → itemized Word tables; optional grouping, sorting, totals, and basic styling.
