---
name: saber-invoice-reconciliation-skill
description: >
    Reconciles a consolidated Excel sheet against multiple mother files (usually PDFs, one per invoice) by comparing key fields such as HS Code, Product Name, Quantity, Unit Price, Country of Origin (COO), and Total Amount. The mother files are the source of truth; the Excel is what gets checked for errors. Use this skill whenever the user wants to cross-check, compare, verify, or reconcile invoices — trigger on phrases like: "compare invoices to excel", "check if invoices match", "verify invoice amounts", "does the PDF match the spreadsheet", or any mention of a mother folder, mother files, or files to be reconciled, or the name of the skill, or when the word "saber" is used in the context of comparing or reconciling documents.
---



## Step 1 — Identify the Files

Locate the two folders:

- **Mother folder** — the source of truth. Usually contains one PDF per invoice, but may contain Excel files instead. May contain multiple files — if so, identify what links them (almost always the Invoice Number).
- **Comparison folder** — the file(s) to be checked for errors. Almost always a single consolidated Excel file, but may contain multiple. If multiple, identify the common thread (usually Invoice Number) before proceeding.

**File location order (try in this order):**
1. Folder paths mentioned by the user / mounted by the user
2. Files uploaded directly in the conversation
3. If neither is clear, ask the user

**Matching mother files to Excel rows:**
- The Invoice Number is the primary key linking the two folders
- Extract the invoice number from **inside** the mother file (header section of the PDF or Excel)
- Cross-reference with the **filename** of the invoice as a secondary hint — filenames may contain the invoice number but not always clearly (e.g. `900370171-804183370-OUNASS-KSA.pdf`)
- Use **fuzzy matching** when comparing invoice numbers — OCR on scanned PDFs can introduce small errors (e.g. `1014000051` vs `1O14000051`)
- If no match is found for an Excel invoice number even after fuzzy matching → **flag it in the report and move on**

**Reading PDFs:**
- Use `pdfplumber` first; fall back to Tesseract OCR if output is empty or garbled
- Line item structure varies — scan the full document before extracting, don't assume a fixed layout
- Watch for OCR character swaps: `0/O`, `1/l`, `5/S`, `8/B`
- Number formats vary by client (e.g. European `1.740,00` vs `1,740.00`) — parse with judgment
 



## Step 2 - Understand What To Compare

Whatever columns exist in the Excel in the Comparison Folder — those are what need to be checked. If a field exists in the Excel but cannot be found in the mother file, flag it rather than skipping it.

Common fields in invoice reconciliation (examples only — the actual columns in the Excel are what drive the comparison, not this list)::
- Delivery Number
- Invoice Number
- HS Code / Tariff Code / Harminized Code
- Product Name / Description
- Country of Origin (COO)
- Quantity (QTY)
- Unit Price
- Currency
- Total Amount
- Destination

## Step 3 — Read and Normalize
 
**Excel:** Use `openpyxl` (for .xlsx) or `pandas` (for .csv). Handle merged cells — unmerge and forward-fill before processing. Read the header row first to understand what columns are present.
 
**Mother files:** Extract all fields that correspond to Excel columns. For each invoice, scan the full document — header fields (Delivery Number, COO, Destination) may appear anywhere, not just beside line items.
 
**Dependencies:** `pdfplumber`, `openpyxl`, `pandas`, `pytesseract`, `pdf2image`, `Pillow`, `thefuzz` — check availability and install only if missing.
 
 
## Step 4 — Match Rows
 
Matching happens at two levels:
 
1. **Document level** — match each Excel invoice number to its corresponding mother file. Use fuzzy matching if required
2. **Row level** — within a matched document, build a composite key from available columns. Start with HS Code + Product Name + QTY. If that's not unique enough to distinguish rows, add Unit Price. Use judgment — the goal is reliable matching, not a fixed formula.
 

**Matching rules:**
- If OCR was used on the mother file, use fuzzy matching to account for character-level errors
- If both sources are clean digital (pdfplumber text + Excel), require an exact match

If a row in the Excel cannot be matched to any line item in the mother file → flag as UNMATCHED.
If a line item exists in the mother file but not in the Excel → flag as MISSING.
 

## Step 5 — Compare and Flag
 
For each matched pair of rows, compare every Excel column against the mother file value.
 
| Status | Meaning | Color |
|---|---|---|
| ✅ MATCH | Values are equal | Green |
| ❌ MISMATCH | Values differ | Red |
| ⚠️ UNMATCHED | Excel row has no corresponding line in mother file | Orange |
| ❓ MISSING | Line item exists in mother file but not in Excel | Yellow |
 
Use fuzzy matching for text fields when OCR was involved. For numeric fields, parse carefully accounting for format differences (e.g. `1.740,00` vs `1,740.00`) — values must be numerically equal after parsing, do not round.

## Step 6 — Output: Reconciliation Report (Excel)
 
### One sheet per invoice
 
Side-by-side layout:
 
```
| [Excel columns...] | — | [Mother file columns...] | Status |
```
 
- Left block: Excel values
- Divider column (labelled `—`) for visual separation
- Right block: corresponding mother file values (same column labels, prefixed `[Source]`)
- Final column: Status (MATCH / MISMATCH / UNMATCHED / MISSING)
 
**Row rules:**
- Excel has data, mother file doesn't → right block cells empty, status = UNMATCHED
- Mother file has data, Excel doesn't → left block cells empty, status = MISSING
 
**Cell color coding** (applies to the Status column and mismatched value cells):
- 🟢 Green: MATCH
- 🔴 Red: MISMATCH
- 🟠 Orange: UNMATCHED
- 🟡 Yellow: MISSING
 
### Final sheet: "Missing From Excel" (only if applicable)
 
Include this sheet only if there are line items in any mother file that have no corresponding row in the Excel. Show the mother file values with empty Excel columns and MISSING status.
 
---
 
After generating the report, provide the summary in chat — not in the Excel file. Include a simple table:
 
| Invoice | Total Rows | Matches | Mismatches | Unmatched | Missing | Result |
|---|---|---|---|---|---|---|
| INV-001 | 12 | 11 | 1 | 0 | 0 | ❌ FAIL |
| INV-002 | 8 | 8 | 0 | 0 | 0 | ✅ PASS |
 
Follow with a short plain-language note on any invoices that failed — what the issues were and which fields were affected.
