---
name: doorcrashers-item-report-updater
description: Cleans a flyer Item Report and updates it from a Doorcrashers source file. Use when a user supplies an Item Report CSV/Excel plus a Doorcrashers links file and wants page items cleaned, SKUs matched to fence IDs, and names and URLs populated without guessing.
---

# Doorcrashers Item Report Updater

This skill is a thin wrapper around `doorcrashers_updater.py` in this repo. **Always run the script; do not hand-edit the CSV.** The script does the work the same way every time, which is why it is safe for untrained users and for any AI tool (Claude, Gemini, ChatGPT, ...) or no AI at all.

## When to use

- The user supplies two CSV/Excel files: an **Item Report** (filename contains "Item Report"; has a `sku` column) and a **Doorcrashers source file** (any filename; has `fence_id`, `friendly_name` or `name`, `url` columns).
- The user asks to clean, reconcile or merge flyer / doorcrasher data.

## How to run

```bash
python3 doorcrashers_updater.py "<file 1>" "<file 2>" [-o OUTPUT_DIR]
```

Files can be given in either order. The Item Report is the file with "Item Report" in its name; the other file (any name) is the source, and must have `fence_id`, `friendly_name` (or `name`) and `url` columns. Needs Python 3.8+ only (`openpyxl` only if reading `.xlsx`).

## What the script does

1. **Clean the Item Report**: keep the header row; delete the `PAGE ITEMS` row; delete the `FLYER ITEMS` row and everything below it; strip trailing blank rows. All columns and their names are kept.
2. **Match**: Item Report `sku` == source `fence_id`, after trimming whitespace and removing BOM characters. Exact match only: no fuzzy matching, no assuming row order. A source `fence_id` that appears twice is reported, not guessed.
3. **Update**: clear every value in `name`, then fill `name` from the source `friendly_name` (or `name`) and `url` from the source `url` for matched rows. All other columns are left untouched.
4. **Output**: `[Source Prefix] - Output - [Report Name] - Cleaned.csv`. The prefix is the source filename as-is (e.g. `WK44_CA_Doorcrashers_Links`).

## Reporting back to the user

Give the user the output file and paste the script's task summary as-is. Row numbers are spreadsheet rows (header = row 1) in the cleaned Item Report and in the source file.

Perfect 1:1 match:

```
Task summary:
- X SKUs successfully matched
```

With discrepancies:

```
Task summary:
- X SKUs successfully matched
- Y SKU(s) did not match a fence_id provided:
  - <SKU> (row <row_number>)
- Z fence_id(s) were provided that did not match SKUs provided:
  - <fence_id> (row <row_number>)
```

Never fix a discrepancy by guessing. Tell the user and let them correct the source.
