# Doorcrashers Item Report Updater

Cleans a flyer **Item Report** and updates it from a **Doorcrashers links file** (run twice weekly or on demand). No Claude/AI licence needed: it is a small Python script with no dependencies.

## What it does

1. Removes the `PAGE ITEMS` row, and the `FLYER ITEMS` row and everything below it.
2. Matches Item Report `sku` to source `fence_id` (exact match, whitespace/BOM ignored, never guesses).
3. Clears `name`, then fills `name` (from `friendly_name`) and `url` for matched rows. All other columns stay as they were.
4. Saves one updated CSV and shows a summary of any mismatches.

## Run it (no coding)

1. Install Python 3.8+ from python.org (Windows: tick "Add Python to PATH").
2. Download this repo (Code -> Download ZIP) and unzip it.
3. **Windows:** drag your two files onto `run_windows.bat`. **Mac:** double-click `run_mac.command` and pick the two files (Cmd-click to select both).
4. The result appears next to your Item Report, e.g. `WK44_CA_Doorcrashers_Links - Output - Item Report-1202599-20261027-063830 - Cleaned.csv`. The task summary is shown in the window when it finishes.

(First time on Mac: right-click `run_mac.command` -> Open.)

## Run it (command line)

```bash
python3 doorcrashers_updater.py "Item Report-1202599-20261027-063830.csv" WK44_CA_Doorcrashers_Links.csv
```

Options: `-o DIR` output folder, `--prefix TEXT` filename prefix, `--excel-bom` write UTF-8 with BOM so Excel shows symbols like (R) correctly. `.xlsx` inputs need `pip install openpyxl`. Without arguments it opens a file picker.

## Reading the summary

```
Task summary:
- 31 SKUs successfully matched
- 1 SKU(s) did not match a fence_id provided:
  - P03-09 (row 12)
- 1 fence_id(s) were provided that did not match SKUs provided:
  - P03-90 (row 12)
```

Row numbers are spreadsheet rows (header = row 1) in the cleaned Item Report / the source file. Unmatched SKUs get a blank `name` and keep whatever `url` they had. Fix the source file and re-run.

## Notes

- Output filename: `[source filename] - Output - [Item Report name] - Cleaned.csv`.
- Using it with an AI assistant: `SKILL.md` is the skill definition (Claude-style frontmatter) that tells an assistant to run the script.
- Tests: `python3 -m unittest discover -s tests` (synthetic data only).
