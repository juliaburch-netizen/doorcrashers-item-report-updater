#!/usr/bin/env python3
"""Doorcrashers Item Report Updater.

Cleans a flyer Item Report and updates it from a Doorcrashers source/link file.

  * Removes the "PAGE ITEMS" marker row and everything from "FLYER ITEMS" down.
  * Matches Item Report `sku` to source `fence_id` (exact match only, no guessing).
  * Replaces `name` with the source `friendly_name` (or `name`) and fills `url`.
  * Writes one updated CSV and prints a task summary.

Needs only Python 3.8+ (standard library). Reading .xlsx/.xls files additionally
needs `openpyxl` (pip install openpyxl); CSV needs nothing.

Usage:
    python doorcrashers_updater.py <file1> <file2> [-o OUTPUT_DIR] [--prefix TEXT]
    python doorcrashers_updater.py            # no arguments: asks you to pick the files

The two files can be given in either order. The Item Report is the one with
"Item Report" in its filename; the other file (any name) is the source.
"""

import argparse
import csv
import re
import sys
from pathlib import Path

BOM = "﻿"
PAGE_MARKER = "PAGE ITEMS"
FLYER_MARKER = "FLYER ITEMS"


class UpdaterError(Exception):
    """A problem the user can fix (wrong file, missing column, ...)."""


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def norm(value):
    """Normalise a key/header for comparison: drop BOM, trim whitespace."""
    return ("" if value is None else str(value)).replace(BOM, "").strip()


def norm_header(value):
    return norm(value).lower()


def is_blank(row):
    return all(norm(c) == "" for c in row)


def _cell_to_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def read_table(path):
    """Read a CSV / XLSX / XLS file into a list of rows (lists of strings)."""
    path = Path(path)
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        try:
            import openpyxl
        except ImportError:
            raise UpdaterError(
                f"'{path.name}' is an Excel file. Install openpyxl "
                "(pip install openpyxl) or save the file as CSV and try again."
            )
        if ext == ".xls":
            raise UpdaterError(
                f"'{path.name}' is an old .xls file. Open it in Excel and "
                "'Save As' .xlsx or .csv, then try again."
            )
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        return [[_cell_to_text(c) for c in row] for row in ws.iter_rows(values_only=True)]

    for encoding in ("utf-8-sig", "cp1252"):
        try:
            with open(path, newline="", encoding=encoding) as fh:
                return [row for row in csv.reader(fh)]
        except UnicodeDecodeError:
            continue
    raise UpdaterError(f"Could not read '{path.name}' as text (unknown encoding).")


def header_index(header):
    """Map normalised header name -> column index (first occurrence wins)."""
    idx = {}
    for i, h in enumerate(header):
        idx.setdefault(norm_header(h), i)
    return idx


def first_nonblank_row(rows):
    for i, r in enumerate(rows):
        if not is_blank(r):
            return i
    raise UpdaterError("The file is empty.")


# --------------------------------------------------------------------------- #
# File identification
# --------------------------------------------------------------------------- #
def has_columns(rows, *needed):
    headers = {norm_header(h) for h in rows[first_nonblank_row(rows)]}
    return all(n in headers for n in needed)


def identify_files(path_a, path_b):
    """Return (report_path, report_rows, source_path, source_rows).

    The Item Report is the file with "Item Report" in its name. The other file is
    the source; its name does not matter, but its headers must include fence_id.
    """
    paths = [Path(path_a), Path(path_b)]
    reports = [p for p in paths if "item report" in p.name.lower()]
    if len(reports) != 1:
        raise UpdaterError(
            "Exactly one of the two files must have 'Item Report' in its filename "
            f"(found {len(reports)}). The other file can be named anything."
        )
    report_path = reports[0]
    source_path = paths[0] if paths[1] == report_path else paths[1]
    report_rows, source_rows = read_table(report_path), read_table(source_path)
    if not has_columns(report_rows, "sku"):
        raise UpdaterError(f"'{report_path.name}' has no 'sku' column, so it doesn't look like an Item Report.")
    if not has_columns(source_rows, "fence_id", "url"):
        raise UpdaterError(
            f"'{source_path.name}' is missing the columns it needs: fence_id, "
            "friendly_name (or name) and url."
        )
    return report_path, report_rows, source_path, source_rows


# --------------------------------------------------------------------------- #
# Step 1: clean the Item Report
# --------------------------------------------------------------------------- #
def _marker_row(row, marker):
    """True if the row's only non-empty cell is exactly `marker` (any case)."""
    filled = [norm(c) for c in row if norm(c) != ""]
    return len(filled) == 1 and filled[0].upper() == marker


def clean_report(rows):
    """Return (header, data_rows, warnings). Header is row 1 and is preserved."""
    start = first_nonblank_row(rows)
    header = [c.replace(BOM, "") for c in rows[start]]
    warnings = []
    body = []
    found_flyer = False
    for row in rows[start + 1:]:
        if _marker_row(row, FLYER_MARKER):
            found_flyer = True
            break
        if _marker_row(row, PAGE_MARKER):
            continue
        body.append(list(row))
    if not found_flyer:
        warnings.append(f"No '{FLYER_MARKER}' row was found, so no rows were cut off the bottom.")
    while body and is_blank(body[-1]):
        body.pop()
    width = len(header)
    body = [(r + [""] * width)[:width] if len(r) < width else r for r in body]
    return header, body, warnings


# --------------------------------------------------------------------------- #
# Step 2 + 3: match and update
# --------------------------------------------------------------------------- #
def load_source(rows):
    """Return (entries, name_col_label, warnings).

    entries: dict fence_id -> list of (row_number, name, url).
    """
    start = first_nonblank_row(rows)
    idx = header_index(rows[start])
    for needed in ("fence_id", "url"):
        if needed not in idx:
            raise UpdaterError(f"The source file has no '{needed}' column.")
    if "friendly_name" in idx:
        name_label = "friendly_name"
    elif "name" in idx:
        name_label = "name"
    else:
        raise UpdaterError("The source file needs a 'friendly_name' (or 'name') column.")

    def cell(row, key):
        i = idx[key]
        return row[i] if i < len(row) else ""

    entries = {}
    for offset, row in enumerate(rows[start + 1:], start=start + 2):  # 1-based row number
        if is_blank(row):
            continue
        key = norm(cell(row, "fence_id"))
        if key == "":
            continue
        entries.setdefault(key, []).append(
            (offset, cell(row, name_label).strip(), cell(row, "url").strip())
        )
    return entries, name_label, []


def update_report(header, data_rows, source_entries):
    """Apply source names/URLs. Returns (updated_rows, result dict)."""
    idx = header_index(header)
    for needed in ("sku", "name", "url"):
        if needed not in idx:
            raise UpdaterError(f"The Item Report has no '{needed}' column.")
    sku_i, name_i, url_i = idx["sku"], idx["name"], idx["url"]

    matched = 0
    unmatched_skus = []   # (sku, row_number)
    ambiguous = []        # (sku, row_number, source_rows)
    used_fences = set()
    updated = []

    for n, row in enumerate(data_rows, start=2):  # row 1 is the header
        row = list(row)
        row[name_i] = ""  # always clear existing names
        sku = norm(row[sku_i])
        hits = source_entries.get(sku) if sku else None
        if not hits:
            unmatched_skus.append((sku or "(blank)", n))
        elif len(hits) > 1:
            ambiguous.append((sku, n, [h[0] for h in hits]))
            used_fences.add(sku)
        else:
            _, name, url = hits[0]
            row[name_i] = name
            row[url_i] = url
            matched += 1
            used_fences.add(sku)
        updated.append(row)

    unmatched_fences = []
    for fence, hits in source_entries.items():
        if fence not in used_fences:
            unmatched_fences.extend((fence, h[0]) for h in hits)
    unmatched_fences.sort(key=lambda t: t[1])

    return updated, {
        "matched": matched,
        "unmatched_skus": unmatched_skus,
        "unmatched_fences": unmatched_fences,
        "ambiguous": ambiguous,
    }


# --------------------------------------------------------------------------- #
# Step 5: output names + summary
# --------------------------------------------------------------------------- #
def derive_prefix(source_path):
    """The source filename without its extension, used as-is."""
    return Path(source_path).stem


def cleaned_name(report_path):
    stem = Path(report_path).stem
    return stem if stem.lower().endswith("- cleaned") else f"{stem} - Cleaned"


def safe_filename(name):
    return re.sub(r'[<>:"/\\|?*]', "-", name).strip()


def build_summary(result, warnings):
    lines = ["Task summary:", f"- {result['matched']} SKUs successfully matched"]
    if result["unmatched_skus"]:
        lines.append(f"- {len(result['unmatched_skus'])} SKU(s) did not match a fence_id provided:")
        lines += [f"  - {s} (row {r})" for s, r in result["unmatched_skus"]]
    if result["unmatched_fences"]:
        lines.append(
            f"- {len(result['unmatched_fences'])} fence_id(s) were provided that did not match SKUs provided:"
        )
        lines += [f"  - {f} (row {r})" for f, r in result["unmatched_fences"]]
    if result["ambiguous"]:
        lines.append(
            f"- {len(result['ambiguous'])} SKU(s) matched more than one source row, "
            "so name and url were left blank (not guessed):"
        )
        lines += [
            f"  - {s} (row {r}; source rows {', '.join(map(str, src))})"
            for s, r, src in result["ambiguous"]
        ]
    if warnings:
        lines.append("Warnings:")
        lines += [f"  - {w}" for w in warnings]
    return "\n".join(lines)


def write_csv(path, header, rows, bom=False):
    with open(path, "w", newline="", encoding="utf-8-sig" if bom else "utf-8") as fh:
        w = csv.writer(fh, lineterminator="\n")
        w.writerow(header)
        w.writerows(rows)


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def run(path_a, path_b, out_dir=None, prefix=None, excel_bom=False):
    """Do the whole job. Returns (output_csv_path, summary_text)."""
    report_path, report_rows, source_path, source_rows = identify_files(path_a, path_b)

    header, data_rows, warnings = clean_report(report_rows)
    entries, _, src_warnings = load_source(source_rows)
    warnings += src_warnings
    updated, result = update_report(header, data_rows, entries)

    out_dir = Path(out_dir) if out_dir else report_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    base = cleaned_name(report_path)
    prefix = prefix or derive_prefix(source_path)
    out_path = out_dir / safe_filename(f"{prefix} - Output - {base}.csv")

    write_csv(out_path, header, updated, excel_bom)

    summary = build_summary(result, warnings)
    return out_path, summary


def pick_files_interactively():
    """Ask for the two files with a file dialog, or fall back to typed paths."""
    try:
        import tkinter
        from tkinter import filedialog

        root = tkinter.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        paths = filedialog.askopenfilenames(
            title="Select the Item Report AND the Doorcrashers file (hold Ctrl/Cmd to pick both)",
            filetypes=[("CSV / Excel", "*.csv *.xlsx *.xlsm"), ("All files", "*.*")],
        )
        root.destroy()
        if len(paths) == 2:
            return paths
        print("Please select exactly two files.")
    except Exception:
        pass
    print("Drag the two files into this window (or type their paths), one per line.")
    a = input("File 1: ").strip().strip("'\"")
    b = input("File 2: ").strip().strip("'\"")
    return a, b


def main(argv=None):
    p = argparse.ArgumentParser(description="Clean an Item Report and update it from a Doorcrashers source file.")
    p.add_argument("files", nargs="*", help="The Item Report and the Doorcrashers source file (either order).")
    p.add_argument("-o", "--output-dir", help="Where to save results (default: next to the Item Report).")
    p.add_argument("--prefix", help="Output filename prefix (default: the source filename).")
    p.add_argument("--excel-bom", action="store_true",
                   help="Write UTF-8 with BOM so Excel shows symbols like ® correctly.")
    args = p.parse_args(argv)

    files = args.files
    if not files:
        files = pick_files_interactively()
    if len(files) != 2:
        p.error("please provide exactly two files: the Item Report and the Doorcrashers source file")

    try:
        out_path, summary = run(files[0], files[1], args.output_dir, args.prefix,
                                args.excel_bom)
    except UpdaterError as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    print(f"\nSaved: {out_path}\n")
    print(summary)
    return 0


if __name__ == "__main__":
    sys.exit(main())
