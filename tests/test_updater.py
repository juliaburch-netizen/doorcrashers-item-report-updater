import csv
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import doorcrashers_updater as du  # noqa: E402

HEADER = ["item_id", "flyer_name", "page", "sku", "name", "description", "url"]


def write(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(rows)


def read(path):
    with open(path, newline="", encoding="utf-8-sig") as fh:
        return list(csv.reader(fh))


class UpdaterTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def make(self, report_rows, source_rows, source_header=("fence_id", "page", "friendly_name", "url"),
             source_name="WK44_CA_Doorcrashers_Links.csv", bom=False):
        r = self.dir / "Item Report-1.csv"
        s = self.dir / source_name
        write(r, [HEADER] + report_rows)
        with open(s, "w", newline="", encoding="utf-8-sig" if bom else "utf-8") as fh:
            csv.writer(fh).writerows([list(source_header)] + source_rows)
        return r, s

    def test_clean_match_and_naming(self):
        rows = [
            ["PAGE ITEMS"] + [""] * 6,
            ["1", "BASE", "1", "P01-01", "old name", "desc, kept", "old-url"],
            ["2", "BASE", "1", "P01-02", "old2", "", ""],
            [""] * 7,
            ["FLYER ITEMS"] + [""] * 6,
            ["9", "BASE", "1", "P01-01", "should vanish", "", ""],
        ]
        src = [["P01-01", "1", "A", "http://a"], ["P01-02", "1", "B", "http://b"]]
        r, s = self.make(rows, src)
        out, summary = du.run(r, s)  # order independent
        self.assertEqual(out.name, "WK44_CA_Doorcrashers_Links - Output - Item Report-1 - Cleaned.csv")
        got = read(out)
        self.assertEqual(len(got), 3)
        self.assertEqual(got[1][4:], ["A", "desc, kept", "http://a"])
        self.assertEqual(got[2][4], "B")
        self.assertEqual(summary, "Task summary:\n- 2 SKUs successfully matched")

    def test_either_file_order(self):
        r, s = self.make([["1", "B", "1", "P01-01", "", "", ""]], [["P01-01", "1", "A", "u"]])
        out, _ = du.run(s, r)
        self.assertEqual(read(out)[1][4], "A")

    def test_discrepancies_no_fuzzy_matching(self):
        rows = [
            ["1", "B", "1", "P01-01", "x", "", ""],
            ["2", "B", "1", "P01-02", "y", "", "keep"],
            ["3", "B", "1", "p01-03", "z", "", ""],   # case differs -> no match
            ["4", "B", "1", "", "w", "", ""],        # blank sku
        ]
        src = [["P01-01", "1", "A", "u"], ["P01-03", "1", "C", "c"], ["P09-09", "9", "Z", "z"]]
        r, s = self.make(rows, src)
        out, summary = du.run(r, s)
        got = read(out)
        self.assertEqual(got[2][4], "")           # name cleared, not guessed
        self.assertEqual(got[2][6], "keep")       # url untouched when unmatched
        self.assertIn("3 SKU(s) did not match", summary)
        self.assertIn("p01-03 (row 4)", summary)
        self.assertIn("(blank) (row 5)", summary)
        self.assertIn("2 fence_id(s)", summary)
        self.assertIn("P01-03 (row 3)", summary)
        self.assertIn("P09-09 (row 4)", summary)

    def test_whitespace_and_bom_normalised(self):
        r, s = self.make([["1", "B", "1", " P01-01 ", "", "", ""]], [["P01-01", "1", "A", "u"]],
                         source_header=("fence_id", "page", "friendly_name", "url"), bom=True)
        out, summary = du.run(r, s)
        self.assertEqual(read(out)[1][4], "A")
        self.assertIn("1 SKUs successfully matched", summary)

    def test_name_column_fallback(self):
        r, s = self.make([["1", "B", "1", "P01-01", "", "", ""]], [["P01-01", "1", "A", "u"]],
                         source_header=("fence_id", "page", "name", "url"))
        out, _ = du.run(r, s)
        self.assertEqual(read(out)[1][4], "A")

    def test_duplicate_fence_not_guessed(self):
        r, s = self.make([["1", "B", "1", "P01-01", "", "", ""]],
                         [["P01-01", "1", "A", "u1"], ["P01-01", "1", "B", "u2"]])
        out, summary = du.run(r, s)
        self.assertEqual(read(out)[1][4], "")
        self.assertIn("matched more than one source row", summary)

    def test_multiline_cells_survive(self):
        r, s = self.make([["1", "B", "1", "P01-01", "", "line1\nline2", ""]], [["P01-01", "1", "A", "u"]])
        out, _ = du.run(r, s)
        self.assertEqual(read(out)[1][5], "line1\nline2")

    def test_missing_flyer_marker_warns(self):
        r, s = self.make([["1", "B", "1", "P01-01", "", "", ""]], [["P01-01", "1", "A", "u"]])
        _, summary = du.run(r, s)
        self.assertIn("Warnings", summary)

    def test_source_name_is_irrelevant(self):
        r, s = self.make([["1", "B", "1", "P01-01", "", "", ""]], [["P01-01", "1", "A", "u"]],
                         source_name="whatever this is called.csv")
        out, summary = du.run(r, s)
        self.assertEqual(out.name, "whatever this is called - Output - Item Report-1 - Cleaned.csv")

    def test_description_never_modified(self):
        r, s = self.make([["1", "B", "1", "P01-01", "", "trailing space ", ""]], [["P01-01", "1", "A", "u"]])
        out, _ = du.run(r, s)
        self.assertEqual(read(out)[1][5], "trailing space ")

    def test_wrong_files_error(self):
        a = self.dir / "a.csv"
        write(a, [["foo", "bar"], ["1", "2"]])
        r, _ = self.make([], [])
        with self.assertRaises(du.UpdaterError):
            du.run(a, r)

    def test_prefix(self):
        self.assertEqual(du.derive_prefix("WK44_CA_Doorcrashers_Links.csv"), "WK44_CA_Doorcrashers_Links")


if __name__ == "__main__":
    unittest.main()
