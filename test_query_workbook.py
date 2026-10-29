"""Synthetic, independent ZIP/XML verification of the streaming export."""

import os
from pathlib import Path
import tempfile
import tracemalloc
import unittest
from unittest.mock import patch
from decimal import Decimal
from xml.etree import ElementTree as ET
from zipfile import ZipFile

try:
    from .query_workbook import write_query_workbook
except ImportError:
    from query_workbook import write_query_workbook


NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
COLUMNS = [{"key": key, "label": label} for key, label in
           [("code", "Code"), ("quantity", "Quantity"), ("boxes", "Boxes")]]


class QueryWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "query.xlsx"

    def sheet(self, number=1):
        with ZipFile(self.path) as book:
            return ET.fromstring(book.read(f"xl/worksheets/sheet{number}.xml"))

    def test_safe_typed_cells_and_package(self):
        rows = [{"code": "00123", "quantity": 250, "boxes": Decimal("1.004")},
                {"code": '=SUM(A1:A2)&<>"\x00\x0b\ud800\ufffe\t\n\r',
                 "quantity": 0, "boxes": 0.008},
                {"code": " literal _x0041_ ", "quantity": None, "boxes": -1.125}]
        result = write_query_workbook(self.path, COLUMNS, iter(rows), ["Filter: all & <none>"])
        self.assertEqual(result, {"rowCount": 3, "sheetCount": 1})
        root = self.sheet()
        cells = {cell.get("r"): cell for cell in root.findall(".//s:c", NS)}
        self.assertEqual(cells["A3"].get("t"), "inlineStr")
        self.assertEqual(cells["A3"].find("s:is/s:t", NS).text, "00123")
        self.assertEqual(cells["A4"].find("s:is/s:t", NS).text,
                         '=SUM(A1:A2)&<>"\t\n\r')
        self.assertEqual(cells["A5"].find("s:is/s:t", NS).text, " literal _x005F_x0041_ ")
        self.assertEqual(cells["B3"].find("s:v", NS).text, "250")
        self.assertEqual(cells["C3"].find("s:v", NS).text, "1.004")
        self.assertEqual(cells["C4"].find("s:v", NS).text, "0.008")
        self.assertEqual(cells["B4"].find("s:v", NS).text, "0")
        self.assertIsNone(cells["B5"].find("s:v", NS))
        self.assertFalse(root.findall(".//s:f", NS))
        self.assertEqual(root.find("s:sheetViews/s:sheetView/s:pane", NS).get("topLeftCell"), "A3")
        self.assertEqual(root.find("s:autoFilter", NS).get("ref"), "A2:C5")
        self.assertEqual(len(root.findall("s:cols/s:col", NS)), 3)
        with ZipFile(self.path) as book:
            self.assertIsNone(book.testzip())
            self.assertNotIn("xl/sharedStrings.xml", book.namelist())
            for name in book.namelist():
                ET.fromstring(book.read(name))
            styles = ET.fromstring(book.read("xl/styles.xml"))
            formats = {item.get("numFmtId"): item.get("formatCode")
                       for item in styles.findall("s:numFmts/s:numFmt", NS)}
            cell_styles = styles.findall("s:cellXfs/s:xf", NS)
            box_format = cell_styles[int(cells["C3"].get("s"))].get("numFmtId")
            self.assertEqual(formats[box_format], "0.000")
            self.assertNotEqual(cells["A2"].get("s"), cells["A3"].get("s"))

    def test_small_splits_repeat_metadata_and_headers_without_trailing_sheet(self):
        for count, expected_sheets in [(0, 1), (1, 1), (2, 1), (3, 2), (4, 2), (5, 3)]:
            with self.subTest(count=count):
                result = write_query_workbook(self.path, COLUMNS,
                    ({"code": f"{i:05}", "quantity": i} for i in range(count)),
                    ["Filter", "Units"], sheet_row_limit=5)
                self.assertEqual(result, {"rowCount": count, "sheetCount": expected_sheets})
                actual = []
                for number in range(1, expected_sheets + 1):
                    root = self.sheet(number)
                    rows = root.findall("s:sheetData/s:row", NS)
                    self.assertLessEqual(len(rows), 5)
                    texts = [r.find("s:c/s:is/s:t", NS).text for r in rows]
                    self.assertEqual(texts[:3], ["Filter", "Units", "Code"])
                    actual.extend(texts[3:])
                    self.assertEqual(root.find("s:autoFilter", NS).get("ref"), f"A3:C{len(rows)}")
                self.assertEqual(actual, [f"{i:05}" for i in range(count)])

    def test_progress_is_periodic_and_iterator_is_not_materialized(self):
        events = []
        def rows():
            for i in range(25001):
                if i == 15000:
                    self.assertTrue(any(e["rowCount"] >= 10000 for e in events))
                yield {"quantity": i}
        result = write_query_workbook(self.path, COLUMNS, rows(), [], progress=events.append)
        self.assertEqual(result, {"rowCount": 25001, "sheetCount": 1})
        self.assertEqual(events[-1], result)
        self.assertGreaterEqual(len(events), 3)
        self.assertEqual([e["rowCount"] for e in events], sorted(e["rowCount"] for e in events))

    def test_failure_removes_partial_file(self):
        def broken_rows():
            yield {"code": "001"}
            raise RuntimeError("synthetic iterator failure")
        def broken_progress(_):
            raise RuntimeError("synthetic callback failure")
        for rows, callback in [(broken_rows(), None), (iter([{}]), broken_progress)]:
            with self.subTest(callback=callback):
                with self.assertRaises(RuntimeError):
                    write_query_workbook(self.path, COLUMNS, rows, [], progress=callback)
                self.assertFalse(self.path.exists())

    def test_zip_write_and_final_callback_failures_remove_output(self):
        with patch.object(ZipFile, "writestr", side_effect=OSError("synthetic disk failure")):
            with self.assertRaisesRegex(OSError, "synthetic disk failure"):
                write_query_workbook(self.path, COLUMNS, [{"code": "001"}], [])
        self.assertFalse(self.path.exists())

        def fail_after_completion(state):
            if state["rowCount"] == 1:
                with ZipFile(self.path) as book:
                    self.assertIsNone(book.testzip())
                raise RuntimeError("final callback failure")

        with self.assertRaisesRegex(RuntimeError, "final callback failure"):
            write_query_workbook(self.path, COLUMNS, [{"code": "001"}], [], progress=fail_after_completion)
        self.assertFalse(self.path.exists())

    def test_invalid_dimensions_and_numbers(self):
        for limit in [0, 1, 2, 1048577, 3.5, True]:
            with self.subTest(limit=limit), self.assertRaises(ValueError):
                write_query_workbook(self.path, COLUMNS, iter([]), ["metadata"], sheet_row_limit=limit)
        for columns in [[], [{"key": "a", "label": "A"}] * 16385]:
            with self.assertRaises(ValueError):
                write_query_workbook(self.path, columns, [], [])
        for value in [float("nan"), float("inf"), Decimal("NaN"), Decimal("Infinity"), "x" * 32768]:
            with self.subTest(value=str(value)[:20]), self.assertRaises(ValueError):
                write_query_workbook(self.path, COLUMNS, [{"boxes": value}], [])
            self.assertFalse(self.path.exists())

    def test_wide_column_references_and_unicode(self):
        columns = [{"key": str(i), "label": "\u5546\u54c1"} for i in range(28)]
        write_query_workbook(self.path, columns, [{"27": "\U0001f600"}], [])
        root = self.sheet()
        self.assertEqual(root.find("s:autoFilter", NS).get("ref"), "A1:AB2")
        self.assertEqual(root.find(".//s:c[@r='AB2']/s:is/s:t", NS).text, "\U0001f600")

    def test_metadata_character_limit_and_200_product_names(self):
        names = "Products: " + ", ".join(f"Synthetic product {i:03}" for i in range(200))
        self.assertLess(len(names), 32767)
        write_query_workbook(self.path, COLUMNS, [], [names, "x" * 32767])
        root = self.sheet()
        self.assertEqual(root.find(".//s:c[@r='A1']/s:is/s:t", NS).text, names)
        self.assertEqual(len(root.find(".//s:c[@r='A2']/s:is/s:t", NS).text), 32767)
        self.path.unlink()
        for line in ["x" * 32768, "\U0001f600" * 16384]:
            with self.subTest(length=len(line)), self.assertRaisesRegex(ValueError, "32767"):
                write_query_workbook(self.path, COLUMNS, [], [line])
            self.assertFalse(self.path.exists())

    def test_memory_does_not_grow_with_rows(self):
        def peak(count):
            tracemalloc.start()
            try:
                write_query_workbook(self.path, COLUMNS,
                    ({"code": f"{i:012}", "quantity": i, "boxes": 0.004} for i in range(count)), [])
                return tracemalloc.get_traced_memory()[1]
            finally:
                tracemalloc.stop()
        small, large = peak(1000), peak(30000)
        self.assertLess(large, small + 2 * 1024 * 1024)
        print(f"Writer traced peak bytes: 1000 rows={small}; 30000 rows={large}")

    @unittest.skipUnless(os.environ.get("RUN_XLSX_BOUNDARY") == "1", "set RUN_XLSX_BOUNDARY=1 for real Excel boundary")
    def test_real_excel_boundary_with_bounded_iterparse(self):
        count = 1048576
        result = write_query_workbook(self.path, COLUMNS,
            ({"code": f"{i:012}", "quantity": i, "boxes": 0.004} for i in range(count)),
            ["Synthetic rows only", "1 box = 250 units"])
        self.assertEqual(result, {"rowCount": count, "sheetCount": 2})
        seen, row_counts = 0, []
        with ZipFile(self.path) as book:
            for sheet_no in (1, 2):
                physical_rows = 0
                with book.open(f"xl/worksheets/sheet{sheet_no}.xml") as stream:
                    stack = []
                    for event, node in ET.iterparse(stream, events=("start", "end")):
                        if event == "start":
                            stack.append(node)
                            continue
                        if node.tag == "{" + NS["s"] + "}row":
                            physical_rows += 1
                            if physical_rows > 3:
                                self.assertEqual(node.find("s:c/s:is/s:t", NS).text, f"{seen:012}")
                                self.assertEqual(node.findall("s:c", NS)[1].find("s:v", NS).text, str(seen))
                                seen += 1
                            stack[-2].remove(node)
                            node.clear()
                        stack.pop()
                row_counts.append(physical_rows)
        self.assertEqual(seen, count)
        self.assertEqual(row_counts, [1048576, 6])
        print(f"Boundary verified: {seen} data rows; physical sheet rows={row_counts}; ZIP bytes={self.path.stat().st_size}")


if __name__ == "__main__":
    unittest.main()
