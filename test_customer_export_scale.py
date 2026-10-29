"""Synthetic scale verification; no customer data or persistent XLSX artifacts.

Run from the project root with RUN_CUSTOMER_SCALE=1:
    py -3.14 -m unittest strategy_excel_tool.test_customer_export_scale -v

Normal discovery runs only the small verifier smoke test. The opt-in export
streams 20,000 customers x 200 products to a TemporaryDirectory and verifies
every row with bounded XML iterparse, including quantities and split headers.
When the engine is present, the large export consumes its actual plan's
columns and lazy iter_rows() directly. Otherwise it uses standalone synthetic
rows and reports that fallback explicitly. The separate fast engine check uses the test fixtures'
public contract; it waits once for 10 seconds if that module is not present,
then explicitly skips only the engine check. To rerun only that check:
    py -3.14 -m unittest strategy_excel_tool.test_customer_export_scale.CustomerExportScaleTests.test_real_engine_summary_and_late_page -v
"""

from decimal import Decimal
import importlib
import json
import os
from pathlib import Path
import sys
import tempfile
import time
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

try:
    from .query_workbook import write_query_workbook
except ImportError:
    from query_workbook import write_query_workbook


_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_LIMIT = 1048576
_COLUMNS = [{"key": key, "label": label} for key, label in [
    ("productCode", "Product code"), ("productName", "Product name"),
    ("customerCode", "Customer code"), ("customerName", "Customer name"),
    ("county", "County"), ("tier", "Tier"), ("market", "Market"),
    ("quantity", "Units"), ("boxes", "Boxes"),
]]


def _peak_process_bytes():
    """OS high-water RSS for this test process, not an allocation estimate."""
    if sys.platform == "win32":
        import ctypes
        from ctypes import wintypes

        class MemoryCounters(ctypes.Structure):
            _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD)] + [
                (name, ctypes.c_size_t) for name in (
                    "PeakWorkingSetSize", "WorkingSetSize", "QuotaPeakPagedPoolUsage",
                    "QuotaPagedPoolUsage", "QuotaPeakNonPagedPoolUsage",
                    "QuotaNonPagedPoolUsage", "PagefileUsage", "PeakPagefileUsage")]

        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        psapi = ctypes.WinDLL("psapi", use_last_error=True)
        kernel.GetCurrentProcess.restype = wintypes.HANDLE
        psapi.GetProcessMemoryInfo.argtypes = [wintypes.HANDLE, ctypes.POINTER(MemoryCounters), wintypes.DWORD]
        psapi.GetProcessMemoryInfo.restype = wintypes.BOOL
        counters = MemoryCounters()
        counters.cb = ctypes.sizeof(counters)
        if not psapi.GetProcessMemoryInfo(kernel.GetCurrentProcess(), ctypes.byref(counters), counters.cb):
            return None
        return counters.PeakWorkingSetSize
    try:
        import resource
        value = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        return value if sys.platform == "darwin" else value * 1024
    except (ImportError, AttributeError):
        return None


def _synthetic_rows(customer_count, product_count):
    for product in range(product_count):
        for customer in range(customer_count):
            tier = customer % 30 + 1
            quantity = (product % 11) * tier
            yield {
                "productCode": f"{product:06}", "productName": f"Synthetic product {product:03}",
                "customerCode": f"{customer:012}", "customerName": f"Synthetic customer {customer:05}",
                "county": "Synthetic county", "tier": tier, "market": "Synthetic market",
                "quantity": quantity, "boxes": quantity / 250,
            }


def _expected_units(customer_count, product_count):
    cycles, remainder = divmod(customer_count, 30)
    tier_sum = cycles * 465 + remainder * (remainder + 1) // 2
    return tier_sum * sum(product % 11 for product in range(product_count))


def _cell_value(cell):
    if cell.get("t") == "inlineStr":
        return cell.find(f"{_NS}is/{_NS}t").text or ""
    return cell.find(f"{_NS}v").text


def _verify_export(case, path, customer_count, product_count, metadata, columns, tiered=True):
    expected_rows = customer_count * product_count
    header_row = len(metadata) + 1
    capacity = _LIMIT - header_row
    sheet_count = max(1, (expected_rows + capacity - 1) // capacity)
    seen = units = 0
    boxes = Decimal(0)
    physical_counts = []
    with ZipFile(path) as book:
        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        case.assertEqual(len(workbook.findall(f"{_NS}sheets/{_NS}sheet")), sheet_count)
        styles = ET.fromstring(book.read("xl/styles.xml"))
        formats = {item.get("numFmtId"): item.get("formatCode")
                   for item in styles.findall(f"{_NS}numFmts/{_NS}numFmt")}
        cell_styles = styles.findall(f"{_NS}cellXfs/{_NS}xf")
        for sheet_no in range(1, sheet_count + 1):
            physical = 0
            with book.open(f"xl/worksheets/sheet{sheet_no}.xml") as stream:
                stack = []
                for event, node in ET.iterparse(stream, events=("start", "end")):
                    if event == "start":
                        stack.append(node)
                        continue
                    if node.tag == f"{_NS}row":
                        physical += 1
                        case.assertEqual(int(node.get("r")), physical)
                        cells = list(node)
                        if physical < header_row:
                            case.assertEqual(_cell_value(cells[0]), metadata[physical - 1])
                        elif physical == header_row:
                            case.assertEqual([_cell_value(c) for c in cells], [c["label"] for c in columns])
                        else:
                            product, customer = divmod(seen, customer_count)
                            quantity = (product % 11) * (customer % 30 + 1 if tiered else 1)
                            case.assertEqual(len(cells), len(columns))
                            case.assertEqual(_cell_value(cells[0]), f"{product:06}")
                            case.assertEqual(_cell_value(cells[2]), f"{customer:012}")
                            case.assertEqual(int(_cell_value(cells[7])), quantity)
                            box_value = Decimal(_cell_value(cells[8]))
                            case.assertEqual(box_value * 250, quantity)
                            if physical == header_row + 1:
                                case.assertEqual(formats[cell_styles[int(cells[8].get("s"))].get("numFmtId")], "0.000")
                            units += quantity
                            boxes += box_value
                            seen += 1
                        # Clearing alone leaves millions of empty child nodes attached.
                        stack[-2].remove(node)
                        node.clear()
                    elif node.tag == f"{_NS}pane":
                        case.assertEqual(node.get("topLeftCell"), f"A{header_row + 1}")
                    elif node.tag == f"{_NS}autoFilter":
                        case.assertEqual(node.get("ref"), f"A{header_row}:I{physical}")
                    stack.pop()
            expected_physical = min(capacity, expected_rows - (sheet_no - 1) * capacity) + header_row
            case.assertEqual(physical, expected_physical)
            physical_counts.append(physical)
            if expected_rows >= 1000000:
                print(f"Verified sheet {sheet_no}/{sheet_count}: {seen:,} cumulative data rows", flush=True)
    case.assertEqual(seen, expected_rows)
    expected_units = (_expected_units(customer_count, product_count) if tiered else
                      customer_count * sum(p % 11 for p in range(product_count)))
    case.assertEqual(units, expected_units)
    case.assertEqual(boxes * 250, units)
    return {"rowCount": seen, "sheetCount": sheet_count, "physicalRows": physical_counts,
            "units": units, "boxes": str(boxes)}


def _exercise_export(case, customer_count, product_count, plan=None):
    started = time.perf_counter()
    metadata = ["Synthetic only; all customers; 1 box = 250 units; complete",
                "Products: " + ", ".join(f"Synthetic product {p:03}" for p in range(product_count))]
    previous = {"rowCount": 0, "sheetCount": 0}
    callbacks = 0
    next_report = 500000

    def progress(state):
        nonlocal previous, callbacks, next_report
        case.assertGreaterEqual(state["rowCount"], previous["rowCount"])
        case.assertGreaterEqual(state["sheetCount"], previous["sheetCount"])
        previous = state
        callbacks += 1
        if state["rowCount"] >= next_report:
            print(f"Written {state['rowCount']:,} rows; {time.perf_counter() - started:.1f}s elapsed", flush=True)
            next_report += 500000

    with tempfile.TemporaryDirectory(prefix="customer-scale-") as directory:
        path = Path(directory) / "synthetic-scale.xlsx"
        columns = plan.columns if plan is not None else _COLUMNS
        rows = plan.iter_rows() if plan is not None else _synthetic_rows(customer_count, product_count)
        result = write_query_workbook(path, columns, rows, metadata, progress=progress)
        written = time.perf_counter()
        case.assertEqual(previous, result)
        case.assertGreaterEqual(callbacks, customer_count * product_count // 10000)
        checked = _verify_export(case, path, customer_count, product_count, metadata, columns,
                                 tiered=plan is None)
        case.assertEqual(result, {key: checked[key] for key in ("rowCount", "sheetCount")})
        if plan is not None:
            case.assertEqual(checked["units"], plan.totals["quantityTiao"])
            case.assertEqual(Decimal(checked["boxes"]), Decimal(str(plan.totals["quantityXiang"])))
        checked.update({"writeSeconds": round(written - started, 3),
                        "verifySeconds": round(time.perf_counter() - written, 3),
                        "zipBytes": path.stat().st_size, "peakProcessBytes": _peak_process_bytes(),
                        "rowSource": "real-engine" if plan is not None else "standalone-synthetic"})
    case.assertFalse(Path(directory).exists())
    checked["temporaryOutputRemoved"] = True
    print("SCALE_RESULT " + json.dumps(checked, sort_keys=True), flush=True)
    return checked


class ExportVerifierTests(unittest.TestCase):
    def test_small_verifier_smoke(self):
        _exercise_export(self, customer_count=31, product_count=3)


@unittest.skipUnless(os.environ.get("RUN_CUSTOMER_SCALE") == "1", "set RUN_CUSTOMER_SCALE=1 for 4 million rows")
class CustomerExportScaleTests(unittest.TestCase):
    def _real_engine_plan(self):
        engine_path = Path(__file__).with_name("customer_allocation.py")
        if not engine_path.is_file():
            print("Engine not available; waiting once for 10 seconds before fast engine verification", flush=True)
            time.sleep(10)
        if not engine_path.is_file():
            self.skipTest("Engine file unavailable after bounded wait; standalone writer scale check remains independent")
        prefix = f"{__package__}." if __package__ else ""
        importlib.invalidate_caches()
        engine_module = importlib.import_module(prefix + "customer_allocation")
        server = importlib.import_module(prefix + "server")
        started = time.perf_counter()
        # Same accepted fields/aliases as test_customer_allocation, no private data.
        county = "\u5b9c\u5dde\u5e02\u8425\u9500\u90e8"
        county_alias = "\u5b9c\u5dde\u8425\u9500\u90e8"
        market = "\u57ce\u5e02"
        customer_columns = [{"key": key, "label": label} for key, label in [
            ("customerCode", "\u5ba2\u6237\u7f16\u7801"),
            ("customerName", "\u5ba2\u6237\u540d\u79f0"),
            ("county", "\u8425\u9500\u90e8"), ("tier", "\u6863\u4f4d"),
            ("market", "\u5e02\u573a\u7c7b\u578b(\u6269\u5c55)"),
        ]]
        customers = ({"customerCode": f"{i:012}", "customerName": "Synthetic customer",
                      "county": county, "tier": 1, "market": market} for i in range(20000))
        with tempfile.TemporaryDirectory(prefix="customer-engine-scale-") as directory:
            path = Path(directory) / "synthetic-customers.xlsx"
            write_query_workbook(path, customer_columns, customers, ["Synthetic customers only"])
            dataset = engine_module.parse_customers(path.read_bytes(), "county", server.XlsxBook)
        self.assertFalse(Path(directory).exists())
        self.assertEqual(dataset.summary["validCustomerCount"], 20000)
        review_rows = [{"sourceRow": p + 3, "productCode": f"{p:06}",
                        "productName": f"Synthetic product {p:03}", "county": county_alias,
                        "marketType": market, "assignments": {1: p % 11}, "warnings": []}
                       for p in range(200)]
        choices = {str(row["sourceRow"]): "keep" for row in review_rows}
        engine = engine_module.AllocationQuery(dataset, {"reviewRows": review_rows}, choices, "county")
        setup_done = time.perf_counter()
        summary = engine.plan_query({}, "summary")
        detail = engine.plan_query({}, "detail")
        self.assertEqual(summary.row_count, 200)
        self.assertEqual(detail.row_count, 4000000)
        expected = 20000 * sum(p % 11 for p in range(200))
        self.assertEqual(summary.totals["quantityTiao"], expected)
        self.assertEqual(summary.totals["quantityXiang"], expected / 250)
        self.assertEqual(summary.totals["customerCount"], 20000)
        self.assertEqual(summary.totals["detailRowCount"], 4000000)
        self.assertFalse(summary.totals["incomplete"])
        summary_rows = summary_units = 0
        for row in summary.iter_rows():
            self.assertEqual(row["matchedCustomers"], 20000)
            summary_units += row["quantityTiao"]
            summary_rows += 1
        self.assertEqual(summary_rows, 200)
        self.assertEqual(summary_units, expected)
        page = detail.page(400000, 10)
        self.assertEqual(len(page["rows"]), 10)
        for offset, row in enumerate(page["rows"], 19990):
            self.assertEqual(row["customerCode"], f"{offset:012}")
            self.assertEqual(row["productCode"], "000199")
            self.assertEqual(row["quantityTiao"], 199 % 11)
        query_seconds = time.perf_counter() - setup_done
        self.assertLess(query_seconds, 8, "Summary and late page must not enumerate all 4 million details")
        print("ENGINE_SCALE_RESULT " + json.dumps({
            "customerCount": 20000, "productCount": 200, "detailRowCount": detail.row_count,
            "quantityTiao": expected, "quantityXiang": expected / 250,
            "setupSeconds": round(setup_done - started, 3), "querySeconds": round(query_seconds, 3),
            "peakProcessBytes": _peak_process_bytes(), "temporaryOutputRemoved": True,
        }, sort_keys=True), flush=True)
        return detail

    def test_real_engine_summary_and_late_page(self):
        self._real_engine_plan()

    def test_four_million_detail_export(self):
        plan = self._real_engine_plan() if Path(__file__).with_name("customer_allocation.py").is_file() else None
        if plan is None:
            print("Engine unavailable: testing standalone synthetic writer fallback", flush=True)
        _exercise_export(self, customer_count=20000, product_count=200, plan=plan)


if __name__ == "__main__":
    unittest.main()
