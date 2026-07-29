# -*- coding: utf-8 -*-
import argparse
import copy
import datetime as dt
import html
import io
import json
import os
import re
import socket
import sys
import time
import uuid
import zipfile
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse
from xml.etree import ElementTree as ET

import xlrd
from xlutils.filter import XLWTWriter, XLRDReader, process


HOST = os.environ.get("STRATEGY_TOOL_HOST", "127.0.0.1")
PORT = int(os.environ.get("STRATEGY_TOOL_PORT", "8010"))
IS_FROZEN = bool(getattr(sys, "frozen", False))
BASE_DIR = Path(sys.executable).resolve().parent if IS_FROZEN else Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR if IS_FROZEN else BASE_DIR.parent


def resolve_template_dir():
    configured = os.environ.get("STRATEGY_TOOL_TEMPLATE_DIR", "").strip()
    if configured:
        path = Path(configured).expanduser()
        if not path.is_absolute():
            path = PROJECT_ROOT / path
        return path.resolve()
    return (PROJECT_ROOT / "private_templates").resolve()


TEMPLATE_DIR = resolve_template_dir()
OUTPUT_DIR = BASE_DIR / "outputs"
INNER_CODE_MAPPING_FILENAME = "商品内码实例.xls"
BUILD_VERSION = "2026-07-27.3-market-type-alias"
SESSIONS = {}


def find_available_port(host="127.0.0.1", start=8010, end=8099):
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind((host, port))
            except OSError:
                continue
            return port
    raise RuntimeError("No available local port found between %s and %s." % (start, end))


def create_http_server(host, preferred_port, handler):
    try:
        return ThreadingHTTPServer((host, preferred_port), handler), preferred_port
    except OSError:
        fallback = find_available_port(host, preferred_port + 1, preferred_port + 200)
        return ThreadingHTTPServer((host, fallback), handler), fallback

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)

TEMPLATES = {
    "district_type": {
        "label": "档位+区倒+市场类型",
        "filename": "档位+区倒+市场类型.xls",
        "analysisFilename": "档位+区倒+市场类型.xlsx",
        "marker": ("M200007", "2"),
        "requiredDims": ["county", "marketType"],
    },
    "county": {
        "label": "档位+区县",
        "filename": "档位+区县.xls",
        "analysisFilename": "档位+区县.xlsx",
        "marker": ("M100003", "1"),
        "requiredDims": ["county"],
    },
    "county_business": {
        "label": "档位+区县+商圈",
        "filename": "档位+区县+商圈.xls",
        "analysisFilename": "档位+区县+商圈.xlsx",
        "marker": ("M200013", "2"),
        "requiredDims": ["county", "businessArea"],
    },
    "market_type": {
        "label": "档位+市场类型",
        "filename": "档位+市场类型.xls",
        "analysisFilename": "档位+市场类型.xlsx",
        "marker": ("M100004", "1"),
        "requiredDims": ["marketType"],
    },
}


def qn(name):
    return "{%s}%s" % (NS_MAIN, name)


def normalize_text(value):
    if value is None:
        return ""
    text = str(value).replace("\u3000", " ")
    return re.sub(r"\s+", "", text).strip()


def normalize_name(value):
    return normalize_text(value).strip("：:;；,.，。")


def split_names(value):
    text = normalize_name(value)
    return [part for part in re.split(r"[、,，/]+", text) if part]


def col_to_index(col):
    total = 0
    for char in col:
        total = total * 26 + ord(char.upper()) - ord("A") + 1
    return total


def index_to_col(index):
    result = ""
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def split_cell_ref(ref):
    match = re.match(r"([A-Z]+)(\d+)$", ref or "")
    if not match:
        return None, None
    return col_to_index(match.group(1)), int(match.group(2))


def row_number(row):
    raw = row.attrib.get("r")
    if raw and raw.isdigit():
        return int(raw)
    cells = list(row.findall(qn("c")))
    if not cells:
        return 0
    _, number = split_cell_ref(cells[0].attrib.get("r"))
    return number or 0


def parse_rule_segments(text):
    rules = []
    pattern = r"(\d{1,2})(?:\s*[-~～至到]\s*(\d{1,2}))?\s*档\s*(\d+)\s*条"
    for start, end, count in re.findall(pattern, text or ""):
        a = int(start)
        b = int(end or start)
        value = int(count)
        if a > b:
            a, b = b, a
        rules.append({"start": a, "end": b, "count": value})
    return rules


ZERO_POLICY_WORDS = ["无库存", "库存不足", "停投", "暂停投放", "不投放", "无货"]


def is_zero_policy_text(text):
    normalized = normalize_name(text)
    return any(word in normalized for word in ZERO_POLICY_WORDS)


def looks_like_county_header(label):
    parts = split_names(label)
    market_words = {"城市", "县城", "乡镇", "农村", "市区", "城区", "镇区"}
    return bool(parts) and all(
        part not in market_words and (part.endswith("县") or part.endswith("区"))
        for part in parts
    )


def looks_like_market_type_label(label):
    market_words = {"城市", "县城", "乡镇", "农村", "市区", "城区", "镇区"}
    parts = split_names(label)
    return bool(parts) and all(part in market_words for part in parts)


def parse_strategy_text(raw_text):
    warnings = []
    rules = []
    current_counties = []
    order = 0

    text = (raw_text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[。；;]+", "\n", text)
    lines = [line.strip() for line in text.split("\n") if line.strip()]

    for line in lines:
        label = None
        body = line
        if "：" in line or ":" in line:
            label, body = re.split(r"[：:]", line, maxsplit=1)
            label = label.strip()
            body = body.strip()

        body_rules = parse_rule_segments(body)

        if label and (looks_like_county_header(label) or not looks_like_market_type_label(label)):
            current_counties = split_names(label)
            if not body:
                continue
            if not body_rules and is_zero_policy_text(body):
                order += 1
                rules.append({
                    "order": order,
                    "counties": current_counties[:],
                    "marketTypes": [],
                    "tiers": [],
                    "zero": True,
                    "source": line,
                })
                continue
            if not body_rules:
                warnings.append("县区 '%s' 后没有识别到档位规则：%s" % (label, line))
                continue
            order += 1
            rules.append({
                "order": order,
                "counties": current_counties[:],
                "marketTypes": [],
                "tiers": body_rules,
                "source": line,
            })
            continue

        if label:
            if not body_rules and is_zero_policy_text(body):
                order += 1
                rules.append({
                    "order": order,
                    "counties": current_counties[:],
                    "marketTypes": split_names(label),
                    "tiers": [],
                    "zero": True,
                    "source": line,
                })
                continue
            if not body_rules:
                warnings.append("类型 '%s' 后没有识别到档位规则：%s" % (label, line))
                continue
            order += 1
            rules.append({
                "order": order,
                "counties": current_counties[:],
                "marketTypes": split_names(label),
                "tiers": body_rules,
                "source": line,
            })
            continue

        direct_rules = parse_rule_segments(body)
        if direct_rules:
            order += 1
            rules.append({
                "order": order,
                "counties": current_counties[:],
                "marketTypes": [],
                "tiers": direct_rules,
                "source": line,
            })
        elif is_zero_policy_text(body):
            order += 1
            rules.append({
                "order": order,
                "counties": current_counties[:],
                "marketTypes": [],
                "tiers": [],
                "zero": True,
                "source": line,
            })
        elif line:
            warnings.append("无法解析策略行：%s" % line)

    return {"rules": rules, "warnings": warnings}


class XlsxBook:
    def __init__(self, content):
        self.content = content
        self.zip_buffer = io.BytesIO(content)
        self.zf = zipfile.ZipFile(self.zip_buffer, "r")
        self.shared_strings = self._read_shared_strings()
        self.sheets = self._read_sheets()

    def _read_shared_strings(self):
        if "xl/sharedStrings.xml" not in self.zf.namelist():
            return []
        root = ET.fromstring(self.zf.read("xl/sharedStrings.xml"))
        values = []
        for item in root.findall(qn("si")):
            values.append("".join(node.text or "" for node in item.iter(qn("t"))))
        return values

    def _read_sheets(self):
        workbook = ET.fromstring(self.zf.read("xl/workbook.xml"))
        rels = ET.fromstring(self.zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {}
        for rel in rels.findall("{%s}Relationship" % NS_PKG_REL):
            rel_map[rel.attrib["Id"]] = rel.attrib["Target"]

        sheets = []
        for sheet in workbook.find(qn("sheets")).findall(qn("sheet")):
            rel_id = sheet.attrib.get("{%s}id" % NS_REL)
            target = rel_map.get(rel_id)
            if not target:
                continue
            path = target.lstrip("/") if target.startswith("/") else "xl/" + target.lstrip("/")
            root = ET.fromstring(self.zf.read(path))
            sheet_data = root.find(qn("sheetData"))
            if sheet_data is None:
                continue
            sheets.append({
                "name": sheet.attrib.get("name", "Sheet"),
                "path": path,
                "root": root,
                "sheetData": sheet_data,
            })
        if not sheets:
            raise ValueError("工作簿没有可读取的工作表。")
        return sheets

    def rows(self, sheet):
        return sheet["sheetData"].findall(qn("row"))

    def row_cells(self, row):
        cells = {}
        for cell in row.findall(qn("c")):
            col, _ = split_cell_ref(cell.attrib.get("r"))
            if col:
                cells[col] = cell
        return cells

    def cell_text(self, cell):
        if cell is None:
            return ""
        value_type = cell.attrib.get("t")
        if value_type == "s":
            node = cell.find(qn("v"))
            if node is None or node.text is None:
                return ""
            try:
                return self.shared_strings[int(node.text)]
            except Exception:
                return ""
        if value_type == "inlineStr":
            node = cell.find(qn("is"))
            if node is None:
                return ""
            return "".join(item.text or "" for item in node.iter(qn("t")))
        node = cell.find(qn("v"))
        if node is not None and node.text is not None:
            return node.text
        formula = cell.find(qn("f"))
        return formula.text if formula is not None and formula.text else ""

    def row_values(self, row):
        return {col: self.cell_text(cell) for col, cell in self.row_cells(row).items()}

    def get_or_create_cell(self, row, col_index, style_source=None):
        row_no = row_number(row)
        ref = "%s%s" % (index_to_col(col_index), row_no)
        cells = self.row_cells(row)
        if col_index in cells:
            return cells[col_index]

        cell = copy.deepcopy(style_source) if style_source is not None else ET.Element(qn("c"))
        cell.attrib["r"] = ref
        for child in list(cell):
            cell.remove(child)

        inserted = False
        for existing in list(row.findall(qn("c"))):
            existing_col, _ = split_cell_ref(existing.attrib.get("r"))
            if existing_col and existing_col > col_index:
                row.insert(list(row).index(existing), cell)
                inserted = True
                break
        if not inserted:
            row.append(cell)
        return cell

    def set_number(self, cell, value):
        for child in list(cell):
            cell.remove(child)
        cell.attrib.pop("t", None)
        ET.SubElement(cell, qn("v")).text = str(int(value))

    def set_text(self, cell, value):
        for child in list(cell):
            cell.remove(child)
        cell.attrib["t"] = "inlineStr"
        inline = ET.SubElement(cell, qn("is"))
        ET.SubElement(inline, qn("t")).text = "" if value is None else str(value)

    def clear_cell(self, cell):
        for child in list(cell):
            cell.remove(child)
        cell.attrib.pop("t", None)

    def clone_row(self, row, new_row_no):
        cloned = copy.deepcopy(row)
        cloned.attrib["r"] = str(new_row_no)
        for cell in cloned.findall(qn("c")):
            col, _ = split_cell_ref(cell.attrib.get("r"))
            if col:
                cell.attrib["r"] = "%s%s" % (index_to_col(col), new_row_no)
        return cloned

    def replace_data_rows(self, sheet, preserve_until_row_no, new_rows):
        sheet_data = sheet["sheetData"]
        for row in list(sheet_data.findall(qn("row"))):
            if row_number(row) > preserve_until_row_no:
                sheet_data.remove(row)
        for row in new_rows:
            sheet_data.append(row)
        self.update_dimension(sheet)

    def update_dimension(self, sheet):
        max_row = 1
        max_col = 1
        for row in self.rows(sheet):
            max_row = max(max_row, row_number(row))
            for col in self.row_cells(row).keys():
                max_col = max(max_col, col)
        dimension = sheet["root"].find(qn("dimension"))
        if dimension is not None:
            dimension.attrib["ref"] = "A1:%s%s" % (index_to_col(max_col), max_row)

    def to_bytes(self):
        sheet_bytes = {
            sheet["path"]: ET.tostring(sheet["root"], encoding="utf-8", xml_declaration=True)
            for sheet in self.sheets
        }
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as out:
            for item in self.zf.infolist():
                data = sheet_bytes.get(item.filename)
                if data is None:
                    data = self.zf.read(item.filename)
                out.writestr(item, data)
        return output.getvalue()


def header_score(values, mode):
    score = 0
    names = [normalize_name(value) for value in values.values()]
    if mode == "strategy":
        for text in names:
            if "商品编码" in text or text in ["编码", "商品代码", "卷烟编码"]:
                score += 5
            if "商品名称" in text or text in ["品规", "卷烟名称", "名称"]:
                score += 5
            if "本期投放策略" in text:
                score += 8
        return score

    for text in names:
        if is_product_code_header(text):
            score += 4
        if is_product_name_header(text):
            score += 4
        if is_product_inner_code_header(text):
            score += 2
        if is_county_header(text):
            score += 3
        if is_market_type_header(text):
            score += 3
        if is_business_area_header(text):
            score += 3
        if tier_from_header(text):
            score += 1
    return score


def is_product_code_header(text):
    return text in ["商品编码", "商品代码", "卷烟编码", "编码"] or "商品编码" in text


def is_product_name_header(text):
    return text in ["商品名称", "卷烟名称", "品规", "名称"] or "商品名称" in text



def is_product_inner_code_header(text):
    return text in ["商品内码", "卷烟内码", "品规内码"] or ("商品" in text and "内码" in text)


def normalize_product_code_key(value):
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value or "").strip()
        if re.fullmatch(r"\d+\.0", text):
            text = text[:-2]
    text = text.replace(" ", "").replace("\u3000", "")
    if text.isdigit() and len(text) < 5:
        text = text.zfill(5)
    return text


def normalize_inner_code(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value or "").strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text.replace(" ", "").replace("\u3000", "")


def inner_code_mapping_path():
    return TEMPLATE_DIR / INNER_CODE_MAPPING_FILENAME


def load_inner_code_mapping():
    path = inner_code_mapping_path()
    if not path.exists():
        raise ValueError(
            "缺少商品内码映射文件：%s。请将其放入 %s。"
            % (INNER_CODE_MAPPING_FILENAME, TEMPLATE_DIR)
        )

    try:
        book = xlrd.open_workbook(str(path))
    except Exception as exc:
        raise ValueError("商品内码映射文件无法读取：%s" % exc)

    mapping = {}
    names = {}
    conflicts = {}
    source_row_count = 0
    matched_sheets = []

    for sheet in book.sheets():
        header_rowx = None
        code_colx = None
        inner_colx = None
        name_colx = None
        for rowx in range(min(sheet.nrows, 80)):
            for colx in range(sheet.ncols):
                text = normalize_name(sheet.cell_value(rowx, colx))
                if code_colx is None and is_product_code_header(text):
                    code_colx = colx
                if inner_colx is None and is_product_inner_code_header(text):
                    inner_colx = colx
                if name_colx is None and is_product_name_header(text):
                    name_colx = colx
            if code_colx is not None and inner_colx is not None:
                header_rowx = rowx
                break
            code_colx = None
            inner_colx = None
            name_colx = None

        if header_rowx is None:
            continue

        matched_sheets.append(sheet.name)
        for rowx in range(header_rowx + 1, sheet.nrows):
            code = normalize_product_code_key(sheet.cell_value(rowx, code_colx))
            inner_code = normalize_inner_code(sheet.cell_value(rowx, inner_colx))
            if not code or not inner_code:
                continue
            source_row_count += 1
            previous = mapping.get(code)
            if previous and previous != inner_code:
                conflicts.setdefault(code, {previous}).add(inner_code)
                continue
            mapping[code] = inner_code
            if name_colx is not None:
                name = str(sheet.cell_value(rowx, name_colx) or "").strip()
                if name:
                    names[code] = name

    if not matched_sheets:
        raise ValueError("商品内码映射文件未找到“商品编码”和“商品内码”表头。")
    if conflicts:
        samples = [
            "%s=%s" % (code, "/".join(sorted(values)))
            for code, values in sorted(conflicts.items())[:20]
        ]
        raise ValueError("商品内码映射存在一对多冲突：%s" % "；".join(samples))
    if not mapping:
        raise ValueError("商品内码映射文件没有可用的数据行。")

    return {
        "path": path,
        "mapping": mapping,
        "names": names,
        "sourceRowCount": source_row_count,
        "matchedSheets": matched_sheets,
    }


def apply_inner_code_mapping(products, mapping_data):
    missing = []
    replaced = []
    mapping = mapping_data["mapping"]
    for product in products:
        key = normalize_product_code_key(product.get("code"))
        inner_code = mapping.get(key)
        if not inner_code:
            product["innerCode"] = ""
            missing.append(product)
            continue
        original = normalize_inner_code(product.get("innerCode"))
        if original and original != inner_code:
            replaced.append((product, original, inner_code))
        product["innerCode"] = inner_code
    return missing, replaced


def is_strategy_header(text):
    return "本期投放策略" in text


def is_county_header(text):
    if "编码" in text or "代码" in text or "内码" in text:
        return False
    return text in ["县区", "区县", "区倒", "县", "区", "所属县区", "区县名称", "县区名称"] or (
        ("县" in text or "区" in text) and "档" not in text and "商圈" not in text
    )


def is_market_type_header(text):
    if "编码" in text or "代码" in text or "内码" in text:
        return False
    return "市场类型" in text or text in ["市场", "类型", "客户类型", "城乡类型"]


def is_business_area_header(text):
    if "编码" in text or "代码" in text or "内码" in text:
        return False
    return "商圈" in text


CHINESE_NUMBERS = {
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
    "十一": 11,
    "十二": 12,
    "十三": 13,
    "十四": 14,
    "十五": 15,
    "十六": 16,
    "十七": 17,
    "十八": 18,
    "十九": 19,
    "二十": 20,
    "二十一": 21,
    "二十二": 22,
    "二十三": 23,
    "二十四": 24,
    "二十五": 25,
    "二十六": 26,
    "二十七": 27,
    "二十八": 28,
    "二十九": 29,
    "三十": 30,
}


def tier_from_header(text):
    value = normalize_name(text)
    match = re.fullmatch(r"(\d{1,2})档", value)
    if match:
        tier = int(match.group(1))
        return tier if 1 <= tier <= 30 else None
    match = re.fullmatch(r"([一二三四五六七八九十]{1,3})档", value)
    if match:
        return CHINESE_NUMBERS.get(match.group(1))
    if re.fullmatch(r"\d+", value):
        tier = int(value[-2:])
        return tier if 1 <= tier <= 30 else None
    match = re.search(r"(\d{1,2})$", value)
    if not match:
        return None
    tier = int(match.group(1))
    return tier if 1 <= tier <= 30 else None


def detect_strategy_table(book):
    for sheet in book.sheets:
        best = None
        for row in book.rows(sheet)[:80]:
            values = book.row_values(row)
            score = header_score(values, "strategy")
            if best is None or score > best["score"]:
                best = {"row": row, "rowNumber": row_number(row), "values": values, "score": score}
        if best and best["score"] >= 12:
            columns = {"code": None, "name": None, "strategy": None, "productInnerCode": None}
            for col, value in best["values"].items():
                text = normalize_name(value)
                if columns["code"] is None and is_product_code_header(text):
                    columns["code"] = col
                if columns["productInnerCode"] is None and is_product_inner_code_header(text):
                    columns["productInnerCode"] = col
                if columns["name"] is None and is_product_name_header(text):
                    columns["name"] = col
                if columns["strategy"] is None and is_strategy_header(text):
                    columns["strategy"] = col
            if columns["code"] and columns["name"] and columns["strategy"]:
                return sheet, best["rowNumber"], columns
    raise ValueError("策略表未找到包含“商品编码、商品名称、本期投放策略”的表头。")


def read_strategy_products(content):
    book = XlsxBook(content)
    sheet, header_row, columns = detect_strategy_table(book)
    products = []
    warnings = []

    for row in book.rows(sheet):
        row_no = row_number(row)
        if row_no <= header_row:
            continue
        cells = book.row_cells(row)
        code = book.cell_text(cells.get(columns["code"])).strip()
        name = book.cell_text(cells.get(columns["name"])).strip()
        strategy_text = book.cell_text(cells.get(columns["strategy"])).strip()
        inner_code = book.cell_text(cells.get(columns.get("productInnerCode"))).strip() if columns.get("productInnerCode") else ""
        if not code and not name:
            continue
        parsed = parse_strategy_text(strategy_text)
        product_warnings = parsed["warnings"][:]
        if not parsed["rules"]:
            product_warnings.append("没有识别到档位规则，导出时所有档位填0。")
        if product_warnings:
            warnings.extend("第%s行 %s：%s" % (row_no, name or code, item) for item in product_warnings)
        products.append({
            "sourceRow": row_no,
            "code": code,
            "name": name,
            "innerCode": inner_code,
            "strategyText": strategy_text,
            "rules": parsed["rules"],
            "warnings": product_warnings,
        })

    if not products:
        raise ValueError("策略表没有读取到商品数据行。")

    return {
        "sheetName": sheet["name"],
        "headerRow": header_row,
        "products": products,
        "warnings": warnings,
    }


def detect_template_columns(book, template_key):
    target = TEMPLATES[template_key]
    sheet = book.sheets[0]
    best = None
    for row in book.rows(sheet)[:80]:
        values = book.row_values(row)
        score = header_score(values, "template")
        if best is None or score > best["score"]:
            best = {"row": row, "rowNumber": row_number(row), "values": values, "score": score}

    if not best or best["score"] <= 0:
        raise ValueError("模板未识别到有效表头。")

    columns = {
        "productCode": None,
        "productName": None,
        "productInnerCode": None,
        "county": None,
        "marketType": None,
        "businessArea": None,
        "tierCols": {},
    }
    for col, value in best["values"].items():
        text = normalize_name(value)
        if columns["productCode"] is None and is_product_code_header(text):
            columns["productCode"] = col
        if columns["productInnerCode"] is None and is_product_inner_code_header(text):
            columns["productInnerCode"] = col
        if columns["productName"] is None and is_product_name_header(text):
            columns["productName"] = col
        if columns["county"] is None and is_county_header(text):
            columns["county"] = col
        if columns["marketType"] is None and is_market_type_header(text):
            columns["marketType"] = col
        if columns["businessArea"] is None and is_business_area_header(text):
            columns["businessArea"] = col
        tier = tier_from_header(text)
        if tier:
            columns["tierCols"][tier] = col

    errors = []
    if columns["productCode"] is None:
        errors.append("模板缺少商品编码列。")
    if columns["productName"] is None:
        errors.append("模板缺少商品名称列。")
    for dim in target["requiredDims"]:
        if columns[dim] is None:
            labels = {"county": "县区/区倒", "marketType": "市场类型", "businessArea": "商圈"}
            errors.append("模板缺少%s列。" % labels[dim])
    if not columns["tierCols"]:
        errors.append("模板缺少 1档-30档列。")
    if errors:
        raise ValueError("；".join(errors))

    missing_tiers = [tier for tier in range(1, 31) if tier not in columns["tierCols"]]
    return {
        "sheet": sheet,
        "headerRow": best["rowNumber"],
        "columns": columns,
        "missingTiers": missing_tiers,
    }


def get_template_data_rows(book, sheet, header_row_no, columns):
    rows = []
    dim_cols = [columns.get("county"), columns.get("marketType"), columns.get("businessArea")]
    dim_cols = [col for col in dim_cols if col]
    tier_cols = list(columns["tierCols"].values())
    for row in book.rows(sheet):
        if row_number(row) <= header_row_no:
            continue
        cells = book.row_cells(row)
        if is_tier_subheader_row(book, cells, dim_cols, tier_cols):
            continue
        has_dim = any(normalize_name(book.cell_text(cells.get(col))) for col in dim_cols)
        has_tier = any(book.cell_text(cells.get(col)) != "" for col in tier_cols)
        if has_dim or has_tier:
            rows.append(row)
    if not rows:
        raise ValueError("模板表头下方没有可复制的维度数据行。")
    return rows


def is_tier_subheader_row(book, cells, dim_cols, tier_cols):
    if any(normalize_name(book.cell_text(cells.get(col))) for col in dim_cols):
        return False
    tier_values = [normalize_name(book.cell_text(cells.get(col))) for col in tier_cols]
    tier_values = [value for value in tier_values if value]
    if not tier_values:
        return False
    return sum(1 for value in tier_values if tier_from_header(value)) >= min(5, len(tier_values))


def template_preserve_until_row(book, sheet, header_row_no, columns):
    preserve_until = header_row_no
    dim_cols = [columns.get("county"), columns.get("marketType"), columns.get("businessArea")]
    dim_cols = [col for col in dim_cols if col]
    tier_cols = list(columns["tierCols"].values())
    for row in book.rows(sheet):
        row_no = row_number(row)
        if row_no <= header_row_no:
            continue
        cells = book.row_cells(row)
        if is_tier_subheader_row(book, cells, dim_cols, tier_cols):
            preserve_until = max(preserve_until, row_no)
            continue
        has_dim = any(normalize_name(book.cell_text(cells.get(col))) for col in dim_cols)
        has_tier = any(book.cell_text(cells.get(col)) != "" for col in tier_cols)
        if has_dim or has_tier:
            break
    return preserve_until


def name_matches(value, options):
    if not options:
        return True
    left = normalize_name(value)
    if not left:
        return False
    for option in options:
        right = normalize_name(option)
        if right == "其他":
            continue
        aliases = NAME_ALIASES.get(right, [])
        candidates = [right] + aliases
        if any(left == item or (item and item in left) or (left and left in item) for item in candidates):
            return True
    return False


NAME_ALIASES = {
    "金城江": ["城区", "城区营销部", "金城江区"],
    "市区": ["城市", "城区"],
    "城区": ["城市", "市区"],
}


def has_other(options):
    return any(normalize_name(option) == "其他" for option in options or [])


def matches_specific_county_rule(dim_county, product):
    for rule in product["rules"]:
        counties = [item for item in rule["counties"] if normalize_name(item) != "其他"]
        if counties and name_matches(dim_county, counties):
            return True
    return False


def compact_assignments(assignments):
    items = sorted((tier, value) for tier, value in assignments.items() if value)
    if not items:
        return "全为0"
    parts = []
    i = 0
    while i < len(items):
        start_tier, value = items[i]
        end_tier = start_tier
        i += 1
        while i < len(items) and items[i][0] == end_tier + 1 and items[i][1] == value:
            end_tier = items[i][0]
            i += 1
        label = "%s档" % start_tier if start_tier == end_tier else "%s-%s档" % (start_tier, end_tier)
        parts.append("%s=%s" % (label, value))
    return "；".join(parts)


def evaluate_product_on_dimension(product, dim, tier_cols):
    assignments = {tier: 0 for tier in tier_cols.keys()}
    matched_rules = []
    conflicts = []

    for rule in product["rules"]:
        if dim["hasCounty"] and rule["counties"]:
            if has_other(rule["counties"]):
                if matches_specific_county_rule(dim["county"], product):
                    continue
            elif not name_matches(dim["county"], rule["counties"]):
                continue
        if dim["hasMarketType"] and rule["marketTypes"] and not name_matches(dim["marketType"], rule["marketTypes"]):
            continue
        matched_rules.append(rule["order"])
        if rule.get("zero"):
            for tier in assignments:
                assignments[tier] = 0
            continue
        for segment in rule["tiers"]:
            for tier in range(segment["start"], segment["end"] + 1):
                if tier not in assignments:
                    continue
                old = assignments[tier]
                new = segment["count"]
                if old not in [0, new]:
                    conflicts.append("%s档 %s->%s" % (tier, old, new))
                assignments[tier] = new

    return assignments, matched_rules, conflicts


def template_path_for_key(template_key):
    item = TEMPLATES.get(template_key)
    if not item:
        raise ValueError("未知导出格式：%s" % template_key)
    return TEMPLATE_DIR / item["filename"]


def analysis_template_path_for_key(template_key):
    item = TEMPLATES.get(template_key)
    if not item:
        raise ValueError("未知导出格式：%s" % template_key)
    return TEMPLATE_DIR / item["analysisFilename"]


def normalize_xls_marker_value(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value or "").strip()


def open_validated_xls_template(template_key):
    path = template_path_for_key(template_key)
    if not path.exists():
        raise ValueError("缺少旧版 Excel 模板：%s" % path.name)

    book = xlrd.open_workbook(str(path), formatting_info=True)
    if not book.sheets():
        raise ValueError("模板没有可读取的工作表：%s" % path.name)
    sheet = book.sheet_by_index(0)
    if sheet.nrows < 1 or sheet.ncols < 2:
        raise ValueError("模板隐藏首行缺少类型标识：%s" % path.name)

    actual = (
        normalize_xls_marker_value(sheet.cell_value(0, 0)),
        normalize_xls_marker_value(sheet.cell_value(0, 1)),
    )
    expected = TEMPLATES[template_key]["marker"]
    if actual != expected:
        raise ValueError(
            "模板类型标识不匹配：%s 应为 %s/%s，实际为 %s/%s"
            % (path.name, expected[0], expected[1], actual[0], actual[1])
        )
    row_info = sheet.rowinfo_map.get(0)
    if row_info is None or not row_info.hidden:
        raise ValueError("模板第1行必须保持隐藏：%s" % path.name)
    return book, sheet


def copy_xls_row_format(source_sheet, target_sheet, source_rowx, target_rowx, hidden=False):
    source_info = source_sheet.rowinfo_map.get(source_rowx)
    target_row = target_sheet.row(target_rowx)
    if source_info is not None:
        target_row.height = source_info.height
        target_row.has_default_height = source_info.has_default_height
        target_row.height_mismatch = source_info.height_mismatch
        target_row.level = source_info.outline_level
        target_row.collapse = source_info.outline_group_starts_ends
        target_row.space_above = source_info.additional_space_above
        target_row.space_below = source_info.additional_space_below
    target_row.hidden = 1 if hidden else 0


def copy_xls_cell(source_book, source_sheet, target_sheet, styles, source_rowx, source_colx, target_rowx, target_colx):
    cell = source_sheet.cell(source_rowx, source_colx)
    style = styles[cell.xf_index] if cell.xf_index is not None else None
    row = target_sheet.row(target_rowx)
    if cell.ctype == xlrd.XL_CELL_EMPTY or cell.ctype == xlrd.XL_CELL_BLANK:
        row.set_cell_blank(target_colx, style)
    elif cell.ctype == xlrd.XL_CELL_TEXT:
        row.set_cell_text(target_colx, cell.value, style)
    elif cell.ctype in (xlrd.XL_CELL_NUMBER, xlrd.XL_CELL_DATE):
        row.set_cell_number(target_colx, cell.value, style)
    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
        row.set_cell_boolean(target_colx, cell.value, style)
    elif cell.ctype == xlrd.XL_CELL_ERROR:
        row.set_cell_error(target_colx, cell.value, style)
    else:
        row.set_cell_text(target_colx, str(cell.value), style)


def build_xls_output(template_key, columns, preserve_until_row_no, output_records, export_title):
    source_book, source_sheet = open_validated_xls_template(template_key)
    if preserve_until_row_no + len(output_records) > 65536:
        raise ValueError(
            "旧版 .xls 最多支持65536行；当前需要%s行，请减少策略商品数量。"
            % (preserve_until_row_no + len(output_records))
        )

    writer = XLWTWriter()
    process(XLRDReader(source_book, TEMPLATES[template_key]["filename"]), writer)
    target_book = writer.output[0][1]
    target_sheet = target_book.get_sheet(0)
    styles = writer.style_list

    # The first hidden row is a system import contract, not display content.
    target_sheet.row(0).hidden = 1
    title_style = styles[source_sheet.cell(1, 0).xf_index]
    target_sheet.write(1, 0, export_title, title_style)

    for offset, record in enumerate(output_records):
        target_rowx = preserve_until_row_no + offset
        source_rowx = record["sourceRowNo"] - 1
        copy_xls_row_format(source_sheet, target_sheet, source_rowx, target_rowx, hidden=False)
        for colx in range(source_sheet.ncols):
            copy_xls_cell(
                source_book,
                source_sheet,
                target_sheet,
                styles,
                source_rowx,
                colx,
                target_rowx,
                colx,
            )

        product = record["product"]
        target_sheet.write(target_rowx, columns["productCode"] - 1, product["code"], styles[source_sheet.cell(source_rowx, columns["productCode"] - 1).xf_index])
        target_sheet.write(target_rowx, columns["productName"] - 1, product["name"], styles[source_sheet.cell(source_rowx, columns["productName"] - 1).xf_index])
        if columns.get("productInnerCode"):
            colx = columns["productInnerCode"] - 1
            style = styles[source_sheet.cell(source_rowx, colx).xf_index]
            if product.get("innerCode"):
                target_sheet.write(target_rowx, colx, product["innerCode"], style)
            else:
                target_sheet.row(target_rowx).set_cell_blank(colx, style)
        for tier, col in columns["tierCols"].items():
            colx = col - 1
            style = styles[source_sheet.cell(source_rowx, colx).xf_index]
            target_sheet.write(target_rowx, colx, record["assignments"][tier], style)

    output = io.BytesIO()
    target_book.save(output)
    data = output.getvalue()
    if not data.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        raise ValueError("导出结果不是真实的 Excel 97-2003 .xls 文件。")

    verify_book = xlrd.open_workbook(file_contents=data, formatting_info=True)
    verify_sheet = verify_book.sheet_by_index(0)
    expected_marker = TEMPLATES[template_key]["marker"]
    actual_marker = (
        normalize_xls_marker_value(verify_sheet.cell_value(0, 0)),
        normalize_xls_marker_value(verify_sheet.cell_value(0, 1)),
    )
    verify_row_info = verify_sheet.rowinfo_map.get(0)
    if actual_marker != expected_marker or verify_row_info is None or not verify_row_info.hidden:
        raise ValueError("导出校验失败：隐藏首行的模板类型标识发生变化。")
    return data


def analyze_and_apply(strategy_bytes, template_key, write_output=False, export_title=None):
    errors = []
    warnings = []
    xls_template_path = template_path_for_key(template_key)
    analysis_template_path = analysis_template_path_for_key(template_key)
    missing_templates = [
        path.name for path in (xls_template_path, analysis_template_path) if not path.exists()
    ]
    if missing_templates:
        return {
            "ok": False,
            "errors": [
                "缺少模板文件：%s。请把原始 .xls 和配套 .xlsx 放入 %s。"
                % ("、".join(missing_templates), TEMPLATE_DIR)
            ],
            "warnings": [],
        }

    # Validate the hidden import marker before doing any expensive parsing.
    open_validated_xls_template(template_key)

    strategy_data = read_strategy_products(strategy_bytes)
    warnings.extend(strategy_data["warnings"])

    try:
        mapping_data = load_inner_code_mapping()
    except ValueError as exc:
        return {"ok": False, "errors": [str(exc)], "warnings": warnings}

    missing_inner_codes, replaced_inner_codes = apply_inner_code_mapping(
        strategy_data["products"], mapping_data
    )
    if missing_inner_codes:
        samples = [
            "第%s行 %s（%s）"
            % (product["sourceRow"], product["name"] or "未命名商品", product["code"] or "无编码")
            for product in missing_inner_codes[:30]
        ]
        suffix = "" if len(missing_inner_codes) <= 30 else "；另有%s个未显示" % (len(missing_inner_codes) - 30)
        warnings.append(
            "以下商品在 %s 中找不到商品内码，将留空并继续导出，请人工复核：%s%s"
            % (INNER_CODE_MAPPING_FILENAME, "；".join(samples), suffix)
        )
    if replaced_inner_codes:
        warnings.append(
            "有%s个商品的策略表内码与映射文件不同，已以 %s 为准。"
            % (len(replaced_inner_codes), INNER_CODE_MAPPING_FILENAME)
        )

    template_book = XlsxBook(analysis_template_path.read_bytes())
    template_info = detect_template_columns(template_book, template_key)
    sheet = template_info["sheet"]
    header_row_no = template_info["headerRow"]
    columns = template_info["columns"]
    tier_cols = columns["tierCols"]
    base_rows = get_template_data_rows(template_book, sheet, header_row_no, columns)
    preserve_until_row_no = template_preserve_until_row(template_book, sheet, header_row_no, columns)

    if template_info["missingTiers"]:
        warnings.append("模板缺少这些档位列：%s" % "、".join("%s档" % tier for tier in template_info["missingTiers"]))
    if columns["county"] is None:
        warnings.append("模板没有县区列，策略中的县区限定无法区分，将按其他维度套用。")
    if columns["marketType"] is None:
        warnings.append("模板没有市场类型列，策略中的市场类型限定无法区分，将按其他维度套用。")

    preview_rows = []
    output_rows = []
    output_records = []
    generated_rows = 0
    matched_dimension_rows = 0
    conflict_count = 0

    for product in strategy_data["products"]:
        for base_row in base_rows:
            generated_rows += 1
            new_row_no = preserve_until_row_no + generated_rows
            row = template_book.clone_row(base_row, new_row_no)
            cells = template_book.row_cells(row)

            dim = {
                "county": template_book.cell_text(cells.get(columns["county"])),
                "marketType": template_book.cell_text(cells.get(columns["marketType"])),
                "businessArea": template_book.cell_text(cells.get(columns["businessArea"])),
                "hasCounty": columns["county"] is not None,
                "hasMarketType": columns["marketType"] is not None,
            }
            assignments, matched_rules, conflicts = evaluate_product_on_dimension(product, dim, tier_cols)
            if matched_rules:
                matched_dimension_rows += 1
            if conflicts:
                conflict_count += 1
                if conflict_count <= 20:
                    warnings.append(
                        "商品 %s 在导出行 %s 有规则覆盖：%s"
                        % (product["name"] or product["code"], new_row_no, "、".join(conflicts[:5]))
                    )

            code_cell = template_book.get_or_create_cell(row, columns["productCode"])
            name_cell = template_book.get_or_create_cell(row, columns["productName"])
            template_book.set_text(code_cell, product["code"])
            template_book.set_text(name_cell, product["name"])
            if columns.get("productInnerCode"):
                inner_cell = template_book.get_or_create_cell(row, columns["productInnerCode"])
                if product.get("innerCode"):
                    template_book.set_text(inner_cell, product["innerCode"])
                else:
                    template_book.clear_cell(inner_cell)

            for tier, col in tier_cols.items():
                style_source = cells.get(col)
                cell = template_book.get_or_create_cell(row, col, style_source)
                template_book.set_number(cell, assignments[tier])

            output_rows.append(row)
            output_records.append({
                "sourceRowNo": row_number(base_row),
                "product": product,
                "assignments": assignments,
            })
            if len(preview_rows) < 150:
                preview_rows.append({
                    "row": new_row_no,
                    "sourceRow": product["sourceRow"],
                    "productCode": product["code"],
                    "productInnerCode": product["innerCode"],
                    "productName": product["name"],
                    "county": dim["county"],
                    "marketType": dim["marketType"],
                    "businessArea": dim["businessArea"],
                    "matched": bool(matched_rules),
                    "matchedRules": matched_rules,
                    "assignment": compact_assignments(assignments),
                })

    summary = {
        "template": TEMPLATES[template_key]["label"],
        "strategySheet": strategy_data["sheetName"],
        "strategyHeaderRow": strategy_data["headerRow"],
        "templateHeaderRow": header_row_no,
        "templatePreserveUntilRow": preserve_until_row_no,
        "productCount": len(strategy_data["products"]),
        "innerCodeMappedProducts": len(strategy_data["products"]) - len(missing_inner_codes),
        "innerCodeMissingProducts": len(missing_inner_codes),
        "innerCodeMappingEntries": len(mapping_data["mapping"]),
        "innerCodeMappingFile": INNER_CODE_MAPPING_FILENAME,
        "templateDimensionRows": len(base_rows),
        "generatedRows": generated_rows,
        "matchedRows": matched_dimension_rows,
        "zeroFilledRows": generated_rows - matched_dimension_rows,
        "ruleCount": sum(len(product["rules"]) for product in strategy_data["products"]),
        "tierColumns": [{"tier": tier, "column": index_to_col(col)} for tier, col in sorted(tier_cols.items(), reverse=True)],
    }

    result = {
        "ok": not errors,
        "errors": errors,
        "warnings": warnings,
        "summary": summary,
        "rules": [
            {
                "sourceRow": product["sourceRow"],
                "productCode": product["code"],
                "productName": product["name"],
                "productInnerCode": product.get("innerCode", ""),
                "strategyText": product["strategyText"],
                "rules": product["rules"],
                "warnings": product["warnings"],
            }
            for product in strategy_data["products"]
        ],
        "previewRows": preview_rows,
    }
    if write_output:
        result["output"] = build_xls_output(
            template_key,
            columns,
            preserve_until_row_no,
            output_records,
            export_title or next_week_default_export_stem(),
        )
    return result


def parse_multipart(headers, body):
    content_type = headers.get("Content-Type", "")
    message = BytesParser(policy=default).parsebytes(
        ("Content-Type: " + content_type + "\r\n\r\n").encode("utf-8") + body
    )
    fields = {}
    files = {}
    if not message.is_multipart():
        return fields, files
    for part in message.iter_parts():
        disposition = part.get("Content-Disposition", "")
        name_match = re.search(r'name="([^"]+)"', disposition)
        filename_match = re.search(r'filename="([^"]*)"', disposition)
        if not name_match:
            continue
        name = name_match.group(1)
        payload = part.get_payload(decode=True) or b""
        if filename_match and filename_match.group(1):
            files[name] = {
                "filename": filename_match.group(1),
                "content_type": part.get_content_type(),
                "content": payload,
            }
        else:
            fields[name] = payload.decode("utf-8", errors="ignore")
    return fields, files


def template_status_payload():
    return [
        {
            "key": key,
            "label": item["label"],
            "filename": item["filename"],
            "exists": (
                template_path_for_key(key).exists()
                and analysis_template_path_for_key(key).exists()
            ),
            "marker": "%s / %s" % item["marker"],
        }
        for key, item in TEMPLATES.items()
    ]


def inner_code_mapping_status_payload():
    path = inner_code_mapping_path()
    return {
        "filename": INNER_CODE_MAPPING_FILENAME,
        "exists": path.exists(),
        "path": str(path),
    }


def export_period_from_filename(filename=None):
    value = Path((filename or "").strip()).stem
    match = re.search(r"\d{6}W\d+\(\d{2}\.\d{2}-\d{2}\.\d{2}\)", value)
    return match.group(0) if match else next_week_default_export_stem()


def build_export_for_session(session, requested_filename=None):
    export_title = export_period_from_filename(requested_filename)
    cache_key = "%s|%s" % (session["templateKey"], export_title)
    if session.get("outputCacheKey") == cache_key and "output" in session:
        return {"ok": True, "output": session["output"]}
    result = analyze_and_apply(
        session["strategy"],
        session["templateKey"],
        write_output=True,
        export_title=export_title,
    )
    if result.get("ok"):
        session["output"] = result["output"]
        session["outputCacheKey"] = cache_key
    return result



def next_week_default_export_stem(today=None):
    today = today or dt.date.today()
    days_until_next_monday = 7 - today.weekday()
    next_monday = today + dt.timedelta(days=days_until_next_monday)
    next_sunday = next_monday + dt.timedelta(days=6)
    week_in_month = ((next_monday.day - 1) // 7) + 1
    return "%04d%02dW%s(%02d.%02d-%02d.%02d)" % (
        next_monday.year,
        next_monday.month,
        week_in_month,
        next_monday.month,
        next_monday.day,
        next_sunday.month,
        next_sunday.day,
    )


def sanitize_export_filename(filename=None):
    final_name = (filename or "").strip() or next_week_default_export_stem()
    final_name = re.sub(r'[<>:"/\\|?*]+', "_", final_name)
    final_name = final_name.rstrip(" .")
    if not final_name:
        final_name = next_week_default_export_stem()
    if final_name.lower().endswith(".xlsx"):
        final_name = final_name[:-5]
    if not final_name.lower().endswith(".xls"):
        final_name += ".xls"
    return final_name


def save_export_bytes(data, filename=None):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    final_name = sanitize_export_filename(filename)
    saved_path = OUTPUT_DIR / final_name
    if saved_path.exists():
        stem = saved_path.stem
        suffix = saved_path.suffix
        saved_path = OUTPUT_DIR / ("%s_%s%s" % (stem, uuid.uuid4().hex[:6], suffix))
    saved_path.write_bytes(data)
    return saved_path


INDEX_HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>货源投放策略转导入Excel</title>
  <style>
    :root { color-scheme: light; --ink:#172033; --muted:#667085; --line:#d8dee8; --bg:#f7f9fc; --brand:#0f766e; --bad:#b42318; --warn:#a15c07; }
    * { box-sizing: border-box; }
    body { margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif; color:var(--ink); background:var(--bg); }
    header { padding:22px 28px; background:#fff; border-bottom:1px solid var(--line); }
    h1 { margin:0; font-size:22px; }
    main { max-width:1220px; margin:0 auto; padding:24px; }
    .panel { background:#fff; border:1px solid var(--line); border-radius:8px; padding:18px; margin-bottom:16px; }
    .grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
    label { display:block; font-weight:600; margin-bottom:8px; }
    input[type=file], input[type=text], select { width:100%; border:1px solid var(--line); border-radius:6px; padding:10px; background:#fff; }
    button { border:0; border-radius:6px; padding:10px 16px; background:var(--brand); color:#fff; cursor:pointer; font-weight:600; }
    button:disabled { background:#9aa4b2; cursor:not-allowed; }
    .actions { display:flex; gap:10px; align-items:center; margin-top:14px; flex-wrap:wrap; }
    .hint { color:var(--muted); font-size:13px; margin-top:8px; line-height:1.5; }
    .status { font-size:14px; color:var(--muted); }
    .messages { display:grid; gap:8px; }
    .msg { border-radius:6px; padding:10px 12px; font-size:14px; line-height:1.5; }
    .error { background:#fff1f0; color:var(--bad); border:1px solid #ffccc7; }
    .warning { background:#fff8eb; color:var(--warn); border:1px solid #ffd89a; }
    .ok { background:#ecfdf3; color:#067647; border:1px solid #abefc6; }
    table { width:100%; border-collapse:collapse; font-size:13px; background:#fff; }
    th, td { border-bottom:1px solid var(--line); padding:8px 10px; text-align:left; vertical-align:top; }
    th { background:#f2f5f9; font-weight:600; position:sticky; top:0; }
    pre { max-height:360px; overflow:auto; background:#111827; color:#e5e7eb; border-radius:6px; padding:14px; font-size:12px; }
    .table-wrap { max-height:430px; overflow:auto; border:1px solid var(--line); border-radius:8px; }
    .summary { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; }
    .metric { border:1px solid var(--line); border-radius:8px; padding:12px; }
    .metric b { display:block; font-size:22px; margin-bottom:2px; }
    @media (max-width:820px) { .grid,.summary { grid-template-columns:1fr; } main { padding:14px; } }
  </style>
</head>
<body>
  <header><h1>货源投放周策略转系统导入 Excel</h1></header>
  <main>
    <section class="panel">
      <div class="grid">
        <div>
          <label for="strategyFile">1. 上传周策略表（.xlsx）</label>
          <input id="strategyFile" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
          <div class="hint">程序按表头名称定位“本期投放策略”，不依赖固定列号，兼容新增“在途”等列。当前版本：<b>__BUILD_VERSION__</b></div>
        </div>
        <div>
          <label for="templateKey">2. 选择导出格式</label>
          <select id="templateKey"></select>
          <div class="hint" id="templateHint">正在检查模板目录...</div>
        </div>
      </div>
      <div style="margin-top:16px;">
        <label for="exportName">3. 导出文件名（可选）</label>
        <input id="exportName" type="text" value="__DEFAULT_EXPORT_NAME__" placeholder="__DEFAULT_EXPORT_NAME__" />
        <div class="hint">不填写时默认使用下周日期，程序会生成真正的 Excel 97-2003 .xls 文件。</div>
      </div>
      <div class="actions">
        <button id="analyzeBtn">解析并预览校验</button>
        <button id="exportBtn" disabled>导出 XLS</button>
        <span id="status" class="status">请先上传周策略表并完成预览。</span>
      </div>
    </section>

    <section class="panel" id="messagePanel" hidden><div class="messages" id="messages"></div></section>

    <section class="panel" id="summaryPanel" hidden>
      <h2 style="margin:0 0 12px;font-size:17px;">校验概览</h2>
      <div class="summary" id="summary"></div>
    </section>

    <section class="panel" id="previewPanel" hidden>
      <h2 style="margin:0 0 12px;font-size:17px;">导出预览（最多显示150行）</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>导出行</th><th>策略行</th><th>商品编码</th><th>商品内码</th><th>商品名称</th><th>区县</th><th>市场类型</th><th>商圈</th><th>命中规则</th><th>档位写入</th></tr></thead>
          <tbody id="previewRows"></tbody>
        </table>
      </div>
    </section>

    <section class="panel" id="rulesPanel" hidden>
      <h2 style="margin:0 0 12px;font-size:17px;">结构化规则 JSON</h2>
      <pre id="rulesJson"></pre>
    </section>
  </main>

  <script>
    const strategyFile = document.getElementById("strategyFile");
    const templateKey = document.getElementById("templateKey");
    const templateHint = document.getElementById("templateHint");
    const exportName = document.getElementById("exportName");
    const analyzeBtn = document.getElementById("analyzeBtn");
    const exportBtn = document.getElementById("exportBtn");
    const statusNode = document.getElementById("status");
    let sessionId = null;

    async function loadTemplates() {
      const resp = await fetch("/api/templates");
      const data = await resp.json();
      templateKey.innerHTML = data.templates.map(item =>
        `<option value="${item.key}" ${item.exists ? "" : "disabled"}>${item.label}${item.exists ? `（${item.marker}）` : "（模板缺失）"}</option>`
      ).join("");
      const missing = data.templates.filter(item => !item.exists).map(item => item.filename);
      const mapping = data.innerCodeMapping;
      if (!mapping.exists) {
        templateHint.textContent = "缺少商品内码映射：" + mapping.filename + "。请放入：" + data.templateDirectory;
      } else {
        templateHint.textContent = missing.length
          ? "缺少私有模板：" + missing.join("、") + "。请放入：" + data.templateDirectory
          : "私有模板和商品内码映射已就绪：" + data.templateDirectory;
      }
    }

    function escapeHtml(value) {
      return String(value ?? "").replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
    }

    function showMessages(data) {
      const panel = document.getElementById("messagePanel");
      const list = document.getElementById("messages");
      const errors = data.errors || [];
      const warnings = data.warnings || [];
      list.innerHTML = "";
      if (!errors.length && !warnings.length) list.innerHTML = '<div class="msg ok">校验通过：没有错误或警告。</div>';
      errors.forEach(item => list.insertAdjacentHTML("beforeend", `<div class="msg error">${escapeHtml(item)}</div>`));
      warnings.forEach(item => list.insertAdjacentHTML("beforeend", `<div class="msg warning">${escapeHtml(item)}</div>`));
      panel.hidden = false;
    }

    function showSummary(summary) {
      const box = document.getElementById("summary");
      const metrics = [
        ["商品数", summary.productCount],
        ["已匹配内码", summary.innerCodeMappedProducts],
        ["缺失内码", summary.innerCodeMissingProducts],
        ["内码映射数", summary.innerCodeMappingEntries],
        ["模板维度行", summary.templateDimensionRows],
        ["生成行", summary.generatedRows],
        ["命中行", summary.matchedRows],
        ["填0行", summary.zeroFilledRows],
        ["规则数", summary.ruleCount],
        ["策略表头行", summary.strategyHeaderRow],
        ["模板表头行", summary.templateHeaderRow],
      ];
      box.innerHTML = metrics.map(([label, value]) => `<div class="metric"><b>${escapeHtml(value)}</b><span>${label}</span></div>`).join("");
      document.getElementById("summaryPanel").hidden = false;
    }

    function showPreview(rows) {
      const body = document.getElementById("previewRows");
      body.innerHTML = rows.map(row => `
        <tr>
          <td>${row.row}</td>
          <td>${row.sourceRow}</td>
          <td>${escapeHtml(row.productCode)}</td>
          <td>${escapeHtml(row.productInnerCode)}</td>
          <td>${escapeHtml(row.productName)}</td>
          <td>${escapeHtml(row.county)}</td>
          <td>${escapeHtml(row.marketType)}</td>
          <td>${escapeHtml(row.businessArea)}</td>
          <td>${escapeHtml((row.matchedRules || []).join(", "))}</td>
          <td>${escapeHtml(row.assignment)}</td>
        </tr>
      `).join("");
      document.getElementById("previewPanel").hidden = false;
    }

    analyzeBtn.addEventListener("click", async () => {
      if (!strategyFile.files[0]) {
        alert("请先上传周策略表 .xlsx");
        return;
      }
      const form = new FormData();
      form.append("strategyWorkbook", strategyFile.files[0]);
      form.append("templateKey", templateKey.value);
      analyzeBtn.disabled = true;
      exportBtn.disabled = true;
      statusNode.textContent = "正在解析周策略表和模板...";
      try {
        const resp = await fetch("/api/analyze", { method: "POST", body: form });
        const data = await resp.json();
        showMessages(data);
        if (!data.ok) {
          statusNode.textContent = "校验未通过，请处理错误后重试。";
          return;
        }
        sessionId = data.sessionId;
        showSummary(data.summary);
        showPreview(data.previewRows || []);
        document.getElementById("rulesJson").textContent = JSON.stringify(data.rules, null, 2);
        document.getElementById("rulesPanel").hidden = false;
        exportBtn.disabled = false;
        statusNode.textContent = "已完成预览校验，可导出 Excel。";
      } catch (err) {
        statusNode.textContent = "解析失败：" + err.message;
      } finally {
        analyzeBtn.disabled = false;
      }
    });

    function filenameFromDisposition(disposition) {
      const fallback = "strategy_import.xlsx";
      if (!disposition) return fallback;
      const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
      if (utf8Match) return decodeURIComponent(utf8Match[1]);
      const asciiMatch = disposition.match(/filename="([^"]+)"/i);
      return asciiMatch ? asciiMatch[1] : fallback;
    }

    exportBtn.addEventListener("click", async () => {
      if (!sessionId) return;
      exportBtn.disabled = true;
      statusNode.textContent = "正在生成 Excel，数据量较大时可能需要几十秒...";
      try {
        const params = new URLSearchParams({ sessionId });
        const requestedName = (exportName.value || "").trim();
        if (requestedName) params.set("filename", requestedName);
        const resp = await fetch("/api/export-save?" + params.toString());
        if (!resp.ok) {
          let data = null;
          try { data = await resp.json(); } catch (_) {}
          showMessages(data || { errors: ["导出失败，HTTP " + resp.status] });
          statusNode.textContent = "导出失败，请查看错误信息。";
          return;
        }
        const data = await resp.json();
        const message = "Excel 已生成并保存到：" + data.savedPath;
        showMessages({
          warnings: [
            message,
            "如果浏览器允许下载，也可以点击下载链接：" + location.origin + data.downloadUrl
          ]
        });
        statusNode.textContent = message;
      } catch (err) {
        showMessages({
          errors: [
            "导出失败：" + err.message,
            "如果解析预览已成功但这里报 Failed to fetch，通常是单位电脑拦截本地下载请求或服务被安全软件中断。请重新打开程序，页面地址会自动换到可用端口。"
          ]
        });
        statusNode.textContent = "导出失败，请查看错误信息。";
      } finally {
        exportBtn.disabled = false;
      }
    });

    loadTemplates();
  </script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    server_version = "StrategyExcelTool/0.2"

    def send_json(self, status, payload):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_bytes(self, status, data, content_type, filename=None, extra_headers=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        if filename:
            safe_name = filename.encode("ascii", errors="ignore").decode("ascii") or "strategy_import.xls"
            self.send_header(
                "Content-Disposition",
                "attachment; filename=\"%s\"; filename*=UTF-8''%s" % (safe_name, quote(filename)),
            )
        for key, value in (extra_headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            page = INDEX_HTML.replace("__BUILD_VERSION__", BUILD_VERSION).replace("__DEFAULT_EXPORT_NAME__", next_week_default_export_stem())
            self.send_bytes(200, page.encode("utf-8"), "text/html; charset=utf-8")
            return
        if parsed.path == "/api/version":
            self.send_json(200, {"ok": True, "version": BUILD_VERSION})
            return
        if parsed.path == "/api/templates":
            self.send_json(200, {
                "ok": True,
                "templates": template_status_payload(),
                "innerCodeMapping": inner_code_mapping_status_payload(),
                "templateDirectory": str(TEMPLATE_DIR),
                "configuredBy": "STRATEGY_TOOL_TEMPLATE_DIR" if os.environ.get("STRATEGY_TOOL_TEMPLATE_DIR") else "default",
            })
            return
        if parsed.path == "/api/download":
            name = parse_qs(parsed.query).get("name", [""])[0]
            path = OUTPUT_DIR / Path(name).name
            if not path.exists() or not path.is_file() or path.suffix.lower() != ".xls":
                self.send_json(404, {"ok": False, "errors": ["导出文件不存在，请重新导出。"]})
                return
            self.send_bytes(
                200,
                path.read_bytes(),
                "application/vnd.ms-excel",
                path.name,
                {"X-Saved-Path": str(path)},
            )
            return
        if parsed.path == "/api/export-save":
            query = parse_qs(parsed.query)
            session_id = query.get("sessionId", [""])[0]
            requested_filename = query.get("filename", [""])[0]
            session = SESSIONS.get(session_id)
            if not session:
                self.send_json(404, {"ok": False, "errors": ["预览会话已失效，请重新解析。"]})
                return
            try:
                result = build_export_for_session(session, requested_filename)
                if not result.get("ok"):
                    self.send_json(400, result)
                    return
                saved_path = save_export_bytes(result["output"], requested_filename)
                self.send_json(200, {
                    "ok": True,
                    "filename": saved_path.name,
                    "savedPath": str(saved_path),
                    "downloadUrl": "/api/download?name=%s" % quote(saved_path.name),
                })
            except Exception as exc:
                self.send_json(500, {"ok": False, "errors": ["导出失败：%s" % exc]})
            return
        if parsed.path == "/api/export":
            query = parse_qs(parsed.query)
            session_id = query.get("sessionId", [""])[0]
            requested_filename = query.get("filename", [""])[0]
            session = SESSIONS.get(session_id)
            if not session:
                self.send_json(404, {"ok": False, "errors": ["预览会话已失效，请重新解析。"]})
                return
            try:
                result = build_export_for_session(session, requested_filename)
            except Exception as exc:
                self.send_json(500, {"ok": False, "errors": [str(exc)]})
                return
            if not result.get("ok"):
                self.send_json(400, result)
                return
            saved_path = save_export_bytes(result["output"], requested_filename)
            self.send_bytes(
                200,
                result["output"],
                "application/vnd.ms-excel",
                saved_path.name,
                {"X-Saved-Path": str(saved_path)},
            )
            return
        self.send_json(404, {"ok": False, "errors": ["not_found"]})

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/api/analyze":
            self.send_json(404, {"ok": False, "errors": ["not_found"]})
            return

        length = int(self.headers.get("Content-Length", "0"))
        fields, files = parse_multipart(self.headers, self.rfile.read(length))
        strategy_file = files.get("strategyWorkbook")
        template_key = fields.get("templateKey", "")

        if not strategy_file:
            self.send_json(400, {"ok": False, "errors": ["缺少周策略 Excel 文件。"]})
            return
        if not strategy_file["filename"].lower().endswith(".xlsx"):
            self.send_json(400, {"ok": False, "errors": ["周策略表必须是 .xlsx 文件。"]})
            return
        if template_key not in TEMPLATES:
            self.send_json(400, {"ok": False, "errors": ["请选择有效的导出格式。"]})
            return

        try:
            result = analyze_and_apply(strategy_file["content"], template_key, write_output=False)
            if result.get("ok"):
                session_id = uuid.uuid4().hex
                SESSIONS[session_id] = {
                    "strategy": strategy_file["content"],
                    "templateKey": template_key,
                    "createdAt": time.time(),
                }
                result["sessionId"] = session_id
                result["version"] = BUILD_VERSION
            self.send_json(200 if result.get("ok") else 400, result)
        except zipfile.BadZipFile:
            self.send_json(400, {"ok": False, "errors": ["文件不是有效的 .xlsx 工作簿。"]})
        except Exception as exc:
            self.send_json(500, {"ok": False, "errors": [str(exc)]})

    def log_message(self, format, *args):
        print("[%s] %s" % (self.log_date_time_string(), format % args))


def xml_escape(value):
    return html.escape("" if value is None else str(value), quote=False)


def build_demo_xlsx(headers, rows):
    shared = []
    shared_index = {}

    def sid(value):
        value = "" if value is None else str(value)
        if value not in shared_index:
            shared_index[value] = len(shared)
            shared.append(value)
        return shared_index[value]

    row_xml = []
    all_rows = [headers] + rows
    for row_no, values in enumerate(all_rows, start=1):
        cells = []
        for col_no, value in enumerate(values, start=1):
            ref = "%s%s" % (index_to_col(col_no), row_no)
            if isinstance(value, (int, float)):
                cells.append('<c r="%s"><v>%s</v></c>' % (ref, value))
            else:
                cells.append('<c r="%s" t="s"><v>%s</v></c>' % (ref, sid(value)))
        row_xml.append('<row r="%s">%s</row>' % (row_no, "".join(cells)))

    shared_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    shared_xml.append('<sst xmlns="%s" count="%s" uniqueCount="%s">' % (NS_MAIN, len(shared), len(shared)))
    for text in shared:
        shared_xml.append("<si><t>%s</t></si>" % xml_escape(text))
    shared_xml.append("</sst>")

    last_ref = "%s%s" % (index_to_col(len(headers)), len(all_rows))
    sheet_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="%s" xmlns:r="%s"><dimension ref="A1:%s"/><sheetData>%s</sheetData></worksheet>""" % (NS_MAIN, NS_REL, last_ref, "".join(row_xml))
    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="%s" xmlns:r="%s"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""" % (NS_MAIN, NS_REL)
    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""" % NS_PKG_REL
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""" % NS_PKG_REL

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        zf.writestr("xl/sharedStrings.xml", "".join(shared_xml))
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return output.getvalue()


def self_test():
    strategy_headers = [
        "厂家", "商品编码", "商品名称", "卷烟类别", "品类", "批发价", "零售价",
        "货源属性分类", "实物库存", "在途", "合同意向", "上期销量", "本期预计销量",
        "本期投放方式", "上期投放策略", "本期投放策略", "备注",
    ]
    def strategy_row(code, name, strategy):
        return [
            "广西中烟", code, name, "", "", "", "", "", "", "", "", "", "", "",
            "上期策略", strategy, "",
        ]
    strategy_rows = [
        strategy_row("01614", "真龙（软海韵）", "城市、县城：20-25档2条，26-30档3条。乡镇、农村：20-25档1条，26-30档2条。"),
        strategy_row("01361", "真龙（海韵）", "F县、G县：1-30档3条。"),
        strategy_row("01497", "真龙（神韵）", "无库存。"),
        strategy_row("01698", "真龙（甲天下中支）", "25-30档1条。"),
    ]
    strategy_bytes = build_demo_xlsx(strategy_headers, strategy_rows)
    strategy_book = XlsxBook(strategy_bytes)
    _, _, strategy_columns = detect_strategy_table(strategy_book)
    assert strategy_columns["strategy"] == 16
    result = analyze_and_apply(
        strategy_bytes,
        "district_type",
        write_output=True,
        export_title="202607W4(07.27-08.02)",
    )
    assert result["ok"], result
    assert result["summary"]["productCount"] == 4, result["summary"]
    assert result["summary"]["innerCodeMappedProducts"] == 4, result["summary"]
    assert result["summary"]["templateDimensionRows"] == 44, result["summary"]
    assert result["summary"]["generatedRows"] == 176, result["summary"]
    assert result["output"].startswith(bytes.fromhex("D0CF11E0A1B11AE1"))

    out = xlrd.open_workbook(file_contents=result["output"], formatting_info=True)
    sheet = out.sheet_by_index(0)
    assert sheet.rowinfo_map[0].hidden
    assert sheet.cell_value(0, 0) == "M200007"
    assert normalize_xls_marker_value(sheet.cell_value(0, 1)) == "2"
    assert sheet.cell_value(1, 0) == "202607W4(07.27-08.02)"
    assert sheet.cell_value(4, 2) == "01614"
    assert sheet.cell_value(4, 1)
    assert sheet.cell_value(4, 8) == 3
    no_stock_rowx = 4 + (2 * 44)
    assert sheet.cell_value(no_stock_rowx, 8) == 0
    global_rule_rowx = 4 + (3 * 44)
    assert sheet.cell_value(global_rule_rowx, 8) == 1
    assert sheet.cell_value(global_rule_rowx, 14) == 0

    missing_product = {"code": "99999", "name": "测试缺失内码", "innerCode": "策略表旧值"}
    missing, replaced = apply_inner_code_mapping([missing_product], {"mapping": {}})
    assert len(missing) == 1
    assert not replaced
    assert missing_product["innerCode"] == ""
    assert name_matches("城市", ["市区"])
    print("self-test passed")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--port", type=int, default=PORT)
    args = parser.parse_args()
    if args.self_test:
        self_test()
        return
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    server, actual_port = create_http_server(HOST, args.port, Handler)
    print("Strategy Excel tool running at http://%s:%s" % (HOST, actual_port))
    if actual_port != args.port:
        print("Preferred port %s was unavailable, switched to %s." % (args.port, actual_port))
    print("Template directory: %s" % TEMPLATE_DIR)
    print("Output directory: %s" % OUTPUT_DIR)
    server.serve_forever()


if __name__ == "__main__":
    main()
