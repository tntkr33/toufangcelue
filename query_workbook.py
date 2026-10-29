"""Streaming XLSX exports for lazy customer-query records, using only stdlib."""

from decimal import Decimal
from html import escape
import io
import math
from pathlib import Path
import re
import zipfile


_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
_INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")
_EXCEL_ESCAPE = re.compile(r"_x[0-9a-fA-F]{4}_")
_PROGRESS_INTERVAL = 10000

_STYLES = _XML + f'''<styleSheet xmlns="{_MAIN}">
<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF286650"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def _column_name(index):
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _text(value):
    value = _INVALID_XML.sub("", str(value))
    if len(value) > 32767 or len(value.encode("utf-16-le")) // 2 > 32767:
        raise ValueError("XLSX cell text exceeds 32767 characters")
    # Protect literal OOXML escape-looking strings, including customer codes.
    value = _EXCEL_ESCAPE.sub(lambda match: "_x005F_" + match.group()[1:], value)
    return escape(value, quote=False).replace("\r", "&#13;")


def _cell(ref, value, style=None):
    if value is None:
        return f'<c r="{ref}"/>'
    if style is None and isinstance(value, bool):
        return f'<c r="{ref}" t="b"><v>{int(value)}</v></c>'
    if style is None and isinstance(value, (int, float, Decimal)):
        if isinstance(value, Decimal):
            finite = value.is_finite()
        else:
            finite = not isinstance(value, float) or math.isfinite(value)
        if not finite:
            raise ValueError("XLSX numeric cells must be finite")
        number_style = 3 if isinstance(value, int) else 2
        return f'<c r="{ref}" s="{number_style}"><v>{value}</v></c>'
    return (f'<c r="{ref}" s="{0 if style is None else style}" t="inlineStr">'
            f'<is><t xml:space="preserve">{_text(value)}</t></is></c>')


def _write_package(book, sheet_count):
    sheets = "".join(f'<sheet name="Query {i}" sheetId="{i}" r:id="rId{i}"/>'
                     for i in range(1, sheet_count + 1))
    book.writestr("xl/workbook.xml", _XML +
                  f'<workbook xmlns="{_MAIN}" xmlns:r="{_REL}"><sheets>{sheets}</sheets></workbook>')
    relationships = "".join(
        f'<Relationship Id="rId{i}" Type="{_REL}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, sheet_count + 1))
    book.writestr("xl/_rels/workbook.xml.rels", _XML +
                  f'<Relationships xmlns="{_PKG_REL}">{relationships}'
                  f'<Relationship Id="rId{sheet_count + 1}" Type="{_REL}/styles" Target="styles.xml"/>'
                  '</Relationships>')
    book.writestr("_rels/.rels", _XML +
                  f'<Relationships xmlns="{_PKG_REL}"><Relationship Id="rId1" '
                  f'Type="{_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, sheet_count + 1))
    book.writestr("[Content_Types].xml", _XML +
                  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                  '<Default Extension="xml" ContentType="application/xml"/>'
                  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                  '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                  + overrides + '</Types>')
    book.writestr("xl/styles.xml", _STYLES)


def write_query_workbook(path, columns, rows, metadata, sheet_row_limit=1048576, progress=None):
    """Write a lazy iterable of dict records to the caller's temporary path.

    Strings remain literal text; int/float/Decimal values are numeric cells,
    with non-integers displayed to exactly three decimal places. Missing/None
    values are blank. Each metadata line occupies one row before the header.
    The row limit includes these rows and must leave room for one data row.
    Even an empty query produces one sheet with metadata and a filter header.

    Progress receives fresh {rowCount, sheetCount} dicts at sheet starts,
    every 10,000 data rows, and after closing the completed ZIP. Any write,
    iterator or callback failure removes the output and propagates the error.
    The caller must not expose this temporary destination until success.
    Memory is bounded by one record, column/metadata definitions and the ZIP
    sheet directory (not the number of records or unique string values).
    """
    columns = tuple(columns)
    metadata = tuple(metadata)
    if not 1 <= len(columns) <= 16384:
        raise ValueError("XLSX requires between 1 and 16384 columns")
    header_row = len(metadata) + 1
    if (type(sheet_row_limit) is not int or
            not header_row < sheet_row_limit <= 1048576):
        raise ValueError("Sheet row limit must include metadata, header and at least one data row, up to 1048576")
    refs = tuple(_column_name(i) for i in range(1, len(columns) + 1))
    keys = tuple(column["key"] for column in columns)
    headers = "".join(_cell(f"{ref}{header_row}", column["label"], 1)
                      for ref, column in zip(refs, columns))
    metadata_xml = "".join(f'<row r="{i}" ht="30" customHeight="1">{_cell(f"A{i}", line, 4)}</row>'
                           for i, line in enumerate(metadata, 1))
    widths = "".join(
        f'<col min="{i}" max="{i}" width="{min(48, max(18, len(str(column["label"])) * 2 + 4))}" customWidth="1"/>'
        for i, column in enumerate(columns, 1))
    preamble = (_XML + f'<worksheet xmlns="{_MAIN}">'
                f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="{header_row}" '
                f'topLeftCell="A{header_row + 1}" activePane="bottomLeft" state="frozen"/>'
                f'<selection pane="bottomLeft" activeCell="A{header_row + 1}" sqref="A{header_row + 1}"/>'
                '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
                f'<cols>{widths}</cols><sheetData>{metadata_xml}'
                f'<row r="{header_row}" ht="32" customHeight="1">{headers}</row>').encode("utf-8")
    merges = ""
    if metadata and len(columns) > 1:
        merges = f'<mergeCells count="{len(metadata)}">' + "".join(
            f'<mergeCell ref="A{i}:{refs[-1]}{i}"/>' for i in range(1, len(metadata) + 1)) + '</mergeCells>'
    target = Path(path)
    row_count = sheet_count = 0
    opened = False

    def report():
        if progress is not None:
            progress({"rowCount": row_count, "sheetCount": sheet_count})

    try:
        with target.open("wb") as output:
            opened = True
            with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True) as book:
                iterator = iter(rows)
                sentinel = object()
                record = next(iterator, sentinel)
                while True:
                    sheet_count += 1
                    report()
                    physical_row = header_row
                    with book.open(f"xl/worksheets/sheet{sheet_count}.xml", "w", force_zip64=True) as raw:
                        with io.BufferedWriter(raw, buffer_size=65536) as sheet:
                            sheet.write(preamble)
                            while record is not sentinel and physical_row < sheet_row_limit:
                                physical_row += 1
                                cells = "".join(_cell(f"{ref}{physical_row}", record.get(key))
                                                for ref, key in zip(refs, keys))
                                sheet.write(f'<row r="{physical_row}">{cells}</row>'.encode("utf-8"))
                                row_count += 1
                                if row_count % _PROGRESS_INTERVAL == 0:
                                    report()
                                record = next(iterator, sentinel)
                            sheet.write((f'</sheetData><autoFilter ref="A{header_row}:{refs[-1]}{physical_row}"/>'
                                         + merges + '</worksheet>').encode("utf-8"))
                    if record is sentinel:
                        break
                _write_package(book, sheet_count)
        report()
        return {"rowCount": row_count, "sheetCount": sheet_count}
    except BaseException:
        if opened:
            target.unlink(missing_ok=True)
        raise
