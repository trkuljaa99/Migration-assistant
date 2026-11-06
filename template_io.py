"""Read a Migration Cockpit template and write it back prefilled.

The cockpit's file templates come as XML Spreadsheet 2003 (.xml) or .xlsx: an introduction sheet, a field list, and
one data sheet per structure (e.g. "Basic Data", "Sales Data"). On a data sheet the header block holds the field
groups, the field descriptions (mandatory ones marked with *), the technical field names and often the data type /
length; the data goes in below it. The exact rows differ by release, so the header is found from its content: the
row of technical names (mostly UPPERCASE_WITH_UNDERSCORES) anchors it, the row above it are the descriptions, the
rows above those the groups, and the data starts at the first empty row below it.

Writing keeps the uploaded file as it is (an .xml template is edited as text, so SAP's styles and hidden rows stay
untouched) and adds the data rows; an .xlsx copy of every template is written too, for Excel and the preview.
"""
import html
import re
import xml.etree.ElementTree as ET
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

NS = "urn:schemas-microsoft-com:office:spreadsheet"
TECH = re.compile(r"^[A-Z][A-Z0-9_/\-]{1,39}$")
TYPE_LEN = re.compile(r"(?i)\b(text|char|string|number|numc|dec|decimal|date|time|boolean|curr|quan)\w*\s*\(?\s*(\d+)?")
SKIP_SHEETS = re.compile(r"(?i)^(introduction|field list|instructions?|help|readme)$")


# ---------------------------------------------------------------- reading

def _xml_grid(path):
    """{sheet name: rows as lists of strings} from an XML Spreadsheet 2003 file (ss:Index gaps honoured)."""
    root = ET.parse(path).getroot()
    q = lambda tag: f"{{{NS}}}{tag}"
    sheets = {}
    for ws in root.iter(q("Worksheet")):
        rows, r = [], 0
        table = ws.find(q("Table"))
        for row in (table.findall(q("Row")) if table is not None else []):
            r = int(row.get(q("Index"), r + 1))
            while len(rows) < r - 1:
                rows.append([])
            cells, c = [], 0
            for cell in row.findall(q("Cell")):
                c = int(cell.get(q("Index"), c + 1))
                while len(cells) < c - 1:
                    cells.append("")
                data = cell.find(q("Data"))
                text = "".join(data.itertext()).strip() if data is not None else ""
                cells.append(text)
                c += int(cell.get(q("MergeAcross"), 0))
                while len(cells) < c:
                    cells.append("")
            rows.append(cells)
        sheets[ws.get(q("Name"))] = rows
    return sheets


def _xlsx_grid(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    return {ws.title: [["" if v is None else str(v).strip() for v in row] for row in ws.iter_rows(values_only=True)]
            for ws in wb.worksheets}


def grid(path):
    return _xml_grid(path) if Path(path).suffix.lower() == ".xml" else _xlsx_grid(path)


def _cell(row, i):
    return row[i] if i < len(row) else ""


def _field_list(sheets):
    """{(sheet, description or technical name): {"importance", "type", "length"}} from a "Field List" sheet."""
    out = {}
    for name, rows in sheets.items():
        if not re.search(r"(?i)field\s*list", name):
            continue
        for h, row in enumerate(rows):
            low = [c.lower() for c in row]
            imp = next((i for i, c in enumerate(low) if "importance" in c or c.startswith("mandatory")), None)
            fld = next((i for i, c in enumerate(low) if "field" in c and ("desc" in c or "name" in c)), None)
            if imp is None or fld is None:
                continue
            sh = next((i for i, c in enumerate(low) if "sheet" in c), None)
            tech = next((i for i, c in enumerate(low) if "technical" in c or c in ("field", "field name")), None)
            typ = next((i for i, c in enumerate(low) if c in ("type", "data type")), None)
            ln = next((i for i, c in enumerate(low) if c in ("length", "max. length", "max length")), None)
            current = ""
            for r in rows[h + 1:]:
                current = _cell(r, sh) or current if sh is not None else ""
                info = {"importance": _cell(r, imp), "type": _cell(r, typ) if typ is not None else "",
                        "length": _cell(r, ln) if ln is not None else ""}
                for key in {_cell(r, fld), _cell(r, tech) if tech is not None else ""} - {""}:
                    out[(current.lower(), key.lower())] = info
            break
    return out


def parse(path):
    """The template's data sheets: [{"name", "tech_row", "data_start", "fields": [field]}] (0-based rows);
    field = {"col", "tech", "label", "group", "mandatory", "importance", "type", "length"}."""
    sheets = grid(path)
    info = _field_list(sheets)
    out = []
    for name, rows in sheets.items():
        if SKIP_SHEETS.match(name.strip()) or re.search(r"(?i)field\s*list", name):
            continue
        tech_row = None
        for i, row in enumerate(rows[:40]):
            vals = [c for c in row if c]
            hits = sum(bool(TECH.match(c)) for c in vals)
            if hits >= 2 and hits >= 0.6 * len(vals):
                tech_row = i
                break
        if tech_row is None:
            continue
        labels = rows[tech_row - 1] if tech_row > 0 else []
        groups = rows[tech_row - 2] if tech_row > 1 else []
        data_start = next((i for i in range(tech_row + 1, len(rows)) if not any(rows[i])), len(rows))
        meta_rows = rows[tech_row + 1:data_start]
        fields, group = [], ""
        for col, tech in enumerate(rows[tech_row]):
            group = _cell(groups, col) or group
            if not TECH.match(tech or ""):
                continue
            label = _cell(labels, col) or tech
            fl = info.get((name.lower(), label.strip("* ").lower())) or info.get((name.lower(), tech.lower())) \
                or info.get(("", label.strip("* ").lower())) or info.get(("", tech.lower())) or {}
            importance = fl.get("importance", "")
            typ, length = fl.get("type", ""), fl.get("length", "")
            for meta in meta_rows:  # a data type / length row under the technical names
                m = TYPE_LEN.search(_cell(meta, col))
                if m:
                    typ = typ or m.group(1)
                    length = length or (m.group(2) or "")
            mandatory = label.strip().startswith("*") or label.strip().endswith("*") or \
                importance.lower().startswith("mandatory") and "for" not in importance.lower()
            fields.append({"col": col, "tech": tech, "label": label.strip("* ").strip(), "group": group,
                           "mandatory": bool(mandatory), "importance": importance or ("Mandatory" if mandatory else ""),
                           "type": typ, "length": int(length) if str(length).isdigit() else None})
        if fields:
            out.append({"name": name, "tech_row": tech_row, "data_start": data_start, "fields": fields})
    return out


# ---------------------------------------------------------------- writing

def _xml_value(v):
    return html.escape(str(v), quote=False)


def _fill_xml(path_in, path_out, sheets, data):
    """The .xml template as text with the data rows added on each sheet (rows from data_start on replaced)."""
    text = Path(path_in).read_text(encoding="utf-8")
    for sheet in sheets:
        rows = data.get(sheet["name"])
        if rows is None:
            continue
        ws = re.search(r'<((?:ss:)?)Worksheet\b[^>]*?ss:Name="' + re.escape(html.escape(sheet["name"], quote=True))
                       + r'"[^>]*>(.*?)</(?:ss:)?Worksheet>', text, re.S)
        if not ws:
            continue
        p = ws.group(1)
        body = ws.group(2)
        tbl = re.search(r"(<(?:ss:)?Table\b[^>]*>)(.*?)(</(?:ss:)?Table>)", body, re.S)
        if not tbl:
            continue
        open_tag = re.sub(r'\s+ss:Expanded(?:Row|Column)Count="\d+"', "", tbl.group(1))
        inner = tbl.group(2)
        first = re.search(r"<(?:ss:)?Row\b", inner)
        head, rows_xml = (inner[:first.start()], inner[first.start():]) if first else (inner, "")
        kept, index = [], 0
        for m in re.finditer(r"<(?:ss:)?Row\b[^>]*/>|<(?:ss:)?Row\b.*?</(?:ss:)?Row>", rows_xml, re.S):
            idx = re.search(r'ss:Index="(\d+)"', m.group(0)[:200])
            index = int(idx.group(1)) if idx else index + 1
            if index <= sheet["data_start"]:
                kept.append(m.group(0))
        new = []
        for r, values in enumerate(rows):
            cells = "".join(
                f'<{p}Cell ss:Index="{f["col"] + 1}"><{p}Data ss:Type="String">{_xml_value(values[f["tech"]])}'
                f"</{p}Data></{p}Cell>" for f in sheet["fields"] if values.get(f["tech"]) not in (None, ""))
            new.append(f'<{p}Row ss:Index="{sheet["data_start"] + r + 1}">{cells}</{p}Row>')
        new_body = body[:tbl.start()] + open_tag + head + "".join(kept) + "".join(new) + tbl.group(3) + body[tbl.end():]
        text = text[:ws.start(2)] + new_body + text[ws.end(2):]
    Path(path_out).write_text(text, encoding="utf-8")


def _fill_xlsx(path_in, path_out, sheets, data):
    wb = openpyxl.load_workbook(path_in)
    for sheet in sheets:
        rows = data.get(sheet["name"])
        if rows is None or sheet["name"] not in wb.sheetnames:
            continue
        ws = wb[sheet["name"]]
        if ws.max_row > sheet["data_start"]:
            ws.delete_rows(sheet["data_start"] + 1, ws.max_row - sheet["data_start"])
        for r, values in enumerate(rows):
            for f in sheet["fields"]:
                v = values.get(f["tech"])
                if v not in (None, ""):
                    ws.cell(row=sheet["data_start"] + r + 1, column=f["col"] + 1, value=str(v))
    wb.save(path_out)


def _xlsx_copy(path_in, path_out, sheets, data):
    """An .xlsx version of the whole template (every sheet's header as text) with the data, for Excel."""
    src = grid(path_in)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    by_name = {s["name"]: s for s in sheets}
    head_fill = PatternFill("solid", fgColor="DCE6F1")
    for name, rows in src.items():
        ws = wb.create_sheet(name[:31])
        sheet = by_name.get(name)
        header = rows[:sheet["data_start"]] if sheet else rows
        for r, row in enumerate(header, 1):
            for c, v in enumerate(row, 1):
                if v:
                    cell = ws.cell(row=r, column=c, value=v)
                    if sheet and r - 1 >= sheet["tech_row"] - 1:
                        cell.font, cell.fill = Font(bold=True), head_fill
        if sheet:
            for r, values in enumerate(data.get(name) or [], sheet["data_start"] + 1):
                for f in sheet["fields"]:
                    v = values.get(f["tech"])
                    if v not in (None, ""):
                        ws.cell(row=r, column=f["col"] + 1, value=str(v))
            ws.freeze_panes = ws.cell(row=sheet["data_start"] + 1, column=1)
    wb.save(path_out)


def fill(path_in, folder, sheets, data):
    """Write the prefilled template: {"xlsx": path} and, for an .xml template, {"xml": path} in its own format.
    data: {sheet name: [{technical name: value}]}."""
    folder = Path(folder)
    out = {"xlsx": str(folder / "prefilled_template.xlsx")}
    if Path(path_in).suffix.lower() == ".xml":
        out["xml"] = str(folder / "prefilled_template.xml")
        _fill_xml(path_in, out["xml"], sheets, data)
        _xlsx_copy(path_in, out["xlsx"], sheets, data)
    else:
        _fill_xlsx(path_in, out["xlsx"], sheets, data)
    return out
