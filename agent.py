"""The migration agent: prefills a Migration Cockpit template from the user's own data file.

1. The source file (Excel or CSV, any layout) is profiled: per sheet its columns, how full they are, and sample
   values (every distinct value for short code lists). Only this profile goes to the model, not the whole file.
2. The LLM (by default Claude, claude-opus-5-5, structured output; see _settings to switch) maps the template's fields to the source: which column (or columns,
   or a constant) fills each field, which transformation it needs (dates to YYYYMMDD, numbers, upper case, a value
   map such as country names to ISO codes, a regex part), which fields identify a record, and notes on what's
   missing.
3. The mapping is applied to every row here, in code, so the result is repeatable and auditable: empty rows
   dropped, duplicates (same key) removed, values transformed, then each row checked: mandatory fields filled,
   lengths, dates and numbers valid, mapped values known.
4. The result: the prefilled template (in its own format and as .xlsx), and a report of every issue.
"""
import json
import math
import re
from datetime import date, datetime
from pathlib import Path

import anthropic
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

MODEL = "claude-opus-5-5"  # the default; LLM_MODEL in the secrets overrides it (see _settings)
SAMPLE_VALUES = 8
FULL_LIST_MAX = 60  # a column with at most this many distinct values is sent in full (for value maps)
TRANSFORMS = ["none", "upper", "date", "number", "map", "pad_zeros"]

FIELD = {
    "type": "object",
    "properties": {
        "field": {"type": "string", "description": "technical field name of the template"},
        "source": {"type": "string", "enum": ["column", "columns_concat", "constant", "none"]},
        "columns": {"type": "array", "items": {"type": "string"},
                    "description": "source column(s); one for 'column', several for 'columns_concat'"},
        "separator": {"type": "string", "description": "for columns_concat"},
        "constant": {"type": "string", "description": "for 'constant'"},
        "regex": {"type": "string", "description": "optional: keep only this part of the value (group 1, else "
                                                   "the whole match); empty for none"},
        "transform": {"type": "string", "enum": TRANSFORMS},
        "date_format": {"type": "string", "description": "strptime format of the source dates for 'date', e.g. "
                                                         "%d.%m.%Y; empty to detect"},
        "value_map": {"type": "array", "items": {
            "type": "object", "properties": {"source_value": {"type": "string"}, "sap_value": {"type": "string"}},
            "required": ["source_value", "sap_value"], "additionalProperties": False}},
        "note": {"type": "string", "description": "short reason or caveat, empty if obvious"},
    },
    "required": ["field", "source", "columns", "separator", "constant", "regex", "transform", "date_format",
                 "value_map", "note"],
    "additionalProperties": False,
}
SCHEMA = {
    "type": "object",
    "properties": {
        "sheets": {"type": "array", "items": {
            "type": "object",
            "properties": {
                "sheet": {"type": "string", "description": "template sheet name"},
                "source_sheet": {"type": "string", "description": "source sheet that fills it; empty if none"},
                "key_fields": {"type": "array", "items": {"type": "string"},
                               "description": "technical fields that identify one record on this sheet"},
                "fields": {"type": "array", "items": FIELD},
            },
            "required": ["sheet", "source_sheet", "key_fields", "fields"], "additionalProperties": False}},
        "notes": {"type": "array", "items": {"type": "string"},
                  "description": "what the consultant should know: data that's missing, assumptions, risks"},
    },
    "required": ["sheets", "notes"],
    "additionalProperties": False,
}

INSTRUCTIONS = """You are an SAP S/4HANA data migration consultant. You map a customer's legacy data file onto an
SAP S/4HANA Migration Cockpit template (file approach) for one migration object.

You get the template's data sheets with their fields (technical name, description, group, whether mandatory, type,
length) and a profile of the source file (its sheets, columns, how full they are, sample values; every distinct
value for short code lists).

Return the mapping:
- For each template sheet that the source can fill: the source sheet, the key fields that identify one record on
  that sheet (used to remove duplicates), and the field mappings.
- List every mandatory field (source "none" when the source doesn't have it - don't invent data) and every
  optional field the source can fill. Leave out optional fields the source has nothing for.
- Use the source's own meaning, not just column names: e.g. a full address column may need a regex for the street,
  "Germany" needs the ISO code DE, a free-text currency "Euro" needs EUR, dates must become YYYYMMDD
  (transform "date", give the source format when it's ambiguous), decimal commas need transform "number".
- Code fields (country, region, currency, language, unit of measure, payment terms, industry, ...) take SAP
  codes: give a value_map from every distinct source value you were shown to the SAP value. Use standard SAP / ISO
  codes where they exist; for customer-specific codes (e.g. payment terms, account groups, sales organisations)
  map only what is clear and say in a note that the customer's configuration decides. A constant is right for
  values that are the same for every record (e.g. a company code or sales organisation given by the project), but
  say so in the note - never guess customer-specific organisational values without saying it.
- Keep notes short and concrete: what's missing for the load, which assumptions you made, data quality risks
  (e.g. mixed date formats, several values in one cell, likely duplicates)."""


# ---------------------------------------------------------------- the source file

def _s(v):
    """A cell as clean text: empty for blanks, dates as YYYY-MM-DD, whole floats without .0."""
    if v is None or (isinstance(v, float) and math.isnan(v)) or (hasattr(pd, "isna") and v is pd.NaT):
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m-%d") if not (v.hour or v.minute or v.second) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_source(path):
    """{sheet: DataFrame of text} from .xlsx / .xls / .csv; the header row is found as the first row that is mostly
    filled with distinct text."""
    path = Path(path)
    if path.suffix.lower() in (".csv", ".txt"):
        raw = {"Data": pd.read_csv(path, header=None, dtype=object, sep=None, engine="python", encoding_errors="replace")}
    else:
        raw = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
    out = {}
    for name, df in raw.items():
        df = df.map(_s) if hasattr(df, "map") else df.applymap(_s)
        df = df.loc[:, (df != "").any(axis=0)]
        df = df[(df != "").any(axis=1)]
        if df.empty:
            continue
        head = 0
        for i in range(min(15, len(df))):
            row = [v for v in df.iloc[i] if v]
            if len(row) >= max(2, 0.6 * df.shape[1]) and len(set(row)) == len(row):
                head = i
                break
        cols, seen = [], {}
        for c in df.iloc[head]:
            c = c or "Column"
            seen[c] = seen.get(c, 0) + 1
            cols.append(c if seen[c] == 1 else f"{c} ({seen[c]})")
        body = df.iloc[head + 1:].reset_index(drop=True)
        body.columns = cols
        out[str(name)] = body
    return out


def profile(sources):
    """The source as the model sees it: per sheet, its columns with fill rate and sample (or all) values."""
    lines = []
    for name, df in sources.items():
        lines.append(f'SOURCE SHEET "{name}": {len(df)} rows')
        for col in df.columns:
            vals = df[col][df[col] != ""]
            distinct = vals.unique()
            shown = distinct if len(distinct) <= FULL_LIST_MAX else distinct[:SAMPLE_VALUES]
            label = "all values" if len(distinct) <= FULL_LIST_MAX else "samples"
            lines.append(f'  - "{col}": {len(vals)}/{len(df)} filled, {len(distinct)} distinct; {label}: '
                         + json.dumps([str(v)[:60] for v in shown], ensure_ascii=False))
    return "\n".join(lines)


def template_text(sheets):
    lines = []
    for s in sheets:
        lines.append(f'TEMPLATE SHEET "{s["name"]}"')
        for f in s["fields"]:
            extra = ", ".join(x for x in (f["importance"] or ("mandatory" if f["mandatory"] else "optional"),
                                          f["type"], f"length {f['length']}" if f["length"] else "") if x)
            lines.append(f'  - {f["tech"]}: {f["label"]} [{f["group"]}] ({extra})')
    return "\n".join(lines)


# ---------------------------------------------------------------- the model's mapping

def _settings():
    """The LLM settings from .streamlit/secrets.toml (else the environment):
    ANTHROPIC_API_KEY, LLM_MODEL (default MODEL) and LLM_BASE_URL (an Anthropic-compatible endpoint, e.g. a company
    gateway; empty = Anthropic's own API). Switching to another LLM is done here and in propose_mapping, the only
    model call: it must return JSON matching SCHEMA."""
    import os

    import streamlit as st
    try:
        secrets = dict(st.secrets)
    except Exception:
        secrets = {}
    get = lambda k: secrets.get(k) or os.environ.get(k) or None
    return {"api_key": get("ANTHROPIC_API_KEY"), "model": get("LLM_MODEL") or MODEL, "base_url": get("LLM_BASE_URL")}


def _client(s):
    kwargs = {k: v for k, v in (("api_key", s["api_key"]), ("base_url", s["base_url"])) if v}
    return anthropic.Anthropic(**kwargs)


def propose_mapping(object_name, target, sheets, sources):
    """The model's mapping (SCHEMA) for one object."""
    s = _settings()
    prompt = (f"Migration object: {object_name}\nTarget: {target}\n\n{template_text(sheets)}\n\n{profile(sources)}")
    with _client(s).beta.messages.stream(
        model=s["model"],
        max_tokens=64000,
        output_config={"effort": "high", "format": {"type": "json_schema", "schema": SCHEMA}},
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        system=[{"type": "text", "text": INSTRUCTIONS, "cache_control": {"type": "ephemeral"}}],
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        final = stream.get_final_message()
    if final.stop_reason == "refusal":
        raise RuntimeError("The model declined this request.")
    if final.stop_reason == "max_tokens":
        raise RuntimeError("The mapping was cut off at the length limit (a very large template?).")
    return json.loads("".join(b.text for b in final.content if b.type == "text"))


# ---------------------------------------------------------------- applying it

def _date(v, fmt):
    if not v:
        return "", None
    for f in ([fmt] if fmt else []) + ["%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y",
                                       "%Y%m%d", "%d-%m-%Y", "%d.%m.%y", "%Y/%m/%d"]:
        try:
            return datetime.strptime(v, f).strftime("%Y%m%d"), None
        except ValueError:
            continue
    return v, "Not a valid date"


def _number(v):
    if not v:
        return "", None
    t = v.replace(" ", "").replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", t) or re.fullmatch(r"-?\d+,\d+", t):
        t = t.replace(".", "").replace(",", ".")  # 1.234,56 / 12,5
    else:
        t = t.replace(",", "")
    try:
        float(t)
        return t, None
    except ValueError:
        return v, "Not a valid number"


def _value(m, row):
    src = m["source"]
    if src == "constant":
        v = m["constant"]
    elif src == "column":
        v = row.get(m["columns"][0], "") if m["columns"] else ""
    elif src == "columns_concat":
        v = (m["separator"] or " ").join(row.get(c, "") for c in m["columns"] if row.get(c, ""))
    else:
        return "", None
    v = str(v).strip()
    if v and m["regex"]:
        found = re.search(m["regex"], v)
        v = (found.group(1) if found and found.groups() else found.group(0) if found else "").strip()
    t = m["transform"]
    if t == "upper":
        return v.upper(), None
    if t == "date":
        return _date(v, m["date_format"])
    if t == "number":
        return _number(v)
    if t == "map":
        table = {x["source_value"].strip().lower(): x["sap_value"] for x in m["value_map"]}
        if not v:
            return "", None
        if v.lower() in table:
            return table[v.lower()], None
        return v, "No SAP value for this source value"
    if t == "pad_zeros" and v.isdigit() and m.get("_length"):
        return v.zfill(m["_length"]), None
    return v, None


def apply(sheets, sources, mapping):
    """The prefilled rows per template sheet, the issues and the summary (see the module docstring)."""
    by_sheet = {s["name"]: s for s in sheets}
    data, issues, removed, summary = {}, [], [], {"sheets": {}}
    for sm in mapping["sheets"]:
        sheet = by_sheet.get(sm["sheet"])
        df = sources.get(sm["source_sheet"])
        if not sheet or df is None:
            continue
        fields = {f["tech"]: f for f in sheet["fields"]}
        maps = [dict(m, _length=(fields.get(m["field"]) or {}).get("length")) for m in sm["fields"]
                if m["field"] in fields]
        mapped = {m["field"] for m in maps if m["source"] != "none"}
        keys = [k for k in sm["key_fields"] if k in fields]
        rows, seen, cell_issues = [], {}, []
        for i, rec in enumerate(df.to_dict("records"), 1):
            out, row_issues = {}, []
            for m in maps:
                v, problem = _value(m, rec)
                out[m["field"]] = v
                if problem and v:
                    row_issues.append((m["field"], problem, v))
            if not any(out.values()):
                continue
            key = tuple(out.get(k, "") for k in keys) if keys and any(out.get(k) for k in keys) else \
                tuple(sorted(out.items()))
            if key in seen:
                removed.append({"Sheet": sm["sheet"], "Source record": i, "Same as source record": seen[key],
                                "Key": ", ".join(f"{k}={out.get(k, '')}" for k in keys) or "whole record"})
                continue
            seen[key] = i
            rows.append(out)
            n = len(rows)
            for f, problem, v in row_issues:
                cell_issues.append({"sheet": sm["sheet"], "row": n, "field": f, "issue": problem, "value": v})
            for f in sheet["fields"]:
                v = out.get(f["tech"], "")
                if f["mandatory"] and not v:
                    cell_issues.append({"sheet": sm["sheet"], "row": n, "field": f["tech"],
                                        "issue": "Mandatory value missing", "value": ""})
                elif v and f["length"] and len(v) > f["length"]:
                    cell_issues.append({"sheet": sm["sheet"], "row": n, "field": f["tech"],
                                        "issue": f"Longer than {f['length']} characters", "value": v})
        data[sm["sheet"]] = rows
        issues += cell_issues
        missing = [f["tech"] for f in sheet["fields"] if f["mandatory"] and f["tech"] not in mapped]
        summary["sheets"][sm["sheet"]] = {
            "source_sheet": sm["source_sheet"], "source_rows": len(df), "rows": len(rows),
            "duplicates_removed": sum(r["Sheet"] == sm["sheet"] for r in removed),
            "fields_filled": len(mapped), "fields": len(sheet["fields"]),
            "mandatory": sum(f["mandatory"] for f in sheet["fields"]), "mandatory_not_in_source": missing,
            "key_fields": keys}
    summary["issues"] = len(issues)
    summary["missing_mandatory_cells"] = sum(x["issue"] == "Mandatory value missing" for x in issues)
    summary["duplicates_removed"] = len(removed)
    summary["rows"] = sum(len(r) for r in data.values())
    return data, issues, removed, summary


def grouped_issues(issues, sheets):
    """Issues per sheet, field and kind: how many rows, and the first rows / values."""
    labels = {(s["name"], f["tech"]): f["label"] for s in sheets for f in s["fields"]}
    groups = {}
    for x in issues:
        g = groups.setdefault((x["sheet"], x["field"], x["issue"]), {"rows": [], "values": []})
        g["rows"].append(x["row"])
        if x["value"] and x["value"] not in g["values"]:
            g["values"].append(x["value"])
    return [{"Sheet": s, "Field": f, "Description": labels.get((s, f), ""), "Issue": kind, "Rows": len(g["rows"]),
             "First rows": ", ".join(map(str, g["rows"][:10])) + (" …" if len(g["rows"]) > 10 else ""),
             "Values": ", ".join(g["values"][:6])}
            for (s, f, kind), g in sorted(groups.items(), key=lambda kv: (-len(kv[1]["rows"]), kv[0]))]


def mapping_rows(mapping, sheets):
    labels = {(s["name"], f["tech"]): f for s in sheets for f in s["fields"]}
    out = []
    for sm in mapping["sheets"]:
        for m in sm["fields"]:
            f = labels.get((sm["sheet"], m["field"]), {})
            source = {"column": ", ".join(m["columns"][:1]), "columns_concat": f" + ".join(m["columns"]),
                      "constant": f'constant "{m["constant"]}"', "none": "– not in the source –"}[m["source"]]
            how = m["transform"] + (f" ({len(m['value_map'])} values)" if m["transform"] == "map" else "") \
                + (f", regex {m['regex']}" if m["regex"] else "")
            out.append({"Sheet": sm["sheet"], "Field": m["field"], "Description": f.get("label", ""),
                        "Mandatory": "yes" if f.get("mandatory") else "", "Source": source,
                        "Transformation": "" if how == "none" else how, "Note": m["note"]})
    return out


def write_report(path, summary, mapping, sheets, issues, removed):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    bold, head = Font(bold=True), PatternFill("solid", fgColor="DCE6F1")
    ws.append(["Migration data analysis"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(["Records prefilled", summary["rows"]])
    ws.append(["Duplicates removed", summary["duplicates_removed"]])
    ws.append(["Missing mandatory values", summary["missing_mandatory_cells"]])
    ws.append(["Issues in total", summary["issues"]])
    ws.append([])
    ws.append(["Sheet", "Source sheet", "Source rows", "Records", "Duplicates removed", "Fields filled",
               "Mandatory fields not in the source"])
    for c in ws[ws.max_row]:
        c.font, c.fill = bold, head
    for name, s in summary["sheets"].items():
        ws.append([name, s["source_sheet"], s["source_rows"], s["rows"], s["duplicates_removed"],
                   f"{s['fields_filled']} of {s['fields']}", ", ".join(s["mandatory_not_in_source"])])
    ws.append([])
    ws.append(["Notes"])
    ws[f"A{ws.max_row}"].font = bold
    for n in mapping.get("notes", []):
        ws.append([n])
    for title, rows in (("Issues", grouped_issues(issues, sheets)), ("Mapping", mapping_rows(mapping, sheets)),
                        ("Duplicates removed", removed),
                        ("All issues", [{"Sheet": x["sheet"], "Record": x["row"], "Field": x["field"],
                                         "Issue": x["issue"], "Value": x["value"]} for x in issues])):
        sh = wb.create_sheet(title)
        if not rows:
            sh.append(["None"])
            continue
        sh.append(list(rows[0]))
        for c in sh[1]:
            c.font, c.fill = bold, head
        for r in rows:
            sh.append([str(v) for v in r.values()])
        sh.freeze_panes = "A2"
    for sh in wb.worksheets:
        for col in sh.columns:
            sh.column_dimensions[col[0].column_letter].width = min(60, max(10, *(len(str(c.value or "")) for c in col)))
    wb.save(path)
