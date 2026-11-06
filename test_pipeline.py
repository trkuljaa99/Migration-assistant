"""End to end on the demo files: parse the template, read the source, map (one real model call), apply, write the
prefilled template and the report, and check the .xml still parses with the data in it."""
import json
import sys
import time
from pathlib import Path

import agent
import template_io

S = Path(__file__).parent / "samples"
out = Path(__file__).parent / "data" / "test"
out.mkdir(parents=True, exist_ok=True)
sheets = template_io.parse(S / "Customer_migration_template_DEMO.xml")
for s in sheets:
    print(s["name"], "tech row", s["tech_row"], "data start", s["data_start"], "fields", len(s["fields"]),
          "mandatory", [f["tech"] for f in s["fields"] if f["mandatory"]])
sources = agent.read_source(S / "Legacy_customers_DEMO.xlsx")
for n, df in sources.items():
    print("source", n, df.shape, list(df.columns)[:6])
if "--no-model" in sys.argv:
    sys.exit()
t = time.time()
mapping = agent.propose_mapping("Customer", "SAP S/4HANA Cloud Public Edition", sheets, sources)
print(f"mapping in {time.time() - t:.0f}s")
(out / "mapping.json").write_text(json.dumps(mapping, indent=1, ensure_ascii=False), encoding="utf-8")
data, issues, removed, summary = agent.apply(sheets, sources, mapping)
files = template_io.fill(S / "Customer_migration_template_DEMO.xml", out, sheets, data)
agent.write_report(out / "analysis_report.xlsx", summary, mapping, sheets, issues, removed)
print(json.dumps(summary, indent=1, ensure_ascii=False)[:1500])
for g in agent.grouped_issues(issues, sheets)[:12]:
    print("  ", g)
print("notes:", *mapping["notes"], sep="\n  - ")
back = template_io.grid(files["xml"])
print("xml rows after fill:", {k: len(v) for k, v in back.items()}, "| first data row:", back["Basic Data"][7][:6])
print("first sales row:", back["Sales Data"][7])
