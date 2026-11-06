"""Headless UI check on a scratch database: sign-in, the projects page, a project with a prefilled object (from
test_pipeline.py's output), and the users page."""
import json
import shutil
from pathlib import Path

import store

ROOT = Path(__file__).parent
store.DATA = ROOT / "data" / "uitest"
store.DB = store.DATA / "app.db"
shutil.rmtree(store.DATA, ignore_errors=True)
store.add_user("admin@example.com", "secret-pass", "Admin", True)
pid = store.create_project("ACME wave 1", "public", "admin@example.com", ["Customer", "Supplier"])
folder = store.folder(pid, "Customer")
import agent, template_io  # noqa: E401,E402
tpl = folder / "template.xml"
shutil.copy(ROOT / "samples" / "Customer_migration_template_DEMO.xml", tpl)
src = folder / "source.xlsx"
shutil.copy(ROOT / "samples" / "Legacy_customers_DEMO.xlsx", src)
sheets = template_io.parse(tpl)
(folder / "template.json").write_text(json.dumps(sheets), encoding="utf-8")
mapping = json.loads((ROOT / "data" / "test" / "mapping.json").read_text(encoding="utf-8"))
sources = agent.read_source(src)
data, issues, removed, summary = agent.apply(sheets, sources, mapping)
files = template_io.fill(tpl, folder, sheets, data)
agent.write_report(folder / "analysis_report.xlsx", summary, mapping, sheets, issues, removed)
(folder / "result.json").write_text(json.dumps({"data": data, "issues": issues, "removed": removed, "mapping": mapping}))
store.update_object(pid, "Customer", template=str(tpl), source=str(src), status="Prefilled",
                    result={"summary": summary, "files": {**files, "report": str(folder / "analysis_report.xlsx")}})

from streamlit.testing.v1 import AppTest  # noqa: E402


def run(state=None):
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=60)  # the app's `store` is this patched module
    for k, v in (state or {}).items():
        at.session_state[k] = v
    return at.run()


user = store.check_login("admin@example.com", "secret-pass")
print("login ok:", bool(user), "| wrong password:", store.check_login("admin@example.com", "nope"))
for label, state in (("sign-in", {}), ("projects", {"user": user}), ("project", {"user": user, "project": pid,
                                                                               "open_obj": "Customer"}),
                     ("users", {"user": user, "page": "users"})):
    at = run(state)
    print(label, "| errors:", [e.value for e in at.exception] or "-", "| headers:",
          [m.value[:40] for m in at.markdown if m.value.startswith("#")][:5],
          "| buttons:", [b.label for b in at.button][:8], "| forms:", len(at.get("form")))
