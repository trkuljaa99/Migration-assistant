"""Migration Assistant: prefills SAP S/4HANA Migration Cockpit templates from a customer's own data files.

Sign in -> projects (each for S/4HANA Cloud Public or Private Edition, with its migration objects) -> per object:
1. upload the migration template downloaded from the system, 2. upload the data to migrate (any Excel / CSV
layout), 3. the agent prefills the template and reports what's wrong or missing (agent.py) -> preview, download.

Run: streamlit run app.py (on the server under a hidden base path, see deploy/sap-migration.service).
"""
import html
import json
from pathlib import Path

import pandas as pd
import streamlit as st

import agent
import objects
import store
import template_io

LOGO = Path(__file__).parent / "assets" / "pwc-logo.svg"  # the PwC logo since 2025 (Wikimedia Commons)
# PwC colours: orange #FD5108 (the logo's), its darker and lighter tints, black text, neutral greys.
st.set_page_config(page_title="Migration Assistant · PwC", layout="wide")
st.markdown("""<style>
.block-container{padding-top:4.5rem;max-width:1200px}
h1,h2,h3,h4{color:#000}
.pill{display:inline-block;padding:2px 10px;border-radius:999px;font-size:.78rem;font-weight:600;white-space:nowrap}
.p0{background:#EEEEEE;color:#2D2D2D}.p1{background:#FFE3D6;color:#A33500}.p2{background:#FFC2A3;color:#7A2700}
.p3{background:#FD5108;color:#fff}
.tgt{display:inline-block;padding:2px 8px;border-radius:6px;font-size:.75rem;font-weight:600;background:#FFF1EA;
color:#A33500;border:1px solid #FFD2BD}
.muted{color:#6B6B6B;font-size:.85rem}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:8px 0 14px}
.kpi{border:1px solid #E6E6E6;border-left:4px solid #FD5108;border-radius:8px;padding:10px 12px}
.kpi b{display:block;font-size:1.35rem;color:#000}.kpi span{color:#6B6B6B;font-size:.8rem}
.ok{background:#FFF6F1;border:1px solid #FFD2BD;border-left:4px solid #FD5108;border-radius:8px;padding:10px 14px;
margin:6px 0 10px;color:#2D2D2D}
.brand{display:flex;align-items:center;gap:14px;margin-bottom:6px}.brand img{height:44px}
.brand span{font-size:1.6rem;font-weight:700;color:#000;border-left:1px solid #DEDEDE;padding-left:14px}
[data-testid="stProgress"] [role="progressbar"] > div > div > div{background-color:#FD5108}
</style>""", unsafe_allow_html=True)


def pill(status):
    i = store.STATUSES.index(status) if status in store.STATUSES else 0
    return f'<span class="pill p{i}">{html.escape(status)}</span>'


def ok(text_html):
    """A done step (an uploaded file) in the PwC palette."""
    st.markdown(f'<div class="ok">{text_html}</div>', unsafe_allow_html=True)


def kpis(items):
    st.markdown('<div class="kpis">' + "".join(f"<div class='kpi'><span>{html.escape(k)}</span><b>{v}</b></div>"
                                                for k, v in items) + "</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------- sign in

def sign_in():
    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        import base64
        logo = base64.b64encode(LOGO.read_bytes()).decode()
        st.markdown(f'<div class="brand"><img src="data:image/svg+xml;base64,{logo}" alt="PwC">'
                    "<span>Migration Assistant</span></div>", unsafe_allow_html=True)
        st.markdown('<p class="muted">Prefill SAP S/4HANA Migration Cockpit templates from your legacy data — '
                    "with an analysis of what's missing, wrong or duplicated.</p>", unsafe_allow_html=True)
        with st.form("signin"):
            email = st.text_input("Email")
            password = st.text_input("Password", type="password")
            if st.form_submit_button("Sign in", type="primary", width="stretch"):
                user = store.check_login(email, password)
                if user:
                    st.session_state.user = user
                    st.rerun()
                st.error("Wrong email or password.")


# ---------------------------------------------------------------- projects

@st.dialog("New project", width="large")
def new_project():
    step = st.session_state.setdefault("np_step", 1)
    if step == 1:
        name = st.text_input("Project name", key="np_name", placeholder="e.g. ACME – wave 1")
        target = st.radio("Target system", list(objects.TARGETS), key="np_target", format_func=objects.label,
                          captions=["Migration objects of the Public Edition (file / staging approach)",
                                    "Migration objects of the Private Edition / on-premise"])
        if st.button("Next: choose migration objects", type="primary", disabled=not (name or "").strip()):
            # widget values are dropped once their widgets aren't drawn: keep them under keys of our own
            st.session_state.np_step, st.session_state.np_v_name, st.session_state.np_v_target = 2, name.strip(), target
            st.rerun(scope="fragment")
        return
    target, name = st.session_state.np_v_target, st.session_state.np_v_name
    st.markdown(f"**{html.escape(name)}** · {objects.label(target)}")
    query = st.text_input("Filter", placeholder="Search migration objects", label_visibility="collapsed")
    chosen = st.session_state.setdefault("np_chosen", set())
    for area, names in objects.catalog(target).items():
        names = [n for n in names if query.lower() in n.lower()]
        if not names:
            continue
        c = st.columns([4, 1])
        c[0].markdown(f"**{area}**")
        if c[1].button("Select all", key=f"np-all-{area}"):
            chosen |= set(names)
            for n in names:
                st.session_state[f"np-{n}"] = True
        cols = st.columns(2)
        for i, n in enumerate(names):
            st.session_state.setdefault(f"np-{n}", n in chosen)
            if cols[i % 2].checkbox(n, key=f"np-{n}"):
                chosen.add(n)
            else:
                chosen.discard(n)
    c = st.columns([1, 1, 2])
    if c[0].button("Back"):
        st.session_state.np_step = 1
        st.rerun(scope="fragment")
    if c[1].button(f"Create project ({len(chosen)})", type="primary", disabled=not chosen):
        all_names = [n for names in objects.catalog(target).values() for n in names]
        pid = store.create_project(name, target, st.session_state.user["email"],
                                   [n for n in all_names if n in chosen])
        for k in [k for k in st.session_state if k.startswith("np")]:
            del st.session_state[k]
        st.session_state.project = pid
        st.rerun()


def projects_page():
    c = st.columns([4, 1])
    c[0].markdown("## Projects")
    if c[1].button("＋ New project", type="primary", width="stretch"):
        for k in [k for k in st.session_state if k.startswith("np")]:
            del st.session_state[k]
        new_project()
    rows = store.projects()
    if not rows:
        st.info("No projects yet. Create one to start.")
        return
    cols = st.columns(3)
    for i, p in enumerate(rows):
        done = sum(o["status"] == "Prefilled" for o in p["objects"])
        with cols[i % 3].container(border=True):
            st.markdown(f"#### {html.escape(p['name'])}")
            st.markdown(f'<span class="tgt">{objects.label(p["target"])}</span>', unsafe_allow_html=True)
            st.progress(done / len(p["objects"]) if p["objects"] else 0.0,
                        text=f"{done} of {len(p['objects'])} objects prefilled")
            st.markdown(f'<p class="muted">Created {p["created"][:10]} by {html.escape(p["owner"])}</p>',
                        unsafe_allow_html=True)
            if st.button("Open", key=f"open-{p['id']}", width="stretch"):
                st.session_state.project = p["id"]
                st.rerun()


# ---------------------------------------------------------------- one project

def _json(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def template_step(pid, obj, folder):
    st.markdown("**1 · Migration template**")
    st.caption("Download it in the system (Migrate Your Data → your project → the object → Download Template) "
               "and upload it here as it is (.xml or .xlsx).")
    sheets = _json(folder / "template.json")
    if sheets:
        ok(f"<b>{html.escape(store.display_name(obj['template']))}</b><br>" + " · ".join(
            f"{html.escape(s['name'])} ({len(s['fields'])} fields, {sum(f['mandatory'] for f in s['fields'])} mandatory)"
            for s in sheets))
    up = st.file_uploader("Template", type=["xml", "xlsx"], key=f"tpl-{obj['name']}", label_visibility="collapsed")
    if up and st.session_state.get(f"tpl-done-{obj['name']}") != up.file_id:
        path = store.save_upload(pid, obj["name"], "template", up.name, up.getvalue())
        try:
            parsed = template_io.parse(path)
        except Exception as err:
            st.error(f"Couldn't read the template: {err}")
            return None
        if not parsed:
            st.error("No data sheet with technical field names found. Is this a Migration Cockpit template?")
            return None
        (folder / "template.json").write_text(json.dumps(parsed), encoding="utf-8")
        status = "Template provided" if not obj["source"] else "Source data uploaded"
        store.update_object(pid, obj["name"], template=path, status=status, result=None)
        st.session_state[f"tpl-done-{obj['name']}"] = up.file_id
        st.session_state.open_obj = obj["name"]  # the object stays open after the reload
        st.rerun()
    return sheets


def source_step(pid, obj, ready):
    st.markdown("**2 · Data to migrate**")
    st.caption("Your legacy extract in any layout (.xlsx, .xls or .csv) – columns don't need to match the template.")
    if obj["source"]:
        ok(f"<b>{html.escape(store.display_name(obj['source']))}</b>")
    up = st.file_uploader("Source data", type=["xlsx", "xls", "csv"], key=f"src-{obj['name']}", disabled=not ready,
                          label_visibility="collapsed")
    if up and st.session_state.get(f"src-done-{obj['name']}") != up.file_id:
        path = store.save_upload(pid, obj["name"], "source", up.name, up.getvalue())
        try:
            sources = agent.read_source(path)
        except Exception as err:
            st.error(f"Couldn't read the file: {err}")
            return
        if not sources:
            st.error("The file has no data.")
            return
        store.update_object(pid, obj["name"], source=path, status="Source data uploaded", result=None)
        st.session_state[f"src-done-{obj['name']}"] = up.file_id
        st.session_state.open_obj = obj["name"]
        st.rerun()
    if obj["source"]:
        with st.popover("Preview the source data"):
            for name, df in agent.read_source(obj["source"]).items():
                st.caption(f"{name}: {len(df)} rows, {df.shape[1]} columns")
                st.dataframe(df.head(50), hide_index=True)


def run_agent(pid, obj, project, sheets, folder):
    sources = agent.read_source(obj["source"])
    with st.status("The agent is working…", expanded=True) as status:
        st.write(f"Read {sum(len(d) for d in sources.values())} rows from {len(sources)} sheet(s).")
        st.write("Mapping the source to the template's fields (this takes a minute or two)…")
        mapping = agent.propose_mapping(obj["name"], objects.label(project["target"]), sheets, sources)
        st.write("Applying the mapping to every record, removing duplicates, checking mandatory fields…")
        _finish(pid, obj, sheets, folder, sources, mapping)
        status.update(label="Done", state="complete", expanded=False)


def _finish(pid, obj, sheets, folder, sources, mapping):
    """Apply a mapping to the source and write the prefilled template, the report and the result."""
    data, issues, removed, summary = agent.apply(sheets, sources, mapping)
    files = template_io.fill(obj["template"], folder, sheets, data)
    report = folder / "analysis_report.xlsx"
    agent.write_report(report, summary, mapping, sheets, issues, removed)
    (folder / "result.json").write_text(json.dumps({"data": data, "issues": issues, "removed": removed,
                                                    "mapping": mapping}), encoding="utf-8")
    store.update_object(pid, obj["name"], status="Prefilled",
                        result={"summary": summary, "files": {**files, "report": str(report)}})


def refill(obj, sheets, folder, full, values):
    """Constant values for mandatory fields the source lacks: the saved mapping again, no new model run."""
    if not values:
        return
    mapping = full["mapping"]
    for sm in mapping["sheets"]:
        have = {m["field"] for m in sm["fields"]}
        for (name, field), v in values.items():
            if sm["sheet"] != name:
                continue
            m = {"field": field, "source": "constant", "columns": [], "separator": "", "constant": v, "regex": "",
                 "transform": "none", "date_format": "", "value_map": [], "note": "Set by the user for every record"}
            sm["fields"] = [m if x["field"] == field else x for x in sm["fields"]] if field in have \
                else sm["fields"] + [m]
    _finish(st.session_state.project, obj, sheets, folder, agent.read_source(obj["source"]), mapping)
    st.session_state.open_obj = obj["name"]
    st.rerun()


def results(obj, sheets, folder):
    res, full = obj["result"], _json(folder / "result.json")
    if not res or not full:
        return
    s = res["summary"]
    kpis([("Records prefilled", s["rows"]), ("Duplicates removed", s["duplicates_removed"]),
          ("Missing mandatory values", s["missing_mandatory_cells"]), ("Issues", s["issues"])])
    missing = [(name, f) for name, x in s["sheets"].items() for f in x["mandatory_not_in_source"]]
    if missing:
        labels = {(sh["name"], f["tech"]): f["label"] for sh in sheets for f in sh["fields"]}
        with st.form(f"fill-{obj['name']}"):
            st.warning("Mandatory fields the source doesn't have. Values that are the same for every record (e.g. "
                       "sales organization, division, BP grouping) can be set here – no new AI run needed.")
            cols, values = st.columns(3), {}
            for i, (name, f) in enumerate(missing):
                values[(name, f)] = cols[i % 3].text_input(f"{labels.get((name, f), f)} ({f}) · {name}",
                                                           key=f"fill-{obj['name']}-{name}-{f}")
            if st.form_submit_button("Apply to every record", type="primary"):
                refill(obj, sheets, folder, full, {k: v.strip() for k, v in values.items() if v.strip()})
    tabs = st.tabs(["Preview", "Issues", "Mapping", "Duplicates removed", "Agent notes"])
    by_sheet = {sh["name"]: sh for sh in sheets}
    with tabs[0]:
        names = [n for n in full["data"] if n in by_sheet]
        if names:
            name = st.selectbox("Sheet", names, key=f"pv-{obj['name']}") if len(names) > 1 else names[0]
            sheet, rows = by_sheet[name], full["data"][name]
            missing = {(x["row"], x["field"]) for x in full["issues"] if x["sheet"] == name}
            cols = [f["tech"] for f in sheet["fields"] if any(r.get(f["tech"]) for r in rows) or f["mandatory"]]
            labels = {f["tech"]: f"{'*' if f['mandatory'] else ''}{f['label']} ({f['tech']})" for f in sheet["fields"]}
            df = pd.DataFrame([{labels[c]: r.get(c, "") for c in cols} for r in rows[:500]])
            if not df.empty:
                marks = pd.DataFrame("", index=df.index, columns=df.columns)
                for r, f in missing:
                    if r - 1 < len(df) and f in cols:
                        marks.iloc[r - 1, cols.index(f)] = "background-color:#fde8e8"
                st.dataframe(df.style.apply(lambda _: marks, axis=None), hide_index=True)
                st.caption(f"{len(rows)} records · * = mandatory · red = an issue (see Issues)"
                           + (" · first 500 shown" if len(rows) > 500 else ""))
    with tabs[1]:
        g = agent.grouped_issues(full["issues"], sheets)
        if g:
            st.dataframe(pd.DataFrame(g), hide_index=True)
        else:
            st.success("No issues found.")
    with tabs[2]:
        st.dataframe(pd.DataFrame(agent.mapping_rows(full["mapping"], sheets)), hide_index=True)
    with tabs[3]:
        if full["removed"]:
            st.dataframe(pd.DataFrame(full["removed"]), hide_index=True)
        else:
            st.caption("None.")
    with tabs[4]:
        for n in full["mapping"].get("notes", []):
            st.markdown(f"- {n}")
    c = st.columns(3)
    files = res["files"]
    if files.get("xml"):
        c[0].download_button("Prefilled template (.xml, for upload)", Path(files["xml"]).read_bytes(),
                             file_name=f"{obj['name']} - prefilled.xml", key=f"dl-xml-{obj['name']}", type="primary")
    c[1 if files.get("xml") else 0].download_button(
        "Prefilled template (.xlsx)", Path(files["xlsx"]).read_bytes(), file_name=f"{obj['name']} - prefilled.xlsx",
        key=f"dl-xlsx-{obj['name']}", type="secondary" if files.get("xml") else "primary")
    c[2].download_button("Analysis report (.xlsx)", Path(files["report"]).read_bytes(),
                         file_name=f"{obj['name']} - analysis.xlsx", key=f"dl-rep-{obj['name']}")


def project_page(pid):
    project = store.project(pid)
    if not project:
        st.session_state.pop("project", None)
        st.rerun()
    if st.button("← All projects"):
        st.session_state.pop("project", None)
        st.rerun()
    rows = store.objects(pid)
    done = sum(o["status"] == "Prefilled" for o in rows)
    st.markdown(f"## {html.escape(project['name'])}")
    st.markdown(f'<span class="tgt">{objects.label(project["target"])}</span> '
                f'<span class="muted">· {len(rows)} migration objects · {done} prefilled</span>',
                unsafe_allow_html=True)
    st.progress(done / len(rows) if rows else 0.0)
    with st.expander("Add migration objects"):
        have = {o["name"] for o in rows}
        more = [n for names in objects.catalog(project["target"]).values() for n in names if n not in have]
        pick = st.multiselect("Objects", more, label_visibility="collapsed")
        if st.button("Add", disabled=not pick):
            store.add_objects(pid, pick)
            st.rerun()
    st.markdown(" ")
    for obj in rows:
        folder = store.folder(pid, obj["name"])
        with st.expander(f"{obj['name']}  ·  {obj['status']}", expanded=st.session_state.get("open_obj") == obj["name"]):
            st.markdown(pill(obj["status"]), unsafe_allow_html=True)
            c = st.columns(2)
            with c[0]:
                sheets = template_step(pid, obj, folder)
            with c[1]:
                source_step(pid, obj, bool(sheets))
            st.markdown("**3 · Prefill**")
            ready = bool(sheets and obj["source"])
            label = "Run the migration agent again" if obj["status"] == "Prefilled" else "Run the migration agent"
            if st.button(label, key=f"run-{obj['name']}", type="primary", disabled=not ready):
                st.session_state.open_obj = obj["name"]
                try:
                    run_agent(pid, obj, project, sheets, folder)
                except Exception as err:
                    st.error(f"The agent failed: {err}")
                else:
                    st.rerun()
            if not ready:
                st.caption("Upload the template and the data first.")
            if obj["status"] == "Prefilled":
                results(obj, sheets, folder)
    if st.session_state.user.get("is_admin"):
        with st.expander("Delete this project"):
            if st.button("Delete project and all its files", type="secondary"):
                store.delete_project(pid)
                st.session_state.pop("project", None)
                st.rerun()


# ---------------------------------------------------------------- users (admin)

def users_page():
    st.markdown("## Users")
    st.dataframe(pd.DataFrame([{"Email": u["email"], "Name": u["name"], "Admin": "yes" if u["is_admin"] else "",
                                "Created": u["created"][:10]} for u in store.users()]), hide_index=True)
    with st.form("add-user", clear_on_submit=True):
        st.markdown("**Add a user**")
        c = st.columns(3)
        email = c[0].text_input("Email")
        name = c[1].text_input("Name")
        pw = c[2].text_input("Password", type="password")
        admin = st.checkbox("Admin")
        if st.form_submit_button("Add user", type="primary"):
            if "@" not in email or len(pw) < 8:
                st.error("Enter an email and a password of at least 8 characters.")
            else:
                store.add_user(email, pw, name, admin)
                st.success(f"Added {email}.")
    others = [u["email"] for u in store.users() if u["email"] != st.session_state.user["email"]]
    if others:
        with st.expander("Remove a user"):
            who = st.selectbox("User", others)
            if st.button("Remove"):
                store.delete_user(who)
                st.rerun()


# ---------------------------------------------------------------- main

def main():
    user = st.session_state.get("user")
    if not user:
        sign_in()  # (with the logo next to the title)
        return
    st.logo(str(LOGO), size="large")
    with st.sidebar:
        st.markdown("### Migration Assistant")
        st.caption(user["email"])
        if st.button("Projects", width="stretch"):
            st.session_state.pop("project", None)
            st.session_state.page = "projects"
            st.rerun()
        if user.get("is_admin") and st.button("Users", width="stretch"):
            st.session_state.page = "users"
            st.rerun()
        if st.button("Sign out", width="stretch"):
            st.session_state.clear()
            st.rerun()
        samples = sorted((Path(__file__).parent / "samples").glob("*DEMO*"))
        if samples:
            with st.expander("Demo files"):
                st.caption("A made-up Customer template and a messy legacy extract, to try the flow.")
                for p in samples:
                    st.download_button(p.name, p.read_bytes(), file_name=p.name, key=f"demo-{p.name}", width="stretch")
    if st.session_state.get("page") == "users" and user.get("is_admin"):
        users_page()
    elif st.session_state.get("project"):
        project_page(st.session_state.project)
    else:
        projects_page()


main()
