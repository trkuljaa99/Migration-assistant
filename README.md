# Migration Assistant

Prefills SAP S/4HANA Migration Cockpit templates (file approach) from a customer's own legacy data, with an
analysis of what's missing, wrong or duplicated.

## Flow

1. **Sign in** – users are kept in a local SQLite database (scrypt-hashed passwords); admins manage users in the app.
2. **Projects** – each project targets **SAP S/4HANA Cloud Public Edition** or **Private Edition** and holds the
   migration objects chosen from that edition's list (`objects.py`).
3. **Per migration object**
   1. Upload the migration template downloaded from the system (XML Spreadsheet 2003 `.xml` or `.xlsx`) →
      *Template provided*. The data sheets, technical field names, mandatory fields, types and lengths are detected
      from the content (`template_io.py`).
   2. Upload the legacy data in any layout (`.xlsx`, `.xls`, `.csv`) → *Source data uploaded*.
   3. **Run the migration agent** (`agent.py`) → *Prefilled*:
      - a profile of the source (columns, fill rate, sample values; full value lists for short code lists) goes to
        the LLM, which returns a field mapping as structured JSON: source column(s) or constant per field,
        transformations (dates → `YYYYMMDD`, numbers, upper case, value maps such as country names → ISO codes,
        regex parts), key fields, and notes;
      - the mapping is applied to every record in code: empty rows dropped, duplicates (same key) removed, values
        transformed, every record checked (mandatory fields, lengths, dates, numbers, unmapped code values);
      - mandatory fields the source doesn't have (sales organisation, division, BP grouping …) can be set once for
        every record in the app, without a new LLM run.
4. **Preview and download** – the prefilled template in its own format (`.xml` for the upload) and as `.xlsx`,
   plus an analysis report (`.xlsx`: summary, issues, mapping, removed duplicates).

## Run locally

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt          # Windows: .venv\Scripts\pip
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # add the API key
echo 'a-strong-password' | .venv/bin/python create_user.py you@example.com "Your Name" --admin
.venv/bin/streamlit run app.py
```

`make_samples.py` writes the demo files in `samples/` (a made-up Customer template and a messy legacy extract);
`test_pipeline.py` runs the whole agent on them (one LLM call), `test_ui.py` renders the pages headlessly.

## LLM

The only model call is `agent.propose_mapping`; settings come from `.streamlit/secrets.toml`:

| Setting | Meaning |
|---|---|
| `ANTHROPIC_API_KEY` | API key |
| `LLM_MODEL` | model name (default `claude-opus-5-5`) |
| `LLM_BASE_URL` | an Anthropic-compatible endpoint, e.g. an internal LLM gateway (empty = Anthropic's API) |

To move to an LLM with a different API, replace `_client` / `propose_mapping` in `agent.py`: the function must
return JSON that matches `agent.SCHEMA` (the rest of the pipeline is plain Python). Only the source *profile*
(column names, counts, sample values) is sent to the model, never the whole file.

## Deploy on Streamlit Community Cloud

1. On [share.streamlit.io](https://share.streamlit.io) sign in with GitHub → **Create app** → *Deploy a public app
   from GitHub*: repository `trkuljaa99/Migration-assistant`, branch `main`, main file `app.py`.
2. **Advanced settings**: Python 3.12, and paste the secrets (see `.streamlit/secrets.toml.example`):
   `ANTHROPIC_API_KEY`, `ADMIN_EMAIL`, `ADMIN_PASSWORD`, `ADMIN_NAME`.
3. **Deploy**. The admin is created from the secrets on start. Community Cloud wipes the disk when the app
   restarts (after a redeploy or a long sleep), so projects, uploads and extra users are temporary there.

## Deploy (Linux, systemd)

Copy the folder to `/opt/sap-migration`, create the venv and `secrets.toml`, then install
`deploy/sap-migration.service.example` as `/etc/systemd/system/sap-migration.service` with `BASE_PATH` set, and
route that path to `127.0.0.1:8502` in the reverse proxy / tunnel.

## Note

The migration object lists follow SAP's object names but depend on the release and scope items – check them
against the "Migration Objects" list of your release. The demo template in `samples/` is not an SAP file.
