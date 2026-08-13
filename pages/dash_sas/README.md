# dash_sas — SAS NC Dashboard

Web dashboard for KDS SAS Emmen nonconformances. Upload the SAP "NC's Overview"
export, get 19 charts across 8 sections. Served at **http://chbs4212/sas**.

Part of the `pages/` product. Uses the shared `~/bgtools/dash/quality.db` (own
tables `sas_nc`, `sas_import`, `feedback`; never touches dash's tables).

## Files
```
pages/dash_sas/
  parse.py          export -> clean DataFrame (batch/status/class/cost rules)
  build_sas.py      creates sas tables in quality.db (WAL)
  ingest_sas.py     export -> sas_nc   [--rebuild default | --append]
  charts.py         builds the 19-block payload from a filtered DataFrame
  app/
    main.py         FastAPI  (/sas)
    page.html       the dashboard UI (dark Beyond Gravity theme)
    plotly.min.js   vendored, offline
  requirements.txt
  dash_sas.service  systemd unit (port 8503)
  nginx_sas.conf    location /sas/ block
```

## Rules baked in (from the spec)
- Scope: `Project Text (Notification) = KDS SAS Emmen`. Totals band row dropped.
- Batch: `W.IC248.Q.NNN` -> Batch N; `.900` = Lager (stock); `.901` = Springs.
- Status: Closed / Open / Deleted (SAP raw).
- Defect class: **Major = 4 or 5; Minor = 0-3; `-` = Unclassified.**
- CoPQ: SAP books it negative; shown as positive cost. Unbooked = WBS `-`.
- Vendor names normalised (case/space) so duplicates merge.

## First deploy (on chbs4212)

```bash
cd ~/bgtools/pages/dash_sas
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt

# create the tables + load the first export
./.venv/bin/python ingest_sas.py /path/to/NC_Overview_export.xlsx      # rebuild (default)

# service
sudo cp dash_sas.service /etc/systemd/system/dash_sas.service
#   edit User= and the paths inside it if the home path differs
sudo systemctl daemon-reload && sudo systemctl enable --now dash_sas

# nginx: paste nginx_sas.conf's two blocks into the existing
#   /etc/nginx/sites-available/quality  server { } block
sudo nginx -t && sudo systemctl reload nginx
```

Open **http://chbs4212/sas**.

## Update the data later
Either drop a new export through the **Load export** button on the page, or:
```bash
cd ~/bgtools/pages/dash_sas
./.venv/bin/python ingest_sas.py /path/to/new_export.xlsx            # replace all
./.venv/bin/python ingest_sas.py /path/to/new_export.xlsx --append   # add/update
```
No restart needed — the app reads the DB on every request.

## Service control
```bash
sudo systemctl restart dash_sas
sudo systemctl status dash_sas
journalctl -u dash_sas -n 40
```

## Feedback
The Feedback button writes to the `feedback` table in `quality.db`; attachments
land in `~/bgtools/dash/feedback_attachments/`. Read them:
```bash
sqlite3 ~/bgtools/dash/quality.db \
 "select ts,category,message from feedback where app='SAS Dashboard (web)' order by id desc"
```
