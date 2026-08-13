# Quality Dashboard — command-line runbook

Everything below runs on the server (`chbs4212`) from the dashboard folder.
No GUI, no Cursor file-dragging — command line only.

```bash
cd ~/bgtools/dash
```

The virtual environment lives in `./quality/`. Always call Python and Streamlit
through it: `./quality/bin/python` and `./quality/bin/streamlit`.

---

## 1. Update the data (run in this order)

Source files live in `data/`. Put the current exports there first:

```bash
# copy the latest files into data/ (adjust the source path/names to yours)
cp /path/to/NCtracker11082026.xlsx                                   ~/bgtools/dash/data/
cp /path/to/2026-08-11_09-37-43.xlsm                                 ~/bgtools/dash/data/
cp "/path/to/USE THIS CAPA 2.0 Tracker_RCA ... KPI.xlsm"             ~/bgtools/dash/data/
ls -la data/*.xls*        # confirm they are all present
```

### Stage 1 — clean the tracker (read-only, writes a cleaned copy)

```bash
./quality/bin/python cleaning/clean_tracker.py
```

Produces `cleaning/cleaned_tracker.xlsx` (sheets: summary · clean · flags).
Open the `flags` sheet, fix what it lists **in the tracker**, and re-run this
step until you are happy. Nothing has touched the database yet.

### Stage 2 — rebuild the database

```bash
./quality/bin/python ingest.py
```

Rebuilds `quality.db` from scratch every run (it drops and recreates every
table — there are never leftover rows). Watch for:

- `Reading NCtracker11082026.xlsx`
- `New-system merge: … overlapped … added as new NCs`
- `CAPA rows across N NCs`
- any `⚠ … creation date IN THE FUTURE` line — fix those in the tracker

### Stage 3 — verify integrity

```bash
./quality/bin/python db_health.py
```

Must print `RESULT: PASS`. If it prints `FAIL`, stop and read the reason
(duplicate id, `nc_id='0'`, or a count mismatch). Do not restart on a FAIL.

### Stage 4 — restart the dashboard (drops its cache, loads new data)

```bash
pkill -f "run app.py"; sleep 2; \
nohup ./quality/bin/streamlit run app.py \
  --server.address 0.0.0.0 --server.port 8501 \
  --server.baseUrlPath dashboard \
  > ~/dash.log 2>&1 & \
sleep 4; tail -5 ~/dash.log
```

The dashboard now serves the new data. No other step.

---

## 2. Start / stop / status (no data change)

Start (or restart) the dashboard:

```bash
pkill -f "run app.py"; sleep 2; \
nohup ./quality/bin/streamlit run app.py \
  --server.address 0.0.0.0 --server.port 8501 \
  --server.baseUrlPath dashboard \
  > ~/dash.log 2>&1 &
```

Check it is running:

```bash
ps aux | grep "run app.py" | grep -v grep
```

Stop it:

```bash
pkill -f "run app.py"
```

Read the log (startup errors, the local URL):

```bash
tail -30 ~/dash.log
```

---

## Notes

- **ADAB** runs separately on port `8502` (`--server.baseUrlPath adab`) from
  `~/bgtools/adab_tool`. The commands above do not touch it.
- **Never commit** `quality.db` or the source exports — they are gitignored.
- A quick full refresh in one line (clean → ingest → health → restart):

```bash
cd ~/bgtools/dash && \
./quality/bin/python cleaning/clean_tracker.py && \
./quality/bin/python ingest.py && \
./quality/bin/python db_health.py && \
pkill -f "run app.py"; sleep 2; \
nohup ./quality/bin/streamlit run app.py --server.address 0.0.0.0 --server.port 8501 --server.baseUrlPath dashboard > ~/dash.log 2>&1 & \
sleep 4; tail -5 ~/dash.log
```