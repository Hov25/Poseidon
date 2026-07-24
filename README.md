# (Streamlit)

Cloud web version of two desktop tools:
- **Billing File Audit** 
- **MIU System Assessment** 

All the original parsing/lookup/Excel-export logic is unchanged — it now
lives in `billing_audit_core.py` and `miu_assessment_core.py` with the
tkinter GUI stripped out. `app.py` is the new Streamlit web front end:
upload files in the browser instead of picking them from disk, preview
results in a table, and click a button to download the Excel report
(instead of `os.startfile` auto-opening it).

## Files
- `app.py` — Streamlit app (two tabs)
- `billing_audit_core.py` — billing audit parsing/export logic
- `miu_assessment_core.py` — MIU lookup/parsing/export logic (includes the
  hardcoded `MIU_RANGES` table, sourced from `Updated_MIU_by_Year_and_model.csv`)
- `requirements.txt` — Python dependencies

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud (free)
1. Push this folder to a GitHub repo (public or private).
2. Go to https://share.streamlit.io → "New app" → point it at the repo, branch,
   and `app.py`.
3. Click Deploy. You'll get a permanent `https://<something>.streamlit.app` link.

## Optional password protection
By default the app is open to anyone with the link. To require a password:
1. In Streamlit Community Cloud, go to your app → **Settings → Secrets**, and add:
   ```
   APP_PASSWORD = "whatever-you-want"
   ```
2. Redeploy. The app will now show a password prompt before the tools load.
3. To remove the password later, just delete that secret.

(Running locally, you can instead set an environment variable:
`export APP_PASSWORD=whatever-you-want` before `streamlit run app.py`.)

## Notes / things worth knowing
- Uploaded files are written to a temporary directory on the server for the
  duration of processing, then processed exactly like the desktop version —
  nothing is stored permanently.
- The `Updated_MIU_by_Year_and_model.csv` you had alongside the desktop tool
  is reference documentation only — the tool doesn't read it at runtime,
  since the ranges are hardcoded into `miu_assessment_core.py`. Keep it as
  documentation of where the ranges came from, or fold future range updates
  directly into `MIU_RANGES` in that file.
- If you add a new company or change the audit/report formatting, edit the
  `_core.py` files — `app.py` shouldn't normally need to change.

