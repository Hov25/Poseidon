"""
Neptune 360 Tools — Streamlit cloud app
Combines:
  1) Billing File Audit  (billing_audit_core.py)
  2) MIU System Assessment (miu_assessment_core.py)

Optional password gate: set an APP_PASSWORD environment variable /
Streamlit secret to require a password before use. If unset, the app
is open to anyone with the link.
"""

import io
import os
import tempfile
from pathlib import Path

import streamlit as st

import billing_audit_core as billing
import miu_assessment_core as miu

st.set_page_config(page_title="Neptune 360 Tools", page_icon="📊", layout="wide")


# ─── Optional password gate ───────────────────────────────────────────────────

def _required_password():
    try:
        return st.secrets.get("APP_PASSWORD", os.environ.get("APP_PASSWORD", ""))
    except Exception:
        return os.environ.get("APP_PASSWORD", "")


def _check_password() -> bool:
    required = _required_password()
    if not required:
        return True  # no password configured -> open access
    if st.session_state.get("_authed"):
        return True
    st.title("🔒 Neptune 360 Tools")
    pw = st.text_input("Password", type="password")
    if st.button("Enter"):
        if pw == required:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


if not _check_password():
    st.stop()


# ─── Shared helpers ────────────────────────────────────────────────────────────

def _save_uploads_to_tmp(uploaded_files) -> list[str]:
    """Write Streamlit UploadedFile objects to a temp dir, return real paths
    (both parsers open files by path, so this keeps their logic untouched)."""
    paths = []
    tmpdir = tempfile.mkdtemp(prefix="n360_")
    for uf in uploaded_files:
        p = Path(tmpdir) / uf.name
        with open(p, "wb") as f:
            f.write(uf.getbuffer())
        paths.append(str(p))
    return paths


st.title("📊 Neptune 360 Tools")
tab_billing, tab_miu = st.tabs(["🧾 Billing File Audit", "📟 MIU System Assessment"])


# ─── TAB 1: Billing File Audit ────────────────────────────────────────────────

with tab_billing:
    st.subheader("N360 Import File Audit Tool")
    st.caption("Upload one or more Neptune 360 .imp import files to check for "
               "meter size/dial conflicts, duplicate collection IDs, duplicate "
               "meter numbers, and missing meter sizes.")

    billing_files = st.file_uploader(
        "Import files (.imp / .txt)",
        type=["imp", "IMP", "txt", "TXT"],
        accept_multiple_files=True,
        key="billing_uploader",
    )

    if billing_files:
        if st.button("▶ Parse & Preview", key="billing_parse"):
            paths = _save_uploads_to_tmp(billing_files)
            with st.spinner(f"Parsing {len(paths)} file(s)…"):
                all_records = []
                combined_meta = {"company": "", "route": "", "read_date": "",
                                  "total_lines": 0, "file_count": len(paths)}
                routes, read_dates = [], []
                for path in paths:
                    records, meta = billing.parse_imp_file(path)
                    all_records.extend(records)
                    combined_meta["total_lines"] += meta.get("total_lines", 0)
                    if not combined_meta["company"]:
                        combined_meta["company"] = meta.get("company", "")
                    r, d = meta.get("route", ""), meta.get("read_date", "")
                    if r and r not in routes:
                        routes.append(r)
                    if d and d not in read_dates:
                        read_dates.append(d)
                combined_meta["route"] = ", ".join(routes)
                combined_meta["read_date"] = ", ".join(read_dates)

            st.session_state["billing_records"] = all_records
            st.session_state["billing_meta"] = combined_meta

    if st.session_state.get("billing_records"):
        records = st.session_state["billing_records"]
        meta = st.session_state["billing_meta"]

        active = sum(1 for r in records if r["account_status"] in ("ACTI", "AWZ"))
        inactive = len(records) - active
        rd = meta.get("read_date", "")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Records", len(records))
        c2.metric("Active", active)
        c3.metric("Inactive", inactive)
        c4.metric("Files", meta.get("file_count", 1))

        dq = billing._compute_quality(records)
        st.info(
            f"**Data Quality** — Meter Size/Dial Conflicts: {dq['meter_dial_conflicts']} | "
            f"Missing Meter Sizes: {dq['missing_meter_sz']} | "
            f"Duplicate Collection IDs: {dq['dup_coll_ids']} | "
            f"Duplicate Meter Numbers: {dq['dup_meter_nums']}"
        )

        st.dataframe(
            [{
                "Account #": r["account_number"], "Customer Name": r["customer_name"],
                "Address": r["address"], "Status": r["account_status"],
                "Meter Number": r["meter_number"], "Collection ID": r["collection_id"],
                "Size": r["meter_size"], "Meter Type": r.get("meter_type", ""),
                "UOM": r["meter_uom"], "Dials": r["dials"], "Prev Read": r["prev_read"],
            } for r in records],
            use_container_width=True,
            height=400,
        )

        buf = io.BytesIO()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            billing.export_xlsx(records, meta, tf.name)
            tf.seek(0)
            buf.write(Path(tf.name).read_bytes())
        buf.seek(0)

        st.download_button(
            "⬇ Download Excel Audit Report",
            data=buf,
            file_name="billing_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="billing_download",
        )


# ─── TAB 2: MIU System Assessment ─────────────────────────────────────────────

with tab_miu:
    st.subheader("MIU System Assessment Tool")
    st.caption("Upload Neptune 360 .imp files, or .xlsx/.csv meter lists, to "
               "classify MIU type and estimate age based on collection ID/serial number.")

    miu_files = st.file_uploader(
        "Import files (.imp / .xlsx / .csv)",
        type=["imp", "IMP", "txt", "TXT", "xlsx", "xlsm", "csv"],
        accept_multiple_files=True,
        key="miu_uploader",
    )

    col_overrides: dict = st.session_state.setdefault("miu_col_overrides", {})

    if miu_files:
        # For tabular files, let the user confirm/override which column holds
        # the MIU serial number (replaces the desktop app's popup dialog).
        for uf in miu_files:
            ext = Path(uf.name).suffix.lower()
            if ext in (".xlsx", ".xlsm", ".csv"):
                tmp_paths = _save_uploads_to_tmp([uf])
                headers, rows = miu._read_tabular(tmp_paths[0])
                if headers:
                    idx = miu._detect_miu_column_index(headers)
                    default = headers[idx] if idx is not None else headers[0]
                    chosen = st.selectbox(
                        f"MIU serial column for **{uf.name}**",
                        options=headers,
                        index=headers.index(default) if default in headers else 0,
                        key=f"colsel_{uf.name}",
                    )
                    col_overrides[uf.name] = chosen

        if st.button("▶ Analyze MIUs", key="miu_parse"):
            paths = _save_uploads_to_tmp(miu_files)
            path_overrides = {p: col_overrides.get(Path(p).name) for p in paths}

            with st.spinner(f"Analyzing {len(paths)} file(s)…"):
                all_records = []
                combined_meta = {"company": "", "route": "", "read_date": "",
                                  "total_lines": 0, "file_count": len(paths),
                                  "has_tabular": False}
                routes, read_dates, col_notes = [], [], []

                for path in paths:
                    ext = Path(path).suffix.lower()
                    if ext in (".xlsx", ".xlsm", ".csv"):
                        records, meta = miu.parse_tabular_file(path, path_overrides.get(path))
                        combined_meta["has_tabular"] = True
                        mc = meta.get("miu_column") or "undetected"
                        col_notes.append(f"{Path(path).name}→{mc}")
                    else:
                        records, meta = miu.parse_imp_file(path)
                        r, d = meta.get("route", ""), meta.get("read_date", "")
                        if not combined_meta["company"]:
                            combined_meta["company"] = meta.get("company", "")
                        if r and r not in routes:
                            routes.append(r)
                        if d and d not in read_dates:
                            read_dates.append(d)

                    combined_meta["total_lines"] += meta.get("total_lines", 0)
                    all_records.extend(records)

                combined_meta["route"] = ", ".join(routes)
                combined_meta["read_date"] = ", ".join(read_dates)
                combined_meta["col_notes"] = col_notes
                all_records.sort(key=lambda r: r["est_year"] if r["est_year"] else 9999)

            st.session_state["miu_records"] = all_records
            st.session_state["miu_meta"] = combined_meta

    if st.session_state.get("miu_records"):
        records = st.session_state["miu_records"]
        meta = st.session_state["miu_meta"]

        from collections import Counter
        cats = Counter(r["age_category"] for r in records)
        missing_count = sum(1 for r in records if r.get("serial_status") == "missing")
        invalid_count = sum(1 for r in records if r.get("serial_status") == "invalid")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("New (≤5 yrs)", cats.get("New (<=5 yrs)", 0))
        c2.metric("Moderate (6-10)", cats.get("Moderate (6-10 yrs)", 0))
        c3.metric("Aging (11-15)", cats.get("Aging (11-15 yrs)", 0))
        c4.metric("End of Life (>15)", cats.get("End of Life (>15 yrs)", 0))
        c5.metric("Missing/Invalid #", missing_count + invalid_count)

        st.dataframe(
            [{
                "Account #": r["account_number"], "Customer Name": r["customer_name"],
                "Address": r["address"],
                "MIU Serial": r["miu_serial"] or "Missing MIU number",
                "System": r["system"], "MIU Type": r["miu_type"],
                "Mfr. Year": r["est_year"] or "--",
                "End Year": r["end_year"] if r["end_year"] else "Present",
                "Age": r["age_label"], "Age Category": r["age_category"],
            } for r in records],
            use_container_width=True,
            height=400,
        )

        has_tabular = meta.get("has_tabular", False)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            if has_tabular:
                miu.export_tabular_xlsx(records, meta, tf.name)
            else:
                miu.export_xlsx(records, meta, tf.name)
            tf.seek(0)
            data = Path(tf.name).read_bytes()

        st.download_button(
            "⬇ Download MIU Age Report",
            data=data,
            file_name="miu_age_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="miu_download",
        )
