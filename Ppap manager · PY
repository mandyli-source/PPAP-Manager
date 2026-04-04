"""
PPAP Manager — Full English UI
Requirements: pip install streamlit pandas openpyxl python-docx watchdog
Run: streamlit run ppap_manager.py
"""

import streamlit as st
import pandas as pd
import os
import re
import subprocess
import platform
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────────────
# APP CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PPAP Manager",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# BLOCK 1 – DATA PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

SUPPORTED_EXTENSIONS = {".xlsx", ".docx", ".xdw"}
# Naming rule: [PI]_[PN]_[CUSTOMER].[ext]
FILE_PATTERN = re.compile(
    r"^(?P<PI>[^_]+)_(?P<PN>[^_]+)_(?P<CUSTOMER>[^_.]+)(?P<EXT>\.[a-zA-Z0-9]+)$",
    re.IGNORECASE,
)


def scan_directory(folder: str) -> pd.DataFrame:
    """Scan folder, parse filenames, return master index DataFrame."""
    records = []
    folder_path = Path(folder)
    if not folder_path.exists():
        return pd.DataFrame()

    for file_path in folder_path.rglob("*"):
        if file_path.is_file() and file_path.suffix.lower() in SUPPORTED_EXTENSIONS:
            match = FILE_PATTERN.match(file_path.name)
            if match:
                stat = file_path.stat()
                records.append(
                    {
                        "PI": match.group("PI"),
                        "PN": match.group("PN"),
                        "CUSTOMER": match.group("CUSTOMER"),
                        "EXT": match.group("EXT").lower(),
                        "FILE_NAME": file_path.name,
                        "FILE_PATH": str(file_path),
                        "CREATED": datetime.fromtimestamp(stat.st_ctime).strftime(
                            "%Y-%m-%d"
                        ),
                        "SIZE_KB": round(stat.st_size / 1024, 1),
                    }
                )

    if not records:
        return pd.DataFrame(
            columns=["PI", "PN", "CUSTOMER", "EXT", "FILE_NAME", "FILE_PATH", "CREATED", "SIZE_KB"]
        )
    df = pd.DataFrame(records)
    df.sort_values(by=["PN", "PI"], ascending=[True, False], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


@st.cache_data(ttl=60, show_spinner="Scanning folder...")
def get_index(folder: str) -> pd.DataFrame:
    """Cached master index — auto-refreshes every 60 seconds."""
    return scan_directory(folder)


def extract_xlsx_summary(file_path: str) -> dict:
    """Extract key-value pairs from the first sheet of an Excel file."""
    try:
        wb = load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        summary = {}
        for row in ws.iter_rows(max_row=30, max_col=10, values_only=True):
            for i, cell in enumerate(row):
                if isinstance(cell, str) and cell.strip():
                    key = cell.strip()
                    val = row[i + 1] if i + 1 < len(row) else None
                    if val is not None and key and len(summary) < 8:
                        summary[key] = val
        wb.close()
        return summary
    except Exception:
        return {}


# ─────────────────────────────────────────────────────────────────────────────
# FILE INTERACTION UTILITIES
# ─────────────────────────────────────────────────────────────────────────────

def open_file_os(file_path: str):
    """Open a file with the OS default application (Windows / Mac / Linux)."""
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", file_path], check=True)
        else:
            subprocess.run(["xdg-open", file_path], check=True)
        return True
    except Exception as e:
        st.error(f"Cannot open file: {e}")
        return False


def open_docuworks(file_path: str):
    """
    Open a .xdw file with DocuWorks Viewer on Windows.
    Adjust exe paths if DocuWorks is installed elsewhere.
    """
    dw_paths = [
        r"C:\Program Files\Fuji Xerox\DocuWorks\deskew.exe",
        r"C:\Program Files (x86)\Fuji Xerox\DocuWorks\deskew.exe",
        r"C:\Program Files\Fujifilm\DocuWorks\deskew.exe",
    ]
    dw_exe = next((p for p in dw_paths if os.path.exists(p)), None)
    if dw_exe:
        subprocess.Popen([dw_exe, file_path])
    else:
        open_file_os(file_path)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 📦 PPAP Manager")
    st.divider()

    folder = st.text_input(
        "PPAP folder",
        value=st.session_state.get("folder", r"C:\PPAP_SHARE"),
        placeholder=r"\\SERVER01\PPAP_SHARE",
        key="folder",
    )

    if st.button("🔄 Re-scan folder", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    df_index = get_index(folder)

    st.divider()
    col1, col2 = st.columns(2)
    col1.metric("Total files", len(df_index))
    col2.metric("PN codes", df_index["PN"].nunique() if not df_index.empty else 0)

    if not df_index.empty:
        n_warn = df_index[df_index["EXT"] == ".docx"]["PN"].nunique()
        st.metric("PN codes with 8D alerts", n_warn)

    st.divider()
    st.caption(f"Index last updated: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    st.caption(f"Folder: `{folder}`")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

st.title("🔍 Search PPAP")

keyword = st.text_input(
    "Search",
    placeholder="Enter PI number, PN code, or Customer Code...",
    label_visibility="collapsed",
)

if not keyword:
    st.info("Enter a PI number, PN code, or Customer Code above to begin searching.")
    st.stop()

if df_index.empty:
    st.warning(
        f"No files found in `{folder}`. "
        "Check the folder path and click **Re-scan folder**."
    )
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# BLOCK 2 – QUERY & ANALYZE
# ─────────────────────────────────────────────────────────────────────────────

kw = keyword.strip().upper()
mask = (
    df_index["PI"].str.upper().str.contains(kw, na=False)
    | df_index["PN"].str.upper().str.contains(kw, na=False)
    | df_index["CUSTOMER"].str.upper().str.contains(kw, na=False)
)
results = df_index[mask].copy()

if results.empty:
    st.warning(f"No results found for `{keyword}`.")
    st.stop()

pn_list = results["PN"].unique()

for pn in pn_list:
    pn_df = results[results["PN"] == pn]
    customer = pn_df["CUSTOMER"].iloc[0]

    xlsx_count   = (pn_df["EXT"] == ".xlsx").sum()
    xdw_count    = (pn_df["EXT"] == ".xdw").sum()
    docx_count   = (pn_df["EXT"] == ".docx").sum()
    update_count = xlsx_count + xdw_count
    has_8d = docx_count > 0

    # ── 8D ALERT ─────────────────────────────────────────────────────────────
    if has_8d:
        st.error(
            f"⚠️ **8D Alert** — **{docx_count}** defect report(s) (.docx) linked to "
            f"**{pn}**. Review root causes and corrective actions."
        )
    else:
        st.success(f"✅ **{pn}** — No 8D reports on record.")

    # ── HEADER ───────────────────────────────────────────────────────────────
    st.subheader(f"📦 {pn}  —  Customer: {customer}")

    c1, c2, c3 = st.columns(3)
    c1.metric("PPAP update count", update_count)
    c2.metric("DocuWorks drawings", xdw_count)
    c3.metric("8D reports (.docx)", docx_count)

    st.divider()

    # ── PPAP HISTORY ──────────────────────────────────────────────────────────
    st.markdown("#### PPAP history")

    for _, row in pn_df.sort_values("PI", ascending=False).iterrows():
        with st.expander(
            f"**{row['PI']}** &nbsp;|&nbsp; `{row['EXT']}` &nbsp;|&nbsp; {row['CREATED']}",
            expanded=False,
        ):
            st.code(row["FILE_PATH"], language=None)

            # ── BLOCK 3 – FILE INTERACTION ────────────────────────────────────
            if row["EXT"] == ".xlsx":
                col_a, col_b = st.columns(2)
                if col_a.button("📊 View measurement data", key=f"view_{row['FILE_NAME']}"):
                    summary = extract_xlsx_summary(row["FILE_PATH"])
                    if summary:
                        st.table(pd.DataFrame(summary.items(), columns=["Parameter", "Value"]))
                    else:
                        st.info("No structured data could be extracted from this file.")
                if col_b.button("⤓ Open / Download", key=f"dl_{row['FILE_NAME']}"):
                    open_file_os(row["FILE_PATH"])
                    st.toast("Opening Excel file...")

            elif row["EXT"] == ".docx":
                st.warning("📄 8D report — review root cause and corrective action.")
                if st.button("📖 Open 8D report", key=f"8d_{row['FILE_NAME']}"):
                    open_file_os(row["FILE_PATH"])
                    st.toast("Opening 8D report...")

            elif row["EXT"] == ".xdw":
                st.info("🖼 Drawing / document (DocuWorks format)")
                if st.button("🖥 Open in DocuWorks Viewer", key=f"dw_{row['FILE_NAME']}"):
                    open_docuworks(row["FILE_PATH"])
                    st.toast("Launching DocuWorks Viewer...")

    st.divider()
