"""
PPAP Manager — Full English UI
Works both locally (e.g. C:\\PPAP) and on Streamlit Cloud (e.g. ./data folder)

Requirements:
    pip install streamlit pandas openpyxl python-docx

Run locally:
    streamlit run ppap_manager.py
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
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PPAP Manager",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
SUPPORTED_EXT = {".xlsx", ".docx", ".xdw"}

# Naming rule: [PI]_[PN]_[CUSTOMER].[ext]
# Example   : PI-2024-091_PN-4471_CUST01.xlsx
FILE_PATTERN = re.compile(
    r"^(?P<PI>[^_]+)_(?P<PN>[^_]+)_(?P<CUSTOMER>[^_.]+)(?P<EXT>\.[a-zA-Z0-9]+)$",
    re.IGNORECASE,
)

# Default folder — works on Windows local and Streamlit Cloud
DEFAULT_FOLDER = r"C:\PPAP" if platform.system() == "Windows" else "data"

# DocuWorks Viewer possible install paths (Windows)
DOCUWORKS_PATHS = [
    r"C:\Program Files\Fuji Xerox\DocuWorks\deskew.exe",
    r"C:\Program Files (x86)\Fuji Xerox\DocuWorks\deskew.exe",
    r"C:\Program Files\Fujifilm\DocuWorks\deskew.exe",
    r"C:\Program Files (x86)\Fujifilm\DocuWorks\deskew.exe",
]

# ─────────────────────────────────────────────────────────────────────────────
# BLOCK 1 — DATA PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def parse_filename(name: str) -> dict | None:
    """
    Parse a filename following the rule [PI]_[PN]_[CUSTOMER].[ext].
    Returns a dict with keys PI, PN, CUSTOMER, EXT or None if no match.
    """
    m = FILE_PATTERN.match(name)
    if not m:
        return None
    ext = m.group("EXT").lower()
    if ext not in SUPPORTED_EXT:
        return None
    return {
        "PI":       m.group("PI"),
        "PN":       m.group("PN"),
        "CUSTOMER": m.group("CUSTOMER"),
        "EXT":      ext,
    }


def scan_directory(folder: str) -> pd.DataFrame:
    """
    Recursively scan *folder*, parse every matching filename,
    and return a master-index DataFrame.
    """
    folder_path = Path(folder)

    if not folder_path.exists():
        return pd.DataFrame()

    records = []
    for fp in folder_path.rglob("*"):
        if not fp.is_file():
            continue
        parsed = parse_filename(fp.name)
        if parsed is None:
            continue
        try:
            stat = fp.stat()
            created = datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d")
            size_kb = round(stat.st_size / 1024, 1)
        except OSError:
            created, size_kb = "—", 0

        records.append({
            **parsed,
            "FILE_NAME": fp.name,
            "FILE_PATH": str(fp),
            "CREATED":   created,
            "SIZE_KB":   size_kb,
        })

    if not records:
        return pd.DataFrame(
            columns=["PI", "PN", "CUSTOMER", "EXT",
                     "FILE_NAME", "FILE_PATH", "CREATED", "SIZE_KB"]
        )

    df = pd.DataFrame(records)
    df.sort_values(["PN", "PI"], ascending=[True, False], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


@st.cache_data(ttl=60, show_spinner="Scanning folder…")
def get_index(folder: str) -> pd.DataFrame:
    """Cached master index — auto-refreshes every 60 s."""
    return scan_directory(folder)


def extract_xlsx_summary(file_path: str) -> dict:
    """
    Read the active sheet and collect up to 8 key-value pairs
    where a string cell is followed by a non-empty value cell.
    """
    try:
        wb = load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        result = {}
        for row in ws.iter_rows(max_row=40, max_col=10, values_only=True):
            for i, cell in enumerate(row):
                if isinstance(cell, str) and cell.strip() and i + 1 < len(row):
                    val = row[i + 1]
                    if val is not None and len(result) < 8:
                        result[cell.strip()] = val
        wb.close()
        return result
    except Exception:
        return {}


# ─────────────────────────────────────────────────────────────────────────────
# BLOCK 3 — FILE INTERACTION HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def is_local() -> bool:
    """True when running on the user's own machine (not Streamlit Cloud)."""
    return os.environ.get("STREAMLIT_SHARING_MODE") is None and \
           os.environ.get("HOME", "").startswith("/home/user") is False


def open_file_os(file_path: str) -> bool:
    """Open a file with the OS default application."""
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)          # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.run(["open", file_path], check=True)
        else:
            subprocess.run(["xdg-open", file_path], check=True)
        return True
    except Exception as e:
        st.error(f"Cannot open file: {e}")
        return False


def open_docuworks(file_path: str):
    """Open a .xdw file with DocuWorks Viewer; fall back to OS default."""
    dw_exe = next((p for p in DOCUWORKS_PATHS if os.path.exists(p)), None)
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
        "PPAP folder path",
        value=st.session_state.get("folder", DEFAULT_FOLDER),
        placeholder=r"C:\PPAP  or  data",
        key="folder",
        help=(
            "Local machine: enter the full Windows path, e.g. C:\\PPAP\n"
            "Streamlit Cloud: enter a relative folder name, e.g. data"
        ),
    )

    if st.button("🔄  Re-scan folder", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # ── Load index ────────────────────────────────────────────────────────────
    df_index = get_index(folder)

    st.divider()

    # ── Summary metrics ───────────────────────────────────────────────────────
    total_files = len(df_index)
    total_pn    = df_index["PN"].nunique() if not df_index.empty else 0
    total_8d    = df_index[df_index["EXT"] == ".docx"]["PN"].nunique() \
                  if not df_index.empty else 0

    col1, col2 = st.columns(2)
    col1.metric("Total files", total_files)
    col2.metric("PN codes",    total_pn)
    st.metric("PN codes with 8D alerts ⚠", total_8d)

    st.divider()
    st.caption(f"Index refreshed: {datetime.now().strftime('%d/%m/%Y  %H:%M')}")
    st.caption(f"Folder: `{folder}`")

    # ── Folder-not-found warning ──────────────────────────────────────────────
    if not Path(folder).exists():
        st.error(
            f"❌ Folder not found:\n\n`{folder}`\n\n"
            "Check the path above and click **Re-scan folder**."
        )

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — HEADER
# ─────────────────────────────────────────────────────────────────────────────

st.title("🔍  Search PPAP")

# ── Folder-not-found guard ────────────────────────────────────────────────────
if not Path(folder).exists():
    st.error(
        f"### Folder not found: `{folder}`\n\n"
        "**If you are running locally on Windows**, make sure the path exists, e.g.:\n"
        "```\nC:\\PPAP\n```\n\n"
        "**If you are on Streamlit Cloud**, the app cannot read your local drive. "
        "Upload your PPAP files to a `data/` folder in the GitHub repo, "
        "then set the folder field to `data`."
    )
    st.stop()

# ── Empty-index guard ─────────────────────────────────────────────────────────
if df_index.empty:
    st.warning(
        f"No matching files found in `{folder}`.\n\n"
        "Make sure files follow the naming rule:\n"
        "```\n[PI]_[PN]_[CUSTOMER].[xlsx | docx | xdw]\n```\n"
        "Example: `PI-2024-091_PN-4471_CUST01.xlsx`"
    )
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — SEARCH BAR
# ─────────────────────────────────────────────────────────────────────────────

keyword = st.text_input(
    "Search",
    placeholder="Enter PI number, PN code, or Customer Code…",
    label_visibility="collapsed",
)

if not keyword:
    # Show quick stats when no search term is entered
    st.info("Enter a PI number, PN code, or Customer Code above to search.")

    with st.expander("📋  Browse all indexed files"):
        st.dataframe(
            df_index[["PI", "PN", "CUSTOMER", "EXT", "CREATED", "SIZE_KB"]],
            use_container_width=True,
            hide_index=True,
        )
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# BLOCK 2 — QUERY & ANALYZE
# ─────────────────────────────────────────────────────────────────────────────

kw = keyword.strip().upper()

mask = (
    df_index["PI"].str.upper().str.contains(kw, na=False)
    | df_index["PN"].str.upper().str.contains(kw, na=False)
    | df_index["CUSTOMER"].str.upper().str.contains(kw, na=False)
)
results = df_index[mask].copy()

if results.empty:
    st.warning(f"No results found for **{keyword}**.")
    st.stop()

st.caption(f"Found **{len(results)}** file(s) matching **{keyword}**")

# ─────────────────────────────────────────────────────────────────────────────
# RESULTS — ONE SECTION PER PN CODE
# ─────────────────────────────────────────────────────────────────────────────

for pn in results["PN"].unique():
    pn_df    = results[results["PN"] == pn]
    customer = pn_df["CUSTOMER"].iloc[0]

    xlsx_count   = int((pn_df["EXT"] == ".xlsx").sum())
    xdw_count    = int((pn_df["EXT"] == ".xdw").sum())
    docx_count   = int((pn_df["EXT"] == ".docx").sum())
    update_count = xlsx_count + xdw_count
    has_8d       = docx_count > 0

    # ── 8D Alert banner ───────────────────────────────────────────────────────
    if has_8d:
        st.error(
            f"⚠️ **8D Alert** — {docx_count} defect report(s) (.docx) linked to "
            f"**{pn}**. Review root causes and corrective actions."
        )
    else:
        st.success(f"✅ **{pn}** — No 8D reports on record.")

    # ── Header ────────────────────────────────────────────────────────────────
    st.subheader(f"📦  {pn}  —  Customer: {customer}")

    c1, c2, c3 = st.columns(3)
    c1.metric("PPAP update count",    update_count)
    c2.metric("DocuWorks drawings",   xdw_count)
    c3.metric("8D reports (.docx)",   docx_count)

    st.divider()

    # ── PPAP History ──────────────────────────────────────────────────────────
    st.markdown("#### PPAP history")

    for _, row in pn_df.sort_values("PI", ascending=False).iterrows():
        label = f"**{row['PI']}** &nbsp;|&nbsp; `{row['EXT']}` &nbsp;|&nbsp; {row['CREATED']}"
        with st.expander(label, expanded=False):

            st.code(row["FILE_PATH"], language=None)

            # ── BLOCK 3 — FILE INTERACTION ────────────────────────────────────
            running_local = platform.system() == "Windows"

            if row["EXT"] == ".xlsx":
                col_a, col_b = st.columns(2)

                # View measurement data
                if col_a.button("📊  View measurement data",
                                key=f"view_{row['FILE_NAME']}"):
                    with st.spinner("Reading Excel file…"):
                        summary = extract_xlsx_summary(row["FILE_PATH"])
                    if summary:
                        st.table(
                            pd.DataFrame(
                                summary.items(),
                                columns=["Parameter", "Value"]
                            )
                        )
                    else:
                        st.info(
                            "No structured key-value data found in this file. "
                            "The sheet may use a custom layout."
                        )

                # Open / Download
                if running_local:
                    if col_b.button("📂  Open file",
                                    key=f"open_{row['FILE_NAME']}"):
                        open_file_os(row["FILE_PATH"])
                        st.toast("Opening Excel file…")
                else:
                    # On Streamlit Cloud: offer a download button
                    try:
                        with open(row["FILE_PATH"], "rb") as f:
                            col_b.download_button(
                                "⤓  Download",
                                data=f,
                                file_name=row["FILE_NAME"],
                                mime="application/vnd.openxmlformats-officedocument"
                                     ".spreadsheetml.sheet",
                                key=f"dl_{row['FILE_NAME']}",
                            )
                    except OSError:
                        col_b.warning("File not accessible.")

            elif row["EXT"] == ".docx":
                st.warning("📄  8D report — review root cause and corrective action.")

                if running_local:
                    if st.button("📖  Open 8D report",
                                 key=f"8d_{row['FILE_NAME']}"):
                        open_file_os(row["FILE_PATH"])
                        st.toast("Opening 8D report…")
                else:
                    try:
                        with open(row["FILE_PATH"], "rb") as f:
                            st.download_button(
                                "⤓  Download 8D report",
                                data=f,
                                file_name=row["FILE_NAME"],
                                mime="application/vnd.openxmlformats-officedocument"
                                     ".wordprocessingml.document",
                                key=f"dl8d_{row['FILE_NAME']}",
                            )
                    except OSError:
                        st.warning("File not accessible.")

            elif row["EXT"] == ".xdw":
                st.info("🖼  Drawing / document (DocuWorks format)")

                if running_local:
                    if st.button("🖥  Open in DocuWorks Viewer",
                                 key=f"dw_{row['FILE_NAME']}"):
                        open_docuworks(row["FILE_PATH"])
                        st.toast("Launching DocuWorks Viewer…")
                else:
                    st.info(
                        "DocuWorks files can only be opened on a local Windows machine. "
                        "Copy the file path above and open it manually."
                    )

    st.divider()
