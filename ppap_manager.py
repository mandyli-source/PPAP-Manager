"""
PPAP Manager — with Debug Mode
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
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PPAP Manager",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

SUPPORTED_EXT = {".xlsx", ".docx", ".xdw"}

# Naming rule: [PI]_[PN]_[CUSTOMER].[ext]
FILE_PATTERN = re.compile(
    r"^(?P<PI>[^_]+)_(?P<PN>[^_]+)_(?P<CUSTOMER>[^_.]+)(?P<EXT>\.[a-zA-Z0-9]+)$",
    re.IGNORECASE,
)

DEFAULT_FOLDER = r"C:\Users\S2234009\Desktop\PPAP"

DOCUWORKS_PATHS = [
    r"C:\Program Files\Fuji Xerox\DocuWorks\deskew.exe",
    r"C:\Program Files (x86)\Fuji Xerox\DocuWorks\deskew.exe",
    r"C:\Program Files\Fujifilm\DocuWorks\deskew.exe",
    r"C:\Program Files (x86)\Fujifilm\DocuWorks\deskew.exe",
]

# ─────────────────────────────────────────────────────────────────────────────
# DATA PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def scan_directory(folder: str):
    """
    Scan folder recursively.
    Returns (DataFrame of matched files, list of skipped filenames).
    """
    folder_path = Path(folder)
    if not folder_path.exists():
        return pd.DataFrame(), [], "Folder does not exist."

    all_files   = []
    matched     = []
    skipped     = []

    for fp in folder_path.rglob("*"):
        if not fp.is_file():
            continue
        if fp.suffix.lower() not in SUPPORTED_EXT:
            continue
        all_files.append(fp.name)
        m = FILE_PATTERN.match(fp.name)
        if not m:
            skipped.append(fp.name)
            continue
        try:
            stat    = fp.stat()
            created = datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d")
            size_kb = round(stat.st_size / 1024, 1)
        except OSError:
            created, size_kb = "—", 0

        matched.append({
            "PI":        m.group("PI"),
            "PN":        m.group("PN"),
            "CUSTOMER":  m.group("CUSTOMER"),
            "EXT":       m.group("EXT").lower(),
            "FILE_NAME": fp.name,
            "FILE_PATH": str(fp),
            "CREATED":   created,
            "SIZE_KB":   size_kb,
        })

    if not matched:
        df = pd.DataFrame(
            columns=["PI","PN","CUSTOMER","EXT","FILE_NAME","FILE_PATH","CREATED","SIZE_KB"]
        )
    else:
        df = pd.DataFrame(matched)
        df.sort_values(["PN","PI"], ascending=[True,False], inplace=True)
        df.reset_index(drop=True, inplace=True)

    msg = f"Found {len(all_files)} supported file(s) total. " \
          f"{len(matched)} matched naming rule. " \
          f"{len(skipped)} skipped (wrong name format)."
    return df, skipped, msg


@st.cache_data(ttl=60, show_spinner="Scanning folder…")
def get_index(folder: str):
    return scan_directory(folder)


def extract_xlsx_summary(file_path: str) -> dict:
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
# FILE OPENERS
# ─────────────────────────────────────────────────────────────────────────────

def open_file_os(file_path: str):
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", file_path], check=True)
        else:
            subprocess.run(["xdg-open", file_path], check=True)
    except Exception as e:
        st.error(f"Cannot open file: {e}")


def open_docuworks(file_path: str):
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
        key="folder",
    )

    debug_mode = st.toggle("🔍 Debug mode", value=False,
                           help="Show all files found in folder and naming issues")

    if st.button("🔄  Re-scan folder", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    df_index, skipped_files, scan_msg = get_index(folder)

    st.divider()
    col1, col2 = st.columns(2)
    col1.metric("Total files", len(df_index))
    col2.metric("PN codes",    df_index["PN"].nunique() if not df_index.empty else 0)

    if not df_index.empty:
        n_8d = df_index[df_index["EXT"] == ".docx"]["PN"].nunique()
        st.metric("PN codes with 8D alerts", n_8d)

    st.divider()
    st.caption(f"Refreshed: {datetime.now().strftime('%d/%m/%Y  %H:%M')}")
    st.caption(f"Folder: `{folder}`")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

st.title("🔍  Search PPAP")

# ── DEBUG MODE ────────────────────────────────────────────────────────────────
if debug_mode:
    st.subheader("🔍 Debug — Folder scan report")

    folder_path = Path(folder)

    # 1. Does folder exist?
    if not folder_path.exists():
        st.error(f"❌ Folder does not exist: `{folder}`")
    else:
        st.success(f"✅ Folder exists: `{folder}`")

        # 2. List ALL files in folder (any extension)
        all_items = list(folder_path.rglob("*"))
        all_actual_files = [f for f in all_items if f.is_file()]

        st.info(f"Total files found (all types): **{len(all_actual_files)}**")

        if all_actual_files:
            file_data = []
            for f in all_actual_files:
                m = FILE_PATTERN.match(f.name)
                parsed_ok = m is not None and f.suffix.lower() in SUPPORTED_EXT
                file_data.append({
                    "File name":    f.name,
                    "Extension":    f.suffix.lower(),
                    "Supported":    "✅" if f.suffix.lower() in SUPPORTED_EXT else "❌",
                    "Name matches rule": "✅" if parsed_ok else "❌",
                    "PI":       m.group("PI")       if parsed_ok else "—",
                    "PN":       m.group("PN")       if parsed_ok else "—",
                    "CUSTOMER": m.group("CUSTOMER") if parsed_ok else "—",
                })
            st.dataframe(pd.DataFrame(file_data), use_container_width=True, hide_index=True)
        else:
            st.warning("No files found inside this folder at all.")

        # 3. Naming rule reminder
        st.markdown("""
**Naming rule required:**
```
[PI]_[PN]_[CUSTOMER].[ext]
```
Example: `PI-2024-091_PN-4471_CUST01.xlsx`

- Exactly **2 underscores `_`** separating 3 parts
- Extension must be `.xlsx`, `.docx`, or `.xdw`
- Spaces in filename are **not allowed**
        """)

        # 4. Skipped files detail
        if skipped_files:
            st.warning(f"⚠️ {len(skipped_files)} file(s) skipped due to naming mismatch:")
            for f in skipped_files:
                st.code(f)
        else:
            if all_actual_files:
                st.success("✅ All supported files passed the naming check.")

    st.divider()

# ── FOLDER / INDEX GUARDS ─────────────────────────────────────────────────────
if not Path(folder).exists():
    st.error(
        f"**Folder not found:** `{folder}`\n\n"
        "Make sure the folder exists on your computer, then click **Re-scan folder**.\n\n"
        "💡 Tip: turn on **Debug mode** in the sidebar to see exactly what the app can find."
    )
    st.stop()

if df_index.empty:
    st.warning(
        f"No matching files found in `{folder}`.\n\n"
        "Files must follow this naming rule:\n"
        "```\n[PI]_[PN]_[CUSTOMER].xlsx\n"
        "[PI]_[PN]_[CUSTOMER].docx\n"
        "[PI]_[PN]_[CUSTOMER].xdw\n```\n"
        "Example: `PI-2024-091_PN-4471_CUST01.xlsx`\n\n"
        "💡 Turn on **Debug mode** in the sidebar to see the full list of files and what's wrong."
    )
    st.stop()

# ── SEARCH BAR ────────────────────────────────────────────────────────────────
keyword = st.text_input(
    "Search",
    placeholder="Enter PI number, PN code, or Customer Code…",
    label_visibility="collapsed",
)

if not keyword:
    st.info("Enter a PI number, PN code, or Customer Code above to search.")
    with st.expander("📋  Browse all indexed files"):
        st.dataframe(
            df_index[["PI","PN","CUSTOMER","EXT","CREATED","SIZE_KB"]],
            use_container_width=True,
            hide_index=True,
        )
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# QUERY & ANALYZE
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
# RESULTS — ONE SECTION PER PN
# ─────────────────────────────────────────────────────────────────────────────

running_local = platform.system() == "Windows"

for pn in results["PN"].unique():
    pn_df    = results[results["PN"] == pn]
    customer = pn_df["CUSTOMER"].iloc[0]

    xlsx_count   = int((pn_df["EXT"] == ".xlsx").sum())
    xdw_count    = int((pn_df["EXT"] == ".xdw").sum())
    docx_count   = int((pn_df["EXT"] == ".docx").sum())
    update_count = xlsx_count + xdw_count

    # 8D alert
    if docx_count > 0:
        st.error(
            f"⚠️ **8D Alert** — {docx_count} defect report(s) (.docx) linked to "
            f"**{pn}**. Review root causes and corrective actions."
        )
    else:
        st.success(f"✅ **{pn}** — No 8D reports on record.")

    st.subheader(f"📦  {pn}  —  Customer: {customer}")

    c1, c2, c3 = st.columns(3)
    c1.metric("PPAP update count",  update_count)
    c2.metric("DocuWorks drawings", xdw_count)
    c3.metric("8D reports (.docx)", docx_count)

    st.divider()
    st.markdown("#### PPAP history")

    for _, row in pn_df.sort_values("PI", ascending=False).iterrows():
        with st.expander(
            f"**{row['PI']}** &nbsp;|&nbsp; `{row['EXT']}` &nbsp;|&nbsp; {row['CREATED']}",
            expanded=False,
        ):
            st.code(row["FILE_PATH"], language=None)

            # ── .xlsx ─────────────────────────────────────────────────────────
            if row["EXT"] == ".xlsx":
                col_a, col_b = st.columns(2)
                if col_a.button("📊  View measurement data",
                                key=f"view_{row['FILE_NAME']}"):
                    with st.spinner("Reading Excel…"):
                        summary = extract_xlsx_summary(row["FILE_PATH"])
                    if summary:
                        st.table(pd.DataFrame(summary.items(),
                                              columns=["Parameter","Value"]))
                    else:
                        st.info("No structured data found in this file.")

                if running_local:
                    if col_b.button("📂  Open file",
                                    key=f"open_{row['FILE_NAME']}"):
                        open_file_os(row["FILE_PATH"])
                        st.toast("Opening Excel file…")
                else:
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

            # ── .docx ─────────────────────────────────────────────────────────
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

            # ── .xdw ─────────────────────────────────────────────────────────
            elif row["EXT"] == ".xdw":
                st.info("🖼  Drawing / document (DocuWorks format)")
                if running_local:
                    if st.button("🖥  Open in DocuWorks Viewer",
                                 key=f"dw_{row['FILE_NAME']}"):
                        open_docuworks(row["FILE_PATH"])
                        st.toast("Launching DocuWorks Viewer…")
                else:
                    st.info("DocuWorks files can only be opened on a local Windows machine.")

    st.divider()
