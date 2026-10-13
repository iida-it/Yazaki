import streamlit as st
import pdfplumber
import pandas as pd
import unicodedata
import io
import os
import subprocess
import time
import zipfile
import shutil
from pathlib import Path
from datetime import datetime
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Yazaki DocuWorks & PDF Processor",
    page_icon="📑",
    layout="wide"
)

st.title("📑 Yazaki Multi-Stage Processing Portal")

# 1. Dynamic base directory resolution (based on app.py location)
APP_DIR = Path(__file__).resolve().parent

CONVERT_DIR = os.path.join(APP_DIR, "converted_pdfs")
PDF_OUTPUT_DIR = os.path.join(APP_DIR, "PDF_Output")
REPORT_DIR = os.path.join(APP_DIR, "Report")

os.makedirs(CONVERT_DIR, exist_ok=True)
os.makedirs(PDF_OUTPUT_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)

# 2. Path detection for DocuWorks Viewer Light
DWVLT_PATHS = [
    r"C:\Program Files (x86)\FUJIFILM\DocuWorks\bin\dwvlt.exe",
    r"C:\Program Files\FUJIFILM\DocuWorks\bin\dwvlt.exe",
    r"C:\Program Files (x86)\Fuji Xerox\DocuWorks\bin\dwvlt.exe"
]

DWVLT_EXE = None
for p in DWVLT_PATHS:
    if os.path.isfile(p):
        DWVLT_EXE = p
        break

# 3. Helper to configure Bullzip silent printing without UI prompts
def configure_bullzip_silent(target_pdf_path):
    """
    Writes Bullzip runonce.ini configuration and patches settings.ini
    so it silently saves directly to target_pdf_path without any popups,
    progress dialogs, or 'Do you want to open the document?' prompts.
    """
    appdata = os.environ.get("APPDATA", "")
    pdf_writer_root = os.path.join(appdata, "PDF Writer")

    config_content = (
        "[PDF Printer]\n"
        f"Output={os.path.abspath(target_pdf_path)}\n"
        "ShowSettings=never\n"
        "ShowSaveAS=never\n"
        "ShowProgress=no\n"
        "ShowProgressFinished=no\n"
        "ShowPDF=no\n"
        "ConfirmOverwrite=yes\n"
    )

    written = False
    if os.path.isdir(pdf_writer_root):
        for item in os.listdir(pdf_writer_root):
            sub_dir = os.path.join(pdf_writer_root, item)
            if os.path.isdir(sub_dir):
                # 1. Write runonce.ini for the current print job
                runonce_file = os.path.join(sub_dir, "runonce.ini")
                with open(runonce_file, "w", encoding="utf-8") as f:
                    f.write(config_content)
                written = True

                # 2. Patch settings.ini to prevent post-print 'Do you want to open the document?' prompt
                settings_file = os.path.join(sub_dir, "settings.ini")
                if os.path.isfile(settings_file):
                    try:
                        with open(settings_file, "r", encoding="utf-8", errors="ignore") as sf:
                            s_data = sf.read()
                        updated = False
                        for old_line, new_line in [
                            ("showpdf=ask", "showpdf=no"),
                            ("showpdf=yes", "showpdf=no"),
                            ("showprogress=yes", "showprogress=no"),
                            ("showprogressfinished=yes", "showprogressfinished=no"),
                        ]:
                            if old_line in s_data:
                                s_data = s_data.replace(old_line, new_line)
                                updated = True
                        if updated:
                            with open(settings_file, "w", encoding="utf-8") as sf:
                                sf.write(s_data)
                    except Exception:
                        pass

    # Fallback to default directory name if none existed yet
    if not written:
        default_dir = os.path.join(pdf_writer_root, "Bullzip PDF Printer")
        os.makedirs(default_dir, exist_ok=True)
        with open(os.path.join(default_dir, "runonce.ini"), "w", encoding="utf-8") as f:
            f.write(config_content)

def convert_single_xdw(file_name, file_bytes):
    """
    Converts a single .xdw file into PDF using DocuWorks Viewer Light and Bullzip PDF Printer.
    Returns target_pdf path if successful, None if timed out/failed.
    """
    if not DWVLT_EXE:
        return None

    base_name = os.path.splitext(file_name)[0]
    temp_xdw = os.path.join(CONVERT_DIR, file_name)
    target_pdf = os.path.join(PDF_OUTPUT_DIR, f"{base_name}.pdf")

    # 1. Clear previous PDF if exists
    if os.path.exists(target_pdf):
        try:
            os.remove(target_pdf)
        except Exception:
            pass

    # 2. Save .xdw locally
    with open(temp_xdw, "wb") as f:
        f.write(file_bytes)

    # 3. Instruct Bullzip to save directly to target_pdf silently
    configure_bullzip_silent(target_pdf)

    # 4. Invoke silent print via dwvlt.exe
    cmd = f'start /wait "" "{DWVLT_EXE}" /p "{temp_xdw}"'
    subprocess.run(cmd, shell=True)

    # 5. Wait for target PDF creation and size stability
    timeout = 35
    start_t = time.time()
    file_done = False
    last_size = -1
    stable_count = 0

    while time.time() - start_t < timeout:
        if os.path.exists(target_pdf) and os.path.getsize(target_pdf) > 0:
            try:
                current_size = os.path.getsize(target_pdf)
                if current_size == last_size:
                    stable_count += 1
                else:
                    stable_count = 0
                    last_size = current_size

                if stable_count >= 2:
                    with open(target_pdf, "rb") as test_f:
                        if len(test_f.read(1024)) > 0:
                            file_done = True
                            break
            except (PermissionError, IOError):
                stable_count = 0
        time.sleep(1)

    # Cleanup source temp .xdw
    try:
        os.remove(temp_xdw)
    except Exception:
        pass

    return target_pdf if file_done else None

# ----------------- Helper Functions for PDF Parsing -----------------
def extract_target_tables(pdf_bytes, keyword="適合"):
    matched_pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            norm_text = unicodedata.normalize('NFKC', text)
            
            if keyword in text or keyword in norm_text or "矢崎ユニット明細書" in text:
                tables = page.extract_tables()
                if tables:
                    matched_pages.append((idx, tables))
                    
            if '付 図 １' in text or '付図1' in norm_text:
                break

    matched_pages.sort(key=lambda x: x[0])
    all_tables = []
    for _, tbls in matched_pages:
        all_tables.extend(tbls)
    return all_tables

def parse_table_to_df(page_tables):
    if not page_tables:
        return pd.DataFrame()

    all_data_rows = []
    for tbl in page_tables:
        normalized_table = []
        for row in tbl:
            if not row:
                continue
            norm_row = [
                unicodedata.normalize('NFKC', str(c)).strip() if c is not None else None
                for c in row
            ]
            normalized_table.append(norm_row)

        if not normalized_table:
            continue

        df_raw = pd.DataFrame(normalized_table)
        header_row_idx, branch_col_idx = None, None

        for r_idx, row in df_raw.iterrows():
            for c_idx, cell_val in enumerate(row):
                if cell_val and "分岐本数" in str(cell_val):
                    header_row_idx = r_idx
                    branch_col_idx = c_idx
                    break
            if header_row_idx is not None:
                break

        start_r = header_row_idx + 1 if header_row_idx is not None else 0
        if branch_col_idx is None:
            branch_col_idx = 3

        plan_col_idx = 0

        for r_idx in range(start_r, len(df_raw)):
            row = df_raw.iloc[r_idx]
            row_str = " ".join([str(x) for x in row if x])

            if "承 認" in row_str or "承認" in row_str or "矢崎エナジーシステム" in row_str:
                break
            if "分岐本数" in row_str or "プランNo" in row_str:
                continue

            plan_val = row[plan_col_idx] if len(row) > plan_col_idx else None
            count_val = row[branch_col_idx] if branch_col_idx < len(row) else None

            if count_val and str(count_val).isdigit():
                all_data_rows.append({
                    "適合プランNo.": plan_val,
                    "分岐本数": int(count_val)
                })

    if not all_data_rows:
        return pd.DataFrame()

    df_clean = pd.DataFrame(all_data_rows)
    # Convert empty or whitespace-only strings to NA so ffill propagates merged cells accurately
    df_clean["適合プランNo."] = df_clean["適合プランNo."].replace(r"^\s*$", pd.NA, regex=True).ffill()
    summary_df = df_clean.groupby("適合プランNo.", as_index=False, sort=False)["分岐本数"].sum()
    summary_df.rename(columns={"分岐本数": "合計 分岐本数(本)"}, inplace=True)
    return summary_df

def extract_single_pdf_to_df(pdf_bytes, file_name, row_counter_start=1):
    """
    Extracts target summary table from raw PDF bytes and appends file name and sequential numbering.
    """
    tables = extract_target_tables(pdf_bytes)
    df_summary = parse_table_to_df(tables)
    if not df_summary.empty:
        df_summary.insert(0, "File Name", file_name)
        df_summary.insert(0, "No.", range(row_counter_start, row_counter_start + len(df_summary)))
        return df_summary
    return pd.DataFrame()

def build_excel_report(master_rows):
    """
    Consolidates parsed DataFrames into an Excel file with formatted group subtotals,
    vertical cell merges for each file group, and pastel styling matching specifications.
    Returns (final_df, excel_bytes, excel_name, excel_path).
    """
    if not master_rows:
        return None, None, None, None

    final_df = pd.concat(master_rows, ignore_index=True)
    # Add file-level sum column for the Streamlit preview table
    if "合計 分岐本数(本)" in final_df.columns and "File Name" in final_df.columns:
        final_df["ファイル合計"] = final_df.groupby("File Name")["合計 分岐本数(本)"].transform("sum")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_name = f"マンション納品明細書_{timestamp}.xlsx"
    excel_path = os.path.join(REPORT_DIR, excel_name)

    # 10 harmonious highlight colors matching the reference screenshot and cycling
    PALETTE_COLORS = [
        "FFFF00",  # Yellow (HZ6H377)
        "FFC000",  # Orange (HZ6H435)
        "92D050",  # Light Green (HZ6H491)
        "00B050",  # Medium Green (HZ6I016)
        "00B0F0",  # Sky Blue (HZ6I026)
        "B2A1C7",  # Soft Purple
        "F4B084",  # Soft Coral
        "A9D08E",  # Sage Green
        "9BC2E6",  # Soft Blue
        "FFD966",  # Goldenrod
    ]

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name="Summary")
        ws = writer.sheets["Summary"]

        total_rows = len(final_df)
        total_cols = len(final_df.columns)

        # Style Header Row (Row 1)
        for col_idx in range(1, total_cols + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = Font(name="MS Pゴシック", size=11, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        # Identify contiguous file groups
        file_groups = []
        current_file = None
        start_row = 2

        for r_idx in range(2, total_rows + 2):
            f_val = ws.cell(row=r_idx, column=2).value
            if current_file is None:
                current_file = f_val
                start_row = r_idx
            elif f_val != current_file:
                file_groups.append((current_file, start_row, r_idx - 1))
                current_file = f_val
                start_row = r_idx
        if current_file is not None:
            file_groups.append((current_file, start_row, total_rows + 1))

        # Format rows and merged sums per file group
        for group_idx, (f_name, s_row, e_row) in enumerate(file_groups):
            fill_hex = PALETTE_COLORS[group_idx % len(PALETTE_COLORS)]
            file_fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")

            for r in range(s_row, e_row + 1):
                # No. (Col 1)
                c_no = ws.cell(row=r, column=1)
                c_no.font = Font(name="MS Pゴシック", size=11)
                c_no.alignment = Alignment(horizontal="center", vertical="center")
                c_no.border = thin_border

                # File Name (Col 2)
                c_fn = ws.cell(row=r, column=2)
                c_fn.font = Font(name="MS Pゴシック", size=11)
                c_fn.alignment = Alignment(horizontal="left", vertical="center")
                c_fn.fill = file_fill
                c_fn.border = thin_border

                # 適合プランNo. (Col 3)
                c_plan = ws.cell(row=r, column=3)
                c_plan.font = Font(name="MS Pゴシック", size=11)
                c_plan.alignment = Alignment(horizontal="left", vertical="center")
                c_plan.border = thin_border

                # 合計 分岐本数(本) (Col 4)
                c_cnt = ws.cell(row=r, column=4)
                c_cnt.font = Font(name="MS Pゴシック", size=11)
                c_cnt.alignment = Alignment(horizontal="right", vertical="center")
                c_cnt.border = thin_border

                # ファイル合計 (Col 5) - ensure borders on every cell in block
                c_e = ws.cell(row=r, column=5)
                c_e.font = Font(name="MS Pゴシック", size=11, bold=True)
                c_e.border = thin_border

            # Clear subcells in column 5 so Excel merges cleanly
            for r in range(s_row + 1, e_row + 1):
                ws.cell(row=r, column=5, value=None)

            # Set Excel SUM formula on the top cell of the group
            top_e_cell = ws.cell(row=s_row, column=5)
            top_e_cell.value = f"=SUM(D{s_row}:D{e_row})"

            # Merge vertically if multiple rows
            if s_row < e_row:
                ws.merge_cells(start_row=s_row, start_column=5, end_row=e_row, end_column=5)

            top_e_cell.alignment = Alignment(horizontal="center", vertical="center")
            top_e_cell.font = Font(name="MS Pゴシック", size=11, bold=True)

        # Auto-adjust column widths with padding for Japanese characters
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or "")
                if not val.startswith("="):
                    w = sum(2 if ord(c) > 127 else 1 for c in val)
                    if w > max_len:
                        max_len = w
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    excel_bytes = excel_buffer.getvalue()
    try:
        with open(excel_path, "wb") as f:
            f.write(excel_bytes)
    except Exception:
        pass

    return final_df, excel_bytes, excel_name, excel_path

def build_pdf_zip(pdf_paths):
    """
    Packages a list of PDF file paths into a zip buffer.
    """
    if not pdf_paths:
        return None
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for pdf_path in pdf_paths:
            if os.path.exists(pdf_path):
                zf.write(pdf_path, os.path.basename(pdf_path))
    return zip_buffer.getvalue()

# ----------------- UI Tabs Structure -----------------
tab_auto, tab1, tab2 = st.tabs([
    "⚡ Full Automation (.xdw → PDF & Excel)",
    "🖨️ Stage 1: Convert .xdw to PDF (Bulk)", 
    "📊 Stage 2: Extract PDF to Excel Report"
])

# Initialize session state keys to ensure downloads and state persist across reruns
# Full Automation State
if "auto_final_df" not in st.session_state:
    st.session_state["auto_final_df"] = None
if "auto_excel_bytes" not in st.session_state:
    st.session_state["auto_excel_bytes"] = None
if "auto_excel_name" not in st.session_state:
    st.session_state["auto_excel_name"] = None
if "auto_excel_path" not in st.session_state:
    st.session_state["auto_excel_path"] = None
if "auto_zip_bytes" not in st.session_state:
    st.session_state["auto_zip_bytes"] = None
if "auto_pdf_count" not in st.session_state:
    st.session_state["auto_pdf_count"] = 0
if "auto_extracted_count" not in st.session_state:
    st.session_state["auto_extracted_count"] = 0

# Stage 1 State
if "stage1_zip_bytes" not in st.session_state:
    st.session_state["stage1_zip_bytes"] = None
if "stage1_converted_count" not in st.session_state:
    st.session_state["stage1_converted_count"] = 0

# Stage 2 State
if "stage2_final_df" not in st.session_state:
    st.session_state["stage2_final_df"] = None
if "stage2_excel_bytes" not in st.session_state:
    st.session_state["stage2_excel_bytes"] = None
if "stage2_excel_name" not in st.session_state:
    st.session_state["stage2_excel_name"] = None
if "stage2_excel_path" not in st.session_state:
    st.session_state["stage2_excel_path"] = None
if "stage2_count" not in st.session_state:
    st.session_state["stage2_count"] = 0

# ==================== FULL AUTOMATION: XDW TO PDF & EXCEL ====================
with tab_auto:
    st.subheader("⚡ 1-Click End-to-End Automation (.xdw ➔ PDF ➔ Excel)")
    st.caption("Upload `.xdw` files. The system converts them to PDF and extracts the tables directly into a consolidated Excel report in one click.")

    if not DWVLT_EXE:
        st.error("⚠️ DocuWorks Viewer Light (`dwvlt.exe`) was not detected on this PC.")
    else:
        uploaded_auto_xdw = st.file_uploader(
            "Upload Multiple .xdw Files for Full Automation", 
            type=["xdw"], 
            accept_multiple_files=True,
            key="auto_xdw_uploader"
        )

        if uploaded_auto_xdw:
            if st.button("🚀 Start Full Automation (.xdw ➔ Excel Report)", key="btn_start_full_auto", type="primary"):
                converted_pdf_list = []
                total_files = len(uploaded_auto_xdw)
                progress_bar = st.progress(0)
                status_text = st.empty()

                # Phase 1: Convert .xdw to PDF
                for idx, file in enumerate(uploaded_auto_xdw):
                    status_text.text(f"Phase 1/2 - Converting ({idx+1}/{total_files}): {file.name} to PDF ...")
                    file.seek(0)
                    pdf_path = convert_single_xdw(file.name, file.read())
                    if pdf_path:
                        converted_pdf_list.append(pdf_path)
                    else:
                        st.warning(f"⚠️ Conversion timed out for: {file.name}")
                    progress_bar.progress((idx + 1) / (total_files * 2))

                # Clean up DWVLT processes
                subprocess.run(["taskkill", "/f", "/im", "DWVLT.exe"], capture_output=True)

                if not converted_pdf_list:
                    st.error("No PDFs were produced. Please ensure Bullzip PDF Printer is installed and set as default printer.")
                else:
                    # Phase 2: Extract data from converted PDFs into Excel
                    master_rows = []
                    row_counter = 1
                    for idx, pdf_path in enumerate(converted_pdf_list):
                        file_name = os.path.basename(pdf_path)
                        status_text.text(f"Phase 2/2 - Extracting tables ({idx+1}/{len(converted_pdf_list)}): {file_name} ...")
                        try:
                            with open(pdf_path, "rb") as pf:
                                pdf_bytes = pf.read()
                            df_summary = extract_single_pdf_to_df(pdf_bytes, file_name, row_counter_start=row_counter)
                            if not df_summary.empty:
                                row_counter += len(df_summary)
                                master_rows.append(df_summary)
                        except Exception as e:
                            st.warning(f"⚠️ Error parsing {file_name}: {e}")

                        progress_bar.progress(0.5 + ((idx + 1) / (len(converted_pdf_list) * 2)))

                    if master_rows:
                        final_df, excel_bytes, excel_name, excel_path = build_excel_report(master_rows)
                        zip_bytes = build_pdf_zip(converted_pdf_list)

                        st.session_state["auto_final_df"] = final_df
                        st.session_state["auto_excel_bytes"] = excel_bytes
                        st.session_state["auto_excel_name"] = excel_name
                        st.session_state["auto_excel_path"] = excel_path
                        st.session_state["auto_zip_bytes"] = zip_bytes
                        st.session_state["auto_pdf_count"] = len(converted_pdf_list)
                        st.session_state["auto_extracted_count"] = len(master_rows)
                        status_text.text("🎉 Full Automation complete! Excel report and PDFs are ready.")
                    else:
                        st.session_state["auto_final_df"] = None
                        st.warning("PDFs were created, but no matching table rows with '適合' or '矢崎ユニット明細書' were detected.")

        if st.session_state.get("auto_final_df") is not None:
            st.success(
                f"✅ Automation completed successfully! Converted {st.session_state['auto_pdf_count']} file(s) and extracted tables from {st.session_state['auto_extracted_count']} document(s)."
            )

            # Download buttons side-by-side
            d_col1, d_col2 = st.columns(2)
            with d_col1:
                st.download_button(
                    label="📥 Download Consolidated Excel (.xlsx)",
                    data=st.session_state["auto_excel_bytes"],
                    file_name=st.session_state["auto_excel_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_auto_excel",
                    type="primary"
                )
            with d_col2:
                if st.session_state.get("auto_zip_bytes"):
                    st.download_button(
                        label="📦 Download All Converted PDFs (.zip)",
                        data=st.session_state["auto_zip_bytes"],
                        file_name=f"Converted_PDFs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        key="download_auto_zip"
                    )

            st.dataframe(st.session_state["auto_final_df"], use_container_width=True)
            st.info(f"💾 Excel report saved locally to: `{st.session_state['auto_excel_path']}`")
            st.info(f"📁 PDF files saved locally to: `{PDF_OUTPUT_DIR}`")

# ==================== STAGE 1: XDW TO PDF ====================
with tab1:
    st.subheader("Step 1: Multi .xdw Conversion to PDF (Standalone)")
    st.caption(f"Working Directory: `{APP_DIR}`")

    if not DWVLT_EXE:
        st.error("⚠️ DocuWorks Viewer Light (`dwvlt.exe`) was not detected on this PC.")
    else:
        uploaded_xdw_files = st.file_uploader(
            "Upload Multiple .xdw Files", 
            type=["xdw"], 
            accept_multiple_files=True,
            key="xdw_uploader"
        )

        if uploaded_xdw_files:
            if st.button("🚀 Start Bulk Conversion to PDF", key="btn_start_xdw_convert"):
                converted_pdf_list = []
                progress_bar = st.progress(0)
                status_text = st.empty()

                for idx, file in enumerate(uploaded_xdw_files):
                    status_text.text(f"Converting ({idx+1}/{len(uploaded_xdw_files)}): {file.name} ...")
                    file.seek(0)
                    target_pdf = convert_single_xdw(file.name, file.read())

                    if target_pdf:
                        converted_pdf_list.append(target_pdf)
                    else:
                        st.warning(f"⚠️ Conversion timed out for: {file.name}")

                    progress_bar.progress((idx + 1) / len(uploaded_xdw_files))

                # Force kill leftover DWVLT processes
                subprocess.run(["taskkill", "/f", "/im", "DWVLT.exe"], capture_output=True)

                if converted_pdf_list:
                    zip_bytes = build_pdf_zip(converted_pdf_list)
                    st.session_state["stage1_zip_bytes"] = zip_bytes
                    st.session_state["stage1_converted_count"] = len(converted_pdf_list)
                    status_text.text("Conversion process finished!")
                else:
                    st.session_state["stage1_zip_bytes"] = None
                    st.error("No PDFs were produced. Please ensure Bullzip PDF Printer is installed and set as the default printer.")

        if st.session_state.get("stage1_zip_bytes"):
            st.success(f"Successfully converted {st.session_state['stage1_converted_count']} file(s) into `{PDF_OUTPUT_DIR}`!")
            st.download_button(
                label="📦 Download All Converted PDFs (.zip)",
                data=st.session_state["stage1_zip_bytes"],
                file_name=f"Converted_PDFs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                key="download_xdw_zip"
            )

# ==================== STAGE 2: PDF TO EXCEL ====================
with tab2:
    st.subheader("Step 2: Multi-PDF Table Extraction to Consolidated Excel (Standalone)")
    st.caption("Upload multiple PDF files to extract branch counts and consolidate them into an Excel report.")

    uploaded_pdf_files = st.file_uploader(
        "Upload Multiple .pdf Files", 
        type=["pdf"], 
        accept_multiple_files=True,
        key="pdf_uploader"
    )

    if uploaded_pdf_files:
        if st.button("🚀 Extract and Consolidate Reports", key="btn_extract_pdf"):
            master_rows = []
            row_counter = 1
            progress_bar = st.progress(0)
            status_text = st.empty()

            for idx, file in enumerate(uploaded_pdf_files):
                status_text.text(f"Extracting ({idx+1}/{len(uploaded_pdf_files)}): {file.name} ...")
                try:
                    file.seek(0)
                    df_summary = extract_single_pdf_to_df(file.read(), file.name, row_counter_start=row_counter)
                    if not df_summary.empty:
                        row_counter += len(df_summary)
                        master_rows.append(df_summary)
                except Exception as e:
                    st.warning(f"⚠️ Error processing {file.name}: {e}")

                progress_bar.progress((idx + 1) / len(uploaded_pdf_files))

            if master_rows:
                final_df, excel_bytes, excel_name, excel_path = build_excel_report(master_rows)
                st.session_state["stage2_final_df"] = final_df
                st.session_state["stage2_excel_bytes"] = excel_bytes
                st.session_state["stage2_excel_name"] = excel_name
                st.session_state["stage2_excel_path"] = excel_path
                st.session_state["stage2_count"] = len(uploaded_pdf_files)
                status_text.text("Extraction completed successfully!")
            else:
                st.session_state["stage2_final_df"] = None
                st.warning("No target table matching '適合' or '矢崎ユニット明細書' was found in the provided PDF(s).")

    if st.session_state.get("stage2_final_df") is not None:
        st.success(f"Compiled summary from {st.session_state.get('stage2_count', 0)} PDF document(s)!")
        st.dataframe(st.session_state["stage2_final_df"], use_container_width=True)

        st.download_button(
            label="📥 Download Consolidated Excel File",
            data=st.session_state["stage2_excel_bytes"],
            file_name=st.session_state["stage2_excel_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_report"
        )
        st.info(f"💾 Report also saved to: `{st.session_state['stage2_excel_path']}`")