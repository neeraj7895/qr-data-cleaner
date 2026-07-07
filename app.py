import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import load_workbook
from typing import List, Tuple, Optional, Dict
import logging
import os

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page configuration
st.set_page_config(
    page_title="QR Data Cleaner Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------------------------
# Simple, OneStack-admin-style CSS (dark sidebar + plain white content area
# with dark bar section headers, instead of the earlier heavy card/gradient UI)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    * {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    /* Main App Background */
    .stApp {
        background: #f4f5f7;
    }

    /* Sidebar - dark, like OneStack nav */
    section[data-testid="stSidebar"] {
        background: #111827;
        border-right: 1px solid rgba(255,255,255,0.06);
    }

    section[data-testid="stSidebar"] > div {
        padding: 1rem 1.25rem;
    }

    section[data-testid="stSidebar"] .stMarkdown {
        color: #cbd5e1;
    }

    .sidebar-logo {
        color: white;
        font-size: 1.15rem;
        font-weight: 700;
        padding: 0.75rem 0 1rem 0;
        border-bottom: 1px solid rgba(255,255,255,0.08);
        margin-bottom: 1rem;
    }

    .sidebar-nav-item {
        padding: 0.6rem 0.25rem;
        color: #cbd5e1;
        font-size: 0.9rem;
        border-bottom: 1px solid rgba(255,255,255,0.05);
    }

    /* Plain top bar, like the OneStack admin header */
    .top-bar {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 0.9rem 1.25rem;
        margin-bottom: 1.25rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    .top-bar-title {
        font-size: 1.05rem;
        font-weight: 600;
        color: #1f2937;
        margin: 0;
    }

    /* Dark bar section headers, like "Active Soundboxes - Last 7 Days" */
    .section-bar {
        background: #1f2937;
        color: white;
        font-weight: 600;
        font-size: 0.95rem;
        padding: 0.85rem 1.1rem;
        border-radius: 6px;
        margin: 1.25rem 0 0.75rem 0;
    }

    /* Plain content panel */
    .panel {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 1.25rem;
    }

    /* Buttons - simple, single color, no gradients */
    .stButton > button {
        background: #1f2937;
        color: white;
        border: none;
        padding: 0.6rem 1.5rem;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.9rem;
        width: 100%;
    }

    .stButton > button:hover {
        background: #111827;
    }

    /* File Uploader - simple bordered box */
    section[data-testid="stFileUploader"] {
        background: #fafafa;
        border-radius: 6px;
        padding: 1.25rem;
        border: 1px dashed #cbd5e1;
    }

    /* Metrics */
    div[data-testid="stMetric"] {
        background: white;
        padding: 1rem;
        border-radius: 6px;
        border: 1px solid #e2e8f0;
    }

    div[data-testid="stMetricLabel"] {
        color: #64748b;
        font-size: 0.8rem;
        font-weight: 600;
    }

    div[data-testid="stMetricValue"] {
        color: #1f2937;
        font-weight: 700;
        font-size: 1.5rem;
    }

    /* Tabs - simple, no gradient */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0;
        background: white;
        border-radius: 6px;
        border: 1px solid #e2e8f0;
        overflow: hidden;
    }

    .stTabs [data-baseweb="tab"] {
        padding: 0.85rem 1.5rem;
        font-weight: 600;
        color: #64748b;
        background: white;
        border-right: 1px solid #e2e8f0;
    }

    .stTabs [aria-selected="true"] {
        background: #1f2937;
        color: white;
    }

    .footer {
        text-align: center;
        color: #94a3b8;
        padding: 1.5rem;
        margin-top: 2rem;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# ALL CLEANING LOGIC BELOW IS UNCHANGED
# ---------------------------------------------------------------------------

def pre_process_dataframe(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """Pre-process dataframe before cleaning"""
    logs = []
    df = df.copy()

    source_file_variations = ["source_file", "sourcefile", "source file"]
    columns_to_drop = []

    for df_col in df.columns:
        df_col_normalized = df_col.lower().replace(" ", "").replace("_", "")
        if df_col_normalized in source_file_variations:
            columns_to_drop.append(df_col)

    if columns_to_drop:
        df = df.drop(columns=columns_to_drop)
        logs.append(f"✓ Deleted Source_File column(s): {', '.join(columns_to_drop)}")

    branch_exists = any(
        col.lower().replace(" ", "").replace("_", "") == "branchname"
        for col in df.columns
    )

    if not branch_exists:
        df["Branch Name"] = "HO Branch"
        logs.append("✓ Added 'Branch Name' column with default value 'HO Branch'")

    return df, logs

def clean_mobile_number(x) -> str:
    x = str(x).strip()
    x = re.sub(r"\D", "", x)
    if len(x) == 12 and x.startswith("91"):
        x = x[2:]
    return x

def format_date(x) -> str:
    if pd.isna(x) or str(x).strip() == "":
        return ""

    if isinstance(x, (int, float)) and not pd.isna(x):
        try:
            dt = pd.to_datetime("1899-12-30") + pd.to_timedelta(int(x), unit="D")
            return "'" + dt.strftime("%d-%m-%Y")
        except Exception:
            pass

    try:
        dt = pd.to_datetime(str(x), dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return str(x)
        return "'" + dt.strftime("%d-%m-%Y")
    except Exception:
        return str(x)

def format_aadhaar(x) -> str:
    if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
        return ""

    x_str = str(x).strip().lstrip("'")
    x_str = re.sub(r'[^0-9]', '', x_str)

    if not x_str:
        return ""

    return "'" + x_str

def format_account_number(x) -> str:
    if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
        return ""

    x_str = str(x).strip().lstrip("'")

    if x_str.endswith('.0'):
        x_str = x_str[:-2]

    return "'" + x_str

def clean_address(x) -> str:
    if pd.isna(x) or str(x).strip() == "":
        return ""

    x_str = str(x)
    special_chars = [',', '.', '/', '&', '-', '"', ';', '(', ')', '\\']
    for char in special_chars:
        x_str = x_str.replace(char, ' ')

    x_str = re.sub(r'\s+', ' ', x_str)
    return x_str.strip()

def clean_name(x) -> str:
    if pd.isna(x) or str(x).strip() in ["", "nan", "NaN", "None"]:
        return ""

    x_str = str(x)
    special_chars = ['-', '/', ':', '|', '(', ')', '&', '#', ',', '.', ';', "'"]
    for char in special_chars:
        x_str = x_str.replace(char, ' ')

    x_str = re.sub(r'\s+', ' ', x_str).strip()
    return x_str

def clean_data(df: pd.DataFrame, source_file: Optional[str] = None) -> Tuple[pd.DataFrame, List[str]]:
    """Main data cleaning function"""
    logs = []
    df = df.copy()

    try:
        df, pre_logs = pre_process_dataframe(df)
        logs.extend(pre_logs)

        if "Mobile No" in df.columns:
            before = len(df)
            df = df.drop_duplicates(subset=["Mobile No"], keep="first")
            after = len(df)
            if before > after:
                logs.append(f"✓ Removed {before - after} duplicate mobile numbers")

        if "Mobile No" in df.columns:
            df["Mobile No"] = df["Mobile No"].apply(clean_mobile_number)
            logs.append("✓ Cleaned mobile numbers")

        date_columns = ["DOB", "DOI", "Account Opening Date"]
        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].apply(format_date)
                logs.append(f"✓ Formatted date column: {col}")

        aadhaar_columns = ["Aadhar No", "Aadhaar No"]
        for col in aadhaar_columns:
            if col in df.columns:
                df[col] = df[col].apply(format_aadhaar)
                logs.append(f"✓ Formatted Aadhaar column: {col}")

        if "Address Line 1" in df.columns:
            df["Address Line 1"] = df["Address Line 1"].apply(clean_address)
            logs.append("✓ Cleaned Address Line 1")

        if "Account No" in df.columns:
            df["Account No"] = df["Account No"].apply(format_account_number)
            logs.append("✓ Formatted Account No")

        if "Branch Name" in df.columns:
            df["Branch Name"] = "HO Branch"
            logs.append("✓ Set Branch Name to 'HO Branch'")

        name_columns = ["First Name", "Middle Name", "Last Name", "Entity Name", "Account Holder Name"]
        for col in name_columns:
            if col in df.columns:
                df[col] = df[col].apply(clean_name)
        logs.append("✓ Cleaned name columns")

        if "Entity Name" in df.columns:
            entity_mask = df["Entity Name"].notna() & (df["Entity Name"].str.strip() != "")
            personal_name_cols = ["First Name", "Middle Name", "Last Name"]
            for col in personal_name_cols:
                if col in df.columns:
                    df.loc[entity_mask, col] = ""
            logs.append("✓ Cleared personal names where Entity Name present")

        if "Account Holder Name" in df.columns and "Entity Name" in df.columns:
            mask = (df["Account Holder Name"].isna()) | (df["Account Holder Name"].str.strip() == "")
            entity_mask = (df["Entity Name"].notna()) & (df["Entity Name"].str.strip() != "")
            df.loc[mask & entity_mask, "Account Holder Name"] = df.loc[mask & entity_mask, "Entity Name"]
            logs.append("✓ Filled Account Holder Names from Entity Name")

        if "Address Line 1" in df.columns and "Address Line 2" in df.columns:
            mask = (df["Address Line 2"].isna()) | (df["Address Line 2"].str.strip() == "")
            has_addr1 = (df["Address Line 1"].notna()) & (df["Address Line 1"].str.strip() != "")
            df.loc[mask & has_addr1, "Address Line 2"] = df.loc[mask & has_addr1, "Address Line 1"]
            logs.append("✓ Copied Address Line 1 to Address Line 2")

        clear_cols = [
            "Turnover Type", "Acceptance Type", "Ownership Type", "MCC",
            "Email ID", "Source_File", "Bank Cust ID", "State Code (GST)",
            "Latitude", "Longitude", "District"
        ]

        for col in clear_cols:
            col_normalized = col.lower().replace(" ", "").replace("_", "")
            for df_col in df.columns:
                df_col_normalized = df_col.lower().replace(" ", "").replace("_", "")
                if df_col_normalized == col_normalized:
                    df[df_col] = ""
                    logs.append(f"✓ Cleared: {df_col}")

        return df, logs

    except Exception as e:
        logger.error(f"Error in clean_data: {str(e)}")
        logs.append(f"❌ Error: {str(e)}")
        return df, logs

def add_dropdowns(buffer: io.BytesIO, sheet_name: str = "Cleaned") -> io.BytesIO:
    try:
        buffer.seek(0)
        wb = load_workbook(buffer)
        ws = wb[sheet_name]

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output
    except Exception as e:
        logger.error(f"Error adding dropdowns: {str(e)}")
        return buffer

def load_excel(file) -> Optional[pd.DataFrame]:
    file_extension = file.name.split('.')[-1].lower()

    try:
        if file_extension == 'csv':
            df = pd.read_csv(file, dtype=str)
        else:
            df = pd.read_excel(file, dtype=str)

        logger.info(f"Successfully loaded file: {file.name}")
        return df
    except Exception as e:
        logger.error(f"Error loading file {file.name}: {str(e)}")
        return None

def extract_bank_name_from_filename(filename: str) -> str:
    name_without_ext = os.path.splitext(filename)[0]
    return name_without_ext.strip()

def get_expected_columns() -> List[str]:
    return [
        "Mobile No", "First Name", "Middle Name", "Last Name",
        "Entity Name", "Account Holder Name", "DOB", "DOI",
        "Account Opening Date", "Aadhar No", "Aadhaar No",
        "Account No", "Address Line 1", "Address Line 2",
        "Branch Name"
    ]

def check_column_match(df: pd.DataFrame, expected_cols: List[str]) -> Tuple[bool, List[str], List[str]]:
    df_cols_normalized = [col.lower().replace(" ", "").replace("_", "") for col in df.columns]
    expected_cols_normalized = [col.lower().replace(" ", "").replace("_", "") for col in expected_cols]

    missing = []
    for exp_col in expected_cols:
        exp_normalized = exp_col.lower().replace(" ", "").replace("_", "")
        if exp_normalized not in df_cols_normalized:
            missing.append(exp_col)

    extra = []
    for df_col in df.columns:
        df_normalized = df_col.lower().replace(" ", "").replace("_", "")
        if df_normalized not in expected_cols_normalized:
            extra.append(df_col)

    has_mobile = any(col.lower().replace(" ", "").replace("_", "") == "mobileno" for col in df.columns)

    return has_mobile, missing, extra

def create_mismatch_report(mismatch_files: List[Dict]) -> pd.DataFrame:
    if not mismatch_files:
        return pd.DataFrame()

    report_data = []
    for item in mismatch_files:
        report_data.append({
            'File Name': item['filename'],
            'Bank Name': item['bank_name'],
            'Error/Issue': item['error'],
            'Missing Columns': item['missing_columns'],
            'Extra Columns': item['extra_columns'],
            'All Columns in File': item['all_columns']
        })

    return pd.DataFrame(report_data)

# ---------------------------------------------------------------------------
# UI COMPONENTS (simplified, English Creator module removed)
# ---------------------------------------------------------------------------

def render_sidebar():
    with st.sidebar:
        st.markdown('<div class="sidebar-logo">📊 QR Cleaner Pro</div>', unsafe_allow_html=True)

        st.markdown("""
        <div class='sidebar-nav-item'>✅ Single / Multiple File Cleaning</div>
        <div class='sidebar-nav-item'>✅ Bulk 350+ Files (Master)</div>
        <div class='sidebar-nav-item'>✅ Bank Name Column Addition</div>
        <div class='sidebar-nav-item'>✅ Mismatch Report Generation</div>
        <div class='sidebar-nav-item'>✅ Column Validation</div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown(
            "<div style='color:#94a3b8; font-size:0.75rem; text-align:center;'>Version 3.0<br>Updated: February 2026</div>",
            unsafe_allow_html=True
        )

def render_data_cleaner_tab():
    st.markdown('<div class="section-bar">📁 Upload Excel Files</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Select one or multiple files (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="file_uploader_normal"
    )

    if uploaded_files:
        st.caption(f"{len(uploaded_files)} file(s) selected")

    st.button("🧹 Clean & Process Files", key="process_normal",
              on_click=lambda: st.session_state.update({"_run_normal": True}))
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_files and st.session_state.get("_run_normal"):
        st.session_state["_run_normal"] = False
        process_files(uploaded_files)

def render_bulk_processor_tab():
    st.markdown('<div class="section-bar">📂 Bulk Master File Creator</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.caption("Upload all your files at once. Bank name is taken from each file name. "
               "Everything is cleaned and merged into one master file, plus a mismatch report for problem files.")

    bulk_files = st.file_uploader(
        "Select ALL files from your folder (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="file_uploader_bulk"
    )

    if bulk_files:
        st.caption(f"{len(bulk_files)} file(s) selected")

    st.button("🧹 Create Master File", key="process_bulk",
              on_click=lambda: st.session_state.update({"_run_bulk": True}))
    st.markdown('</div>', unsafe_allow_html=True)

    if bulk_files and st.session_state.get("_run_bulk"):
        st.session_state["_run_bulk"] = False
        process_bulk_to_master(bulk_files)

def process_files(uploaded_files):
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        all_logs = []
        total_steps = len(uploaded_files) + 2
        current_step = 0

        if len(uploaded_files) == 1:
            status_text.text("Loading file...")
            df = load_excel(uploaded_files[0])
            current_step += 1
            progress_bar.progress(current_step / total_steps)

            if df is None:
                st.error("❌ Failed to load file")
                return

            status_text.text("Cleaning data...")
            cleaned_df, logs = clean_data(df, uploaded_files[0].name)
            all_logs.extend(logs)
            current_step += 1
            progress_bar.progress(current_step / total_steps)

            status_text.text("Creating Excel file...")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                cleaned_df.to_excel(writer, index=False, sheet_name="Cleaned")

            final_output = add_dropdowns(output, sheet_name="Cleaned")
            current_step += 1
            progress_bar.progress(1.0)

            status_text.text("✅ Processing complete!")
            st.success("✅ File processed successfully!")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", len(cleaned_df))
            with col2:
                st.metric("Total Columns", len(cleaned_df.columns))
            with col3:
                st.metric("Operations", len(all_logs))

            st.download_button(
                "⬇️ Download Cleaned File",
                data=final_output.getvalue(),
                file_name=f"Cleaned_{uploaded_files[0].name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            all_dfs = []

            for idx, file in enumerate(uploaded_files):
                status_text.text(f"Processing {idx + 1}/{len(uploaded_files)}: {file.name}")
                df = load_excel(file)

                if df is not None:
                    cleaned_df, logs = clean_data(df, file.name)
                    all_dfs.append(cleaned_df)
                    all_logs.extend(logs)

                current_step += 1
                progress_bar.progress(current_step / total_steps)

            if not all_dfs:
                st.error("❌ No files could be loaded")
                return

            status_text.text("Merging files...")
            merged_df = pd.concat(all_dfs, ignore_index=True, sort=False)
            current_step += 1
            progress_bar.progress(current_step / total_steps)

            if "Mobile No" in merged_df.columns:
                before = len(merged_df)
                merged_df = merged_df.drop_duplicates(subset=["Mobile No"], keep="first")
                after = len(merged_df)
                if before > after:
                    all_logs.append(f"✓ Removed {before - after} duplicates from merged data")

            status_text.text("Creating merged Excel file...")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                merged_df.to_excel(writer, index=False, sheet_name="Cleaned_Merged")

            final_output = add_dropdowns(output, sheet_name="Cleaned_Merged")
            current_step += 1
            progress_bar.progress(1.0)

            status_text.text("✅ Processing complete!")
            st.success("✅ Files processed and merged successfully!")

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Files Merged", len(uploaded_files))
            with col2:
                st.metric("Total Rows", len(merged_df))
            with col3:
                st.metric("Columns", len(merged_df.columns))
            with col4:
                st.metric("Operations", len(all_logs))

            st.download_button(
                "⬇️ Download Merged File",
                data=final_output.getvalue(),
                file_name="Cleaned_Merged.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with st.expander("📝 View Cleaning Logs"):
            for log in all_logs:
                st.write(log)

        progress_bar.empty()
        status_text.empty()

    except Exception as e:
        logger.error(f"Error processing files: {str(e)}")
        st.error(f"❌ Error: {str(e)}")
        progress_bar.empty()
        status_text.empty()

def process_bulk_to_master(uploaded_files):
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        status_text.text("🔄 Starting bulk processing...")

        results = {
            'master_data': None,
            'mismatch_files': [],
            'statistics': {
                'total_files': len(uploaded_files),
                'successful': 0,
                'failed': 0,
                'total_rows_processed': 0,
                'unique_banks': 0
            }
        }

        all_cleaned_dfs = []
        expected_cols = get_expected_columns()
        bank_names_set = set()

        for idx, file in enumerate(uploaded_files):
            progress = (idx + 1) / len(uploaded_files)
            progress_bar.progress(progress)
            status_text.text(f"Processing {idx + 1}/{len(uploaded_files)}: {file.name}")

            try:
                bank_name = extract_bank_name_from_filename(file.name)
                bank_names_set.add(bank_name)

                df = load_excel(file)
                if df is None:
                    results['mismatch_files'].append({
                        'filename': file.name,
                        'bank_name': bank_name,
                        'error': 'Failed to load file',
                        'missing_columns': 'N/A',
                        'extra_columns': 'N/A',
                        'all_columns': 'N/A'
                    })
                    results['statistics']['failed'] += 1
                    continue

                has_mobile, missing_cols, extra_cols = check_column_match(df, expected_cols)

                if not has_mobile:
                    results['mismatch_files'].append({
                        'filename': file.name,
                        'bank_name': bank_name,
                        'error': 'Missing required column: Mobile No',
                        'missing_columns': ', '.join(missing_cols) if missing_cols else 'None',
                        'extra_columns': ', '.join(extra_cols) if extra_cols else 'None',
                        'all_columns': ', '.join(df.columns.tolist())
                    })
                    results['statistics']['failed'] += 1
                    continue

                cleaned_df, logs = clean_data(df, file.name)
                cleaned_df.insert(0, 'Bank Name', bank_name)

                all_cleaned_dfs.append(cleaned_df)

                results['statistics']['successful'] += 1
                results['statistics']['total_rows_processed'] += len(cleaned_df)

                if missing_cols or extra_cols:
                    results['mismatch_files'].append({
                        'filename': file.name,
                        'bank_name': bank_name,
                        'error': 'Column mismatch (but processed successfully)',
                        'missing_columns': ', '.join(missing_cols) if missing_cols else 'None',
                        'extra_columns': ', '.join(extra_cols) if extra_cols else 'None',
                        'all_columns': ', '.join(df.columns.tolist())
                    })

            except Exception as e:
                bank_name = extract_bank_name_from_filename(file.name)
                results['mismatch_files'].append({
                    'filename': file.name,
                    'bank_name': bank_name,
                    'error': f'Error: {str(e)}',
                    'missing_columns': 'N/A',
                    'extra_columns': 'N/A',
                    'all_columns': 'N/A'
                })
                results['statistics']['failed'] += 1

        if all_cleaned_dfs:
            status_text.text("Creating master file...")
            master_df = pd.concat(all_cleaned_dfs, ignore_index=True, sort=False)

            if "Mobile No" in master_df.columns:
                before = len(master_df)
                master_df = master_df.drop_duplicates(subset=["Mobile No"], keep="first")
                after = len(master_df)
                if before > after:
                    results['statistics']['duplicates_removed'] = before - after

            results['master_data'] = master_df
            results['statistics']['unique_banks'] = len(bank_names_set)

        progress_bar.progress(1.0)
        status_text.text("✅ Complete!")

        st.success("🎉 Master file created successfully!")

        st.markdown('<div class="section-bar">📊 Processing Summary</div>', unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Total Files", results['statistics']['total_files'])
        with col2:
            st.metric("✅ Successful", results['statistics']['successful'])
        with col3:
            st.metric("❌ Failed", results['statistics']['failed'])
        with col4:
            st.metric("🏦 Unique Banks", results['statistics']['unique_banks'])

        col5, col6, col7, col8 = st.columns(4)
        with col5:
            st.metric("Total Rows", results['statistics']['total_rows_processed'])
        with col6:
            if 'duplicates_removed' in results['statistics']:
                st.metric("Duplicates Removed", results['statistics']['duplicates_removed'])
        with col7:
            if results['master_data'] is not None:
                st.metric("Final Rows", len(results['master_data']))
        with col8:
            if results['master_data'] is not None:
                st.metric("Columns", len(results['master_data'].columns))

        if results['master_data'] is not None:
            st.markdown('<div class="section-bar">📄 Master File Preview</div>', unsafe_allow_html=True)
            st.dataframe(results['master_data'].head(10), use_container_width=True)

            st.markdown('<div class="section-bar">🏦 Bank Distribution</div>', unsafe_allow_html=True)
            bank_counts = results['master_data']['Bank Name'].value_counts()
            st.dataframe(
                pd.DataFrame({
                    'Bank Name': bank_counts.index,
                    'Row Count': bank_counts.values
                }),
                use_container_width=True
            )

        if results['mismatch_files']:
            st.warning(f"⚠️ {len(results['mismatch_files'])} files have issues")
            with st.expander("📋 View Mismatch Report"):
                mismatch_df = create_mismatch_report(results['mismatch_files'])
                st.dataframe(mismatch_df, use_container_width=True)

        st.markdown('<div class="section-bar">📥 Download Files</div>', unsafe_allow_html=True)
        col_dl1, col_dl2 = st.columns(2)

        if results['master_data'] is not None:
            with col_dl1:
                master_buffer = io.BytesIO()
                with pd.ExcelWriter(master_buffer, engine='openpyxl') as writer:
                    results['master_data'].to_excel(writer, index=False, sheet_name='All_Banks_Master')

                st.download_button(
                    "⬇️ Download Master File",
                    data=master_buffer.getvalue(),
                    file_name=f"All_Banks_Master_Cleaned_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        if results['mismatch_files']:
            with col_dl2:
                mismatch_df = create_mismatch_report(results['mismatch_files'])
                mismatch_buffer = io.BytesIO()
                with pd.ExcelWriter(mismatch_buffer, engine='openpyxl') as writer:
                    mismatch_df.to_excel(writer, index=False, sheet_name='Mismatch_Report')

                st.download_button(
                    "⬇️ Download Mismatch Report",
                    data=mismatch_buffer.getvalue(),
                    file_name=f"Mismatch_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        progress_bar.empty()
        status_text.empty()

    except Exception as e:
        logger.error(f"Error: {str(e)}")
        st.error(f"❌ Error: {str(e)}")
        progress_bar.empty()
        status_text.empty()

def main():
    st.markdown("""
    <div class="top-bar">
        <p class="top-bar-title">📊 QR Data Cleaner Pro</p>
    </div>
    """, unsafe_allow_html=True)

    render_sidebar()

    tab1, tab2 = st.tabs(["📁 QR Data Cleaner", "📂 Bulk Master Creator"])

    with tab1:
        render_data_cleaner_tab()

    with tab2:
        render_bulk_processor_tab()

    st.markdown("""
    <div class="footer">
        <p>QR Data Cleaner Pro v3.0 | © 2026</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
