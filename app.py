import streamlit as st
import pandas as pd
import re
import io
import requests
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from typing import List, Tuple, Optional, Dict
import logging
import os

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page configuration
st.set_page_config(
    page_title="QR Data Cleaner Pro",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS - Enhanced PhonePe/Razorpay style
st.markdown("""
<style>
    /* Main background */
    .stApp {
        background: #f8f9fa;
    }
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background: #1f2937;
        padding-top: 2rem;
    }
    
    section[data-testid="stSidebar"] > div {
        padding: 1.5rem;
    }
    
    /* Header styling */
    .main-header {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin-bottom: 2rem;
    }
    
    .main-title {
        color: #1f2937;
        font-size: 2rem;
        font-weight: 700;
        margin: 0;
    }
    
    .subtitle {
        color: #6b7280;
        font-size: 0.95rem;
        margin-top: 0.5rem;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #5f72bd 0%, #9921e8 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 10px;
        font-weight: 600;
        width: 100%;
        box-shadow: 0 4px 12px rgba(95, 114, 189, 0.3);
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        box-shadow: 0 6px 16px rgba(95, 114, 189, 0.4);
        transform: translateY(-2px);
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: white;
        padding: 0.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        color: #64748b;
        background: transparent;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #5f72bd 0%, #9921e8 100%);
        color: white;
    }
    
    /* Card styling */
    div[data-testid="stExpander"] {
        background: white;
        border-radius: 12px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    
    /* Text area styling */
    .stTextArea textarea {
        border-radius: 10px;
        border: 2px solid #e5e7eb;
        background: white;
        font-family: 'Courier New', monospace;
    }
    
    .stTextArea textarea:focus {
        border-color: #5f72bd;
        box-shadow: 0 0 0 3px rgba(95, 114, 189, 0.1);
    }
    
    /* File uploader */
    section[data-testid="stFileUploader"] {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        border: 2px dashed #e5e7eb;
    }
    
    /* Metrics */
    div[data-testid="stMetricValue"] {
        color: #1f2937;
        font-weight: 700;
    }
    
    /* Info/Success boxes */
    .stAlert {
        border-radius: 10px;
        border: none;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(135deg, #5f72bd 0%, #9921e8 100%);
    }
</style>
""", unsafe_allow_html=True)


# ============= DATA CLEANING FUNCTIONS (SAME AS BEFORE - NO CHANGES) =============

def pre_process_dataframe(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """Pre-process dataframe before cleaning"""
    logs = []
    df = df.copy()
    
    # Delete Source_File column
    source_file_variations = ["source_file", "sourcefile", "source file"]
    columns_to_drop = []
    
    for df_col in df.columns:
        df_col_normalized = df_col.lower().replace(" ", "").replace("_", "")
        if df_col_normalized in source_file_variations:
            columns_to_drop.append(df_col)
    
    if columns_to_drop:
        df = df.drop(columns=columns_to_drop)
        logs.append(f"✓ Deleted Source_File column(s): {', '.join(columns_to_drop)}")
    
    # Add Branch Name column if doesn't exist
    branch_exists = any(
        col.lower().replace(" ", "").replace("_", "") == "branchname" 
        for col in df.columns
    )
    
    if not branch_exists:
        df["Branch Name"] = "HO Branch"
        logs.append("✓ Added 'Branch Name' column with default value 'HO Branch'")
    
    return df, logs


def clean_mobile_number(x) -> str:
    """Clean and format mobile number"""
    x = str(x).strip()
    x = re.sub(r"\D", "", x)
    if len(x) == 12 and x.startswith("91"):
        x = x[2:]
    return x


def format_date(x) -> str:
    """Format date to dd-mm-yyyy format"""
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
    """Format Aadhaar number"""
    if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
        return ""
    
    x_str = str(x).strip().lstrip("'")
    x_str = re.sub(r'[^0-9]', '', x_str)
    
    if not x_str:
        return ""
    
    return "'" + x_str


def format_account_number(x) -> str:
    """Format account number"""
    if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
        return ""
    
    x_str = str(x).strip().lstrip("'")
    
    if x_str.endswith('.0'):
        x_str = x_str[:-2]
    
    return "'" + x_str


def clean_address(x) -> str:
    """Clean address"""
    if pd.isna(x) or str(x).strip() == "":
        return ""
    
    x_str = str(x)
    special_chars = [',', '.', '/', '&', '-', '"', ';', '(', ')', '\\']
    for char in special_chars:
        x_str = x_str.replace(char, ' ')
    
    x_str = re.sub(r'\s+', ' ', x_str)
    return x_str.strip()


def clean_name(x) -> str:
    """Clean name"""
    if pd.isna(x) or str(x).strip() in ["", "nan", "NaN", "None"]:
        return ""
    
    x_str = str(x)
    special_chars = ['-', '/', ':', '|', '(', ')', '&', '#', ',', '.', ';', "'"]
    for char in special_chars:
        x_str = x_str.replace(char, ' ')
    
    x_str = re.sub(r'\s+', ' ', x_str).strip()
    return x_str


def clean_data(df: pd.DataFrame, source_file: Optional[str] = None) -> Tuple[pd.DataFrame, List[str]]:
    """Main data cleaning function - EXACT SAME AS BEFORE"""
    logs = []
    df = df.copy()
    
    try:
        # Pre-processing
        df, pre_logs = pre_process_dataframe(df)
        logs.extend(pre_logs)
        
        # Remove duplicates by Mobile No
        if "Mobile No" in df.columns:
            before = len(df)
            df = df.drop_duplicates(subset=["Mobile No"], keep="first")
            after = len(df)
            if before > after:
                logs.append(f"✓ Removed {before - after} duplicate mobile numbers")
        
        # Clean mobile numbers
        if "Mobile No" in df.columns:
            df["Mobile No"] = df["Mobile No"].apply(clean_mobile_number)
            logs.append("✓ Cleaned mobile numbers")
        
        # Format date columns
        date_columns = ["DOB", "DOI", "Account Opening Date"]
        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].apply(format_date)
                logs.append(f"✓ Formatted date column: {col}")
        
        # Format Aadhaar columns
        aadhaar_columns = ["Aadhar No", "Aadhaar No"]
        for col in aadhaar_columns:
            if col in df.columns:
                df[col] = df[col].apply(format_aadhaar)
                logs.append(f"✓ Formatted Aadhaar column: {col}")
        
        # Clean Address Line 1
        if "Address Line 1" in df.columns:
            df["Address Line 1"] = df["Address Line 1"].apply(clean_address)
            logs.append("✓ Cleaned Address Line 1")
        
        # Format Account No
        if "Account No" in df.columns:
            df["Account No"] = df["Account No"].apply(format_account_number)
            logs.append("✓ Formatted Account No")
        
        # Replace Branch Name values
        if "Branch Name" in df.columns:
            df["Branch Name"] = "HO Branch"
            logs.append("✓ Set Branch Name to 'HO Branch'")
        
        # Clean name columns
        name_columns = ["First Name", "Middle Name", "Last Name", "Entity Name", "Account Holder Name"]
        for col in name_columns:
            if col in df.columns:
                df[col] = df[col].apply(clean_name)
        logs.append("✓ Cleaned name columns")
        
        # Clear personal names if entity present
        if "Entity Name" in df.columns:
            entity_mask = df["Entity Name"].notna() & (df["Entity Name"].str.strip() != "")
            personal_name_cols = ["First Name", "Middle Name", "Last Name"]
            for col in personal_name_cols:
                if col in df.columns:
                    df.loc[entity_mask, col] = ""
            logs.append("✓ Cleared personal names where Entity Name present")
        
        # Account Holder Name fallback
        if "Account Holder Name" in df.columns and "Entity Name" in df.columns:
            mask = (df["Account Holder Name"].isna()) | (df["Account Holder Name"].str.strip() == "")
            entity_mask = (df["Entity Name"].notna()) & (df["Entity Name"].str.strip() != "")
            df.loc[mask & entity_mask, "Account Holder Name"] = df.loc[mask & entity_mask, "Entity Name"]
            logs.append("✓ Filled Account Holder Names from Entity Name")
        
        # Address Line 2 fallback
        if "Address Line 1" in df.columns and "Address Line 2" in df.columns:
            mask = (df["Address Line 2"].isna()) | (df["Address Line 2"].str.strip() == "")
            has_addr1 = (df["Address Line 1"].notna()) & (df["Address Line 1"].str.strip() != "")
            df.loc[mask & has_addr1, "Address Line 2"] = df.loc[mask & has_addr1, "Address Line 1"]
            logs.append("✓ Copied Address Line 1 to Address Line 2")
        
        # Clear unwanted columns
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
    """Add dropdown validations to Excel file"""
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
    """Load Excel/CSV file"""
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


# ============= BULK FOLDER PROCESSING FUNCTIONS (NEW) =============

def extract_bank_name_from_filename(filename: str) -> str:
    """
    Extract bank name from filename
    Examples: 
    - "Bharat Bank.xlsx" -> "Bharat Bank"
    - "HDFC Bank LTD.csv" -> "HDFC Bank LTD"
    """
    name_without_ext = os.path.splitext(filename)[0]
    return name_without_ext.strip()


def get_expected_columns() -> List[str]:
    """Define expected columns for validation"""
    return [
        "Mobile No", "First Name", "Middle Name", "Last Name",
        "Entity Name", "Account Holder Name", "DOB", "DOI",
        "Account Opening Date", "Aadhar No", "Aadhaar No",
        "Account No", "Address Line 1", "Address Line 2",
        "Branch Name"
    ]


def check_column_match(df: pd.DataFrame, expected_cols: List[str]) -> Tuple[bool, List[str], List[str]]:
    """
    Check if DataFrame columns match expected columns
    
    Returns:
        Tuple of (has_mobile_no, missing_columns, extra_columns)
    """
    df_cols_normalized = [col.lower().replace(" ", "").replace("_", "") for col in df.columns]
    expected_cols_normalized = [col.lower().replace(" ", "").replace("_", "") for col in expected_cols]
    
    # Find missing columns
    missing = []
    for exp_col in expected_cols:
        exp_normalized = exp_col.lower().replace(" ", "").replace("_", "")
        if exp_normalized not in df_cols_normalized:
            missing.append(exp_col)
    
    # Find extra columns
    extra = []
    for df_col in df.columns:
        df_normalized = df_col.lower().replace(" ", "").replace("_", "")
        if df_normalized not in expected_cols_normalized:
            extra.append(df_col)
    
    # Check if has Mobile No (required)
    has_mobile = any(col.lower().replace(" ", "").replace("_", "") == "mobileno" for col in df.columns)
    
    return has_mobile, missing, extra


def process_bulk_files_to_master(uploaded_files) -> Dict:
    """
    Process all files and create ONE master file with all banks
    
    Returns:
        Dictionary with master dataframe and mismatch report
    """
    results = {
        'master_data': None,  # Single DataFrame with all banks
        'mismatch_files': [],  # List of files with column mismatches
        'statistics': {
            'total_files': len(uploaded_files),
            'successful': 0,
            'failed': 0,
            'total_rows_processed': 0,
            'unique_banks': 0
        }
    }
    
    all_cleaned_dfs = []  # Collect all cleaned dataframes
    expected_cols = get_expected_columns()
    bank_names_set = set()
    
    for file in uploaded_files:
        try:
            # Extract bank name from filename
            bank_name = extract_bank_name_from_filename(file.name)
            bank_names_set.add(bank_name)
            
            # Load file
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
            
            # Check column match
            has_mobile, missing_cols, extra_cols = check_column_match(df, expected_cols)
            
            if not has_mobile:
                # Cannot process - no Mobile No column
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
            
            # Clean data
            cleaned_df, logs = clean_data(df, file.name)
            
            # Add Bank Name column as FIRST column
            cleaned_df.insert(0, 'Bank Name', bank_name)
            
            # Add to master list
            all_cleaned_dfs.append(cleaned_df)
            
            results['statistics']['successful'] += 1
            results['statistics']['total_rows_processed'] += len(cleaned_df)
            
            # If column mismatch (but processed), add to mismatch report
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
    
    # Merge all cleaned dataframes into ONE master file
    if all_cleaned_dfs:
        # Concatenate all dataframes
        master_df = pd.concat(all_cleaned_dfs, ignore_index=True, sort=False)
        
        # Remove duplicates across all banks by Mobile No
        if "Mobile No" in master_df.columns:
            before = len(master_df)
            master_df = master_df.drop_duplicates(subset=["Mobile No"], keep="first")
            after = len(master_df)
            if before > after:
                removed = before - after
                results['statistics']['duplicates_removed'] = removed
        
        results['master_data'] = master_df
        results['statistics']['unique_banks'] = len(bank_names_set)
    
    return results


def create_mismatch_report(mismatch_files: List[Dict]) -> pd.DataFrame:
    """Create mismatch report DataFrame"""
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


# ============= TRANSLATION FUNCTIONS (SAME AS BEFORE - NO CHANGES) =============

def translate_to_english(text: str) -> dict:
    """Translate to professional English"""
    try:
        url = "https://api.mymemory.translated.net/get"
        params = {'q': text, 'langpair': 'hi|en'}
        
        response = requests.get(url, params=params, timeout=10)
        data = response.json()
        
        if response.status_code == 200 and 'responseData' in data:
            base_text = data['responseData']['translatedText'].strip()
            
            options = {
                'option1': generate_simple_professional(base_text),
                'option2': generate_polite_formal(base_text),
                'option3': generate_crisp_professional(base_text)
            }
            
            return options
        else:
            raise Exception("Translation API returned unexpected response")
    
    except Exception as e:
        logger.error(f"Translation error: {str(e)}")
        raise


def generate_simple_professional(text: str) -> str:
    """Generate simple professional version"""
    text = text.strip()
    if text and not text[0].isupper():
        text = text.capitalize()
    if text and not text.endswith(('.', '!', '?')):
        text += '.'
    if text and not text.lower().startswith(('hi', 'hello')):
        text = "Hi, " + text[0].lower() + text[1:] if len(text) > 1 else text
    return text


def generate_polite_formal(text: str) -> str:
    """Generate polite formal version"""
    text = text.strip()
    if text and not text[0].isupper():
        text = text.capitalize()
    if text and not text.endswith(('.', '!', '?')):
        text += '.'
    if text and not text.lower().startswith('hello'):
        text = "Hello, " + text[0].lower() + text[1:] if len(text) > 1 else text
    if 'thank you' not in text.lower():
        text = text.rstrip('.') + '. Thank you.'
    return text


def generate_crisp_professional(text: str) -> str:
    """Generate crisp professional version"""
    text = text.strip()
    if text and not text[0].isupper():
        text = text.capitalize()
    if text and not text.endswith(('.', '!', '?')):
        text += '.'
    text = text.replace('I am ', "I'm ")
    text = text.replace('could you please', 'please')
    return text


# ============= UI COMPONENTS =============

def render_sidebar():
    """Render sidebar content"""
    with st.sidebar:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #5f72bd 0%, #9921e8 100%); 
                    padding: 1.5rem; 
                    border-radius: 15px; 
                    margin-bottom: 2rem;
                    text-align: center;'>
            <h2 style='color: white; margin: 0; font-size: 1.5rem; font-weight: 700;'>
                🔧 QR Cleaner Pro
            </h2>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.markdown("""
        <div style='background: #10b981; 
                    color: white; 
                    padding: 0.75rem 1rem; 
                    border-radius: 10px; 
                    text-align: center;
                    font-weight: 600;
                    margin-bottom: 1.5rem;'>
            🟢 System Active
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.markdown("""
        <div style='padding: 0.5rem 0;'>
            <h3 style='color: #e5e7eb; font-size: 1rem; font-weight: 700; margin-bottom: 1rem;'>
                Features
            </h3>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div style='background: rgba(255,255,255,0.1); 
                    padding: 1rem; 
                    border-radius: 12px;'>
            <div style='margin-bottom: 0.75rem; color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Single/Multiple files
            </div>
            <div style='margin-bottom: 0.75rem; color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Bulk 350+ files (Master)
            </div>
            <div style='margin-bottom: 0.75rem; color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Bank Name column
            </div>
            <div style='margin-bottom: 0.75rem; color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Mismatch report
            </div>
            <div style='margin-bottom: 0.75rem; color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Column validation
            </div>
            <div style='color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Hindi to English
            </div>
        </div>
        """, unsafe_allow_html=True)


def render_data_cleaner_tab():
    """Render the Data Cleaner tab - ORIGINAL FUNCTIONALITY (NO CHANGES)"""
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📁 Upload Excel Files")
        uploaded_files = st.file_uploader(
            "Select one or multiple files (.xlsx, .xls, .csv)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            help="Upload Excel or CSV files to clean and process",
            key="file_uploader_normal"
        )
    
    with col2:
        st.markdown("### 📊 Processing Info")
        if uploaded_files:
            st.metric("Files Uploaded", len(uploaded_files))
            total_size = sum([f.size for f in uploaded_files]) / 1024
            st.metric("Total Size", f"{total_size:.1f} KB")
        else:
            st.info("No files uploaded yet")
    
    if uploaded_files:
        st.markdown("---")
        
        with st.expander("📋 View Uploaded Files", expanded=True):
            for idx, file in enumerate(uploaded_files, 1):
                col_a, col_b, col_c = st.columns([3, 1, 1])
                with col_a:
                    st.write(f"**{idx}.** {file.name}")
                with col_b:
                    st.write(f"{file.size / 1024:.1f} KB")
                with col_c:
                    st.write("✅ Ready")
        
        st.markdown("---")
        
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 Clean & Process Files", use_container_width=True, type="primary", key="process_normal"):
                process_files(uploaded_files)
        
        with st.expander("🔍 Cleaning Operations", expanded=False):
            st.markdown("""
**Operations performed:**
1. ✅ Remove duplicate mobile numbers
2. ✅ Clean mobile numbers (remove '91' prefix)
3. ✅ Standardize dates to dd-mm-yyyy
4. ✅ Format Aadhaar numbers
5. ✅ Format Account numbers
6. ✅ Clean addresses and names
7. ✅ Add/update Branch Name
8. ✅ Merge multiple files (if applicable)
""")


def render_bulk_processor_tab():
    """Render the Bulk Folder Processor tab - NEW FUNCTIONALITY"""
    st.markdown("### 📂 Bulk Master File Creator")
    st.markdown("Process 350+ Excel files → Create ONE Master File with all banks")
    
    st.info("""
    💡 **How it works:**
    - Upload all your Excel files (350+) at once
    - Bank name extracted from filename (e.g., "HDFC Bank.xlsx" → Bank Name: "HDFC Bank")
    - All files cleaned and processed
    - **"Bank Name" column added as FIRST column**
    - **ONE master Excel file** with all 350 files merged
    - **ONE mismatch report** for files with column issues
    """)
    
    st.markdown("---")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("#### 📤 Upload All 350+ Excel Files")
        bulk_files = st.file_uploader(
            "Select ALL files from your folder (.xlsx, .xls, .csv)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            help="Bank name will be extracted from filename",
            key="file_uploader_bulk"
        )
    
    with col2:
        st.markdown("#### 📊 Upload Statistics")
        if bulk_files:
            st.metric("Files Uploaded", len(bulk_files))
            total_size_mb = sum([f.size for f in bulk_files]) / (1024 * 1024)
            st.metric("Total Size", f"{total_size_mb:.1f} MB")
            
            # Extract unique bank names
            unique_banks = set([extract_bank_name_from_filename(f.name) for f in bulk_files])
            st.metric("Unique Banks", len(unique_banks))
        else:
            st.info("📁 No files uploaded yet")
    
    if bulk_files:
        st.markdown("---")
        
        # Show sample of uploaded files
        with st.expander("📋 Preview Uploaded Files (First 10)", expanded=True):
            for idx, file in enumerate(bulk_files[:10], 1):
                bank_name = extract_bank_name_from_filename(file.name)
                col_a, col_b, col_c, col_d = st.columns([2, 2, 1, 1])
                with col_a:
                    st.write(f"**{idx}.** {file.name}")
                with col_b:
                    st.write(f"🏦 {bank_name}")
                with col_c:
                    st.write(f"{file.size / 1024:.1f} KB")
                with col_d:
                    st.write("✅")
            
            if len(bulk_files) > 10:
                st.write(f"... and **{len(bulk_files) - 10} more files**")
        
        st.markdown("---")
        
        # Processing button
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 Create Master File from All Banks", use_container_width=True, type="primary", key="process_bulk"):
                process_bulk_to_master(bulk_files)
        
        # Instructions
        with st.expander("ℹ️ What Will Be Created", expanded=False):
            st.markdown("""
**Output Files:**

1. **Master Cleaned File** (All_Banks_Master_Cleaned.xlsx)
   - ONE Excel file with ALL 350 files merged
   - "Bank Name" column as FIRST column
   - All data from all banks in one file
   - Example:
   ```
   | Bank Name    | Mobile No  | First Name | ... |
   |--------------|------------|------------|-----|
   | HDFC Bank    | 9876543210 | John       | ... |
   | HDFC Bank    | 9123456789 | Jane       | ... |
   | ICICI Bank   | 9988776655 | Bob        | ... |
   | Bharat Bank  | 9876541230 | Alice      | ... |
   ```

2. **Mismatch Report** (Mismatch_Report.xlsx)
   - Lists all files with column issues
   - Shows missing/extra columns
   - Shows all columns in each file
   - Only created if issues found

**Processing Details:**
- Files with "Mobile No" column: Processed ✅
- Files without "Mobile No": Listed in mismatch report ❌
- Column differences: Noted in mismatch report ⚠️
- Duplicates removed across ALL banks
""")


def process_files(uploaded_files):
    """Process uploaded files - ORIGINAL FUNCTIONALITY (NO CHANGES)"""
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
            st.balloons()
            
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
                status_text.text(f"Processing file {idx + 1}/{len(uploaded_files)}: {file.name}")
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
            st.success("✅ Multiple files processed and merged successfully!")
            st.balloons()
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Files Merged", len(uploaded_files))
            with col2:
                st.metric("Total Rows", len(merged_df))
            with col3:
                st.metric("Total Columns", len(merged_df.columns))
            with col4:
                st.metric("Operations", len(all_logs))
            
            st.download_button(
                "⬇️ Download Merged Cleaned File",
                data=final_output.getvalue(),
                file_name="Cleaned_Merged.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with st.expander("📝 View Detailed Cleaning Logs", expanded=False):
            for log in all_logs:
                st.write(log)
        
        progress_bar.empty()
        status_text.empty()
                    
    except Exception as e:
        logger.error(f"Error processing files: {str(e)}")
        st.error(f"❌ Error processing files: {str(e)}")
        progress_bar.empty()
        status_text.empty()


def process_bulk_to_master(uploaded_files):
    """Process bulk files to create ONE master file - NEW FUNCTIONALITY"""
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
                
                # Add Bank Name column as FIRST column
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
        
        # Merge all into ONE master file
        if all_cleaned_dfs:
            status_text.text("Creating master file with all banks...")
            master_df = pd.concat(all_cleaned_dfs, ignore_index=True, sort=False)
            
            # Remove duplicates across ALL banks
            if "Mobile No" in master_df.columns:
                before = len(master_df)
                master_df = master_df.drop_duplicates(subset=["Mobile No"], keep="first")
                after = len(master_df)
                if before > after:
                    results['statistics']['duplicates_removed'] = before - after
            
            results['master_data'] = master_df
            results['statistics']['unique_banks'] = len(bank_names_set)
        
        progress_bar.progress(1.0)
        status_text.text("✅ Processing complete!")
        
        st.success("🎉 Master file created successfully!")
        st.balloons()
        
        # Display statistics
        st.markdown("### 📊 Processing Summary")
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
        
        st.markdown("---")
        
        # Show preview of master file
        if results['master_data'] is not None:
            st.markdown("### 📄 Master File Preview")
            st.markdown("**First 10 rows of master file:**")
            st.dataframe(results['master_data'].head(10), use_container_width=True)
            
            # Show bank distribution
            st.markdown("### 🏦 Bank Distribution in Master File")
            bank_counts = results['master_data']['Bank Name'].value_counts()
            st.dataframe(
                pd.DataFrame({
                    'Bank Name': bank_counts.index,
                    'Row Count': bank_counts.values
                }),
                use_container_width=True
            )
        
        # Display mismatch report
        if results['mismatch_files']:
            st.markdown("---")
            st.warning(f"⚠️ {len(results['mismatch_files'])} files have issues or column mismatches")
            
            with st.expander("📋 View Mismatch Report Details", expanded=False):
                mismatch_df = create_mismatch_report(results['mismatch_files'])
                st.dataframe(mismatch_df, use_container_width=True)
        
        # Create download section
        st.markdown("---")
        st.markdown("### 📥 Download Files")
        
        col_dl1, col_dl2 = st.columns(2)
        
        # Download Master File
        if results['master_data'] is not None:
            with col_dl1:
                st.markdown("#### 📄 Master Cleaned File")
                master_buffer = io.BytesIO()
                with pd.ExcelWriter(master_buffer, engine='openpyxl') as writer:
                    results['master_data'].to_excel(writer, index=False, sheet_name='All_Banks_Master')
                
                st.download_button(
                    "⬇️ Download Master File (All Banks)",
                    data=master_buffer.getvalue(),
                    file_name=f"All_Banks_Master_Cleaned_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.info(f"""
                **Master File Contains:**
                - {len(results['master_data'])} total rows
                - {results['statistics']['unique_banks']} unique banks
                - Bank Name as first column
                - All 350 files merged into one
                """)
        
        # Download Mismatch Report
        if results['mismatch_files']:
            with col_dl2:
                st.markdown("#### 📋 Mismatch Report")
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
                
                st.warning(f"""
                **Mismatch Report Contains:**
                - {len(results['mismatch_files'])} files with issues
                - Column mismatch details
                - Error descriptions
                - All columns in each file
                """)
        
        progress_bar.empty()
        status_text.empty()
        
    except Exception as e:
        logger.error(f"Error in bulk processing: {str(e)}")
        st.error(f"❌ Error during bulk processing: {str(e)}")
        progress_bar.empty()
        status_text.empty()


def render_english_creator_tab():
    """Render the English Creator tab - ORIGINAL (NO CHANGES)"""
    st.markdown("### 🌐 Hindi/Hinglish to Professional English")
    st.markdown("Perfect for emails, tasks, and formal communication")
    
    if 'translation_results' not in st.session_state:
        st.session_state.translation_results = None
    
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        st.markdown("#### 📝 Input Text")
        input_text = st.text_area(
            "Enter your text (Hindi/English/Hinglish)",
            height=300,
            placeholder="Example:\nमुझे यह काम जल्दी चाहिए",
            key="input_text_translator"
        )
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            convert_button = st.button("✨ Convert to English", use_container_width=True, type="primary")
        with col_btn2:
            if st.button("🗑️ Clear All", use_container_width=True):
                st.session_state.translation_results = None
                st.rerun()
    
    with col_right:
        st.markdown("#### ✅ Professional Options")
        
        if convert_button and input_text.strip():
            with st.spinner("🔄 Converting..."):
                try:
                    results = translate_to_english(input_text)
                    st.session_state.translation_results = results
                    st.success("✅ Converted successfully!")
                except Exception as e:
                    st.error(f"❌ Translation failed: {str(e)}")
        
        if st.session_state.translation_results:
            results = st.session_state.translation_results
            
            st.markdown("**Option 1** (Simple & Professional)")
            st.text_area("", value=results['option1'], height=80, key="out1", label_visibility="collapsed")
            if st.button("📋 Copy Option 1", key="copy1", use_container_width=True):
                st.code(results['option1'], language=None)
            
            st.markdown("---")
            
            st.markdown("**Option 2** (Polite & Formal)")
            st.text_area("", value=results['option2'], height=80, key="out2", label_visibility="collapsed")
            if st.button("📋 Copy Option 2", key="copy2", use_container_width=True):
                st.code(results['option2'], language=None)
            
            st.markdown("---")
            
            st.markdown("**Option 3** (Crisp & Professional)")
            st.text_area("", value=results['option3'], height=80, key="out3", label_visibility="collapsed")
            if st.button("📋 Copy Option 3", key="copy3", use_container_width=True):
                st.code(results['option3'], language=None)
        else:
            st.info("👈 Enter text and click Convert")


# ============= MAIN APP =============

def main():
    """Main application entry point"""
    st.markdown("""
    <div class="main-header">
        <h1 class="main-title">🔧 QR Data Cleaner Pro</h1>
        <p class="subtitle">Clean, merge & standardize your QR code data with ease</p>
    </div>
    """, unsafe_allow_html=True)
    
    render_sidebar()
    
    # Main tabs - 3 tabs
    tab1, tab2, tab3 = st.tabs(["📁 QR Data Cleaner", "📂 Bulk Master Creator", "🌐 English Creator"])
    
    with tab1:
        render_data_cleaner_tab()
    
    with tab2:
        render_bulk_processor_tab()
    
    with tab3:
        render_english_creator_tab()
    
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #64748b; padding: 2rem;'>
        <p style='font-size: 0.9rem;'>Made with ❤️ for operations team | QR Data Cleaner Pro v3.0</p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
