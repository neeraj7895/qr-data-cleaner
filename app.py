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
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Professional Modern CSS - Inspired by OneStack Admin Panel
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global Styles */
    * {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }
    
    /* Main App Background */
    .stApp {
        background: #f5f7fa;
    }
    
    /* Sidebar Modern Design */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%);
        border-right: 1px solid rgba(255,255,255,0.1);
    }
    
    section[data-testid="stSidebar"] > div {
        padding: 1rem 1.5rem;
    }
    
    /* Sidebar Text Styling */
    section[data-testid="stSidebar"] .stMarkdown {
        color: #e2e8f0;
    }
    
    /* Header Card */
    .header-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem 2.5rem;
        border-radius: 16px;
        box-shadow: 0 10px 40px rgba(102, 126, 234, 0.25);
        margin-bottom: 2rem;
        border: 1px solid rgba(255,255,255,0.1);
    }
    
    .header-title {
        color: white;
        font-size: 2rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.02em;
    }
    
    .header-subtitle {
        color: rgba(255,255,255,0.9);
        font-size: 1rem;
        margin-top: 0.5rem;
        font-weight: 400;
    }
    
    /* Modern Card Design */
    .card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05), 0 1px 2px rgba(0,0,0,0.1);
        border: 1px solid #e2e8f0;
        margin-bottom: 1.5rem;
    }
    
    .card-title {
        font-size: 1.25rem;
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 1rem;
    }
    
    /* Stats Card */
    .stats-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border-left: 4px solid #667eea;
        transition: all 0.3s ease;
    }
    
    .stats-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    
    /* Button Styling */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 10px;
        font-weight: 600;
        font-size: 0.95rem;
        width: 100%;
        box-shadow: 0 4px 14px rgba(102, 126, 234, 0.4);
        transition: all 0.3s ease;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .stButton > button:hover {
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.5);
        transform: translateY(-2px);
    }
    
    .stButton > button:active {
        transform: translateY(0);
    }
    
    /* Secondary Button */
    .stButton > button[kind="secondary"] {
        background: white;
        color: #667eea;
        border: 2px solid #667eea;
        box-shadow: none;
    }
    
    /* Tab Styling - Professional */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0;
        background: white;
        padding: 0;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
        overflow: hidden;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 0;
        padding: 1rem 2rem;
        font-weight: 600;
        color: #64748b;
        background: white;
        border-right: 1px solid #e2e8f0;
        transition: all 0.2s ease;
    }
    
    .stTabs [data-baseweb="tab"]:last-child {
        border-right: none;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background: #f8fafc;
        color: #667eea;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    
    /* Expander Styling */
    div[data-testid="stExpander"] {
        background: white;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
    
    div[data-testid="stExpander"] summary {
        font-weight: 600;
        color: #1e293b;
        padding: 1rem;
    }
    
    /* File Uploader Modern Style */
    section[data-testid="stFileUploader"] {
        background: white;
        border-radius: 12px;
        padding: 2rem;
        border: 2px dashed #cbd5e1;
        transition: all 0.3s ease;
    }
    
    section[data-testid="stFileUploader"]:hover {
        border-color: #667eea;
        background: #f8fafc;
    }
    
    /* Metrics Styling */
    div[data-testid="stMetric"] {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
    }
    
    div[data-testid="stMetricLabel"] {
        color: #64748b;
        font-size: 0.875rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    div[data-testid="stMetricValue"] {
        color: #1e293b;
        font-weight: 700;
        font-size: 2rem;
    }
    
    /* Alert Boxes */
    .stAlert {
        border-radius: 12px;
        border: none;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    
    /* Success Alert */
    .stSuccess {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
    }
    
    /* Info Alert */
    .stInfo {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
    }
    
    /* Warning Alert */
    .stWarning {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        color: white;
    }
    
    /* Error Alert */
    .stError {
        background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
        color: white;
    }
    
    /* Progress Bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
    }
    
    /* Text Input */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea {
        border-radius: 10px;
        border: 2px solid #e2e8f0;
        padding: 0.75rem 1rem;
        font-size: 0.95rem;
        transition: all 0.2s ease;
    }
    
    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
    }
    
    /* Select Box */
    .stSelectbox > div > div {
        border-radius: 10px;
        border: 2px solid #e2e8f0;
    }
    
    /* DataFrame Styling */
    .dataframe {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    
    /* Table Header */
    .dataframe thead tr th {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 1rem !important;
        text-transform: uppercase;
        font-size: 0.85rem;
        letter-spacing: 0.5px;
    }
    
    /* Table Rows */
    .dataframe tbody tr:nth-child(even) {
        background: #f8fafc;
    }
    
    .dataframe tbody tr:hover {
        background: #f1f5f9;
    }
    
    /* Sidebar Logo Section */
    .sidebar-logo {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 4px 14px rgba(102, 126, 234, 0.3);
    }
    
    .sidebar-logo h2 {
        color: white;
        margin: 0;
        font-size: 1.5rem;
        font-weight: 700;
    }
    
    /* Status Badge */
    .status-badge {
        background: #10b981;
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 20px;
        display: inline-block;
        font-weight: 600;
        font-size: 0.875rem;
        box-shadow: 0 2px 8px rgba(16, 185, 129, 0.3);
    }
    
    /* Feature List */
    .feature-list {
        background: rgba(255,255,255,0.05);
        padding: 1rem;
        border-radius: 12px;
        backdrop-filter: blur(10px);
    }
    
    .feature-item {
        padding: 0.75rem;
        color: #e2e8f0;
        font-size: 0.9rem;
        border-bottom: 1px solid rgba(255,255,255,0.1);
        transition: all 0.2s ease;
    }
    
    .feature-item:last-child {
        border-bottom: none;
    }
    
    .feature-item:hover {
        background: rgba(255,255,255,0.05);
        padding-left: 1rem;
    }
    
    /* Upload Section */
    .upload-section {
        background: white;
        border-radius: 12px;
        padding: 2rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
    }
    
    .section-title {
        font-size: 1.25rem;
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* File Badge */
    .file-badge {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 0.25rem 0.75rem;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 600;
        display: inline-block;
    }
    
    /* Bank Card */
    .bank-card {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #667eea;
        margin: 0.5rem 0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: all 0.2s ease;
    }
    
    .bank-card:hover {
        transform: translateX(4px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    
    /* Error Card */
    .error-card {
        background: #fef2f2;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #ef4444;
        margin: 0.5rem 0;
    }
    
    /* Download Section */
    .download-section {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 12px;
        color: white;
        margin-top: 2rem;
        box-shadow: 0 10px 40px rgba(102, 126, 234, 0.25);
    }
    
    /* Footer */
    .footer {
        text-align: center;
        color: #64748b;
        padding: 2rem;
        margin-top: 3rem;
        font-size: 0.9rem;
    }
    
    /* Scrollbar Styling */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: #f1f5f9;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: #764ba2;
    }
</style>
""", unsafe_allow_html=True)

# [ALL YOUR EXISTING FUNCTIONS GO HERE - NO CHANGES]
# I'll include them but keep them the same as before

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

def translate_to_english(text: str) -> dict:
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
    text = text.strip()
    if text and not text[0].isupper():
        text = text.capitalize()
    if text and not text.endswith(('.', '!', '?')):
        text += '.'
    if text and not text.lower().startswith(('hi', 'hello')):
        text = "Hi, " + text[0].lower() + text[1:] if len(text) > 1 else text
    return text

def generate_polite_formal(text: str) -> str:
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
    text = text.strip()
    if text and not text[0].isupper():
        text = text.capitalize()
    if text and not text.endswith(('.', '!', '?')):
        text += '.'
    text = text.replace('I am ', "I'm ")
    text = text.replace('could you please', 'please')
    return text

# UI COMPONENTS

def render_sidebar():
    with st.sidebar:
        st.markdown("""
        <div class="sidebar-logo">
            <h2>📊 QR Cleaner Pro</h2>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div style='text-align: center; margin-bottom: 1.5rem;'>
            <span class="status-badge">🟢 System Online</span>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.markdown("""
        <div style='padding: 0.5rem 0; margin-bottom: 1rem;'>
            <h3 style='color: #e2e8f0; font-size: 0.875rem; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 1rem;'>
                FEATURES
            </h3>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class='feature-list'>
            <div class='feature-item'>✅ Single/Multiple File Processing</div>
            <div class='feature-item'>✅ Bulk 350+ Files (Master)</div>
            <div class='feature-item'>✅ Bank Name Column Addition</div>
            <div class='feature-item'>✅ Mismatch Report Generation</div>
            <div class='feature-item'>✅ Column Validation</div>
            <div class='feature-item'>✅ Hindi to English Translation</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.markdown("""
        <div style='text-align: center; color: #94a3b8; font-size: 0.75rem; margin-top: 2rem;'>
            Version 3.0<br>
            Updated: February 2026
        </div>
        """, unsafe_allow_html=True)

def render_data_cleaner_tab():
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown('<p class="section-title">📁 Upload Excel Files</p>', unsafe_allow_html=True)
        uploaded_files = st.file_uploader(
            "Select one or multiple files (.xlsx, .xls, .csv)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            help="Upload Excel or CSV files to clean and process",
            key="file_uploader_normal",
            label_visibility="collapsed"
        )
    
    with col2:
        st.markdown('<p class="section-title">📊 Upload Stats</p>', unsafe_allow_html=True)
        if uploaded_files:
            st.metric("Files Uploaded", len(uploaded_files))
            total_size = sum([f.size for f in uploaded_files]) / 1024
            st.metric("Total Size", f"{total_size:.1f} KB")
        else:
            st.info("📂 No files uploaded yet")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_files:
        st.markdown("---")
        
        with st.expander("📋 View Uploaded Files", expanded=True):
            for idx, file in enumerate(uploaded_files, 1):
                col_a, col_b, col_c = st.columns([3, 1, 1])
                with col_a:
                    st.markdown(f'**{idx}.** {file.name} <span class="file-badge">Ready</span>', unsafe_allow_html=True)
                with col_b:
                    st.write(f"{file.size / 1024:.1f} KB")
                with col_c:
                    st.write("✅")
        
        st.markdown("---")
        
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 CLEAN & PROCESS FILES", use_container_width=True, type="primary", key="process_normal"):
                process_files(uploaded_files)
        
        with st.expander("ℹ️ Cleaning Operations Info"):
            st.markdown("""
**Operations that will be performed:**

1. ✅ Remove duplicate mobile numbers
2. ✅ Clean mobile numbers (remove '91' prefix)
3. ✅ Standardize dates to dd-mm-yyyy format
4. ✅ Format Aadhaar numbers with prefix
5. ✅ Format Account numbers with prefix
6. ✅ Clean special characters from addresses
7. ✅ Clean special characters from names
8. ✅ Add/update Branch Name column
9. ✅ Merge multiple files (if applicable)
""")

def render_bulk_processor_tab():
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    
    st.markdown('<p class="section-title">📂 Bulk Master File Creator</p>', unsafe_allow_html=True)
    st.info("""
    💡 **How it works:** Upload all your Excel files (350+) at once. Bank names are extracted from filenames. 
    All files are cleaned, and a "Bank Name" column is added as the first column. Creates ONE master Excel with all banks merged + ONE mismatch report for problematic files.
    """)
    
    st.markdown("---")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown('<p class="section-title">📤 Upload All Files</p>', unsafe_allow_html=True)
        bulk_files = st.file_uploader(
            "Select ALL files from your folder (.xlsx, .xls, .csv)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            help="Bank name will be extracted from filename",
            key="file_uploader_bulk",
            label_visibility="collapsed"
        )
    
    with col2:
        st.markdown('<p class="section-title">📊 Upload Statistics</p>', unsafe_allow_html=True)
        if bulk_files:
            st.metric("Files Uploaded", len(bulk_files))
            total_size_mb = sum([f.size for f in bulk_files]) / (1024 * 1024)
            st.metric("Total Size", f"{total_size_mb:.1f} MB")
            
            unique_banks = set([extract_bank_name_from_filename(f.name) for f in bulk_files])
            st.metric("Unique Banks", len(unique_banks))
        else:
            st.info("📂 No files uploaded yet")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    if bulk_files:
        st.markdown("---")
        
        with st.expander("📋 Preview Files (First 10)", expanded=True):
            for idx, file in enumerate(bulk_files[:10], 1):
                bank_name = extract_bank_name_from_filename(file.name)
                col_a, col_b, col_c, col_d = st.columns([2, 2, 1, 1])
                with col_a:
                    st.write(f"**{idx}.** {file.name}")
                with col_b:
                    st.markdown(f'<span class="file-badge">🏦 {bank_name}</span>', unsafe_allow_html=True)
                with col_c:
                    st.write(f"{file.size / 1024:.1f} KB")
                with col_d:
                    st.write("✅")
            
            if len(bulk_files) > 10:
                st.info(f"... and **{len(bulk_files) - 10} more files**")
        
        st.markdown("---")
        
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 CREATE MASTER FILE", use_container_width=True, type="primary", key="process_bulk"):
                process_bulk_to_master(bulk_files)
        
        with st.expander("ℹ️ What Will Be Created"):
            st.markdown("""
**Output Files:**

1. **Master Cleaned File** (All_Banks_Master_Cleaned.xlsx)
   - ONE Excel file with ALL files merged
   - "Bank Name" column as FIRST column
   - All data cleaned and formatted
   
2. **Mismatch Report** (Mismatch_Report.xlsx)
   - Lists all files with column issues
   - Shows missing/extra columns
   - Only created if issues found
""")

def render_english_creator_tab():
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    
    st.markdown('<p class="section-title">🌐 Hindi/Hinglish to Professional English</p>', unsafe_allow_html=True)
    st.info("Perfect for emails, tasks, and formal communication")
    
    if 'translation_results' not in st.session_state:
        st.session_state.translation_results = None
    
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        st.markdown("#### 📝 Input Text")
        input_text = st.text_area(
            "Enter your text",
            height=300,
            placeholder="Example:\nमुझे यह काम जल्दी चाहिए",
            key="input_text_translator",
            label_visibility="collapsed"
        )
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            convert_button = st.button("✨ CONVERT", use_container_width=True, type="primary")
        with col_btn2:
            if st.button("🗑️ CLEAR", use_container_width=True):
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
            if st.button("📋 COPY OPTION 1", key="copy1", use_container_width=True):
                st.code(results['option1'], language=None)
            
            st.markdown("---")
            
            st.markdown("**Option 2** (Polite & Formal)")
            st.text_area("", value=results['option2'], height=80, key="out2", label_visibility="collapsed")
            if st.button("📋 COPY OPTION 2", key="copy2", use_container_width=True):
                st.code(results['option2'], language=None)
            
            st.markdown("---")
            
            st.markdown("**Option 3** (Crisp & Professional)")
            st.text_area("", value=results['option3'], height=80, key="out3", label_visibility="collapsed")
            if st.button("📋 COPY OPTION 3", key="copy3", use_container_width=True):
                st.code(results['option3'], language=None)
        else:
            st.info("👈 Enter text and click Convert")
    
    st.markdown('</div>', unsafe_allow_html=True)

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
            st.balloons()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", len(cleaned_df))
            with col2:
                st.metric("Total Columns", len(cleaned_df.columns))
            with col3:
                st.metric("Operations", len(all_logs))
            
            st.download_button(
                "⬇️ DOWNLOAD CLEANED FILE",
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
            st.balloons()
            
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
                "⬇️ DOWNLOAD MERGED FILE",
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
        st.balloons()
        
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
        
        if results['master_data'] is not None:
            st.markdown("### 📄 Master File Preview")
            st.dataframe(results['master_data'].head(10), use_container_width=True)
            
            st.markdown("### 🏦 Bank Distribution")
            bank_counts = results['master_data']['Bank Name'].value_counts()
            st.dataframe(
                pd.DataFrame({
                    'Bank Name': bank_counts.index,
                    'Row Count': bank_counts.values
                }),
                use_container_width=True
            )
        
        if results['mismatch_files']:
            st.markdown("---")
            st.warning(f"⚠️ {len(results['mismatch_files'])} files have issues")
            
            with st.expander("📋 View Mismatch Report"):
                mismatch_df = create_mismatch_report(results['mismatch_files'])
                st.dataframe(mismatch_df, use_container_width=True)
        
        st.markdown("---")
        st.markdown('<div class="download-section">', unsafe_allow_html=True)
        st.markdown("### 📥 Download Files")
        
        col_dl1, col_dl2 = st.columns(2)
        
        if results['master_data'] is not None:
            with col_dl1:
                st.markdown("#### 📄 Master File")
                master_buffer = io.BytesIO()
                with pd.ExcelWriter(master_buffer, engine='openpyxl') as writer:
                    results['master_data'].to_excel(writer, index=False, sheet_name='All_Banks_Master')
                
                st.download_button(
                    "⬇️ DOWNLOAD MASTER FILE",
                    data=master_buffer.getvalue(),
                    file_name=f"All_Banks_Master_Cleaned_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        if results['mismatch_files']:
            with col_dl2:
                st.markdown("#### 📋 Mismatch Report")
                mismatch_df = create_mismatch_report(results['mismatch_files'])
                mismatch_buffer = io.BytesIO()
                with pd.ExcelWriter(mismatch_buffer, engine='openpyxl') as writer:
                    mismatch_df.to_excel(writer, index=False, sheet_name='Mismatch_Report')
                
                st.download_button(
                    "⬇️ DOWNLOAD MISMATCH REPORT",
                    data=mismatch_buffer.getvalue(),
                    file_name=f"Mismatch_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        progress_bar.empty()
        status_text.empty()
        
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        st.error(f"❌ Error: {str(e)}")
        progress_bar.empty()
        status_text.empty()

def main():
    st.markdown("""
    <div class="header-card">
        <h1 class="header-title">📊 QR Data Cleaner Pro</h1>
        <p class="header-subtitle">Professional data cleaning & processing solution for banking operations</p>
    </div>
    """, unsafe_allow_html=True)
    
    render_sidebar()
    
    tab1, tab2, tab3 = st.tabs(["📁 QR Data Cleaner", "📂 Bulk Master Creator", "🌐 English Creator"])
    
    with tab1:
        render_data_cleaner_tab()
    
    with tab2:
        render_bulk_processor_tab()
    
    with tab3:
        render_english_creator_tab()
    
    st.markdown("""
    <div class="footer">
        <p>Made with ❤️ for operations team | QR Data Cleaner Pro v3.0 | © 2026</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
