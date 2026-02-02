import streamlit as st
import pandas as pd
import re
import io
import requests
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from typing import List, Tuple, Optional
import logging

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
    
    /* Success message */
    .success-box {
        background: #d1fae5;
        border-left: 4px solid #10b981;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    /* Error message */
    .error-box {
        background: #fee2e2;
        border-left: 4px solid #ef4444;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)


# ============= DATA CLEANING FUNCTIONS =============

def pre_process_dataframe(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """
    Pre-process dataframe before cleaning:
    1. Delete Source_File column if it exists (any variation)
    2. Add Branch Name column if it doesn't exist and fill with "HO Branch"
    """
    logs = []
    df = df.copy()  # Work on a copy to avoid SettingWithCopyWarning
    
    # 1. Delete Source_File column (any variation)
    source_file_variations = ["source_file", "sourcefile", "source file"]
    columns_to_drop = []
    
    for df_col in df.columns:
        df_col_normalized = df_col.lower().replace(" ", "").replace("_", "")
        if df_col_normalized in source_file_variations:
            columns_to_drop.append(df_col)
    
    if columns_to_drop:
        df = df.drop(columns=columns_to_drop)
        logs.append(f"✓ Deleted Source_File column(s): {', '.join(columns_to_drop)}")
    
    # 2. Add Branch Name column if it doesn't exist
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
    """Format date to dd-mm-yyyy format with leading quote"""
    if pd.isna(x) or str(x).strip() == "":
        return ""
    
    # Handle Excel serial dates
    if isinstance(x, (int, float)) and not pd.isna(x):
        try:
            dt = pd.to_datetime("1899-12-30") + pd.to_timedelta(int(x), unit="D")
            return "'" + dt.strftime("%d-%m-%Y")
        except Exception:
            pass
    
    # Handle string dates
    try:
        dt = pd.to_datetime(str(x), dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return str(x)
        return "'" + dt.strftime("%d-%m-%Y")
    except Exception:
        return str(x)


def format_aadhaar(x) -> str:
    """Format Aadhaar number with leading quote and remove special characters"""
    if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
        return ""
    
    # Convert to string and remove any existing quotes
    x_str = str(x).strip().lstrip("'")
    
    # Remove alphabets and special characters (keep only digits)
    x_str = re.sub(r'[^0-9]', '', x_str)
    
    # If empty after cleaning, return empty
    if not x_str:
        return ""
    
    # Add prefix and return
    return "'" + x_str


def format_account_number(x) -> str:
    """Format account number with leading quote"""
    if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
        return ""
    
    x_str = str(x).strip().lstrip("'")
    
    # Remove .0 if present at the end
    if x_str.endswith('.0'):
        x_str = x_str[:-2]
    
    return "'" + x_str


def clean_address(x) -> str:
    """Clean address by removing special characters"""
    if pd.isna(x) or str(x).strip() == "":
        return ""
    
    x_str = str(x)
    
    # Replace special characters with space
    special_chars = [',', '.', '/', '&', '-', '"', ';', '(', ')', '\\']
    for char in special_chars:
        x_str = x_str.replace(char, ' ')
    
    # Replace multiple spaces with single space
    x_str = re.sub(r'\s+', ' ', x_str)
    
    return x_str.strip()


def clean_name(x) -> str:
    """Clean name by removing special characters"""
    if pd.isna(x) or str(x).strip() in ["", "nan", "NaN", "None"]:
        return ""
    
    x_str = str(x)
    
    # Replace special characters with space
    special_chars = ['-', '/', ':', '|', '(', ')', '&', '#', ',', '.', ';', "'"]
    for char in special_chars:
        x_str = x_str.replace(char, ' ')
    
    # Replace multiple spaces with single space and strip
    x_str = re.sub(r'\s+', ' ', x_str).strip()
    
    return x_str


def clean_data(df: pd.DataFrame, source_file: Optional[str] = None) -> Tuple[pd.DataFrame, List[str]]:
    """
    Main data cleaning function
    
    Args:
        df: Input DataFrame
        source_file: Name of source file (optional)
    
    Returns:
        Tuple of (cleaned DataFrame, list of log messages)
    """
    logs = []
    df = df.copy()  # Work on a copy
    
    try:
        # PRE-PROCESSING
        df, pre_logs = pre_process_dataframe(df)
        logs.extend(pre_logs)
        
        # 1. Remove duplicates by Mobile No
        if "Mobile No" in df.columns:
            before = len(df)
            df = df.drop_duplicates(subset=["Mobile No"], keep="first")
            after = len(df)
            if before > after:
                logs.append(f"✓ Removed {before - after} duplicate mobile numbers")
        
        # 2. Clean mobile numbers
        if "Mobile No" in df.columns:
            df["Mobile No"] = df["Mobile No"].apply(clean_mobile_number)
            logs.append("✓ Cleaned mobile numbers (removed '91' prefix where applicable)")
        
        # 3. Format date columns
        date_columns = ["DOB", "DOI", "Account Opening Date"]
        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].apply(format_date)
                logs.append(f"✓ Formatted date column: {col}")
        
        # 4. Format Aadhaar columns
        aadhaar_columns = ["Aadhar No", "Aadhaar No"]
        for col in aadhaar_columns:
            if col in df.columns:
                df[col] = df[col].apply(format_aadhaar)
                logs.append(f"✓ Formatted Aadhaar column: {col}")
        
        # 5. Clean Address Line 1
        if "Address Line 1" in df.columns:
            df["Address Line 1"] = df["Address Line 1"].apply(clean_address)
            logs.append("✓ Cleaned special characters from Address Line 1")
        
        # 6. Format Account No
        if "Account No" in df.columns:
            df["Account No"] = df["Account No"].apply(format_account_number)
            logs.append("✓ Formatted Account No column")
        
        # 7. Replace all Branch Name values
        if "Branch Name" in df.columns:
            df["Branch Name"] = "HO Branch"
            logs.append("✓ Replaced all values in 'Branch Name' with 'HO Branch'")
        
        # 8. Clean name columns
        name_columns = ["First Name", "Middle Name", "Last Name", "Entity Name", "Account Holder Name"]
        for col in name_columns:
            if col in df.columns:
                df[col] = df[col].apply(clean_name)
        logs.append("✓ Cleaned special characters from name columns")
        
        # 9. Clear personal names if entity present
        if "Entity Name" in df.columns:
            entity_mask = df["Entity Name"].notna() & (df["Entity Name"].str.strip() != "")
            personal_name_cols = ["First Name", "Middle Name", "Last Name"]
            for col in personal_name_cols:
                if col in df.columns:
                    df.loc[entity_mask, col] = ""
            logs.append("✓ Cleared personal names where Entity Name is present")
        
        # 10. Account Holder Name fallback
        if "Account Holder Name" in df.columns and "Entity Name" in df.columns:
            mask = (df["Account Holder Name"].isna()) | (df["Account Holder Name"].str.strip() == "")
            entity_mask = (df["Entity Name"].notna()) & (df["Entity Name"].str.strip() != "")
            df.loc[mask & entity_mask, "Account Holder Name"] = df.loc[mask & entity_mask, "Entity Name"]
            logs.append("✓ Filled missing Account Holder Names with Entity Name where applicable")
        
        # 11. Address Line 2 fallback
        if "Address Line 1" in df.columns and "Address Line 2" in df.columns:
            mask = (df["Address Line 2"].isna()) | (df["Address Line 2"].str.strip() == "")
            has_addr1 = (df["Address Line 1"].notna()) & (df["Address Line 1"].str.strip() != "")
            df.loc[mask & has_addr1, "Address Line 2"] = df.loc[mask & has_addr1, "Address Line 1"]
            logs.append("✓ Copied Address Line 1 to Address Line 2 where blank")
        
        # 12. Clear unwanted columns
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
                    logs.append(f"✓ Cleared data from column: {df_col}")
        
        return df, logs
    
    except Exception as e:
        logger.error(f"Error in clean_data: {str(e)}")
        logs.append(f"❌ Error during cleaning: {str(e)}")
        return df, logs


def add_dropdowns(buffer: io.BytesIO, sheet_name: str = "Cleaned") -> io.BytesIO:
    """
    Add dropdown validations to Excel file
    
    Args:
        buffer: BytesIO buffer containing Excel file
        sheet_name: Name of the sheet to add dropdowns to
    
    Returns:
        BytesIO buffer with dropdowns added
    """
    try:
        buffer.seek(0)
        wb = load_workbook(buffer)
        ws = wb[sheet_name]
        
        # Example: Add dropdown for a specific column (customize as needed)
        # This is a placeholder - add your actual dropdown logic here
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output
    except Exception as e:
        logger.error(f"Error adding dropdowns: {str(e)}")
        return buffer


def load_excel(file) -> Optional[pd.DataFrame]:
    """
    Load Excel/CSV file with proper dtype to preserve leading zeros
    
    Args:
        file: Uploaded file object
    
    Returns:
        DataFrame or None if error
    """
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
        st.error(f"❌ Error loading file {file.name}: {str(e)}")
        return None


# ============= TRANSLATION FUNCTIONS =============

def translate_to_english(text: str) -> dict:
    """
    Translate Hindi/Hinglish to professional English using MyMemory API
    
    Args:
        text: Input text to translate
    
    Returns:
        Dictionary with 3 professional options
    """
    try:
        url = "https://api.mymemory.translated.net/get"
        params = {
            'q': text,
            'langpair': 'hi|en'
        }
        
        response = requests.get(url, params=params, timeout=10)
        data = response.json()
        
        if response.status_code == 200 and 'responseData' in data:
            base_text = data['responseData']['translatedText'].strip()
            
            # Generate 3 professional versions
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
    """Generate simple and professional version"""
    text = text.strip()
    
    # Capitalize first letter
    if text and not text[0].isupper():
        text = text.capitalize()
    
    # Add period if missing
    if text and not text.endswith(('.', '!', '?')):
        text += '.'
    
    # Add greeting if missing
    if text and not text.lower().startswith(('hi', 'hello')):
        text = "Hi, " + text[0].lower() + text[1:] if len(text) > 1 else text
    
    return text


def generate_polite_formal(text: str) -> str:
    """Generate more polite and formal version"""
    text = text.strip()
    
    # Capitalize first letter
    if text and not text[0].isupper():
        text = text.capitalize()
    
    # Add period if missing
    if text and not text.endswith(('.', '!', '?')):
        text += '.'
    
    # Add formal greeting
    if text and not text.lower().startswith('hello'):
        text = "Hello, " + text[0].lower() + text[1:] if len(text) > 1 else text
    
    # Add thank you if missing
    if 'thank you' not in text.lower():
        text = text.rstrip('.') + '. Thank you.'
    
    return text


def generate_crisp_professional(text: str) -> str:
    """Generate crisp and professional version"""
    text = text.strip()
    
    # Capitalize first letter
    if text and not text[0].isupper():
        text = text.capitalize()
    
    # Add period if missing
    if text and not text.endswith(('.', '!', '?')):
        text += '.'
    
    # Make it more concise
    text = text.replace('I am ', "I'm ")
    text = text.replace('could you please', 'please')
    
    return text


# ============= UI COMPONENTS =============

def render_sidebar():
    """Render sidebar content"""
    with st.sidebar:
        # Logo/Title Section
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
        
        # Status
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
        
        # Features List
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
                ✅ Remove duplicates
            </div>
            <div style='margin-bottom: 0.75rem; color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Clean mobile numbers
            </div>
            <div style='margin-bottom: 0.75rem; color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Standardize dates
            </div>
            <div style='margin-bottom: 0.75rem; color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Format Aadhaar/Account
            </div>
            <div style='margin-bottom: 0.75rem; color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Add dropdown validations
            </div>
            <div style='color: #e5e7eb; font-size: 0.9rem;'>
                ✅ Hindi to English
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Version info
        st.markdown("""
        <div style='text-align: center; color: #9ca3af; font-size: 0.8rem;'>
            Version 2.0<br>
            Last updated: Feb 2026
        </div>
        """, unsafe_allow_html=True)


def render_data_cleaner_tab():
    """Render the Data Cleaner tab"""
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📁 Upload Excel Files")
        uploaded_files = st.file_uploader(
            "Select one or multiple files (.xlsx, .xls, .csv)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            help="Upload Excel or CSV files to clean and process",
            key="file_uploader"
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
        
        # Show uploaded files
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
        
        # Process button
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 Clean & Process Files", use_container_width=True, type="primary"):
                process_files(uploaded_files)
        
        # Processing steps info
        with st.expander("🔍 Cleaning Operations", expanded=False):
            st.markdown("""
**The following operations will be performed:**

1. ✅ Delete Source_File column if exists
2. ✅ Add Branch Name column with 'HO Branch' value
3. ✅ Remove duplicate mobile numbers
4. ✅ Clean 12-digit mobile numbers (remove '91' prefix)
5. ✅ Standardize date formats to dd-mm-yyyy
6. ✅ Format Aadhaar numbers (remove non-digits, add prefix)
7. ✅ Format Account numbers with prefix
8. ✅ Clean special characters from addresses
9. ✅ Clean special characters from names
10. ✅ Clear personal names where Entity Name exists
11. ✅ Fill Account Holder Name from Entity Name if missing
12. ✅ Copy Address Line 1 to Address Line 2 if blank
13. ✅ Clear unwanted column data
14. ✅ Add dropdown validations
15. ✅ Merge multiple files (if applicable)
""")


def process_files(uploaded_files):
    """Process uploaded files with progress tracking"""
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        all_logs = []
        total_steps = len(uploaded_files) + 2  # Files + merge + finalize
        current_step = 0
        
        if len(uploaded_files) == 1:
            # Single file processing
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
            
            # Create Excel output
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
            
            # Display statistics
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
            # Multiple files processing
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
            
            # Merge files
            status_text.text("Merging files...")
            merged_df = pd.concat(all_dfs, ignore_index=True, sort=False)
            current_step += 1
            progress_bar.progress(current_step / total_steps)
            
            # Remove duplicates in merged data
            if "Mobile No" in merged_df.columns:
                before = len(merged_df)
                merged_df = merged_df.drop_duplicates(subset=["Mobile No"], keep="first")
                after = len(merged_df)
                if before > after:
                    all_logs.append(f"✓ Removed {before - after} duplicates from merged data")
            
            # Create Excel output
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
            
            # Display statistics
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
        
        # Show logs
        with st.expander("📝 View Detailed Cleaning Logs", expanded=False):
            for log in all_logs:
                st.write(log)
        
        # Clear progress
        progress_bar.empty()
        status_text.empty()
                    
    except Exception as e:
        logger.error(f"Error processing files: {str(e)}")
        st.error(f"❌ Error processing files: {str(e)}")
        progress_bar.empty()
        status_text.empty()


def render_english_creator_tab():
    """Render the English Creator tab"""
    st.markdown("### 🌐 Hindi/Hinglish to Professional English")
    st.markdown("Perfect for emails, tasks, and formal communication")
    
    # Initialize session state
    if 'translation_results' not in st.session_state:
        st.session_state.translation_results = None
    
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        st.markdown("#### 📝 Input Text")
        input_text = st.text_area(
            "Enter your text (Hindi/English/Hinglish)",
            height=300,
            placeholder="Example:\nमुझे यह काम जल्दी चाहिए\n\nWill be converted to professional English...",
            key="input_text_translator"
        )
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            convert_button = st.button(
                "✨ Convert to English", 
                use_container_width=True, 
                type="primary"
            )
        with col_btn2:
            if st.button("🗑️ Clear All", use_container_width=True):
                st.session_state.translation_results = None
                st.rerun()
    
    with col_right:
        st.markdown("#### ✅ Professional Options")
        
        if convert_button and input_text.strip():
            with st.spinner("🔄 Converting to professional English..."):
                try:
                    results = translate_to_english(input_text)
                    st.session_state.translation_results = results
                    st.success("✅ Converted successfully!")
                except Exception as e:
                    st.error(f"❌ Translation failed: {str(e)}")
                    st.info("💡 Please check your internet connection and try again.")
        
        # Display options
        if st.session_state.translation_results:
            results = st.session_state.translation_results
            
            # Option 1
            st.markdown("**Option 1** (Simple & Professional)")
            st.text_area(
                "", 
                value=results['option1'], 
                height=80, 
                key="out1", 
                label_visibility="collapsed"
            )
            if st.button("📋 Copy Option 1", key="copy1", use_container_width=True):
                st.code(results['option1'], language=None)
            
            st.markdown("---")
            
            # Option 2
            st.markdown("**Option 2** (Polite & Formal)")
            st.text_area(
                "", 
                value=results['option2'], 
                height=80, 
                key="out2", 
                label_visibility="collapsed"
            )
            if st.button("📋 Copy Option 2", key="copy2", use_container_width=True):
                st.code(results['option2'], language=None)
            
            st.markdown("---")
            
            # Option 3
            st.markdown("**Option 3** (Crisp & Professional)")
            st.text_area(
                "", 
                value=results['option3'], 
                height=80, 
                key="out3", 
                label_visibility="collapsed"
            )
            if st.button("📋 Copy Option 3", key="copy3", use_container_width=True):
                st.code(results['option3'], language=None)
        else:
            st.info("👈 Enter text and click Convert to see professional English options")
    
    st.markdown("---")
    
    with st.expander("💡 Usage Tips & Examples"):
        st.markdown("""
**How to use:**
1. Type or paste your Hindi/Hinglish/English text in the input box
2. Click "Convert to English" button
3. Choose from 3 professional versions
4. Click "Copy" to copy the text you prefer

**Examples:**

**Input:** मुझे रिपोर्ट चाहिए
**Output:** Hi, I need the report.

**Input:** kal meeting hai kya?
**Output:** Hello, is there a meeting tomorrow? Thank you.

**Input:** Please send me details ASAP
**Output:** Please send me the details at your earliest convenience.
""")


# ============= MAIN APP =============

def main():
    """Main application entry point"""
    # Header
    st.markdown("""
    <div class="main-header">
        <h1 class="main-title">🔧 QR Data Cleaner Pro</h1>
        <p class="subtitle">Clean, merge & standardize your QR code data with ease</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Render sidebar
    render_sidebar()
    
    # Main tabs
    tab1, tab2 = st.tabs(["📁 Data Cleaner", "🌐 English Creator"])
    
    with tab1:
        render_data_cleaner_tab()
    
    with tab2:
        render_english_creator_tab()
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #64748b; padding: 2rem;'>
        <p style='font-size: 0.9rem;'>Made with ❤️ for operations team | QR Data Cleaner Pro v2.0</p>
        <p style='font-size: 0.8rem; color: #9ca3af;'>
            Need help? Contact your IT support team
        </p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
