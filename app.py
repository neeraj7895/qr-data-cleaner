import streamlit as st
import pandas as pd
import re
import io
import requests
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

# Page configuration
st.set_page_config(
    page_title="QR Data Cleaner Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS - PhonePe/Razorpay style (Blue & Grey theme)
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
    }
    
    .stButton > button:hover {
        box-shadow: 0 6px 16px rgba(95, 114, 189, 0.4);
        transform: translateY(-2px);
        transition: all 0.3s ease;
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
</style>
""", unsafe_allow_html=True)

# ============= YOUR EXISTING FUNCTIONS =============
# (Keep all your existing functions here: clean_data, add_dropdowns, load_excel, etc.)

def clean_data(df, source_file=None):
    """Your existing clean_data function"""
    logs = []
    
    # 1. Remove duplicates by Mobile No
    if "Mobile No" in df.columns:
        before = len(df)
        df = df.drop_duplicates(subset=["Mobile No"], keep="first").copy()
        after = len(df)
        if before > after:
            logs.append(f"Removed {before - after} duplicate mobile numbers")
    
    # 2. Clean 12-digit mobile numbers starting with '91'
    if "Mobile No" in df.columns:
        def clean_mobile(x):
            x = str(x).strip()
            x = re.sub(r"\D", "", x)
            if len(x) == 12 and x.startswith("91"):
                x = x[2:]
            return x
        
        df["Mobile No"] = df["Mobile No"].apply(clean_mobile)
        logs.append("Cleaned 12-digit mobile numbers by removing '91' prefix where applicable")
    
    # 3. Dates formatting
    for col in ["DOB", "DOI", "Account Opening Date"]:
        if col in df.columns:
            def format_date(x):
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
            
            df[col] = df[col].apply(format_date)
            logs.append(f"Formatted date column: {col}")
    
    # 4. Aadhaar formatting
    for col in ["Aadhar No", "Aadhaar No"]:
        if col in df.columns:
            def format_aadhaar(x):
                if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
                    return ""
                
                # Convert to string and remove any existing quotes
                x_str = str(x).strip().lstrip("'")
                
                # If it's a float with .0, remove only the .0 part
                if '.' in x_str and x_str.endswith('.0'):
                    x_str = x_str[:-2]
                
                # Add prefix and return
                return "'" + x_str
            
            df[col] = df[col].apply(format_aadhaar)
            logs.append(f"Formatted Aadhaar column: {col}")
    
    # 5. Account No formatting
    if "Account No" in df.columns:
        def format_account(x):
            if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
                return ""
            
            # Convert to string (already string from load_excel dtype)
            x_str = str(x).strip().lstrip("'")
            
            # Remove .0 if present at the end
            if x_str.endswith('.0'):
                x_str = x_str[:-2]
            
            # Add prefix and return
            return "'" + x_str
        
        df["Account No"] = df["Account No"].apply(format_account)
        logs.append("Formatted Account No column")
    
    # 6. Branch Name → Replace all values with "HO Branch"
    if "Branch Name" in df.columns:
        df["Branch Name"] = "HO Branch"
        logs.append("Replaced all values in 'Branch Name' with 'HO Branch'")
    
    # 7. Add Source File column if multiple uploads
    if source_file:
        df["Source_File"] = source_file
        logs.append(f"Added Source_File column: {source_file}")
    
    # 8. Clear unwanted columns (keep header, clear data)
    clear_cols = [
        "Turnover Type", "Acceptance Type", "Ownership Type", "MCC", 
        "Email ID", "Source_File", "Bank Cust ID", "State Code (GST)", 
        "Latitude", "Longitude", "District"
    ]
    
    for col in clear_cols:
        # Find matching columns (case-insensitive, space-insensitive)
        col_normalized = col.lower().replace(" ", "").replace("_", "")
        for df_col in df.columns:
            df_col_normalized = df_col.lower().replace(" ", "").replace("_", "")
            if df_col_normalized == col_normalized:
                df[df_col] = ""
                logs.append(f"Cleared data from column: {df_col}")
    
    return df, logs

def add_dropdowns(buffer, sheet_name="Cleaned"):
    """Your existing add_dropdowns function"""
    buffer.seek(0)
    wb = load_workbook(buffer)
    ws = wb[sheet_name]
    
    # Add your dropdown logic here
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def load_excel(file):
    """Load Excel file with proper dtype to preserve leading zeros"""
    # Read with Account No as string to preserve leading zeros
    return pd.read_excel(file, dtype={'Account No': str, 'Aadhar No': str, 'Aadhaar No': str})

# Function to convert Hindi/Hinglish to English
async def convert_to_english(text):
    """Convert Hindi/Hinglish to professional English using Claude API"""
    try:
        response = await fetch('https://api.anthropic.com/v1/messages', {
            'method': 'POST',
            'headers': {'Content-Type': 'application/json'},
            'body': {
                'model': 'claude-sonnet-4-20250514',
                'max_tokens': 1000,
                'messages': [{
                    'role': 'user',
                    'content': f"""Convert the following text to professional corporate English. If it's in Hindi or Hinglish, translate it to English. If it's already in English, improve it for professional communication. Provide ONLY the converted text without any explanation:

"{text}"

Output:"""
                }]
            }
        })
        
        data = await response.json()
        result = data.get('content', [{}])[0].get('text', '')
        return result.strip()
    except Exception as e:
        return f"Error: {str(e)}"

# ============= STREAMLIT UI =============

# Header
st.markdown("""
<div class="main-header">
    <h1 class="main-title">📊 QR Data Cleaner Pro</h1>
    <p class="subtitle">Clean, merge & standardize your data</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    # Logo/Title Section
    st.markdown("""
    <div style='background: linear-gradient(135deg, #5f72bd 0%, #9921e8 100%); 
                padding: 1.5rem; 
                border-radius: 15px; 
                margin-bottom: 2rem;
                text-align: center;'>
        <h2 style='color: white; margin: 0; font-size: 1.5rem; font-weight: 700;'>
            📊 QR Cleaner Pro
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
        <h3 style='color: #1f2937; font-size: 1rem; font-weight: 700; margin-bottom: 1rem;'>
            Features
        </h3>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div style='background: white; 
                padding: 1rem; 
                border-radius: 12px; 
                box-shadow: 0 2px 8px rgba(0,0,0,0.08);'>
        <div style='margin-bottom: 0.75rem; color: #374151; font-size: 0.9rem;'>
            ✅ Remove duplicates
        </div>
        <div style='margin-bottom: 0.75rem; color: #374151; font-size: 0.9rem;'>
            ✅ Clean mobile numbers
        </div>
        <div style='margin-bottom: 0.75rem; color: #374151; font-size: 0.9rem;'>
            ✅ Standardize dates
        </div>
        <div style='margin-bottom: 0.75rem; color: #374151; font-size: 0.9rem;'>
            ✅ Format Aadhaar/Account
        </div>
        <div style='margin-bottom: 0.75rem; color: #374151; font-size: 0.9rem;'>
            ✅ Add dropdowns
        </div>
        <div style='color: #374151; font-size: 0.9rem;'>
            ✅ Hindi to English
        </div>
    </div>
    """, unsafe_allow_html=True)

# Main tabs
tab1, tab2 = st.tabs(["📁 Data Cleaner", "🌐 English Creator"])

# ============= DATA CLEANER TAB =============
with tab1:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📤 Upload Excel Files")
        uploaded_files = st.file_uploader(
            "Select one or multiple Excel files (.xlsx, .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="Upload Excel files to clean and process"
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
            if st.button("🚀 Clean & Process Files", use_container_width=True):
                with st.spinner("Processing files..."):
                    try:
                        all_logs = []
                        
                        if len(uploaded_files) == 1:
                            # Single file processing
                            df = load_excel(uploaded_files[0])
                            cleaned_df, logs = clean_data(df, uploaded_files[0].name)
                            all_logs.extend(logs)
                            
                            # Create Excel output
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                                cleaned_df.to_excel(writer, index=False, sheet_name="Cleaned")
                            
                            final_output = add_dropdowns(output, sheet_name="Cleaned")
                            
                            st.success("✅ File processed successfully!")
                            st.balloons()
                            
                            st.download_button(
                                "⬇️ Download Cleaned File",
                                data=final_output.getvalue(),
                                file_name="Cleaned_Single.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            # Multiple files processing
                            all_dfs = []
                            for file in uploaded_files:
                                df = load_excel(file)
                                cleaned_df, logs = clean_data(df, file.name)
                                all_dfs.append(cleaned_df)
                                all_logs.extend(logs)
                            
                            # Merge with same columns only (no blank columns)
                            merged_df = pd.concat(all_dfs, ignore_index=True, sort=False)
                            
                            # Remove duplicates in merged
                            if "Mobile No" in merged_df.columns:
                                before = len(merged_df)
                                merged_df = merged_df.drop_duplicates(subset=["Mobile No"], keep="first").copy()
                                after = len(merged_df)
                                all_logs.append(f"Removed {before - after} duplicates from merged data")
                            
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                                merged_df.to_excel(writer, index=False, sheet_name="Cleaned_Merged")
                            
                            final_output = add_dropdowns(output, sheet_name="Cleaned_Merged")
                            
                            st.success("✅ Multiple files processed and merged successfully!")
                            st.balloons()
                            
                            st.download_button(
                                "⬇️ Download Merged Cleaned File",
                                data=final_output.getvalue(),
                                file_name="Cleaned_Merged.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        
                        # Show logs
                        with st.expander("📝 View Cleaning Logs"):
                            for log in all_logs:
                                st.write("✔️", log)
                                
                    except Exception as e:
                        st.error(f"❌ Error processing files: {str(e)}")
        
        # Processing steps info
        with st.expander("🔍 Cleaning Operations", expanded=False):
            st.markdown("""
**The following operations will be performed:**

1. ✅ Remove duplicate mobile numbers
2. ✅ Clean 12-digit mobile numbers (remove '91' prefix)
3. ✅ Standardize date formats to dd-mm-yyyy
4. ✅ Format Aadhaar numbers with prefix
5. ✅ Format Account numbers with prefix
6. ✅ Add dropdown validations for specific columns
7. ✅ Merge multiple files (if applicable)
""")

# ============= ENGLISH CREATOR TAB =============
with tab2:
    st.markdown("### 🌐 Hindi/English to Professional English")
    st.markdown("Perfect for emails, tasks, and formal communication")
    
    # Initialize session state
    if 'option1' not in st.session_state:
        st.session_state.option1 = ""
    if 'option2' not in st.session_state:
        st.session_state.option2 = ""
    if 'option3' not in st.session_state:
        st.session_state.option3 = ""
    
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        st.markdown("#### 📝 Input Text")
        input_text = st.text_area(
            "Enter your text (Hindi/English/English)",
            height=250,
            placeholder="Example:\nHi",
            key="input_text"
        )
        
        convert_button = st.button("✨ Convert to Professional English", use_container_width=True, type="primary")
    
    with col_right:
        st.markdown("#### ✅ Professional Options")
        
        if convert_button and input_text.strip():
            with st.spinner("🔄 Converting..."):
                try:
                    # Use MyMemory Free Translation API
                    url = "https://api.mymemory.translated.net/get"
                    params = {
                        'q': input_text,
                        'langpair': 'hi|en'
                    }
                    
                    response = requests.get(url, params=params, timeout=10)
                    data = response.json()
                    
                    if response.status_code == 200 and 'responseData' in data:
                        base_text = data['responseData']['translatedText']
                        
                        # Generate 3 professional versions
                        # Option 1: Simple & Professional
                        opt1 = base_text.strip()
                        if opt1 and not opt1[0].isupper():
                            opt1 = opt1.capitalize()
                        if opt1 and not opt1.endswith(('.', '!', '?')):
                            opt1 += '.'
                        if opt1 and not opt1.lower().startswith(('hi', 'hello')):
                            opt1 = "Hi, " + opt1[0].lower() + opt1[1:] if len(opt1) > 1 else opt1
                        st.session_state.option1 = opt1
                        
                        # Option 2: More Polite & Formal
                        opt2 = base_text.strip()
                        if opt2 and not opt2[0].isupper():
                            opt2 = opt2.capitalize()
                        if opt2 and not opt2.endswith(('.', '!', '?')):
                            opt2 += '.'
                        if opt2 and not opt2.lower().startswith('hello'):
                            opt2 = "Hello, " + opt2[0].lower() + opt2[1:] if len(opt2) > 1 else opt2
                        if 'thank you' not in opt2.lower():
                            opt2 = opt2.rstrip('.') + '. Thank you.'
                        st.session_state.option2 = opt2
                        
                        # Option 3: Crisp & Professional
                        opt3 = base_text.strip()
                        if opt3 and not opt3[0].isupper():
                            opt3 = opt3.capitalize()
                        if opt3 and not opt3.endswith(('.', '!', '?')):
                            opt3 += '.'
                        opt3 = opt3.replace('I am ', "I'm ")
                        opt3 = opt3.replace('could you please', 'please')
                        st.session_state.option3 = opt3
                        
                        st.success("✅ Converted!")
                    else:
                        st.error("❌ Translation failed. Please try again.")
                    
                except Exception as e:
                    st.error("❌ Translation failed. Check internet connection.")
        
        # Display options
        if st.session_state.option1:
            st.markdown("**Option 1** (Simple & Professional)")
            st.text_area("", value=st.session_state.option1, height=80, key="out1", label_visibility="collapsed")
            if st.button("📋 Copy Option 1", key="copy1", use_container_width=True):
                st.code(st.session_state.option1)
        
        if st.session_state.option2:
            st.markdown("**Option 2** (More Polite & Formal)")
            st.text_area("", value=st.session_state.option2, height=80, key="out2", label_visibility="collapsed")
            if st.button("📋 Copy Option 2", key="copy2", use_container_width=True):
                st.code(st.session_state.option2)
        
        if st.session_state.option3:
            st.markdown("**Option 3** (Crisp & Professional)")
            st.text_area("", value=st.session_state.option3, height=80, key="out3", label_visibility="collapsed")
            if st.button("📋 Copy Option 3", key="copy3", use_container_width=True):
                st.code(st.session_state.option3)
        
        if st.session_state.option1:
            st.markdown("---")
            if st.button("🗑️ Clear All", use_container_width=True):
                st.session_state.option1 = ""
                st.session_state.option2 = ""
                st.session_state.option3 = ""
                st.rerun()
        
        if not st.session_state.option1:
            st.info("👈 Enter text and click Convert")
    
    st.markdown("---")
    
    with st.expander("💡 Examples"):
        st.markdown("""
**Example:**

Input: Hi

Output: Professional English versions will be generated automatically.
""")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #64748b; padding: 2rem;'>
    <p style='font-size: 0.9rem;'>Made with ❤️ for your team</p>
</div>
""", unsafe_allow_html=True)
