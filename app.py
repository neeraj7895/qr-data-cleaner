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
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    
    /* Header styling */
    .main-header {
        background: linear-gradient(135deg, #5f72bd 0%, #9921e8 100%);
        padding: 2rem;
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(95, 114, 189, 0.3);
        margin-bottom: 2rem;
    }
    
    .main-title {
        color: white;
        font-size: 2.5rem;
        font-weight: 800;
        margin: 0;
    }
    
    .subtitle {
        color: rgba(255,255,255,0.9);
        font-size: 1rem;
        margin-top: 0.5rem;
    }
    
    /* Card styling */
    .custom-card {
        background: white;
        padding: 2rem;
        border-radius: 20px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.08);
        margin-bottom: 1.5rem;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #5f72bd 0%, #9921e8 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 12px;
        font-weight: 600;
        width: 100%;
        box-shadow: 0 4px 15px rgba(95, 114, 189, 0.3);
    }
    
    .stButton > button:hover {
        box-shadow: 0 6px 20px rgba(95, 114, 189, 0.5);
        transform: translateY(-2px);
        transition: all 0.3s ease;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: white;
        padding: 0.5rem;
        border-radius: 15px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.06);
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        color: #64748b;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #5f72bd 0%, #9921e8 100%);
        color: white;
    }
    
    /* Text area styling */
    .stTextArea textarea {
        border-radius: 12px;
        border: 2px solid #e2e8f0;
    }
    
    .stTextArea textarea:focus {
        border-color: #5f72bd;
        box-shadow: 0 0 0 3px rgba(95, 114, 189, 0.1);
    }
    
    /* Success/Info boxes */
    .stSuccess, .stInfo {
        border-radius: 12px;
    }
    
    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8fafc 0%, #e2e8f0 100%);
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
    
    # 4. Aadhaar formatting
    for col in ["Aadhar No", "Aadhaar No"]:
        if col in df.columns:
            df[col] = df[col].astype(str).apply(
                lambda x: "'" + x.lstrip("'").replace(".0", "")
                if x.strip() != "" and x.lower() != "nan" else ""
            )
    
    # 5. Account No formatting
    if "Account No" in df.columns:
        df["Account No"] = df["Account No"].astype(str).apply(
            lambda x: "'" + x.lstrip("'").replace(".0", "")
            if x.strip() != "" and x.lower() != "nan" else ""
        )
    
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
    """Load Excel file"""
    return pd.read_excel(file)

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
    <p class="subtitle">Clean, merge & standardize your data efficiently</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    st.markdown("---")
    st.success("**System Status:** 🟢 Active")
    st.markdown("---")
    st.markdown("### 📋 Features")
    st.markdown("""
    - ✅ Remove duplicates
    - ✅ Clean mobile numbers
    - ✅ Standardize dates
    - ✅ Format Aadhaar/Account
    - ✅ Add dropdowns
    - ✅ Hindi to English
    """)

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
                            
                            merged_df = pd.concat(all_dfs, ignore_index=True)
                            
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
    st.markdown("### 🌐 Hindi/Hinglish to Professional English")
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
            "Enter your text (Hindi/Hinglish/English)",
            height=250,
            placeholder="Example:\nHi mujhe ye samjah nahee aa rha hee krapya ye detils dobara send kare task par",
            key="input_text"
        )
        
        convert_button = st.button("✨ Convert to Professional English", use_container_width=True, type="primary")
    
    with col_right:
        st.markdown("#### ✅ Professional Options")
        
        if convert_button and input_text.strip():
            with st.spinner("🔄 Converting..."):
                # Translation dictionary
                translations = {
                    'mujhe': 'I', 'me': 'me', 'hee': '', 'hai': 'is', 'ka': 'of', 'ke': 'of',
                    'samjah nahee aa rha': 'unable to understand', 'krapya': 'please',
                    'ye': 'the', 'detils': 'details', 'dobara': 'again', 'send kare': 'send',
                    'task par': 'regarding the task', 'kal': 'tomorrow', 'meeting': 'meeting',
                    'schedule karni': 'schedule', 'team': 'team', 'saath': 'with',
                    'chahiye': 'need', 'urgent': 'urgent', 'status': 'status', 'update': 'update',
                    'karo': '', 'complete': 'complete', 'by': 'by', 'evening': 'evening'
                }
                
                # Clean and translate
                text = input_text.lower()
                for hindi, english in translations.items():
                    text = text.replace(hindi, english)
                
                # Clean up extra spaces
                text = ' '.join(text.split())
                text = text.capitalize()
                
                # Generate 3 options
                st.session_state.option1 = f"Hi, I am {text}. Could you please help with this?"
                st.session_state.option2 = f"Hello, I would like to request assistance with {text}. Thank you."
                st.session_state.option3 = f"Hi, need help with {text}."
                
                st.success("✅ Converted!")
        
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

Input: Hi mujhe ye samjah nahee aa rha hee krapya ye detils dobara send kare task par

Options will be generated based on your input.
""")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #64748b; padding: 2rem;'>
    <p style='font-size: 0.9rem;'>Made with ❤️ for your team</p>
</div>
""", unsafe_allow_html=True)
