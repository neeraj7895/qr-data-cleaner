import streamlit as st
import pandas as pd
import re
import io
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

# Custom CSS for modern UI
st.markdown("""
<style>
    /* Main background gradient */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Header styling */
    .main-header {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin-bottom: 2rem;
    }
    
    .main-title {
        color: #1f2937;
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
    }
    
    .subtitle {
        color: #6b7280;
        font-size: 1rem;
        margin-top: 0.5rem;
    }
    
    /* Card styling */
    .custom-card {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin-bottom: 1.5rem;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 10px;
        font-weight: 600;
        width: 100%;
    }
    
    .stButton > button:hover {
        box-shadow: 0 6px 12px rgba(102, 126, 234, 0.4);
        transform: translateY(-2px);
        transition: all 0.3s ease;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: white;
        padding: 0.5rem;
        border-radius: 10px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
    }
    
    /* Upload box styling */
    .uploadedFile {
        background: #f3f4f6;
        border-radius: 8px;
        padding: 1rem;
    }
    
    /* Success/Info boxes */
    .stSuccess, .stInfo {
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ============= YOUR EXISTING FUNCTIONS =============
# (Keep all your existing functions here: clean_data, add_dropdowns, load_excel, etc.)

def clean_data(df, source_file=None):
    """Your existing clean_data function - paste it here"""
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
    st.info("**System Status:** 🟢 Active")
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
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("#### 📝 Input Text")
        input_text = st.text_area(
            "Enter your text (Hindi/Hinglish/English)",
            height=250,
            placeholder="Example:\nMujhe kal meeting schedule karni hai team ke saath...\n\nOr in English:\ni need meeting tomorrow with team",
            help="Type or paste your text in Hindi, Hinglish, or English"
        )
        
        convert_button = st.button("✨ Convert to Professional English", use_container_width=True)
    
    with col_right:
        st.markdown("#### ✅ Professional Output")
        
        if convert_button and input_text.strip():
            st.info("🔧 **API Integration Required**\n\nTo enable this feature, you need to add the translation API. For now, here's a demo output:")
            
            # Demo output
            demo_output = """I would like to schedule a meeting with the team tomorrow.

Please let me know your availability so we can coordinate accordingly.

Thank you."""
            
            st.text_area(
                "Professional English Result (Demo)",
                value=demo_output,
                height=250,
                help="This is a demo output. Integrate with translation API for real conversions."
            )
            
            st.code(demo_output, language=None)
            st.warning("⚠️ To enable real-time translation, integrate with Claude API or Google Translate API")
        
        elif convert_button:
            st.warning("⚠️ Please enter some text first!")
        else:
            st.info("👈 Enter your text and click Convert")
    
    # Info section
    st.markdown("---")
    with st.expander("ℹ️ How it works"):
        st.markdown("""
        **This tool helps you:**
        
        - 📧 **Translate** Hindi/Hinglish to English
        - 💼 **Improve** existing English to corporate standard
        - ✨ **Polish** informal text for professional use
        - 📝 **Perfect** for emails, tasks, and formal communication
        
        **Example:**
        - Input: "Mujhe kal meeting rakhni hai"
        - Output: "I need to schedule a meeting tomorrow"
        
        **Note:** Currently showing demo mode. To enable real translation:
        1. Add translation API (Claude/Google Translate)
        2. Update the conversion logic in the code
        3. Deploy with API credentials
        """)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: white; padding: 2rem;'>
    <p style='font-size: 0.9rem;'>Made with ❤️ for your team | Powered by AI</p>
</div>
""", unsafe_allow_html=T
