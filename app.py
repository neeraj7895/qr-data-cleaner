import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
import anthropic

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

# Header
st.markdown("""
<div class="main-header">
    <h1 class="main-title">📊 QR Data Cleaner Pro</h1>
    <p class="subtitle">Clean, merge & standardize your data efficiently</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.image("https://via.placeholder.com/150x50/667eea/ffffff?text=QR+Cleaner", use_column_width=True)
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
                    # YOUR EXISTING CLEANING LOGIC HERE
                    # (Keep your clean_data, add_dropdowns functions)
                    
                    # Example placeholder
                    import time
                    time.sleep(2)
                    
                    st.success("✅ Files processed successfully!")
                    st.balloons()
                    
                    # Download button would go here
                    st.download_button(
                        "⬇️ Download Cleaned File",
                        data=b"",  # Your processed file bytes
                        file_name="Cleaned_Data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
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
            with st.spinner("Converting to professional English..."):
                try:
                    # Initialize Anthropic client (no API key needed in claude.ai)
                    client = anthropic.Anthropic()
                    
                    # Call Claude API
                    message = client.messages.create(
                        model="claude-sonnet-4-20250514",
                        max_tokens=1000,
                        messages=[
                            {
                                "role": "user",
                                "content": f"""Convert the following text to professional corporate English. If it's in Hindi/Hinglish, translate it. If it's already in English, improve it for professional communication:

"{input_text}"

Provide only the professional English version without any explanations."""
                            }
                        ]
                    )
                    
                    # Extract response
                    output_text = message.content[0].text
                    
                    # Display output
                    st.text_area(
                        "Professional English Result",
                        value=output_text,
                        height=250,
                        help="Copy this text for your email or communication"
                    )
                    
                    # Copy button
                    st.code(output_text, language=None)
                    st.success("✅ Conversion complete! You can copy the text above.")
                    
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    st.info("Make sure the Anthropic API is accessible.")
        
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
        """)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: white; padding: 2rem;'>
    <p style='font-size: 0.9rem;'>Made with ❤️ for your team | Powered by AI</p>
</div>
""", unsafe_allow_html=True)
