import streamlit as st
import pandas as pd
import re
import io
import os
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation


# ------------------ CLEANING FUNCTION ------------------ #
def clean_data(df, source_file=None):
    logs = []

    # Normalize column names (strip spaces/newlines)
    df.rename(columns=lambda x: str(x).strip().replace("\n", " ").strip(), inplace=True)

    # 1. Remove rows where Mobile No is blank
    if "Mobile No" in df.columns:
        before = len(df)
        df = df[df["Mobile No"].notna() & (df["Mobile No"].astype(str).str.strip() != "")]
        after = len(df)
        logs.append(f"Removed {before - after} rows with blank Mobile No")
    
    # 2. Remove duplicate Mobile numbers → Keep first occurrence
    if "Mobile No" in df.columns:
        before = len(df)
        df = df.drop_duplicates(subset=["Mobile No"], keep="first").copy()
        after = len(df)
        logs.append(f"Removed {before - after} duplicate row(s) by Mobile No (kept first occurrence)")

    # 3. Dates → format 'dd-mm-yyyy with prefix '
    for col in ["DOB", "DOI", "Account Opening Date"]:
        if col in df.columns:
            def format_date(x):
                if pd.isna(x) or str(x).strip() == "":
                    return ""
                dt = pd.to_datetime(x, dayfirst=True, errors="coerce")
                if pd.isna(dt):
                    return str(x)   # keep original if invalid
                return "'" + dt.strftime("%d-%m-%Y")
            df[col] = df[col].apply(format_date)

    # 4. Aadhaar No → add prefix `'`, skip NaN/blank, remove .0
    for col in ["Aadhar No", "Aadhaar No"]:
        if col in df.columns:
            df[col] = df[col].astype(str).apply(
                lambda x: "'" + x.replace(".0", "") if x.strip() != "" and x.lower() != "nan" else ""
            )

    # 5. Account No → add prefix `'`, skip NaN/blank, remove .0
    if "Account No" in df.columns:
        df["Account No"] = df["Account No"].astype(str).apply(
            lambda x: "'" + x.lstrip("'").replace(".0", "") 
            if x.strip() != "" and x.lower() != "nan" else ""
        )

    # 6. Address cleanup
    for col in ["Address Line 1", "Address Line 2"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(
                r"[,\.\#\&\):\(]", " ", regex=True
            ).str.strip()
            df[col] = df[col].replace("nan", "").replace("NaN", "").replace("None", "")

    # If Address Line 2 is blank/NaN → copy from Address Line 1
    if "Address Line 1" in df.columns and "Address Line 2" in df.columns:
        df["Address Line 2"] = df.apply(
            lambda r: r["Address Line 1"]
            if (pd.isna(r["Address Line 2"]) or str(r["Address Line 2"]).strip() == "")
               and str(r["Address Line 1"]).strip() != ""
            else r["Address Line 2"],
            axis=1
        )

    # 7. Names cleanup
    name_cols = ["First Name", "Middle Name", "Last Name", "Entity Name", "Account Holder Name"]
    for col in name_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r"[\/&#,.;']", " ", regex=True).str.strip()
            df[col] = df[col].replace("nan", "").replace("NaN", "").replace("None", "")

    # 8. Entity vs Personal Names
    if "Entity Name" in df.columns:
        entity_mask = df["Entity Name"].notna() & (df["Entity Name"].str.strip() != "")
        for col in ["First Name", "Middle Name", "Last Name"]:
            if col in df.columns:
                df.loc[entity_mask, col] = ""   # clear personal names if entity present

    # 9. Branch Name → Replace all values with "HO Branch"
    if "Branch Name" in df.columns:
        df["Branch Name"] = "HO Branch"
        logs.append("Replaced all values in 'Branch Name' with 'HO Branch'")

    # 10. Add Source File column if multiple uploads
    if source_file:
        df["Source_File"] = source_file

    # 11. Clear unwanted columns
    clear_cols = [
        "Turnover Type", "Acceptance Type", "Ownership Type", "MCC", "Email ID", "Source_File", "Bank Cust ID",
        "State Code (GST)", "Latitude", "Longitude", "District"
    ]
    for col in clear_cols:
        matches = [c for c in df.columns if c.lower().replace(" ", "") == col.lower().replace(" ", "")]
        for m in matches:
            df[m] = ""   # clear values but keep header
            logs.append(f"Cleared all data from column: {m}")

    return df, logs


# ---------------- ADD DROPDOWNS TO EXCEL ---------------- #
def add_dropdowns(excel_bytes, sheet_name="Cleaned"):
    excel_bytes.seek(0)
    wb = load_workbook(excel_bytes)
    ws = wb[sheet_name]

    # Dropdown options
    account_types = ["Savings", "Current", "Loan", "Fixed Deposit"]
    sub_types = ["Regular", "Premium", "Zero Balance", "Overdraft"]

    # Account Type dropdown
    if "Account Type" in [c.value for c in ws[1]]:
        col_idx = [c.value for c in ws[1]].index("Account Type") + 1
        dv_type = DataValidation(type="list", formula1='"' + ",".join(account_types) + '"', allow_blank=True)
        ws.add_data_validation(dv_type)
        dv_type.add(f"{ws.cell(row=2, column=col_idx).coordinate}:{ws.cell(row=ws.max_row, column=col_idx).coordinate}")

    # Account Sub Type dropdown
    if "Account Sub Type" in [c.value for c in ws[1]]:
        col_idx = [c.value for c in ws[1]].index("Account Sub Type") + 1
        dv_sub = DataValidation(type="list", formula1='"' + ",".join(sub_types) + '"', allow_blank=True)
        ws.add_data_validation(dv_sub)
        dv_sub.add(f"{ws.cell(row=2, column=col_idx).coordinate}:{ws.cell(row=ws.max_row, column=col_idx).coordinate}")

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output


# --- Custom CSS ---
st.markdown("""
    <style>
    .title {
        font-size:28px !important;
        font-weight:bold;
        color:#FF9800;
        background:#1E88E5;
        padding:12px;
        border-radius:8px;
        text-align:center;
        display:flex;
        align-items:center;
        justify-content:center;
        gap:10px;
    }
    .broom {
        display:inline-block;
        animation: sweep 3.2s infinite;
    }
    .streamlit-expanderHeader {
        font-size: 18px !important;
        font-weight: bold !important;
        color: white !important;
        background: linear-gradient(90deg, #1E88E5, #42A5F5);
        border-radius: 6px;
        padding: 8px 12px;
    }
    .streamlit-expanderContent {
        background: #f9f9f9;
        border-left: 3px solid #1E88E5;
        padding: 10px;
        border-radius: 4px;
    }
    </style>
""", unsafe_allow_html=True)


# --- Header ---
st.markdown('<div class="title"><span class="broom">🧹</span> Operations QR Data Cleaner</div>', unsafe_allow_html=True)
st.write("Upload either a **single Excel file** or **multiple files** below for cleaning.")


col1, col2 = st.columns(2)

with col1:
    single_file = st.file_uploader("📂 Upload Single Excel File", type=["xlsx","xls"], accept_multiple_files=False)

with col2:
    multiple_files = st.file_uploader("📂 Upload Multiple Excel Files", type=["xlsx","xls"], accept_multiple_files=True)


# ------------------ PROCESSING ------------------ #
def load_excel(file):
    ext = os.path.splitext(file.name)[1].lower()
    if ext == ".xls":
        return pd.read_excel(file, sheet_name=0, dtype={"Account No": str}, engine="xlrd")
    else:
        return pd.read_excel(file, sheet_name=0, dtype={"Account No": str}, engine="openpyxl")


if single_file:
    df = load_excel(single_file)
    cleaned_df, logs = clean_data(df)

    st.success("✅ Single file processed successfully!")

    # Download with Dropdowns
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cleaned_df.to_excel(writer, index=False, sheet_name="Cleaned")

    final_output = add_dropdowns(output, sheet_name="Cleaned")
    st.download_button("⬇️ Download Cleaned File", final_output.getvalue(), file_name="Cleaned_Single.xlsx")

    # Logs in Expander
    with st.expander("📝 View Cleaning Logs"):
        for log in logs:
            st.write("✔️", log)

elif multiple_files:
    all_dfs = []
    all_logs = []
    for f in multiple_files:
        df = load_excel(f)
        cleaned_df, logs = clean_data(df, source_file=f.name)
        all_dfs.append(cleaned_df)
        all_logs.extend([f"[{f.name}] {log}" for log in logs])

    merged_df = pd.concat(all_dfs, ignore_index=True)

    # Remove duplicate Mobile Nos in merged output
    if "Mobile No" in merged_df.columns:
        before = len(merged_df)
        merged_df = merged_df.drop_duplicates(subset=["Mobile No"], keep="first").copy()
        after = len(merged_df)
        all_logs.append(f"(Merged) Removed {before - after} duplicate row(s) by Mobile No (kept first occurrence)")

    st.success("✅ Multiple files processed and merged successfully!")

    # Download with Dropdowns
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        merged_df.to_excel(writer, index=False, sheet_name="Cleaned_Merged")

    final_output = add_dropdowns(output, sheet_name="Cleaned_Merged")
    st.download_button("⬇️ Download Merged Cleaned File", final_output.getvalue(), file_name="Cleaned_Merged.xlsx")

    # Logs in Expander
    with st.expander("📝 View Cleaning Logs"):
        for log in all_logs:
            st.write("✔️", log)


