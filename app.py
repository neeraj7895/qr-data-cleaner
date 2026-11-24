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
# Simple clean BLACK theme (CSS only - no logic depends on this)
# Tip: also add .streamlit/config.toml with [theme] base="dark" so tables
# and inputs render dark too.
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, .stApp, .stMarkdown, p, label, li, h1, h2, h3, h4 {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    /* Main background */
    .stApp {
        background: #0a0a0a;
        color: #e5e5e5;
    }

    header[data-testid="stHeader"] {
        background: #0a0a0a;
    }

    .stApp p, .stApp label, .stApp li, .stApp span,
    div[data-testid="stCaptionContainer"] {
        color: #d4d4d4;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #000000;
        border-right: 1px solid #1f1f1f;
    }

    section[data-testid="stSidebar"] > div {
        padding: 1rem 1.25rem;
    }

    .sidebar-logo {
        color: #ffffff;
        font-size: 1.15rem;
        font-weight: 700;
        padding: 0.75rem 0 1rem 0;
        border-bottom: 1px solid #1f1f1f;
        margin-bottom: 1rem;
    }

    .sidebar-nav-item {
        padding: 0.6rem 0.25rem;
        color: #a3a3a3;
        font-size: 0.9rem;
        border-bottom: 1px solid #171717;
    }

    /* Top bar */
    .top-bar {
        background: #111111;
        border: 1px solid #262626;
        border-radius: 8px;
        padding: 0.9rem 1.25rem;
        margin-bottom: 1.25rem;
    }

    .top-bar-title {
        font-size: 1.05rem;
        font-weight: 600;
        color: #ffffff !important;
        margin: 0;
    }

    /* Section headers */
    .section-bar {
        background: #111111;
        color: #ffffff;
        font-weight: 600;
        font-size: 0.95rem;
        padding: 0.8rem 1.1rem;
        border: 1px solid #262626;
        border-left: 3px solid #ffffff;
        border-radius: 6px;
        margin: 1.25rem 0 0.75rem 0;
    }

    .panel {
        background: #111111;
        border: 1px solid #262626;
        border-radius: 6px;
        padding: 1.25rem;
    }

    /* Buttons - white on black, simple */
    .stButton > button, .stDownloadButton > button {
        background: #ffffff;
        color: #000000;
        border: 1px solid #ffffff;
        padding: 0.6rem 1.5rem;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.9rem;
        width: 100%;
    }

    .stButton > button p, .stDownloadButton > button p {
        color: #000000 !important;
    }

    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #d4d4d4;
        border-color: #d4d4d4;
        color: #000000;
    }

    /* File uploader */
    section[data-testid="stFileUploader"],
    div[data-testid="stFileUploaderDropzone"] {
        background: #111111;
        border-radius: 6px;
    }

    div[data-testid="stFileUploaderDropzone"] {
        border: 1px dashed #404040;
    }

    section[data-testid="stFileUploader"] button {
        background: #262626;
        color: #ffffff;
        border: 1px solid #404040;
    }

    /* Metrics */
    div[data-testid="stMetric"] {
        background: #111111;
        padding: 1rem;
        border-radius: 6px;
        border: 1px solid #262626;
    }

    div[data-testid="stMetricLabel"] p {
        color: #a3a3a3 !important;
        font-size: 0.8rem;
        font-weight: 600;
    }

    div[data-testid="stMetricValue"] {
        color: #ffffff;
        font-weight: 700;
        font-size: 1.5rem;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.25rem;
        background: transparent;
        border-bottom: 1px solid #262626;
    }

    .stTabs [data-baseweb="tab"] {
        padding: 0.75rem 1.25rem;
        background: transparent;
    }

    .stTabs [data-baseweb="tab"] p {
        color: #a3a3a3;
        font-weight: 600;
    }

    .stTabs [aria-selected="true"] p {
        color: #ffffff;
    }

    .stTabs [data-baseweb="tab-highlight"] {
        background-color: #ffffff;
    }

    /* Expander */
    div[data-testid="stExpander"] details {
        background: #111111;
        border: 1px solid #262626;
        border-radius: 6px;
    }

    .footer {
        text-align: center;
        padding: 1.5rem;
        margin-top: 2rem;
    }

    .footer p {
        color: #737373 !important;
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

    # FIX: year-first dates (e.g. '1990-03-05 00:00:00') - read as year-month-day
    # so day and month are not swapped when the day is 12 or lower
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T].*)?$", str(x).strip())
    if m:
        try:
            dt = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return "'" + dt.strftime("%d-%m-%Y")
        except ValueError:
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

# ---------------------------------------------------------------------------
# NEW (QR Data Cleaner tab only): add missing acceptance headers at the end
# ---------------------------------------------------------------------------
QR_ACCEPTANCE_HEADERS = [
    "Mobile No", "First Name", "Middle Name", "Last Name", "Entity Name", "DOB", "PAN",
    "Aadhar No", "GST", "Address Line 1", "Address Line 2", "Pincode", "City", "State",
    "DOI", "Account No", "Account Holder Name", "Account Type", "Account Sub Type",
    "IFSC Code", "Account Opening Date", "MCC", "Turnover Type", "Acceptance Type",
    "Ownership Type", "State Code (GST)", "Latitude", "Longitude", "District",
    "Branch Name", "Center Name", "Email ID", "Bank Cust  ID", "Card Number", "Card Name",
    "Card Expiry", "Merchant ID", "Store ID", "Terminal ID", "Brand", "Legal", "Franchise",
    "Merchant Type", "Merchant Genre", "Onboarding Type"
]


def add_missing_qr_headers(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """
    Add any header from QR_ACCEPTANCE_HEADERS that is not in the file.
    Existing columns are untouched. Missing ones are added at the END,
    in the order of the list above, with empty data (header only).
    """
    def _norm(c):
        return re.sub(r"[\s_]+", "", str(c)).lower()

    existing = {_norm(c) for c in df.columns}
    added = []
    df = df.copy()
    for header in QR_ACCEPTANCE_HEADERS:
        if _norm(header) not in existing:
            df[header] = ""
            existing.add(_norm(header))
            added.append(header)

    logs = [f"✓ Added missing header(s) at end: {', '.join(added)}"] if added else []
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
# NEW: P2PM FORMAT CLEANING (separate from the old QR logic above)
# ---------------------------------------------------------------------------

P2PM_SHEET_NAME = "Customer Import"

# Columns whose values are cleared (column is kept, data emptied)
P2PM_CLEAR_COLUMNS = [
    "Spouse Name", "Seller Identifier", "MCC", "GST", "Date of Incorporation",
    "Card Number", "Card Name", "Card Expiry", "Merchant ID", "Store ID",
    "Terminal ID", "Brand", "Legal", "Franchise", "Merchant Type",
    "Merchant Genre", "Email ID", "Aadhar Number", "Father Name",
    "Branch Name", "Center Name", "Bank Cust ID"
]


def norm_col(c) -> str:
    """Normalize a column name: lowercase, remove spaces/underscores (handles double spaces)."""
    return re.sub(r"[\s_]+", "", str(c)).lower()


def find_col(df: pd.DataFrame, name: str) -> Optional[str]:
    target = norm_col(name)
    for c in df.columns:
        if norm_col(c) == target:
            return c
    return None


def find_header_row(raw: pd.DataFrame, max_scan: int = 25) -> Optional[int]:
    """P2PM files have note rows on top. Find the row that contains 'Mobile No'."""
    for i in range(min(max_scan, len(raw))):
        values = [norm_col(v) for v in raw.iloc[i].tolist() if pd.notna(v)]
        if "mobileno" in values:
            return i
    return None


def load_p2pm(file) -> Tuple[Optional[pd.DataFrame], Optional[dict], str]:
    """
    Load a P2PM file. Returns (dataframe, template_info, error).
    template_info = {'bytes', 'sheet', 'header_row'} for .xlsx files (used to keep the format).
    """
    file_extension = file.name.split('.')[-1].lower()
    try:
        file.seek(0)
        sheet = None
        if file_extension == 'csv':
            raw = pd.read_csv(file, dtype=str, header=None)
        else:
            xls = pd.ExcelFile(file)
            sheet = next(
                (s for s in xls.sheet_names if norm_col(s) == norm_col(P2PM_SHEET_NAME)),
                xls.sheet_names[0]
            )
            raw = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=None)

        header_idx = find_header_row(raw)
        if header_idx is None:
            return None, None, "Header row with 'Mobile No' not found"

        headers = raw.iloc[header_idx].tolist()
        keep_positions = [i for i, h in enumerate(headers) if pd.notna(h) and str(h).strip() != ""]
        df = raw.iloc[header_idx + 1:, keep_positions].copy()
        df.columns = [re.sub(r"\s+", " ", str(headers[i])).strip() for i in keep_positions]
        df = df.dropna(how="all").reset_index(drop=True)

        template = None
        if file_extension == 'xlsx':
            template = {
                'bytes': file.getvalue(),
                'sheet': sheet,
                'header_row': header_idx + 1  # Excel row number (1-based)
            }

        logger.info(f"Loaded P2PM file: {file.name} (header at row {header_idx + 1})")
        return df, template, ""
    except Exception as e:
        logger.error(f"Error loading P2PM file {file.name}: {str(e)}")
        return None, None, str(e)


# ----- P2PM-only rules -------------------------------------------------------

P2PM_TITLES = ["MR.", "MS.", "MRS."]

# Words people type in the Title column / in front of a name -> one of the 3 titles
TITLE_WORD_MAP = {
    "MR": "MR.", "SHRI": "MR.", "SHREE": "MR.", "SRI": "MR.", "SH": "MR.", "MASTER": "MR.",
    "MRS": "MRS.", "SMT": "MRS.", "SHRIMATI": "MRS.", "SHREEMATI": "MRS.",
    "MS": "MS.", "MISS": "MS.", "KUM": "MS.", "KUMARI": "MS.",
}

# Name-ending hints (checked on the FIRST word of the name only)
FEMALE_ENDINGS = ("ben", "bahen", "behen", "bai", "devi", "kumari", "kaur")
MALE_ENDINGS = ("bhai", "kumar", "singh", "lal", "chand", "bhau", "das", "nath", "ram", "dev")
MALE_NAMES_ENDING_A_I = {
    "ravi", "hari", "rishi", "mani", "muni", "giri", "ali", "sai", "krishna", "rama",
    "shiva", "aditya", "chaitanya", "surya", "satya", "kanha", "bhanu", "joshi", "ji",
    "abhi", "sunny", "bunty", "monty", "nikki", "vicky", "ricky", "pintu", "raja", "mohammadi",
}
FEMALE_NAMES_OTHER = {
    "kiran", "jasmin", "yasmin", "nasrin", "parveen", "shirin", "mehrunnisa", "sharmin",
    "noor", "nur", "gul", "roshan", "nirmal", "kamal", "sonal", "payal", "kajal", "komal",
    "hetal", "minal", "rupal", "bijal", "chhaya", "dimple", "jinal", "khushbu", "mittal",
    "nidhi", "foram", "dhwani", "urvashi", "trupti", "pooja", "rinkal", "sejal", "shital",
    "vaishali", "dharmishtha", "jagruti", "jyoti", "falguni",
}

PAN_PATTERN = re.compile(r"^[A-Z]{5}[0-9]{4}[A-Z]$")


def is_blank(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip().lower() in ("", "nan", "none")


def strict_clean_text(x) -> str:
    """Only letters, numbers and single spaces. No comma, dot, slash or any other sign."""
    if is_blank(x):
        return ""
    x_str = re.sub(r"[^A-Za-z0-9 ]", " ", str(x))
    return re.sub(r"\s+", " ", x_str).strip()


def normalize_title(x) -> str:
    """Return MR. / MS. / MRS. if the value can be understood, else ''."""
    if is_blank(x):
        return ""
    raw = str(x).strip().upper()
    if re.match(r"^M\s*/\s*S\b", raw):        # M/S = firm, not a person title
        return ""
    letters = re.sub(r"[^A-Z]", "", raw)
    return TITLE_WORD_MAP.get(letters, "")


def guess_gender(first_word: str) -> str:
    """Rough guess from the first word of a name: 'F', 'M' or ''."""
    w = re.sub(r"[^a-z]", "", str(first_word).lower())
    if len(w) < 2:
        return ""
    if w in FEMALE_NAMES_OTHER:
        return "F"
    if w in MALE_NAMES_ENDING_A_I:
        return "M"
    if w.endswith(FEMALE_ENDINGS):
        return "F"
    if w.endswith(MALE_ENDINGS):
        return "M"
    if w.endswith(("a", "i", "ee")):
        return "F"
    return "M"


def p2pm_format_date(x) -> str:
    """
    P2PM date rule -> 'dd-mm-yyyy (text with leading apostrophe).
    Handles Excel dates like '1966-04-15 00:00:00' (year first) without
    swapping day/month, Excel serial numbers, and dd/mm/yyyy text.
    """
    if is_blank(x):
        return ""
    s = str(x).strip().lstrip("'")

    # Year-first: 1966-04-15 / 1966-04-15 00:00:00 / 1966/04/15
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T].*)?$", s)
    if m:
        try:
            dt = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return "'" + dt.strftime("%d-%m-%Y")
        except ValueError:
            return s

    # Excel serial number (e.g. 36000 or 36000.0)
    if re.match(r"^\d{4,6}(\.0+)?$", s):
        try:
            dt = pd.to_datetime("1899-12-30") + pd.to_timedelta(int(float(s)), unit="D")
            return "'" + dt.strftime("%d-%m-%Y")
        except Exception:
            return s

    # Anything else (16/08/1999, 16-08-1999, 16.08.1999 ...) -> day first
    return format_date(s)


def clean_p2pm_data(df: pd.DataFrame, file_name: str = "") -> Tuple[pd.DataFrame, List[str], List[Dict]]:
    """
    Clean P2PM data.
    Returns (df, logs, issues). df has a helper column '__row_id' used for the
    issues report; it is not written to Excel.
    """
    logs = []
    issues = []
    df = df.copy()

    try:
        df["__row_id"] = [f"{file_name}|{i}" for i in range(len(df))]

        def add_issue(row_id, mobile, column, original, issue):
            issues.append({
                "__row_id": row_id, "File Name": file_name, "Mobile No": mobile,
                "Column": column, "Original Value": "" if is_blank(original) else str(original),
                "Issue": issue
            })

        # Mobile No - remove duplicates, then clean (same rule as QR)
        mobile_col = find_col(df, "Mobile No")
        if mobile_col:
            before = len(df)
            df = df.drop_duplicates(subset=[mobile_col], keep="first")
            after = len(df)
            if before > after:
                logs.append(f"✓ Removed {before - after} duplicate mobile numbers")
            df[mobile_col] = df[mobile_col].apply(
                lambda v: clean_mobile_number(re.sub(r"\.0$", "", str(v).strip()))
            )
            logs.append("✓ Cleaned mobile numbers")

        def mobile_of(idx):
            return df.at[idx, mobile_col] if mobile_col else ""

        # Dates -> 'dd-mm-yyyy
        for date_name in ["DOB", "Account Opening Date"]:
            date_col = find_col(df, date_name)
            if date_col:
                df[date_col] = df[date_col].apply(p2pm_format_date)
                logs.append(f"✓ Formatted date column: {date_col}")

        # Account Number - same rule as QR 'Account No'
        acc_col = find_col(df, "Account Number")
        if acc_col:
            df[acc_col] = df[acc_col].apply(format_account_number)
            logs.append(f"✓ Formatted {acc_col}")

        ent_col = find_col(df, "Entity Name")
        first_col = find_col(df, "First Name")
        middle_col = find_col(df, "Middle Name")
        last_col = find_col(df, "Last Name")
        accname_col = find_col(df, "Account Name")
        title_col = find_col(df, "Title")
        marital_col = find_col(df, "Marital Status")
        personal_cols = [c for c in [first_col, middle_col, last_col] if c]

        if ent_col is None:
            ent_col = "Entity Name"
            df[ent_col] = ""

        # Keep original title & find a title written in front of the name (e.g. "Mrs Dhruvi")
        original_titles = df[title_col].copy() if title_col else None
        name_prefix_title = {}
        for idx in df.index:
            for c in [first_col, ent_col]:
                if not c or is_blank(df.at[idx, c]):
                    continue
                words = strict_clean_text(df.at[idx, c]).split(" ")
                if len(words) > 1 and words[0].upper() in TITLE_WORD_MAP:
                    name_prefix_title.setdefault(idx, TITLE_WORD_MAP[words[0].upper()])
                    df.at[idx, c] = " ".join(words[1:])

        # Names - strict clean: no special characters at all
        for c in personal_cols + [ent_col] + ([accname_col] if accname_col else []):
            df[c] = df[c].apply(strict_clean_text)
        logs.append("✓ Cleaned names (only letters, numbers and spaces)")

        # Entity Name rule
        filled_entity = 0
        cleared_personal = 0
        for idx in df.index:
            if df.at[idx, ent_col] != "":
                if any(df.at[idx, c] != "" for c in personal_cols):
                    cleared_personal += 1
                for c in personal_cols:
                    df.at[idx, c] = ""
            else:
                full = " ".join(df.at[idx, c] for c in personal_cols if df.at[idx, c] != "")
                full = strict_clean_text(full)
                if full:
                    df.at[idx, ent_col] = full
                    filled_entity += 1
                else:
                    add_issue(df.at[idx, "__row_id"], mobile_of(idx), ent_col, "",
                              "Entity Name missing (mandatory) - no First/Middle/Last Name to copy")
        if cleared_personal:
            logs.append(f"✓ Entity Name present: cleared First/Middle/Last Name in {cleared_personal} rows")
        if filled_entity:
            logs.append(f"✓ Entity Name empty: filled from First/Middle/Last Name in {filled_entity} rows")

        # Account Name - fill from Entity Name if empty
        if accname_col:
            mask = df[accname_col] == ""
            df.loc[mask, accname_col] = df.loc[mask, ent_col]
            logs.append("✓ Account Name cleaned (no special characters), empty ones filled from Entity Name")

        # Title - only MR. / MS. / MRS.
        if title_col:
            fixed, guessed = 0, 0
            for idx in df.index:
                orig = original_titles.at[idx]
                title = normalize_title(orig) or name_prefix_title.get(idx, "")
                if title:
                    if str(orig).strip() != title:
                        fixed += 1
                else:
                    first_word = df.at[idx, ent_col].split(" ")[0] if df.at[idx, ent_col] else ""
                    gender = guess_gender(first_word)
                    married = marital_col and str(df.at[idx, marital_col]).strip().lower() == "married"
                    if gender == "F":
                        title = "MRS." if married else "MS."
                    else:
                        title = "MR."
                    guessed += 1
                    add_issue(df.at[idx, "__row_id"], mobile_of(idx), title_col, orig,
                              f"Title guessed from name '{first_word}' as {title} - please check")
                df.at[idx, title_col] = title
            logs.append(f"✓ Title set to MR./MS./MRS. ({fixed} fixed, {guessed} guessed from name)")

        # IFSC Code - trim spaces, uppercase (unchanged)
        ifsc_col = find_col(df, "IFSC Code")
        if ifsc_col:
            df[ifsc_col] = df[ifsc_col].fillna("").astype(str).str.replace(r"\s+", "", regex=True).str.upper()
            logs.append(f"✓ Cleaned {ifsc_col}")

        # PAN - mandatory, 5 letters + 4 numbers + 1 letter, else blank
        pan_col = find_col(df, "PAN")
        if pan_col:
            valid, removed, missing = 0, 0, 0
            for idx in df.index:
                orig = df.at[idx, pan_col]
                pan = "" if is_blank(orig) else re.sub(r"[^A-Za-z0-9]", "", str(orig)).upper()
                if PAN_PATTERN.match(pan):
                    df.at[idx, pan_col] = pan
                    valid += 1
                else:
                    df.at[idx, pan_col] = ""
                    if pan:
                        removed += 1
                        add_issue(df.at[idx, "__row_id"], mobile_of(idx), pan_col, orig,
                                  "Invalid PAN removed (must be 5 letters + 4 numbers + 1 letter)")
                    else:
                        missing += 1
                        add_issue(df.at[idx, "__row_id"], mobile_of(idx), pan_col, "",
                                  "PAN missing (mandatory)")
            logs.append(f"✓ PAN checked: {valid} valid, {removed} invalid removed, {missing} missing")

        # Clear not-needed columns (header stays, data emptied)
        for name in P2PM_CLEAR_COLUMNS:
            c = find_col(df, name)
            if c:
                df[c] = ""
                logs.append(f"✓ Cleared: {c}")

        return df, logs, issues

    except Exception as e:
        logger.error(f"Error in clean_p2pm_data: {str(e)}")
        logs.append(f"❌ Error: {str(e)}")
        return df, logs, issues


def read_extra_dropdowns(xlsx_bytes: bytes, sheet_name: str) -> List[Tuple[str, str]]:
    """
    Read the dropdowns that point to another sheet (e.g. 'Account Type'!$AI$2:$AI$4).
    Excel stores these in a special section that openpyxl drops, so we read them
    from the original file and add them back. Returns [(formula, cell_range)].
    """
    import zipfile
    try:
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
            workbook_xml = z.read("xl/workbook.xml").decode("utf-8")
            rels_xml = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")

            rid = None
            for m in re.finditer(r"<sheet\b[^>]*>", workbook_xml):
                tag = m.group(0)
                name = re.search(r'name="([^"]*)"', tag)
                if name and name.group(1) == sheet_name:
                    r = re.search(r'r:id="([^"]*)"', tag)
                    rid = r.group(1) if r else None
                    break
            if rid is None:
                return []

            target = None
            for m in re.finditer(r"<Relationship\b[^>]*>", rels_xml):
                tag = m.group(0)
                if f'Id="{rid}"' in tag:
                    target = re.search(r'Target="([^"]*)"', tag).group(1)
                    break
            if target is None:
                return []
            path = target.lstrip("/") if target.startswith("/") else "xl/" + target
            sheet_xml = z.read(path).decode("utf-8")

        result = []
        for m in re.finditer(r"<x14:dataValidation\b.*?</x14:dataValidation>", sheet_xml, re.S):
            block = m.group(0)
            if 'type="list"' not in block:
                continue
            f = re.search(r"<xm:f>(.*?)</xm:f>", block, re.S)
            ref = re.search(r"<xm:sqref>(.*?)</xm:sqref>", block, re.S)
            if f and ref:
                formula = (f.group(1).replace("&amp;", "&").replace("&quot;", '"')
                           .replace("&apos;", "'").replace("&lt;", "<").replace("&gt;", ">"))
                result.append((formula, ref.group(1).strip()))
        return result
    except Exception as e:
        logger.error(f"Could not read extra dropdowns: {str(e)}")
        return []


def write_p2pm_excel(df: pd.DataFrame, template: Optional[dict],
                     highlight: Optional[Dict[int, List[str]]] = None) -> io.BytesIO:
    """
    Write cleaned data back into the original P2PM template. Keeps note rows,
    header row, helper sheets and ALL dropdowns. Cells with issues are
    highlighted yellow so they are easy to check. Falls back to a plain sheet.
    highlight = {dataframe position: [column names]}
    """
    from openpyxl.styles import PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    highlight = highlight or {}
    df = df.drop(columns=["__row_id"], errors="ignore")

    if template is not None:
        try:
            wb = load_workbook(io.BytesIO(template['bytes']))
            ws = wb[template['sheet']]
            header_row = template['header_row']

            col_map = {}
            for cell in ws[header_row]:
                if cell.value is not None and str(cell.value).strip() != "":
                    col_map[norm_col(cell.value)] = cell.column

            # remove old data rows
            if ws.max_row > header_row:
                ws.delete_rows(header_row + 1, ws.max_row - header_row)

            columns = list(df.columns)
            for pos, row in enumerate(df.itertuples(index=False)):
                r_idx = header_row + 1 + pos
                flagged = {norm_col(c) for c in highlight.get(pos, [])}
                for col_name, value in zip(columns, row):
                    target = col_map.get(norm_col(col_name))
                    if target is None:
                        continue
                    if is_blank(value):
                        value = None
                    cell = ws.cell(row=r_idx, column=target, value=value)
                    if norm_col(col_name) in flagged:
                        cell.fill = yellow

            # add back the dropdowns that openpyxl drops
            for formula, cell_range in read_extra_dropdowns(template['bytes'], template['sheet']):
                dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                                    showErrorMessage=True, showInputMessage=True)
                dv.add(cell_range)
                ws.add_data_validation(dv)

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            return output
        except Exception as e:
            logger.error(f"Template write failed, using plain sheet: {str(e)}")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.fillna("").to_excel(writer, index=False, sheet_name=P2PM_SHEET_NAME)
    output.seek(0)
    return output


def render_p2pm_tab():
    st.markdown('<div class="section-bar">🆕 New P2PM Upload</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.caption("Upload one or multiple P2PM files. One file gives one cleaned file; "
               "multiple files are cleaned and merged into one file in the same P2PM format "
               "(dropdowns kept). Cells that need checking are highlighted yellow.")

    p2pm_files = st.file_uploader(
        "Select P2PM file(s) (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="file_uploader_p2pm"
    )

    if p2pm_files:
        st.caption(f"{len(p2pm_files)} file(s) selected")

    st.button("🧹 Clean P2PM Files", key="process_p2pm",
              on_click=lambda: st.session_state.update({"_run_p2pm": True}))
    st.markdown('</div>', unsafe_allow_html=True)

    if p2pm_files and st.session_state.get("_run_p2pm"):
        st.session_state["_run_p2pm"] = False
        process_p2pm_files(p2pm_files)


def process_p2pm_files(uploaded_files):
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        all_logs = []
        all_dfs = []
        all_issues = []
        failed = []
        template = None

        for idx, file in enumerate(uploaded_files):
            status_text.text(f"Processing {idx + 1}/{len(uploaded_files)}: {file.name}")
            df, file_template, error = load_p2pm(file)

            if df is None:
                failed.append({'File Name': file.name, 'Error/Issue': error})
            else:
                cleaned_df, logs, issues = clean_p2pm_data(df, file.name)
                all_dfs.append(cleaned_df)
                all_issues.extend(issues)
                all_logs.append(f"📄 {file.name} — {len(cleaned_df)} rows")
                all_logs.extend(logs)
                if template is None and file_template is not None:
                    template = file_template

            progress_bar.progress((idx + 1) / (len(uploaded_files) + 1))

        if not all_dfs:
            st.error("❌ No files could be processed")
            if failed:
                st.dataframe(pd.DataFrame(failed), use_container_width=True)
            progress_bar.empty()
            status_text.empty()
            return

        if len(all_dfs) == 1:
            final_df = all_dfs[0].reset_index(drop=True)
        else:
            status_text.text("Merging files...")
            final_df = pd.concat(all_dfs, ignore_index=True, sort=False)
            mobile_col = find_col(final_df, "Mobile No")
            if mobile_col:
                before = len(final_df)
                final_df = final_df.drop_duplicates(subset=[mobile_col], keep="first").reset_index(drop=True)
                if before > len(final_df):
                    all_logs.append(f"✓ Removed {before - len(final_df)} duplicates from merged data")

        # Map issues to their row in the output file
        header_row = template['header_row'] if template else 1
        pos_of = {rid: pos for pos, rid in enumerate(final_df["__row_id"])}
        highlight = {}
        issue_rows = []
        for item in all_issues:
            pos = pos_of.get(item["__row_id"])
            if pos is None:   # row removed as duplicate
                continue
            highlight.setdefault(pos, []).append(item["Column"])
            row = {k: v for k, v in item.items() if k != "__row_id"}
            row = {"Excel Row": header_row + 1 + pos, **row}
            issue_rows.append(row)
        issues_df = pd.DataFrame(issue_rows)

        status_text.text("Creating Excel file...")
        final_output = write_p2pm_excel(final_df, template, highlight)
        preview_df = final_df.drop(columns=["__row_id"], errors="ignore")
        progress_bar.progress(1.0)

        if len(uploaded_files) == 1:
            base = os.path.splitext(uploaded_files[0].name)[0]
            out_name = f"Cleaned_{base}.xlsx"
        else:
            out_name = f"P2PM_Cleaned_Merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        st.success("✅ P2PM file(s) processed successfully!")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Files Processed", len(all_dfs))
        with col2:
            st.metric("Failed", len(failed))
        with col3:
            st.metric("Total Rows", len(preview_df))
        with col4:
            st.metric("Cells to Check", len(issues_df))

        st.markdown('<div class="section-bar">📄 Preview</div>', unsafe_allow_html=True)
        st.dataframe(preview_df.head(10).fillna(""), use_container_width=True)

        if failed:
            st.warning(f"⚠️ {len(failed)} file(s) could not be processed")
            st.dataframe(pd.DataFrame(failed), use_container_width=True)

        if not issues_df.empty:
            st.markdown('<div class="section-bar">⚠️ Please Check (highlighted yellow in file)</div>',
                        unsafe_allow_html=True)
            st.dataframe(issues_df, use_container_width=True)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                "⬇️ Download Cleaned P2PM File",
                data=final_output.getvalue(),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        if not issues_df.empty:
            with col_dl2:
                issues_buffer = io.BytesIO()
                with pd.ExcelWriter(issues_buffer, engine="openpyxl") as writer:
                    issues_df.to_excel(writer, index=False, sheet_name="Check_Report")
                st.download_button(
                    "⬇️ Download Check Report",
                    data=issues_buffer.getvalue(),
                    file_name=f"P2PM_Check_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        with st.expander("📝 View Cleaning Logs"):
            for log in all_logs:
                st.write(log)

        progress_bar.empty()
        status_text.empty()

    except Exception as e:
        logger.error(f"Error processing P2PM files: {str(e)}")
        st.error(f"❌ Error: {str(e)}")
        progress_bar.empty()
        status_text.empty()

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
        <div class='sidebar-nav-item'>🆕 New P2PM Format Cleaning</div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown(
            "<div style='color:#737373; font-size:0.75rem; text-align:center;'>Version 3.1<br>Updated: November 2026</div>",
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
            cleaned_df, header_logs = add_missing_qr_headers(cleaned_df)
            all_logs.extend(header_logs)
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

            merged_df, header_logs = add_missing_qr_headers(merged_df)
            all_logs.extend(header_logs)

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

    tab1, tab2, tab3 = st.tabs(["📁 QR Data Cleaner", "📂 Bulk Master Creator", "🆕 New P2PM"])

    with tab1:
        render_data_cleaner_tab()

    with tab2:
        render_bulk_processor_tab()

    with tab3:
        render_p2pm_tab()

    st.markdown("""
    <div class="footer">
        <p>QR Data Cleaner Pro v3.1 | © 2026</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
