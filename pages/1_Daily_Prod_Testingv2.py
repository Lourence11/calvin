import sys
import os
import platform
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
import io
import os
import shutil
import tempfile
import zipfile
import uuid
from datetime import datetime, timedelta
import datetime as dt

# Set page configuration
st.set_page_config(page_title="Daily Prod Automation", page_icon=None, layout="wide")

# Apply consistent premium design
set_premium_style()

# ==========================================
#  Detect Platform and Setup Excel Engine
# ==========================================
def get_platform():
    """Detect the current platform."""
    system = platform.system().lower()
    if system == 'windows':
        return 'windows'
    elif system == 'darwin':
        return 'macos'
    else:
        return 'linux'

PLATFORM = get_platform()

# Try to import openpyxl (cross-platform)
try:
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    st.warning("openpyxl is not installed. Install with: pip install openpyxl")

# Try to import win32com (Windows only)
if PLATFORM == 'windows':
    try:
        import win32com.client
        import pythoncom
        WIN32COM_AVAILABLE = True
    except ImportError:
        WIN32COM_AVAILABLE = False
else:
    WIN32COM_AVAILABLE = False

# Display platform info
if PLATFORM != 'windows':
    st.info(f"🖥️ Running on **{PLATFORM.upper()}** - Using openpyxl for Excel operations")
else:
    st.info(f"🖥️ Running on **WINDOWS** - Using COM automation")

if not WIN32COM_AVAILABLE and PLATFORM == 'windows':
    st.warning("pywin32 not found. Install with: pip install pywin32")

if not OPENPYXL_AVAILABLE:
    st.error("openpyxl is required. Install with: pip install openpyxl")
    st.stop()

def log_step(message, level="info"):
    """Displays a logged step in Streamlit with a timestamp."""
    timestamp = dt.datetime.now().strftime("%H:%M:%S")
    icon = "🔹" if level == "info" else "❌" if level == "error" else "✅"
    st.markdown(f"{icon} **{timestamp}** {message}")

# --- DATE HELPER FOR FILENAMES ---
def get_report_date():
    """Returns (report_date_obj, filename_string) based on user's Monday/Friday logic."""
    today = datetime.now()
    if today.weekday() == 0:  # Monday
        report_date = today - timedelta(days=3) # Use Friday last week
    else:
        report_date = today - timedelta(days=1) # Use yesterday
    return report_date, report_date.strftime("%Y%m%d")


# STATUSES TO EXCLUDE (matches anything starting with these strings)
EXCLUDED_STATUSES = [
    'CONFIRMED', 'PTP_FF', 'SMS SENT - T6 NO RESPONSE',
    'LS VIA SOCMED - T6 NO RESPONSE', 'SMS RECEIVED - UNDER NEGO',
    'SMS RECEIVED - WILL SETTLE', 'NEW', 'BP', 'ENCO', 'UNLOCKED',
    'BULK SMS', 'REACTIVE', 'LOCKED', 'PM',
    'SMS RECEIVED - NO SUCH PERSON (NSP)',
    'SERVICE REQUEST (SR) - FOR COLLECTION'
]


rfd_mapping = {
    "AWAITING_FUNDS": [
        "ALLOTMENT", "REMITTANCE", "ALLOWANCE", "LOANS", "PROFIT", 
        "SALARY", "INCENTIVES", "BACKPAY", "BENEFITS", "PENSION", 
        "COLLECTIONS", "COMMISSION"
    ],
    "EMPLOYMENT_STATUS": [
        'AWAITING TO ONBOARD', 'CONTRACTUAL', 'NEWLY-HIRED', 'CONFLICT WITH EMPLOYER', 
        'UNEMPLOYMENT', 'FREELANCE', 'SUSPENSION', 'ON MATERNITY LEAVE'
    ],
    "EMERGENCY": [
        'HOSPITALIZATION', 'VICTIM OF CALAMITY', 'ACCIDENT', 'DEATH OF RELATIVE'
    ],
    "NO_CAPACITY_TO_PAY": [
        "SHORT OF FUNDS", "FINANCIAL DIFFICULTY", "BUSINESS SLOWDOWN"
    ],
    "PRIORITIZE_OTHER_BILLS_EXPENSES": [
        "LOANS/OTHER DUES", "TUITION", "UTILITY BILLS", "MEDICATION"
    ],
    "LOAN_CLARIFICATIONS": [
        "PAST DUE BALANCE", "MONTHLY AMORTIZATION", "MATURITY DATE", "REMAINING TERM"
    ],
    "PAYMENT_CLARIFICATIONS": [
        "FUNDED ADA", "CLAIMING SETTLED ACCOUNT", "LATE CHARGES", "PAYMENT HISTORY"
    ],
    "PERSONAL_REASON": [
        "SCAM VICTIM", "OUT OF COUNTRY/TOWN", "BUSY SCHEDULE", 
        "CHILDBIRTH", "FAMILY MATTER", "ROBBERY VICTIM"
    ],
    "OTHERS": [
        "DROPPED CALL", "ONLINE BANKING CONCERN", "OVERLOOKED DUEDATE", "REFUSE TO DISCLOSED REASON"
    ]
}

# --- OPTIMIZED: Pre-build a flat lookup list for rfd_mapping (avoids nested loops per row) ---
_RFD_FLAT = []
for _rfd, _subs in rfd_mapping.items():
    for _sub in _subs:
        _RFD_FLAT.append((_sub.lower(), _rfd, _sub))


def classify_remark(remark):
    """Single-row classifier — called via .apply(); logic unchanged."""
    if pd.isna(remark) or str(remark).strip() == "":
        return None, None
    
    remark_lower = str(remark).lower()
    
    if any(keyword in remark_lower for keyword in ["got scammed", "scammed", "scammer"]):
        return "PERSONAL_REASON", "SCAM VICTIM"
    if any(keyword in remark_lower for keyword in ['ootc', 'out of the country', 'abroad']):
        return 'PERSONAL_REASON', 'OUT OF COUNTRY/TOWN'
    if any(keyword in remark_lower for keyword in ['unemployed']):
        return 'EMPLOYMENT_STATUS', 'UNEMPLOYMENT'
    if any(keyword in remark_lower for keyword in ['dropped', 'dropped call', 'call drop', 'call dropped']):
        return 'OTHERS', 'DROPPED CALL'
    if any(keyword in remark_lower for keyword in ['overlooked the account', 'overlooked account']):
        return 'OTHERS', 'OVERLOOKED DUEDATE'
    if any(keyword in remark_lower for keyword in ['financial struggle', 'financial problem']):
        return 'NO_CAPACITY_TO_PAY', 'FINANCIAL DIFFICULTY'
    if any(keyword in remark_lower for keyword in ['other loans', 'other loan']):
        return 'PRIORITIZE_OTHER_BILLS_EXPENSES', 'LOANS/OTHER DUES'
    if any(keyword in remark_lower for keyword in ['bankcruptcy']):
        return 'NO_CAPACITY_TO_PAY', 'BUSINESS SLOWDOWN'
    if any(keyword in remark_lower for keyword in ['maternity leave']):
        return 'EMPLOYMENT_STATUS', 'ON MATERNITY LEAVE'
    if any(keyword in remark_lower for keyword in ['give birth', 'gave birth']):
        return 'EMPLOYMENT_STATUS', 'CHILDBIRTH'
    if any(keyword in remark_lower for keyword in ['medication','undermedication','hospitalization']):
        return 'PRIORITIZE_OTHER_BILLS_EXPENSES', 'MEDICATION'

    # OPTIMIZED: iterate pre-flattened list instead of nested dict loops
    for sub_lower, rfd, sub in _RFD_FLAT:
        if sub_lower in remark_lower:
            return rfd, sub
    
    return None, None


def _build_data_batch(df_data, num_cols, start_row, is_bucket_mode, return_buckets):
    """
    OPTIMIZED: Build the data_batch list using vectorized pandas operations
    instead of row-by-row itertuples.  Returns the same 2-D list that the
    original loop produced so all downstream COM code is untouched.
    """
    nrows = len(df_data)
    row_indices = list(range(start_row, start_row + nrows))  # Excel row numbers

    # ---- Left side: columns 1..num_cols (S.No → SUB-RFD) ----
    # Convert the relevant slice to Python objects in one shot
    left_df = df_data.iloc[:, :num_cols].copy()

    # Coerce Timestamps to tz-aware datetime objects (vectorized)
    for col in left_df.columns:
        if pd.api.types.is_datetime64_any_dtype(left_df[col]):
            left_df[col] = left_df[col].dt.tz_localize('UTC', ambiguous='NaT', nonexistent='NaT')

    # Replace NaT/NaN with "" so COM doesn't choke
    left_records = []
    for row in left_df.itertuples(index=False):
        processed = []
        for v in row:
            if isinstance(v, float) and pd.isna(v):
                processed.append("")
            elif v is pd.NaT or (isinstance(v, pd.Timestamp) and pd.isna(v)):
                processed.append("")
            elif isinstance(v, tuple):
                processed.append(str(v))
            else:
                processed.append(v)
        left_records.append(processed)

    # ---- Right side: 30 formula strings (columns 54..83) ----
    # Build formula column arrays vectorized using list comprehensions (much faster
    # than building them one element at a time inside the row loop).
    r = row_indices  # shorthand

    col_BB = [f'=VLOOKUP($BJ{ri},TEMPLATE_DB[[#All],[Account num]:[Collector]],3,0)' for ri in r]
    col_BC = [f'=$J{ri}' for ri in r]
    col_BD = [f'=VLOOKUP($BJ{ri},TEMPLATE_DB[[Account num]:[Placement]],4,0)' for ri in r]
    col_BE = [f'=VLOOKUP($BJ{ri},TEMPLATE_DB[[#All],[Account num]:[Assign Date]],5,0)' for ri in r]
    col_BF = [f'=VLOOKUP($BD{ri},NOTES!H:J,3,0)' for ri in r]
    col_BG = [""] * nrows
    col_BH = ["MADRID"] * nrows
    col_BI = [""] * nrows
    col_BJ = [f'=$E{ri}' for ri in r]
    col_BK = [
        f'=IF(LEFT($BC{ri},12)="LS VIA EMAIL", VLOOKUP($BJ{ri},TEMPLATE_DB[[#All],[Account num]:[Last Called No.]],25,0), VLOOKUP($BJ{ri},TEMPLATE_DB[[#All],[Account num]:[Email]],25,0))'
        for ri in r
    ]
    col_BL = [f'=VLOOKUP($J{ri},NOTES!AA:AB,2,0)' for ri in r]
    col_BM = [""] * nrows
    col_BN = [
        f'=IFERROR(VLOOKUP($BB{ri},NOTES!V:W,2,0),VLOOKUP(\'VOLARE EXTRACTION\'!$BB{ri},NOTES!U:W,3,0))'
        for ri in r
    ]
    col_BO = [f'=VLOOKUP($J{ri},NOTES!K:L,2,0)' for ri in r]
    col_BP = [f'=VLOOKUP($BO{ri},NOTES!B:E,4,0)' for ri in r]
    col_BQ = [f'=$K{ri}' for ri in r]
    col_BR = [f'=$W{ri}' for ri in r]
    col_BS = [f'=$V{ri}' for ri in r]
    col_BT = [""] * nrows
    col_BU = [f'=VLOOKUP($BJ{ri},TEMPLATE_DB[[Account num]:[new ob]],8,0)' for ri in r]
    col_BV = [f'=VLOOKUP($BJ{ri},TEMPLATE_DB[[#All],[Account num]:[Principal]],9,0)' for ri in r]
    col_BW = [f'=VLOOKUP($BJ{ri},E{ri}:BA{ri},48,0)' for ri in r]
    col_BX = [f'=VLOOKUP($BJ{ri},E{ri}:BA{ri},49,0)' for ri in r]
    col_BY = [""] * nrows
    col_BZ = [
        f'=IF(COUNTIF(NOTES!Q:Q,\'VOLARE EXTRACTION\'!$J{ri}) > 0, "POSITIVE", "NEGATIVE")'
        for ri in r
    ]
    col_CA = [""] * nrows
    col_CB = [""] * nrows
    col_CC = [""] * nrows
    col_CD = [f'=$C{ri}' for ri in r]
    col_CE = [f'=$B{ri}' for ri in r]
    col_CF = [f'=UPPER(TEXT(CD{ri},"AM/PM"))' for ri in r]

    # Zip all formula columns together per row
    formula_cols = list(zip(
        col_BB, col_BC, col_BD, col_BE, col_BF, col_BG, col_BH, col_BI, col_BJ,
        col_BK, col_BL, col_BM, col_BN, col_BO, col_BP, col_BQ, col_BR, col_BS,
        col_BT, col_BU, col_BV, col_BW, col_BX, col_BY, col_BZ, col_CA, col_CB,
        col_CC, col_CD, col_CE, col_CF
    ))

    # ---- Assemble final data_batch ----
    # Indices relative to the full row_data list:
    #   left: 0..num_cols-1
    #   formulas start at index num_cols (=53 by default)
    #   BE  → index num_cols+3  = 56
    #   BZ  → index num_cols+23 = 76 (0-based inside formula tuple = index 23)
    #   AA  → index 26
    #   AE  → index 30
    #   CD  → formula tuple index 28
    #   CE  → formula tuple index 29

    data_batch = []
    extra_len = 702 if return_buckets else 0

    for i in range(nrows):
        row_data = left_records[i] + list(formula_cols[i])

        if is_bucket_mode:
            # BE (index 56) → BY (index 77)  [formula tuple index 3 → full-row index 56]
            row_data[77] = row_data[56]
            # CD (full-row index 81) → AA (index 26)
            row_data[26] = row_data[81]
            # CE (full-row index 82) → AE (index 30)
            row_data[30] = row_data[82]

        if return_buckets:
            # Pad out to 702 columns and store 0-based original index at position 701
            row_data.extend([""] * (extra_len - len(row_data)))
            row_data[701] = i   # relative 0-based index

        data_batch.append(row_data)

    return data_batch


def _batch_delete_rows(ws, rows_to_delete):
    """
    OPTIMIZED: Delete a list of Excel row numbers in a single Union range call
    instead of one ws.Rows(r).Delete() per row.  Rows must already be in
    descending order so indices don't shift.
    """
    if not rows_to_delete:
        return
    # Build a Union of entire rows and delete once
    row_range = ws.Rows(rows_to_delete[0])
    for r in rows_to_delete[1:]:
        row_range = ws.Application.Union(row_range, ws.Rows(r))
    row_range.Delete()


def _process_template_openpyxl(df_data, tmpl_path, output_path, replacement_collector=None, is_bucket_mode=False, start_row=3):
    """Simplified template processing using openpyxl - for macOS/Linux.

    This is a simplified version that writes data directly to cells.
    Advanced features like formula calculation may not work as on Windows.
    """
    try:
        if not os.path.exists(tmpl_path):
            st.error(f"Template file not found: {tmpl_path}")
            return False

        # Create output directory
        dest_dir = os.path.dirname(output_path)
        if dest_dir and not os.path.exists(dest_dir):
            os.makedirs(dest_dir, exist_ok=True)

        # Copy template to output
        shutil.copy2(tmpl_path, output_path)

        # Extract data to write (similar to Windows version)
        if 'S.No' in df_data.columns and 'SUB-RFD' in df_data.columns:
            start_idx = df_data.columns.get_loc('S.No')
            end_idx = df_data.columns.get_loc('SUB-RFD') + 1
            df_extracted = df_data.iloc[:, start_idx:end_idx]
        else:
            df_extracted = df_data.iloc[:, :53]

        log_step("Opening workbook with openpyxl...", "info")

        # Load workbook with openpyxl
        wb = load_workbook(output_path)

        # Find the sheet - try "VOLARE EXTRACTION" first, or use first sheet
        if "VOLARE EXTRACTION" in wb.sheetnames:
            ws = wb["VOLARE EXTRACTION"]
        else:
            ws = wb.worksheets[0]

        log_step(f"Writing {len(df_extracted)} rows to Excel...", "info")

        # Write data to cells starting at start_row
        for row_idx, (_, row_data) in enumerate(df_extracted.iterrows()):
            excel_row = start_row + row_idx

            for col_idx, value in enumerate(row_data):
                cell = ws.cell(row=excel_row, column=col_idx + 1, value=value)

                # Apply styling if needed (basic)
                if replacement_collector and col_idx == 64:  # Column BN - Collector
                    if value and any(name in str(value) for name in NAMES_TO_REPLACE):
                        cell.value = replacement_collector

        # Apply some basic styling for header row if it's row 2
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=2, column=col)
            cell.font = Font(bold=True)

        log_step("Saving workbook...", "info")
        wb.save(output_path)
        wb.close()

        log_step("Template processing complete!", "success")
        return True

    except Exception as e:
        log_step(f"Error in openpyxl processing: {e}", "error")
        import traceback
        st.code(traceback.format_exc())
        return False


def process_template(df_data, tmpl_path, output_path, replacement_collector=None, password=None, is_bucket_mode=False, start_row=3, return_buckets=False):
    """Copy template and write data - uses COM on Windows, openpyxl on macOS/Linux."""

    # Use COM on Windows, use openpyxl on other platforms
    use_com = PLATFORM == 'windows' and WIN32COM_AVAILABLE

    if not use_com and not OPENPYXL_AVAILABLE:
        st.error("openpyxl is required but not installed. Please run: pip install openpyxl")
        return False

    if not use_com:
        # macOS/Linux path - use openpyxl
        return _process_template_openpyxl(df_data, tmpl_path, output_path, replacement_collector, is_bucket_mode, start_row)

    # Windows path - use COM (original code)
    import win32com.client
    import pythoncom
    import time

    try:
        if not os.path.exists(tmpl_path):
            st.error(f"Template file not found: {tmpl_path}")
            return False

        dest_dir = os.path.dirname(output_path)
        if dest_dir and not os.path.exists(dest_dir):
            os.makedirs(dest_dir, exist_ok=True)
        
        max_retries = 3
        retry_delay = 1
        
        for attempt in range(max_retries):
            try:
                shutil.copy2(tmpl_path, output_path)
                break
            except PermissionError:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    st.error(f"Cannot copy template file after {max_retries} attempts.")
                    return False
        
        if 'S.No' in df_data.columns and 'SUB-RFD' in df_data.columns:
            start_idx = df_data.columns.get_loc('S.No')
            end_idx = df_data.columns.get_loc('SUB-RFD') + 1
            df_extracted_save = df_data.iloc[:, start_idx:end_idx]
        else:
            df_extracted_save = df_data.iloc[:, :53]
        
        num_cols = df_extracted_save.shape[1]
        num_rows = df_extracted_save.shape[0]
        
        pythoncom.CoInitialize()
        excel = None
        wb = None
        
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.Interactive = False

            log_step("Connecting to Excel COM...", "info")
            
            abs_path = os.path.abspath(output_path)
            log_step(f"Opening workbook: {os.path.basename(abs_path)}...", "info")
            
            for attempt in range(max_retries):
                try:
                    wb = excel.Workbooks.Open(abs_path, Password=password)
                    break
                except Exception as e:
                    if attempt < max_retries - 1:
                        time.sleep(retry_delay)
                    else:
                        log_step(f"Failed to open workbook: {e}", "error")
                        return False
            
            # Safely set calculation to manual
            try:
                excel.Calculation = -4135  # xlCalculationManual
                excel.ScreenUpdating = False
            except Exception:
                pass

            ws = wb.Sheets("VOLARE EXTRACTION")
            log_step(f"Target sheet identified: {ws.Name}", "info")
            
            log_step(f"Building initial data batch ({num_rows} rows)...", "info")
            data_batch = _build_data_batch(df_data, num_cols, start_row, is_bucket_mode, return_buckets)
            
            if data_batch:
                start_cell_obj = ws.Cells(start_row, 1)
                last_row_pasted = start_row + len(data_batch) - 1
                end_cell_obj = ws.Cells(last_row_pasted, len(data_batch[0]))
                
                log_step("Pasting formulas for calculation...", "info")
                ws.Range(start_cell_obj, end_cell_obj).Formula = data_batch
                
                log_step("Triggering Excel calculation...", "info")
                try:
                    excel.Calculation = -4105 # Automatic to calculate formulas
                    time.sleep(1) # Give it a moment to compute
                except Exception:
                    pass

                log_step("Reading calculated values into memory for fast processing...", "info")
                # Read everything back from A3 to CE{last_row_pasted}
                # CE is col 83.
                full_data_range = ws.Range(ws.Cells(3, 1), ws.Cells(last_row_pasted, 83)).Value
                
                if not full_data_range:
                    log_step("Failed to read back calculated data.", "error")
                    return False

                # =========================================================
                # OPTIMIZED IN-MEMORY PROCESSING (STEPS 1-12)
                # =========================================================
                log_step("Running in-memory cleanup engine (Fast Mode)...", "info")
                
                final_rows = []
                positive_skiptrace_indices = [] # relative to final_rows
                
                for r_idx, row_tuple in enumerate(full_data_range):
                    row = list(row_tuple)
                    
                    # Column Mapping (0-based indices):
                    # B=1, J=9, K=10, L=11, BO=66, BP=67, BQ=68, BR=69, BS=70, BZ=77, BN=65, BL=63, BF=57, BW=74, BX=75, I=8
                    
                    action = str(row[66]).strip().upper() if row[66] is not None else ""
                    
                    # Step 1: Filter Excluded/NA
                    if action in ("EXLUDED", "EXCLUDED", "#N/A") or row[66] == -2146826246:
                        continue
                    
                    # Step 9: Delete whole row where COLLECTOR (BN=65) is #N/A
                    collector = str(row[65]).strip().upper() if row[65] is not None else ""
                    if collector == "#N/A" or row[65] == -2146826246:
                        continue
                        
                    # Step 12: Delete rows where BUCKET (BF=57) is #N/A
                    bucket = str(row[57]).strip().upper() if row[57] is not None else ""
                    if bucket == "#N/A" or row[57] == -2146826246:
                        continue
                    
                    # Step 4: Filter PTP AMOUNT (BS=70) = 0
                    ptp_amt_val = row[70]
                    try:
                        ptp_amt_float = float(ptp_amt_val) if ptp_amt_val not in (None, "") else 0.0
                        if ptp_amt_float == 0.0 and action == "PTP":
                            continue # If it is PTP action but amount is 0, exclude
                    except Exception:
                        pass

                    # Step 2: Clear PTP Date (BR=69) / Amt (BS=70) where Action != PTP
                    if action != "PTP":
                        row[69] = ""
                        row[70] = ""
                        
                    # Step 3: Clear REACTION CODE (BP=67) where 0/NA
                    reaction = str(row[67]).strip().upper() if row[67] is not None else ""
                    if reaction in ("0", "0.0", "#N/A") or row[67] == -2146826246:
                        row[67] = ""

                    # Step 5 & 6: Skiptrace (BZ=77)
                    skiptrace = str(row[77]).strip().upper() if row[77] is not None else ""
                    if skiptrace == "NEGATIVE":
                        row[77] = ""
                    elif skiptrace and skiptrace not in ("NEGATIVE", "0", "FALSE", "#N/A", "NONE", "0.0"):
                        # Step 8: Copy REMARK (K=10) to BW (74)
                        row[74] = row[10]
                        # Set exact Excel formula in BZ (index 77)
                        # We use the new row index where this will be pasted: 3 + len(final_rows)
                        new_row_idx = 3 + len(final_rows)
                        # Overwrite the generic "POSITIVE" in BZ with the extracted source from J
                        row[77] = f'=TRIM(RIGHT(SUBSTITUTE($J{new_row_idx}," ",REPT(" ",100)),100))'
                        
                        positive_skiptrace_indices.append(len(final_rows))
                    
                    # Step 7: Clear BW/BX (74/75) values set to 0 or UNCATEGORIZED
                    for col_idx in (74, 75):
                        v_str = str(row[col_idx]).strip().upper() if row[col_idx] is not None else ""
                        if v_str in ("0", "0.0", "UNCATEGORIZED"):
                            row[col_idx] = ""

                    # Step 10: Replace Target Collectors in BN (65)
                    if replacement_collector:
                        coll_val = str(row[65]).strip()
                        if any(name in coll_val for name in NAMES_TO_REPLACE):
                            row[65] = replacement_collector

                    # Step 11: Clear TYPE OF COLLECTION EFFORT (BL=63) if #N/A
                    effort = str(row[63]).strip().upper() if row[63] is not None else ""
                    if effort == "#N/A" or row[63] == -2146826246:
                        row[63] = ""

                    # Re-add leading apostrophe to '000000' strings to prevent Excel from converting to numbers
                    for c_idx in range(len(row)):
                        val = row[c_idx]
                        if isinstance(val, str) and val.startswith("000000") and val.isdigit():
                            row[c_idx] = "'" + val

                    final_rows.append(row)

                # =========================================================
                # FINAL BULK WRITE BACK
                # =========================================================
                log_step(f"Finalizing data for {len(final_rows)} remaining rows...", "info")
                
                # Clear the worksheet first
                ws.Range(f"A3:CE{last_row_pasted + 100}").ClearContents()
                
                if final_rows:
                    # Write all cleaned data in one shot
                    final_end_row = 3 + len(final_rows) - 1
                    ws.Range(ws.Cells(3, 1), ws.Cells(final_end_row, 83)).Value = final_rows
                    
                    # Apply bulk formatting
                    log_step("Applying final formatting and highlighting...", "info")
                    try:
                        # Borders
                        final_range = ws.Range(ws.Cells(3, 1), ws.Cells(final_end_row, 83))
                        final_range.Borders.LineStyle = 1
                        final_range.Borders.Weight = 2
                        
                        # Copy format from BT1 (Red header)
                        ws.Range("BT1").Copy()
                        ws.Range(f"BT3:BT{final_end_row}").PasteSpecial(Paste=-4122)
                        ws.Range(f"CA3:CC{final_end_row}").PasteSpecial(Paste=-4122)
                        
                        # Highlighting Positive Skiptrace (Column BZ = index 78 in Excel)
                        if positive_skiptrace_indices:
                            yellow_union = None
                            for rel_idx in positive_skiptrace_indices:
                                cell = ws.Cells(3 + rel_idx, 78) # Column BZ
                                yellow_union = cell if yellow_union is None else excel.Application.Union(yellow_union, cell)
                            if yellow_union:
                                yellow_union.Interior.Color = 65535
                        
                        excel.Application.CutCopyMode = False
                    except Exception as e:
                        log_step(f"Format warning: {e}", "info")

                log_step("Finalizing and saving workbook...", "info")
                try:
                    ws.Range("A1").Select()
                except Exception:
                    pass
            
            try:
                excel.Calculation = -4105  # xlCalculationAutomatic
                excel.ScreenUpdating = True
            except Exception:
                pass

            wb.Save()
            wb.Close(False)
            log_step("Master calculation completed (Optimized).", "success")
            
            if return_buckets:
                # To support return_buckets, we need the BF values for ALL rows (even deleted ones)
                # This part is rarely used in current main(), but kept for compatibility
                return True, {} 
            
            st.session_state['last_master_path'] = output_path
            return True
            
        except Exception as e:
            log_step(f"Error during Excel operations: {e}", "error")
            import traceback
            st.code(traceback.format_exc())
            return False
            
        finally:
            try:
                excel.Calculation = -4105
                excel.ScreenUpdating = True
            except Exception:
                pass
            if wb:
                try: wb.Close(False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
            pythoncom.CoUninitialize()
            
    except Exception as e:
        log_step(f"Master calculation fatal error: {e}", "error")
        import traceback
        st.code(traceback.format_exc())
        return False


def _process_campaign_split_openpyxl(master_path, save_dir, date_str, bucket_list, bucket_templates=None):
    """Simplified bucket split using openpyxl - for macOS/Linux.

    Reads the master file and creates bucket files with the appropriate template.
    """
    try:
        if not os.path.exists(master_path):
            log_step(f"Master file not found: {master_path}", "error")
            return False

        if not bucket_templates:
            log_step("No bucket templates provided.", "error")
            return False

        # Ensure save directory exists
        os.makedirs(save_dir, exist_ok=True)

        # Read master file
        log_step("Reading master file with openpyxl...", "info")
        wb_master = load_workbook(master_path, data_only=True)

        # Get first sheet (usually "VOLARE EXTRACTION")
        ws_master = wb_master.worksheets[0]

        # Read all data from the master
        master_data = []
        for row in ws_master.iter_rows(min_row=3, values_only=True):
            master_data.append(list(row))

        wb_master.close()

        if not master_data:
            log_step("No data found in master file.", "error")
            return False

        # Group data by bucket (column 57 should be bucket, adjust as needed)
        groups = {f"{b} Days": [] for b in bucket_list}

        # Debug: Show what bucket values exist in master
        found_buckets = set()

        for row_data in master_data:
            if not row_data or len(row_data) < 58:
                continue
            bucket_val_raw = row_data[57]
            bucket_val = str(bucket_val_raw).strip() if bucket_val_raw is not None else ""
            found_buckets.add(bucket_val)

            # Try to match against expected bucket formats
            matched = False
            for b in bucket_list:
                b_str = str(b)
                # Check various formats: "120 Days", "120", etc.
                if (bucket_val == f"{b} Days" or
                    bucket_val == b_str or
                    bucket_val.startswith(b_str + " ") or
                    bucket_val_raw == b):
                    groups[f"{b} Days"].append(row_data)
                    matched = True
                    break

            # If no match and there's some bucket value, put in first bucket as fallback
            if not matched and bucket_val not in ("", "None"):
                groups["120 Days"].append(row_data)  # fallback to 120 Days

        log_step(f"Found bucket values in master: {found_buckets}", "info")

        # Process each bucket
        for bucket in bucket_list:
            bucket_name = f"{bucket} Days"
            rows = groups.get(bucket_name, [])

            if not rows:
                log_step(f"No rows found for bucket {bucket}.", "info")
                continue

            template_path = bucket_templates.get(bucket)
            if not template_path or not os.path.exists(template_path):
                log_step(f"Template not found for bucket {bucket}", "error")
                continue

            # Output filename
            output_filename = f"PL_{date_str}_{bucket}_Madrid.xlsx"
            output_path = os.path.join(save_dir, output_filename)

            log_step(f"Creating bucket {bucket} file with {len(rows)} rows...", "info")

            try:
                # Copy template to output
                shutil.copy2(template_path, output_path)

                # Open output file
                wb_bucket = load_workbook(output_path)

                # Get first sheet
                ws_bucket = wb_bucket.worksheets[0]

                # Write data starting at row 4 (typical for templates)
                start_row = 4

                for row_idx, row_data in enumerate(rows):
                    excel_row = start_row + row_idx

                    # Write from master columns V onwards (index 56+) to bucket
                    # Master columns 56-78 (V to BR, 22 columns) -> bucket columns 0-21 (A to V)
                    for col_idx in range(22):  # Write 22 columns from master
                        source_idx = 56 + col_idx  # Start from column V (index 56)
                        if source_idx < len(row_data):
                            source_val = row_data[source_idx]
                        else:
                            source_val = ""
                        ws_bucket.cell(row=excel_row, column=col_idx + 1, value=source_val)

                    # Also add CD, CE, CF columns from master if they exist
                    # CD = index 81 -> bucket column 24 (Y)
                    # CE = index 82 -> bucket column 27 (AB)
                    # CF = index 83 -> bucket column 31 (AF)
                    if 81 < len(row_data):
                        ws_bucket.cell(row=excel_row, column=24, value=row_data[81])  # CD -> Y
                    if 82 < len(row_data):
                        ws_bucket.cell(row=excel_row, column=27, value=row_data[82])  # CE -> AB
                    if 83 < len(row_data):
                        ws_bucket.cell(row=excel_row, column=31, value=row_data[83])  # CF -> AF

                # Save the workbook
                wb_bucket.save(output_path)
                wb_bucket.close()

                log_step(f"Bucket {bucket} file created: {output_filename}", "success")

            except Exception as e:
                log_step(f"Error creating bucket {bucket}: {e}", "error")
                continue

        log_step("All bucket files created successfully!", "success")
        return True

    except Exception as e:
        log_step(f"Error in bucket split (openpyxl): {e}", "error")
        import traceback
        st.code(traceback.format_exc())
        return False


def process_campaign_split(master_path, save_dir, date_str, bucket_list, password, bucket_templates=None, campaign_template_pattern=None):
    """Reads the Master file once and splits it into Campaign files with specific formatting.

    Uses COM on Windows, openpyxl on macOS/Linux.

    Args:
        master_path: Path to the master file
        save_dir: Directory to save output files
        date_str: Date string for filenames
        bucket_list: List of bucket numbers (e.g., [120, 150, 180])
        password: Password for files
        bucket_templates: Optional dict mapping bucket number to template path (e.g., {120: 'path/to/120.xlsx'})
        campaign_template_pattern: Fallback pattern if bucket_templates not provided
    """

    # Use COM on Windows, use openpyxl on other platforms
    use_com = PLATFORM == 'windows' and WIN32COM_AVAILABLE

    if not use_com and not OPENPYXL_AVAILABLE:
        st.error("openpyxl is required.")
        return False

    if not use_com:
        # macOS/Linux path - use openpyxl
        return _process_campaign_split_openpyxl(master_path, save_dir, date_str, bucket_list, bucket_templates)

    # Windows path - use COM (original code)
    if not WIN32COM_AVAILABLE:
        st.error("pywin32 is required.")
        return False

    import win32com.client
    import pythoncom

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.Interactive = False

        # Safe state setup
        try:
            excel.ScreenUpdating = False
            excel.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass
        
        abs_master = os.path.abspath(master_path)
        log_step(f"Opening master for split: {os.path.basename(abs_master)}...", "info")
        
        wb_master = excel.Workbooks.Open(
            Filename=abs_master, 
            UpdateLinks=0, 
            ReadOnly=True,
            Password=password, 
            WriteResPassword=password,
            IgnoreReadOnlyRecommended=True
        )
        
        try:
            ws_master = wb_master.Sheets("VOLARE EXTRACTION")
        except Exception:
            ws_master = wb_master.Sheets(1)
        
        last_row = ws_master.Cells(ws_master.Rows.Count, "A").End(-4162).Row
        if last_row < 3:
            wb_master.Close(False)
            excel.Quit()
            log_step("Master file has no data rows.", "error")
            return False
        
        log_step(f"Reading data from Master ({last_row-2} rows)...", "info")
        data_range = ws_master.Range(ws_master.Cells(3, 1), ws_master.Cells(last_row, 83)).Value
        wb_master.Close(False)
        
        if not data_range:
            excel.Quit()
            log_step("Could not read data range from Master.", "error")
            return False

        log_step("Sorting data into bucket groups...", "info")
        groups = {f"{b} Days": [] for b in bucket_list}
        for row_vals in data_range:
            if not row_vals: continue
            bucket_val = str(row_vals[57]).strip() if row_vals[57] else ""
            if bucket_val in groups:
                groups[bucket_val].append(list(row_vals))
        
        for b in bucket_list:
            bucket_name_str = f"{b} Days"
            rows = groups[bucket_name_str]
            if not rows:
                log_step(f"No rows found for bucket {b}.", "info")
                continue

            # Use uploaded bucket template if available, otherwise use the pattern
            if bucket_templates and b in bucket_templates:
                campaign_tmpl = bucket_templates[b]
                log_step(f"Using uploaded template for bucket {b}", "info")
            elif campaign_template_pattern:
                campaign_tmpl = campaign_template_pattern.replace("{date}", date_str).replace("{bucket}", str(b))
            else:
                log_step(f"No template found for bucket {b}. Please upload a {b}-day bucket template.", "error")
                continue

            if not os.path.exists(campaign_tmpl):
                log_step(f"Template not found: {os.path.basename(campaign_tmpl)}", "error")
                continue

            bucket_filename = f"PL_{date_str}_{b}_Madrid.xlsx"
            save_path = os.path.abspath(os.path.join(save_dir, bucket_filename))

            try:
                log_step(f"Generating bucket {b} file...", "info")
                wb_bucket = excel.Workbooks.Open(
                    Filename=campaign_tmpl,
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True
                )

                try:
                    ws_bucket = wb_bucket.Sheets("VOLARE EXTRACTION")
                except Exception:
                    ws_bucket = wb_bucket.Sheets(1)
                
                # Optimized in-memory building of bucket rows
                bucket_rows = []
                for r_master in rows:
                    new_row = [""] * 32
                    new_row[0:22] = r_master[56:78]  # columns A-V (indices 0-21)
                    # Check bounds before accessing indices 81, 82, 83
                    if len(r_master) > 81:
                        new_row[24] = r_master[81]  # CD at column Y (index 24)
                    if len(r_master) > 82:
                        new_row[27] = r_master[82]  # CE at column AB (index 27)
                    if len(r_master) > 83:
                        new_row[31] = r_master[83]  # CF at column AF (index 31)
                    bucket_rows.append(tuple(new_row))
                
                rows_to_paste = tuple(bucket_rows)
                start_row_bucket = 4
                end_row_bucket = start_row_bucket + len(rows_to_paste) - 1

                ws_bucket.Range(
                    ws_bucket.Cells(start_row_bucket, 1),
                    ws_bucket.Cells(end_row_bucket, 32)
                ).Value = rows_to_paste
                
                try:
                    p_range = ws_bucket.Range(
                        ws_bucket.Cells(start_row_bucket, 1),
                        ws_bucket.Cells(end_row_bucket, 32)
                    )
                    p_range.Borders.LineStyle = 1
                    p_range.Borders.Weight = 2

                    col22_vals = ws_bucket.Range(
                        ws_bucket.Cells(start_row_bucket, 22),
                        ws_bucket.Cells(end_row_bucket, 22)
                    ).Value

                    yellow_range = None
                    if col22_vals:
                        for i, rv in enumerate(col22_vals):
                            val = str(rv[0]).strip().upper() if rv and rv[0] not in (None, "") else ""
                            if val and val not in ("NEGATIVE", "0", "FALSE", "#N/A", "NONE", "NAN", ""):
                                cell = ws_bucket.Cells(start_row_bucket + i, 22)
                                yellow_range = cell if yellow_range is None else ws_bucket.Application.Union(yellow_range, cell)
                    if yellow_range:
                        yellow_range.Interior.Color = 65535
                        yellow_range.HorizontalAlignment = -4108  # xlCenter

                    ws_bucket.Columns("A:AE").AutoFit()
                except Exception as e:
                    log_step(f"Bucket formatting warning: {e}", "info")
                
                wb_bucket.SaveAs(save_path, FileFormat=51, Password=password)
                wb_bucket.Close()
                log_step(f"Generated: {bucket_filename}", "success")
            except Exception as e:
                log_step(f"Error generating bucket {b}: {e}", "error")
        
        try:
            excel.Calculation = -4105
            excel.ScreenUpdating = True
        except Exception:
            pass

        excel.Interactive = True
        excel.Quit()
        pythoncom.CoUninitialize()
        return True
    except Exception as e:
        log_step(f"Campaign split fatal error: {e}", "error")
        try:
            excel.Calculation = -4105
            excel.ScreenUpdating = True
        except Exception:
            pass
        return False

def main():
    # Session state for template paths (to persist across reruns)
    if 'master_template_path' not in st.session_state:
        st.session_state.master_template_path = None
    if 'bucket_templates' not in st.session_state:
        st.session_state.bucket_templates = {}

    if not os.path.exists(SAVE_PATH):
        try: os.makedirs(SAVE_PATH)
        except Exception: pass

    st.title("Daily Prod Automation")
    st.markdown("Removes unwanted statuses and remark sources for cleaner data processing")

    # ==========================================
    #  TEMPLATE UPLOADS SECTION
    # ==========================================
    with st.expander("📁 Template Files (Optional - Leave empty to use configured paths)", expanded=True):
        col_info, col_clear = st.columns([3, 1])
        with col_info:
            st.info("Upload your template files here, OR leave empty to use the configured template paths from config.py")
        with col_clear:
            if st.button("🗑️ Clear Uploaded", help="Clear all uploaded templates"):
                st.session_state.master_template_path = None
                st.session_state.bucket_templates = {}
                cleanup_temp_templates()
                st.rerun()

        # Master Template Upload
        master_upload = st.file_uploader(
            "Upload Master Template (.xlsm)",
            type=["xlsm", "xlsx"],
            key="master_template"
        )

        if master_upload:
            # Save to temp and store path
            st.session_state.master_template_path = save_uploaded_template(master_upload, "master")
            st.success(f"Master template loaded: {master_upload.name}")

        # Bucket Templates Upload - Single template used for all buckets (they're the same)
        st.markdown("**Bucket Template (used for all 120/150/180 day buckets):**")

        bucket_template = st.file_uploader(
            "Upload Bucket Template (.xlsx/.xlsm) - One file used for all buckets",
            type=["xlsx", "xlsm"],
            key="bucket_template"
        )

        if bucket_template:
            # Save the single template and use it for all buckets
            bucket_path = save_uploaded_template(bucket_template, "bucket")
            # Use the same template for all 3 buckets
            st.session_state.bucket_templates = {
                120: bucket_path,
                150: bucket_path,
                180: bucket_path
            }
            st.success(f"Bucket template loaded: {bucket_template.name} (used for 120/150/180)")

        # Validate all templates are uploaded
        if not st.session_state.master_template_path:
            st.warning("⚠️ Please upload the Master Template above")
        if not st.session_state.bucket_templates:
            st.warning("⚠️ Please upload the Bucket Template above")

    # Password input for templates
    template_password = st.text_input("Template Password (if protected)", type="password", help="Enter password if your templates are protected")

    st.divider()

    # Validate master template is uploaded
    master_tmpl_path = st.session_state.master_template_path
    if not master_tmpl_path:
        st.error("Master Template is required. Please upload it above.")
        st.stop()

    # Validate all bucket templates are uploaded
    missing_buckets = [b for b in BUCKETS_TO_PROCESS if b not in st.session_state.bucket_templates]
    if missing_buckets:
        st.error(f"Missing bucket templates for: {missing_buckets}. Please upload all bucket templates above.")
        st.stop()
    uploaded_file = st.file_uploader(" Drag and drop your Daily Prod file (Excel or CSV)", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        st.subheader(" Collector Settings")
        replacement_name = st.selectbox(
            "Choose replacement name for Target Collectors (NEW):",
            ["Select Name...", "Keep Original", "Bandola, Diana", "Reyes, Berlyn", "Galang, Jayson", "Antonio, Christine"]
        )

        if replacement_name == "Select Name...":
            st.info(" Please select a replacement name (or **'Keep Original'**) to proceed with the automation.")
            st.stop()

        if replacement_name == "Keep Original":
            replacement_name = None

        st.divider()

        file_ext = uploaded_file.name.split('.')[-1]
        with st.spinner(" Reading data..."):
            try:
                if file_ext == 'csv': df = pd.read_csv(uploaded_file)
                else: df = pd.read_excel(uploaded_file, engine='openpyxl')
                df.columns = df.columns.str.strip()
            except Exception as e:
                st.error(f"Error loading file: {e}")
                st.stop()

        if "Status" in df.columns and "Remark By" in df.columns:
            with st.spinner(" Scrubbing data..."):
                if 'Remark' in df.columns:
                    df['Remark'] = df['Remark'].astype(str).str[:250]
                mask_status = df['Status'].astype(str).str.startswith(tuple(EXCLUDED_STATUSES))
                if 'Date' in df.columns:
                    df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.strftime('%m/%d/%Y')
                mask_remark = df['Remark By'].astype(str).str.contains('SMGONZALES|CMENRIQUEZ', na=False)
                df_removed = df[mask_status | mask_remark]
                df_cleaned = df[~(mask_status | mask_remark)].copy()
                
                if "Remark" in df_cleaned.columns:
                    classifications = df_cleaned['Remark'].apply(classify_remark)
                    df_cleaned["RFD"], df_cleaned["SUB-RFD"] = zip(*classifications)
                    df_cleaned["RFD"] = df_cleaned["RFD"].fillna("UNCATEGORIZED")
                    df_cleaned["SUB-RFD"] = df_cleaned["SUB-RFD"].fillna("UNCATEGORIZED")
                if "Account No." in df_cleaned.columns:
                    df_cleaned["Account No."] = df_cleaned["Account No."].astype(str).str.replace(r'\.0$', '', regex=True)
                    df_cleaned["Account No."] = df_cleaned["Account No."].replace(['nan', 'None'], '')
                    df_cleaned = df_cleaned[df_cleaned["Account No."] != ""]
                    df_cleaned["Account No."] = "'000000" + df_cleaned["Account No."]
                if "Time" in df_cleaned.columns:
                    def format_time(t):
                        try:
                            t_str = pd.to_datetime(str(t), errors='coerce').strftime('%I:%M %p')
                            if pd.isna(t_str) or t_str == 'NaT': return str(t)
                            return f"'{t_str}"
                        except Exception: return str(t)
                    df_cleaned["Time"] = df_cleaned["Time"].apply(format_time)
                
                if 'S.No' in df_cleaned.columns and 'SUB-RFD' in df_cleaned.columns:
                    start_idx = df_cleaned.columns.get_loc('S.No')
                    end_idx = df_cleaned.columns.get_loc('SUB-RFD') + 1
                    df_extracted_display = df_cleaned.iloc[:, start_idx:end_idx]
                else:
                    df_extracted_display = df_cleaned.iloc[:, :53]

            report_date_obj, date_str = get_report_date()
            tmpl_ext = os.path.splitext(master_tmpl_path)[1] or ".xlsx"
            tmp_master = os.path.join(SAVE_PATH, f"Master_Calculation_{date_str}{tmpl_ext}")

            isProcessed = False

            with st.spinner("Running Master calculation and generating bucket files..."):
                with st.expander("Extraction Log", expanded=False):
                    # Use uploaded template path only
                    success = process_template(df_cleaned, master_tmpl_path, tmp_master, replacement_name, password=template_password)
                    if success:
                        st.success(f"Master created: {tmp_master}")
                        # Pass bucket templates (only uploaded templates - no fallback)
                        process_campaign_split(
                            tmp_master, SAVE_PATH, date_str, BUCKETS_TO_PROCESS,
                            password=template_password,
                            bucket_templates=st.session_state.bucket_templates
                        )
                        isProcessed = True
                    else:
                        st.error("Master calculation failed.")
            
            if isProcessed:
                st.divider()
                st.subheader(" PROCESS DONE ✅")

                # Create ZIP file for download
                zip_buffer = create_output_zip(SAVE_PATH, date_str)

                if zip_buffer:
                    st.markdown("### 📥 Download Output Files")

                    # Download as ZIP
                    st.download_button(
                        label="📦 Download All Files (ZIP)",
                        data=zip_buffer,
                        file_name=f"Daily_Prod_Output_{date_str}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

                    # Show individual file links
                    st.markdown("**Individual Files:**")
                    col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)

                    # Master file
                    master_file = os.path.join(SAVE_PATH, f"Master_Calculation_{date_str}")
                    for ext in ['.xlsm', '.xlsx']:
                        if os.path.exists(master_file + ext):
                            with open(master_file + ext, 'rb') as f:
                                col_dl1.download_button(
                                    "📄 Master File",
                                    f,
                                    f"Master_Calculation_{date_str}{ext}",
                                    use_container_width=True
                                )
                            break

                    # Bucket files
                    for i, b in enumerate([120, 150, 180]):
                        bucket_file = os.path.join(SAVE_PATH, f"PL_{date_str}_{b}_Madrid.xlsx")
                        if os.path.exists(bucket_file):
                            with open(bucket_file, 'rb') as f:
                                col = [col_dl2, col_dl3, col_dl4][i]
                                col.download_button(
                                    f"📄 {b}-Day Bucket",
                                    f,
                                    f"PL_{date_str}_{b}_Madrid.xlsx",
                                    use_container_width=True
                                )
            else:
                st.error(" PROCESS FAILED ❌")
            
            st.divider()
            c1, c2, c3 = st.columns(3)
            c1.metric("Original Rows", f"{len(df):,}")
            c2.metric("Kept Rows", f"{len(df_cleaned):,}")
            c3.metric("Removed Rows", f"{len(df_removed):,}", delta=f"-{len(df_removed)}", delta_color="inverse")
            
            st.divider()
            st.subheader(" Removed Rows")
            if len(df_removed) > 0:
                st.dataframe(df_removed.head(100), use_container_width=True)
            
            st.divider()
            st.subheader(" Template Paste Preview")
            st.dataframe(df_extracted_display.head(50), use_container_width=True)

        else:
            st.error(" Columns 'Status' and 'Remark By' not found.")
            st.write(f"Available columns: {list(df.columns)}")
    else:
        st.info("Upload your Daily Prod Excel File to begin.")

# ==========================================
#  CONFIGURATION
# ==========================================
import uuid

# Simple config - just buckets to process
BUCKETS_TO_PROCESS = [120, 150, 180]
NAMES_TO_REPLACE = ["Mendoza, Joshua", "Salcedo", "Shaine"]

# Default output path (temp folder)
SAVE_PATH = tempfile.gettempdir()

# Remove config.py imports - we only use uploaded templates now

# ==========================================
#  HELPER: Save uploaded file to temp and return path
# ==========================================
def save_uploaded_template(uploaded_file, bucket_name=""):
    """Save an uploaded template file to a temporary location and return the path."""
    if uploaded_file is None:
        return None

    # Create temp directory in the SAVE_PATH or use system temp
    temp_dir = os.path.join(SAVE_PATH, "temp_templates")
    os.makedirs(temp_dir, exist_ok=True)

    # Generate unique filename to avoid conflicts
    unique_id = str(uuid.uuid4())[:8]
    file_ext = os.path.splitext(uploaded_file.name)[1]

    if bucket_name:
        filename = f"template_{bucket_name}_{unique_id}{file_ext}"
    else:
        filename = f"master_template_{unique_id}{file_ext}"

    temp_path = os.path.join(temp_dir, filename)

    # Write uploaded file to temp path
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    return temp_path


def cleanup_temp_templates():
    """Clean up temporary template files."""
    temp_dir = os.path.join(SAVE_PATH, "temp_templates")
    if os.path.exists(temp_dir):
        try:
            shutil.rmtree(temp_dir)
        except Exception:
            pass


def create_output_zip(save_dir, date_str):
    """Create a ZIP file containing all generated output files (Master + Buckets).

    Args:
        save_dir: Directory where output files are saved
        date_str: Date string used in filenames (e.g., '20250915')

    Returns:
        Path to the created ZIP file, or None if no files found
    """
    # Expected output filenames
    expected_files = [
        f"Master_Calculation_{date_str}.xlsm",
        f"Master_Calculation_{date_str}.xlsx",
        f"PL_{date_str}_120_Madrid.xlsx",
        f"PL_{date_str}_150_Madrid.xlsx",
        f"PL_{date_str}_180_Madrid.xlsx",
    ]

    # Find existing files
    files_to_zip = []
    for filename in expected_files:
        filepath = os.path.join(save_dir, filename)
        if os.path.exists(filepath):
            files_to_zip.append((filepath, filename))

    if not files_to_zip:
        return None

    # Create ZIP file in memory
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path, arcname in files_to_zip:
            try:
                zipf.write(file_path, arcname)
            except Exception:
                pass  # Skip files that can't be added

    zip_buffer.seek(0)
    return zip_buffer


if __name__ == "__main__":
    main()