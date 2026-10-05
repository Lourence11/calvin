import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Collection Efforts Clean", layout="wide")

# Apply consistent premium design
set_premium_style()


# PAGE_MAP and rfd_mapping remain unchanged
PAGE_MAP = {
    "Home": "Main.py",
    "Daily Prod": "pages/1_Daily_Prod.py",
    "Collection Efforts Clean": "pages/2_CE_Clean.py",
    "Collection Efforts Callout": "pages/3_CE_CALL.py",
    "Collection Efforts Skip": "pages/3_CE_SKIP.py",
    "Daily Prods accu": "pages/dailyprod.py",
}

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

def classify_remark(remark):
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

    for rfd, sub_rfds in rfd_mapping.items():
        for sub_rfd in sub_rfds:
            if sub_rfd.lower() in remark_lower:
                return rfd, sub_rfd
    
    return None, None

def main():
    st.title(" Collection Efforts Cleaner & RFD Remark Classifier")
    st.markdown("Upload a file to clean and classify remarks into RFD/SUB-RFD categories.")
    
    uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx", "xls"])

    if uploaded_file is not None:
        try:
            # 1. Load Data
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            # 2. Data Cleaning
            df['Remark'] = df['Remark'].astype(str).str[:250]
            # Updated Filtering Status List
            excluded_statuses = [

                "ABORT", "BP", "BULK SMS SENT", "CONFIRMED", "ENCO", "DEBIT REQUEST", "LETTER RECEIVED", 
                "FIELD RESULT", "FIELD VISIT", "LOCKED", "LS VIA EMAIL", "LS VIA SOCMED", "NEW", "PM",
                "POS 3RD PARTY - DECEASED", "RPC_REPLY FROM SOCMED", "SERVICE REQUEST (SR)", "SMS FAILED", "SMS RECEIVED", "SMS SENT", "UNLOCKED", "VM",
                "QUALIFIED FOR RETURN", "CONFIRMED_", "FIELD VISITATION", "DL REQUEST", 




                "ABORT", "BP", "BULK SMS SENT", "CONFIRMED SPM - PAID OB FULLY PAID",
                "CONFIRMED SPM - PERENNIAL", "CONFIRMED VIA CALLS - EPA COMPLIED",
                "CONFIRMED VIA CALLS - EPA DEFAULTED", "CONFIRMED VIA CALLS - EPA DOWNPAYMENT",
                "CONFIRMED VIA CALLS - EPA DOWNPAYMENT FULL", "CONFIRMED VIA CALLS - EPA DOWNPAYMENT SPLIT",
                "CONFIRMED VIA CALLS - EPA FULLY PAID", "CONFIRMED VIA CALLS - ONE TIME (FULL)",
                "CONFIRMED VIA CALLS - ONE TIME (SPLIT)", "CONFIRMED VIA CALLS - PERENNIAL",
                "CONFIRMED VIA EMAIL - DOWNPAYMENT FULL", "CONFIRMED VIA EMAIL - EPA COMPLIED",
                "CONFIRMED VIA EMAIL - EPA DEFAULTED", "CONFIRMED VIA EMAIL - EPA DOWNPAYMENT",
                "CONFIRMED VIA EMAIL - EPA DOWNPAYMENT SPLIT", "CONFIRMED VIA EMAIL - EPA PRE TERMINATION",
                "CONFIRMED VIA EMAIL - ONE TIME (FULL)", "CONFIRMED VIA EMAIL - ONE TIME (SPLIT)",
                "CONFIRMED VIA EMAIL - PERENNIAL", "CONFIRMED VIA SKIPS - EPA COMPLIED",
                "CONFIRMED VIA SKIPS - EPA DEFAULTED", "CONFIRMED VIA SKIPS - EPA FULLY PAID",
                "CONFIRMED VIA SKIPS - PERENNIAL", "CONFIRMED VIA SMS - DOWNPAYMENT FULL",
                "CONFIRMED VIA SMS - PERENNIAL", "CONFIRMED VIA VIBER - EPA COMPLIED",
                "CONFIRMED VIA VIBER - EPA DEFAULTED", "CONFIRMED VIA VIBER - EPA DOWNPAYMENT",
                "CONFIRMED VIA VIBER - EPA DOWNPAYMENT FULL", "CONFIRMED VIA VIBER - EPA DOWNPAYMENT SPLIT",
                "CONFIRMED VIA VIBER - EPA FULLY PAID", "CONFIRMED VIA VIBER - ONE TIME (FULL)",
                "CONFIRMED VIA VIBER - ONE TIME (SPLIT)", "CONFIRMED VIA VIBER - PERENNIAL",
                "CONFIRMED VIA VISIT - EPA COMPLIED", "CONFIRMED VIA VISIT - EPA DEFAULTED",
                "CONFIRMED VIA VISIT - EPA DOWNPAYMENT SPLIT", "CONFIRMED VIA VISIT - PERENNIAL",
                "CONFIRMED_ VIBER - PERENNIAL_IN", "DEBIT REQUEST - REQUESTED", "ENCO - DONE ENCO",
                "FIELD RESULT - NEG", "FIELD RESULT - POS", "FIELD VISIT - INCOMPLETE ADDRESS",
                "LETTER RECEIVED - FIELD VISIT", "LETTER RECEIVED - LBC OR REGULAR MAIL",
                "LETTER RECEIVED - THRU EMAIL", "LOCKED", "LS VIA EMAIL - OTHERS",
                "LS VIA EMAIL - T1 NOTIFICATION", "LS VIA EMAIL - T12 THIRD PARTY TEMPLATE",
                "LS VIA EMAIL - T2 DEBT MANAGEMENT ASSISTANCE", "LS VIA EMAIL - T3 PTP REMINDER",
                "LS VIA EMAIL - T4 BROKEN PTP EPA", "LS VIA EMAIL - T5 BROKEN PTP SPLIT AND OTP",
                "LS VIA EMAIL - T6 NO RESPONSE (SMS & EMAIL)", "LS VIA EMAIL - T7 PROMO OFFER LETTER",
                "LS VIA EMAIL - T9 RESTRUCTURING", "LS VIA SOCMED - OTHERS", "LS VIA SOCMED - T1 NOTIFICATION",
                "LS VIA SOCMED - T12 THIRD PARTY TEMPLATE", "LS VIA SOCMED - T2 DEBT MANAGEMENT ASSISTANCE",
                "LS VIA SOCMED - T6 NO RESPONSE (SMS & EMAIL)", "LS VIA SOCMED - T7 PROMO OFFER LETTER",
                "LS VIA SOCMED - T8 AMNESTY PROMO TEMPLATE", "LS VIA SOCMED - T9 RESTRUCTURING",
                "NEW", "PM", "POS 3RD PARTY - DECEASED", "QUALIFIED FOR RETURN - CLIENT DECEASED", "REACTIVE",
                "RPC_REPLY FROM SOCMED - FACEBOOK", "RPC_REPLY FROM SOCMED - OTHER SOCMED PLAN",
                "RPC_REPLY FROM SOCMED - VIBER", "SERVICE REQUEST (SR) - FOR COLLECTION",
                "SMS FAILED", "SMS RECEIVED - NO SUCH PERSON (NSP)", "SMS RECEIVED - TFIP",
                "SMS RECEIVED - UNDER NEGO", "SMS RECEIVED - WILL SETTLE", "SMS SENT",
                "SMS SENT - OTHERS", "SMS SENT - T1 NOTIFICATION", "SMS SENT - T2 DEBT MANAGEMENT ASSISTANCE",
                "SMS SENT - T3 PTP REMINDER", "SMS SENT - T4 BROKEN PTP EPA",
                "SMS SENT - T5 BROKEN PTP SPLIT AND OTP", "SMS SENT - T6 NO RESPONSE FROM SMS AND EMAIL",
                "SMS SENT - T7 PROMO OFFER LETTER", "SMS SENT - T9 RESTRUCTURING", "UNLOCKED", "CONFIRMED VIA EMAIL - EPA FULLY PAID",
                "CONFIRMED VIA SMS - ONE TIME (SPLIT)","CONFIRMED VIA VISIT - ONE TIME (SPLIT)","CONFIRMED_ DEBIT - EPA COMPLIED_IN",
                "CONFIRMED_ DEBIT - PERENNIAL_IN","CONFIRMED_ EMAIL - PERENNIAL_IN","CONFIRMED_ VISIT - EPA COMPLIED_IN",
                "FIELD VISIT - VST RESULT_POS CH","LS VIA EMAIL - T11 4MOS PAST DUE TEMPLATE", "LS VIA SOCMED - T11 4MOS PAST DUE TEMPLATE",
                "SMS SENT - T8 AMNESTY PROMO TEMPLATE", "FIELD VISITATION - VST RESULT_POS THIRD PARTY", "CONFIRMED_ VIBER - PERENNIAL_OUT", "FIELD VISIT - VST RESULT_NEG",
                "FIELD VISIT - OUT OF AREA COVERED", "CONFIRMED_ VIBER - EPA COMPLIED_OUT", "FIELD VISITATION - VST RESULT_POS THIRD PARTY", "CONFIRMED_ EMAIL - EPA DP_IN", 
                "CONFIRMED_ EMAIL - PERENNIAL_OUT", "CONFIRMED_ VIBER - EPA DP_OUT", "CONFIRMED_ VIBER - ONE TIME (SPLIT)_OUT", "DL REQUEST - SPECIAL VISIT", "FIELD RESULT - LBC NEG",
                "VM",

            ]
            #"PTP_FF UP - ANSWERED RENEGO", "PTP_FF UP - ANSWERED WILL SETTLE", "PTP_FF UP - BUSY", "PTP_FF UP - KEEPS ON RINGING (KOR)", "PTP_FF UP - UNCONTACTABLE", "PTP_FF",




            
            # Filter out excluded statuses AND blanks (NaN, empty, or whitespace-only)
            status_str = df['Status'].astype(str).str.strip()
            df_cleaned = df[
                (~df['Status'].astype(str).str.startswith(tuple(excluded_statuses), na=False)) &
                (status_str != "") &
                (status_str != "nan")  # Handles NaN converted to string
            ].copy()
            
            # 3. Classification Logic
            with st.spinner("Classifying remarks..."):
                classifications = df_cleaned['Remark'].apply(classify_remark)
                df_cleaned["RFD"], df_cleaned["SUB-RFD"] = zip(*classifications)
                df_cleaned["RFD"] = df_cleaned["RFD"].fillna("UNCATEGORIZED")
                df_cleaned["SUB-RFD"] = df_cleaned["SUB-RFD"].fillna("UNCATEGORIZED")

            # --- DISPLAY SECTION ---
            st.divider()
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Rows", len(df_cleaned))
            col2.metric("Classified", len(df_cleaned[df_cleaned["RFD"] != "UNCATEGORIZED"]))
            col3.metric("Uncategorized", len(df_cleaned[df_cleaned["RFD"] == "UNCATEGORIZED"]))

            tab1, tab2 = st.tabs([" Classification Stats", " Data Preview"])
            
            with tab1:
                c1, c2 = st.columns(2)
                with c1:
                    st.write("**Top RFD Categories**")
                    st.bar_chart(df_cleaned["RFD"].value_counts())
                with c2:
                    st.write("**Top Sub-RFD Reasons**")
                    st.dataframe(df_cleaned["SUB-RFD"].value_counts(), use_container_width=True)

            with tab2:
                st.dataframe(df_cleaned, use_container_width=True)

            # 4. Export Logic
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_cleaned.to_excel(writer, index=False, sheet_name="Processed Data")
            output.seek(0)

            st.sidebar.header("Actions")
            st.sidebar.download_button(
                label=" Download Processed File",
                data=output,
                file_name=f"CE_Clean_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error processing file: {e}")
    else:
        st.info("👆 Upload a file to begin.")

        

if __name__ == "__main__":
    main()


