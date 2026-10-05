import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
from io import BytesIO


st.set_page_config(page_title="Daily Prod", layout="wide")

# Apply consistent premium design
set_premium_style()


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
    st.title(" DRR Cleaner & RFD Remark Classifier")
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
            df['Remark'] = df['Remark'].str[:250]
            
            # Filtering Status
            excluded_statuses = [
                'CONFIRMED', 'PTP_FF', 'SMS SENT - T6 NO RESPONSE',
                'LS VIA SOCMED - T6 NO RESPONSE', 'SMS RECEIVED - UNDER NEGO', 
                'SMS RECEIVED - WILL SETTLE', 'NEW', 'BP', 'ENCO', 'UNLOCKED',
                'BULK SMS', 'REACTIVE', 'LOCKED', 'PM',
                'SMS RECEIVED - NO SUCH PERSON (NSP)',
                'SERVICE REQUEST (SR) - FOR COLLECTION'
            ]
            df = df[~df['Status'].astype(str).str.startswith(tuple(excluded_statuses))]

            if 'Remark By' in df.columns:
                df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%m/%d/%Y')
                df_cleaned = df[~df['Remark By'].str.contains('SMGONZALES|CMENRIQUEZ', na=False)].copy()
            else:
                st.error("'Remark By' column not found.")
                st.stop()

            # 3. Classification Logic
            with st.spinner("Classifying remarks..."):
                classifications = df_cleaned['Remark'].apply(classify_remark)
                df_cleaned["RFD"], df_cleaned["SUB-RFD"] = zip(*classifications)
                df_cleaned["RFD"] = df_cleaned["RFD"].fillna("UNCATEGORIZED")
                df_cleaned["SUB-RFD"] = df_cleaned["SUB-RFD"].fillna("UNCATEGORIZED")

            # --- NEW DISPLAY SECTION ---
            st.divider()
            
            # Top Level Metrics
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Rows", len(df_cleaned))
            col2.metric("Classified", len(df_cleaned[df_cleaned["RFD"] != "UNCATEGORIZED"]))
            col3.metric("Uncategorized", len(df_cleaned[df_cleaned["RFD"] == "UNCATEGORIZED"]))

            # Visual Summaries
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
                file_name="one_prod_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error processing file: {e}")
    else:
        st.info("👆 Upload a file to begin.")

if __name__ == "__main__":
    main()