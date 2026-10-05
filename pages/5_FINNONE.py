import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
from io import BytesIO
import io

try:
    import msoffcrypto
    HAS_MSOFFCRYPTO = True
except ImportError:
    HAS_MSOFFCRYPTO = False

st.set_page_config(page_title="FINNONE AUTOMATION", layout="wide")

# Apply consistent premium design
set_premium_style()


FINNONE_PASSWORD = "MAD_2Q2026"

def decrypt_excel(uploaded_file, password):
    if not HAS_MSOFFCRYPTO:
        st.error(" The 'msoffcrypto-tool' library is required to automatically unlock password-protected Excel files. Please open your terminal and run:")
        st.code("pip install msoffcrypto-tool")
        return uploaded_file
        
    decrypted = io.BytesIO()
    try:
        office_file = msoffcrypto.OfficeFile(uploaded_file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted
    except Exception as e:
        # Fallback in case it's not encrypted or password doesn't match
        uploaded_file.seek(0)
        return uploaded_file

def main():
    st.title(" FINNONE Automation")
    st.markdown("Upload your FINNONE file to extract and process data according to the action code rules.")
    
    uploaded_file = st.file_uploader("Upload encrypted FINNONE Excel file", type=["xlsx", "xls", "xlrd"])

    if uploaded_file is not None:
        try:
            with st.spinner("Decrypting and loading data..."):
                # Always attempt to decrypt/load the file using the preset password
                decrypted_file = decrypt_excel(uploaded_file, FINNONE_PASSWORD)
                df = pd.read_excel(decrypted_file)
            
            # Helper function to find actual column names safely regardless of case or spaces
            def get_actual_col(expected_names):
                for actual_col in df.columns:
                    if str(actual_col).strip().lower() in [e.lower() for e in expected_names]:
                        return actual_col
                return expected_names[0] # Default if not found

            customer_id_col = get_actual_col(["customer_id", "customer id", "account number"])
            user_id_col = get_actual_col(["user_id", "user id", "userid"])
            action_code_col = get_actual_col(["action_code", "action code", "actioncode"])
            remarks_col = get_actual_col(["remarks", "remark"])
            
            # Data Cleaning: ensure string and strip whitespaces
            df[action_code_col] = df[action_code_col].astype(str).str.strip()
            
            # Filter condition for Action Codes
            # ('action_code' = (all that starts with "CALL") "LR" "LS" (all that starts with "SKIP") "SMS")
            condition = (
                df[action_code_col].str.upper().str.startswith("CALL") |
                (df[action_code_col].str.upper() == "LR") |
                (df[action_code_col].str.upper() == "LS") |
                df[action_code_col].str.upper().str.startswith("SKIP") |
                (df[action_code_col].str.upper() == "SMS")
            )
            
            # Also exclude rows where remarks start with "Predictive"
            predictive_cond = df[remarks_col].astype(str).str.startswith("Predictive", na=False)
            df_filtered = df[condition & ~predictive_cond].copy()
            
            # Generate the new output format
            output_df = pd.DataFrame()
            output_df["ACCOUNT NUMBER"] = df_filtered[customer_id_col].values
            
            # Formula generation: =XLOOKUP(...) starting from row 2
            # A2 references the "ACCOUNT NUMBER" column in the exported document
            formulas = [
                f"=VLOOKUP(A{i+2},'BPI DATABASE PL 2025 MINE.xlsx'!Table_Query_from_dbBCRM2[[#All],[account number]:[chname]],3,0)"
                for i in range(len(df_filtered))
            ]
            output_df["NAME"] = formulas
            output_df["USER ID"] = df_filtered[user_id_col].values
            output_df["ACTION CODE"] = df_filtered[action_code_col].values
            
            # Remarks length limit to 250
            output_df["REMARKS"] = df_filtered[remarks_col].astype(str).str[:250].values

            # --- DISPLAY SECTION ---
            st.divider()
            st.metric("Total Extracted Rows", len(output_df))

            st.write("### Data Preview")
            st.dataframe(output_df, use_container_width=True)

            # Export Logic
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name="Extracted Data")
            output.seek(0)

            st.sidebar.header("Actions")
            st.sidebar.download_button(
                label=" Download Extracted Data",
                data=output,
                file_name=f"FINNONE ENCODING {(pd.Timestamp.now() - pd.Timedelta(days=1)).strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.warning("Please ensure the uploaded file has the expected columns: customer_id, user_id, action_code, and remarks.")
    else:
        st.info("👆 Upload an encrypted FINNONE file to begin.")

if __name__ == "__main__":
    main()
