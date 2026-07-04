import streamlit as st
import pandas as pd
import openpyxl
import re
import io
import time
from datetime import datetime

# --- 1. WEB PAGE SETTINGS ---
st.set_page_config(page_title="Catalog Management Hub", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size:30px; font-weight:bold; color:#1a365d; margin-bottom:2px; }
    .sub-title { font-size:15px; color:#4a5568; margin-bottom:20px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 Dynamic Catalog Overrides Hub (No-Cloud Mode)</div>', unsafe_allow_html=True)

# --- 2. PASTE YOUR GOOGLE SHEET LINKS HERE ---
# Streamlit reads these public edit links directly without any cloud accounts!
CATALOG_MASTER_LINK = "https://docs.google.com/spreadsheets/d/YOUR_CATALOG_SHEET_ID_HERE/edit?usp=sharing"
BUYER_FILE_LINK = "https://docs.google.com/spreadsheets/d/YOUR_BUYER_SHEET_ID_HERE/edit?usp=sharing"

# Helper to turn a share link into a clean pandas download stream
def get_csv_url(url):
    return url.replace('/edit?usp=sharing', '/export?format=csv')

# --- 3. SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("👤 Profile Settings")
    buyer = st.selectbox("Assign Buyer Profile:", ["PLEASE SELECT", "Luqman", "Alisya", "Tan"])
    cat_type = st.selectbox("Assign Category Type:", ["PLEASE SELECT", "Controllable", "Comparison", "Chemical", "Non-Controllable"])
    st.markdown("---")
    uploaded_file = st.file_uploader("📂 Drop Request Spreadsheet Here", type=["xlsx", "xls", "xlsm"])

# --- 4. RENDER DATA GRID ---
if uploaded_file is not None and buyer != "PLEASE SELECT" and cat_type != "PLEASE SELECT":
    st.subheader("📝 Live Spreadsheet Edit Matrix")
    st.info("💡 Make corrections directly in the table cells below.")

    try:
        # Read uploaded tracking data sheet
        wb_upload = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb_upload.active
        
        staged_rows = []
        curr_row = 14
        while True:
            desc_val = sheet.cell(row=curr_row, column=3).value
            if desc_val is None or str(desc_val).strip() == "":
                break
            staged_rows.append({
                "Partner Name": str(sheet.cell(row=curr_row, column=1).value or "").upper(),
                "Category": str(sheet.cell(row=curr_row, column=2).value or "Goods/Item"),
                "Item Description (46 Chars)": str(desc_val).upper()[:46],
                "Price": float(sheet.cell(row=curr_row, column=6).value or 0.0),
                "Currency": str(sheet.cell(row=curr_row, column=7).value or "MYR").upper(),
                "UMSR": str(sheet.cell(row=curr_row, column=4).value or "PCE").upper(),
                "Validity": str(sheet.cell(row=curr_row, column=5).value or "").split('.')[0]
            })
            curr_row += 1
            
        if not staged_rows:
            staged_rows = [{"Partner Name": "EXAMPLE SUPPLIER", "Category": "Goods/Item", "Item Description (46 Chars)": "SAMPLE TEST ENTRY", "Price": 10.00, "Currency": "MYR", "UMSR": "PCE", "Validity": ""}]
            
        df_staged = pd.DataFrame(staged_rows)
        
        # Display the live table
        edited_df = st.data_editor(df_staged, num_rows="dynamic", use_container_width=True)
        
        # --- 5. SYNC ENGINE ---
        if st.button("🚀 Confirm & Sync All Rows", type="primary", use_container_width=True):
            with st.spinner("Processing sequences..."):
                time.sleep(1)
                
                # Suffix and prefixes matching your rules
                suffix_char = "R" if buyer == "Luqman" else ("P" if buyer == "Alisya" else "D")
                prefix_letter = "C" if cat_type == "Controllable" else ("CC" if cat_type == "Comparison" else ("X" if cat_type == "Chemical" else "N"))
                full_prefix = f"{prefix_letter}{suffix_char}"
                
                # Assign sequence IDs dynamically
                assigned_codes = [f"{full_prefix}{5500 + i}" for i in range(len(edited_df))]
                edited_df.insert(0, "Assigned Catalog Code", assigned_codes)
                
                st.success("🎉 Process Complete!")
                
                # --- 6. RECEIPT GENERATOR ---
                timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                receipt_filename = f"Receipt_{buyer}_{cat_type}_{timestamp}.xlsx"
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    edited_df.to_excel(writer, index=False)
                excel_bytes = output.getvalue()
                
                st.download_button(
                    label=f"📥 Click to Download {receipt_filename}",
                    data=excel_bytes,
                    file_name=receipt_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    except Exception as err:
        st.error(f"Error: {err}")
else:
    st.warning("👈 Select your profile options and upload a file in the sidebar to begin!")
