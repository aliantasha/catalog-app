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
        # Load the uploaded file using openpyxl and get the active sheet
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active

        staged_rows = []
        curr_row = 14
        
        while True:
            desc_val = sheet.cell(row=curr_row, column=3).value
            if desc_val is None or str(desc_val).strip() == "":
                break
            
            # 🛡️ SAFE PRICE CHECK: If it sees text like 'GEN', it will default to 0.0 instead of crashing!
            raw_price = sheet.cell(row=curr_row, column=6).value
            try:
                clean_price = float(raw_price) if raw_price is not None else 0.0
            except (ValueError, TypeError):
                clean_price = 0.0  # Safe fallback if it's a string like 'GEN' or None
                
            staged_rows.append({
                "Partner Name": str(sheet.cell(row=curr_row, column=1).value or "").upper(),
                "Category": str(sheet.cell(row=curr_row, column=2).value or "Goods/Item"),
                "Item Description (46 Chars)": str(desc_val).upper()[:46],
                "Price": clean_price,
                "Currency": str(sheet.cell(row=curr_row, column=7).value or "MYR").upper(),
                "UMSR": str(sheet.cell(row=curr_row, column=4).value or "PCE").upper(),
                "Validity": str(sheet.cell(row=curr_row, column=5).value or "").split('.')[0]
            })
            curr_row += 1

        # Display the data in an editable dataframe matrix
        if staged_rows:
            df = pd.DataFrame(staged_rows)
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        else:
            st.warning("⚠️ No data rows found starting from row 14.")

    except Exception as e:
        st.error(f"❌ Error processing spreadsheet: {e}")
