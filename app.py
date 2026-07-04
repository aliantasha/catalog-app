import streamlit as st
import pandas as pd
import openpyxl
import re
import io

# --- 1. WEB PAGE SETTINGS ---
st.set_page_config(page_title="Catalog Management Hub", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size:30px; font-weight:bold; color:#1a365d; margin-bottom:2px; }
    .sub-title { font-size:15px; color:#4a5568; margin-bottom:20px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 Dynamic Catalog Overrides Hub (No-Cloud Mode)</div>', unsafe_allow_html=True)

# --- 2. SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("👤 Profile Settings")
    buyer = st.selectbox("Assign Buyer Profile:", ["PLEASE SELECT", "Luqman", "Alisya", "Tan"])
    cat_type = st.selectbox("Assign Category Type:", ["PLEASE SELECT", "Controllable", "Comparison", "Chemical", "Non-Controllable"])
    st.markdown("---")
    uploaded_file = st.file_uploader("📂 Drop Request Spreadsheet Here", type=["xlsx", "xls", "xlsm"])

# --- 3. RENDER DATA GRID ---
if uploaded_file is not None and buyer != "PLEASE SELECT" and cat_type != "PLEASE SELECT":
    st.subheader("📝 Live Spreadsheet Edit Matrix")
    st.info("💡 Make corrections directly in the table cells below.")

    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active

        # --- DYNAMIC HEADER MAPPER (Restored from your original logic) ---
        col_map = {}
        # Scans row 13 for matching headers
        for col in range(1, sheet.max_column + 1):
            header_val = str(sheet.cell(row=13, column=col).value or "").upper()
            if "PARTNER NAME" in header_val: col_map["partner"] = col
            elif "CATEGORY" in header_val: col_map["category"] = col
            elif "ITEM DESCRIPTION" in header_val and "46" in header_val: col_map["desc"] = col
            elif "PURCHASE UMSR" in header_val: col_map["umsr"] = col
            elif "VALIDITY" in header_val: col_map["validity"] = col
            elif "FINAL U/PRICE" in header_val: col_map["final_price"] = col
            elif "CURRENCY" in header_val: col_map["currency"] = col

        staged_rows = []
        curr_row = 14
        
        while True:
            # Look up values using the dynamic column map, fallback safely if column not found
            desc_col = col_map.get("desc", 3) 
            desc_val = sheet.cell(row=curr_row, column=desc_col).value
            
            if desc_val is None or str(desc_val).strip() == "":
                break
                
            def get_row_val(key, fallback_col, default=""):
                col_idx = col_map.get(key, fallback_col)
                v = sheet.cell(row=curr_row, column=col_idx).value
                return default if v is None else str(v).strip()

            # Safe Price Conversion
            raw_price = sheet.cell(row=curr_row, column=col_map.get("final_price", 6)).value
            try:
                clean_price = float(raw_price) if raw_price is not None else 0.0
            except (ValueError, TypeError):
                clean_price = 0.0  
                
            staged_rows.append({
                "Partner Name": get_row_val("partner", 1).upper(),
                "Category": get_row_val("category", 2, "Goods/Item"),
                "Item Description (46 Chars)": str(desc_val).upper()[:46],
                "Price": clean_price,
                "Currency": get_row_val("currency", 7, "MYR").upper(),
                "UMSR": get_row_val("umsr", 4, "PCE").upper(),
                "Validity": get_row_val("validity", 5).split('.')[0]
            })
            curr_row += 1

        # Display editable UI matrix
        if staged_rows:
            df = pd.DataFrame(staged_rows)
            
            # st.data_editor lets users change, add, or delete items instantly
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
            
            st.markdown("---")
            
            # --- 4. EXPORT ENGINE ---
            # Because you are in "No-Cloud Mode", let users download their corrections as a fresh file
            st.subheader("💾 4. Export Staged Data")
            
            # Convert updated dataframe back into a streamable Excel bytes array
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                edited_df.to_excel(writer, index=False, sheet_name="Corrected Catalog")
            
            st.download_button(
                label="📥 Download Corrected Spreadsheet",
                data=buffer.getvalue(),
                file_name=f"Corrected_Catalog_{buyer}_{cat_type}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        else:
            st.warning("⚠️ No data rows found starting from row 14 matching header targets.")

    except Exception as e:
        st.error(f"❌ Error processing spreadsheet: {e}")
else:
    st.info("👋 Please assign your Buyer Profile, Category Type, and drop a Request Spreadsheet in the sidebar to begin editing.")
   
