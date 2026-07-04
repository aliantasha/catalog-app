import io
import os
import re
import time
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Catalog Sync Hub", layout="wide")

# --- 1. TITLE & FILE MANAGEMENT ---
st.title("⚙️ Dynamic Multi-Row Catalog Hub")
st.markdown("---")

st.sidebar.header("📁 Base Repository Files")
st.sidebar.markdown(
    "Upload your master sheets here. If running locally, you can also place them in the project folder."
)

# File pickers: check sidebar upload first, fall back to local directory
uploaded_master = st.sidebar.file_uploader("Upload Catalog Master.xlsx", type=["xlsx"])
uploaded_buyer = st.sidebar.file_uploader("Upload Buyer File.xlsx", type=["xlsx"])

# Set paths or binary streams
MASTER_FILE = uploaded_master if uploaded_master else "Catalog Master.xlsx"
BUYER_FILE = uploaded_buyer if uploaded_buyer else "Buyer File.xlsx"

# Validate files exist before proceeding
files_ready = True
if isinstance(MASTER_FILE, str) and not os.path.exists(MASTER_FILE):
    st.sidebar.warning("⚠️ Missing 'Catalog Master.xlsx' layout.")
    files_ready = False
if isinstance(BUYER_FILE, str) and not os.path.exists(BUYER_FILE):
    st.sidebar.warning("⚠️ Missing 'Buyer File.xlsx' layout.")
    files_ready = False


# --- Helper: Sequential Sequence Finder ---
def get_highest_sequence_number(sheet, prefix):
    highest_num = 0
    r = 3
    empty_row_allowance = 0
    while True:
        cell_val = sheet.cell(row=r, column=2).value
        if cell_val is None or str(cell_val).strip() == "" or "PLEASE SELECT" in str(cell_val).upper():
            empty_row_allowance += 1
            if empty_row_allowance > 15: 
                break
        else:
            empty_row_allowance = 0
            cell_str = str(cell_val).strip().upper()
            match = re.match(rf"^{re.escape(prefix)}\s*(\d+)", cell_str)
            if match:
                val_num = int(match.group(1))
                if val_num > highest_num: 
                    highest_num = val_num
        r += 1
    return highest_num

def get_base_fallback(full_prefix):
    fallbacks = {"CP": 239, "CR": 8974, "NP": 12514, "NR": 16802, "CCP": 364, "CCR": 1575, "XP": 58, "XR": 110}
    return fallbacks.get(full_prefix, 2600)


if files_ready:
    # --- 2. LOAD PARTNER STRUCTURES (Cached for performance) ---
    @st.cache_data(ttl=60)
    def load_partner_database(master_source):
        # Handle bytes buffer or string paths seamlessly
        source = io.BytesIO(master_source.getvalue()) if hasattr(master_source, 'getvalue') else master_source
        wb_master = load_workbook(source, data_only=True)
        sheet_names = wb_master.sheetnames
        vendor_sheet_name = "PivVendorMaster" if "PivVendorMaster" in sheet_names else sheet_names[0]
        sheet_v = wb_master[vendor_sheet_name]

        partner_db = {}
        for row in range(1, sheet_v.max_row + 1):
            p_name = sheet_v.cell(row=row, column=1).value
            p_code = sheet_v.cell(row=row, column=3).value
            curr = sheet_v.cell(row=row, column=4).value
            ship_t = sheet_v.cell(row=row, column=6).value
            b_code = sheet_v.cell(row=row, column=16).value
            if p_name is not None:
                p_name_str = str(p_name).strip()
                p_name_upper = p_name_str.upper()
                if p_name_str == "" or "VENDOR_NAME" in p_name_upper or "PLEASE SELECT" in p_name_upper:
                    continue
                partner_db[p_name_upper] = {
                    "original_name": p_name_str,
                    "partner_code": str(p_code).strip() if p_code else "",
                    "buyer_code": str(b_code).strip().upper() if b_code else "ZN",
                    "currency": str(curr).strip().upper() if curr else "MYR",
                    "shipping_term": str(ship_t).strip().upper() if ship_t else "PLEASE SELECT"
                }
        wb_master.close()
        return partner_db

    partner_db = load_partner_database(MASTER_FILE)
    sorted_partners = sorted(list(set([v["original_name"] for v in partner_db.values()])))
    umsr_options = ["BT", "BUL", "BX", "CT", "DAY", "DRM", "GRM", "HR", "JOB", "KGM", "LOT", "LTR", "M3", "MAN", "MLT", "MM", "MTH", "NRL", "PCE", "PKT", "PR", "PT", "SET", "SHT", "TON", "TRP", "UT", "PLEASE SELECT"]
    shipping_options = ["CIF", "CIP", "DAP", "DAT", "DDP", "DDU", "EXW", "FCA", "FOB", "OTHERS", "PLEASE SELECT"]

    # --- 3. ASSIGN PROFILE METADATA ---
    col_ui1, col_ui2 = st.columns(2)
    with col_ui1:
        operator = st.selectbox("👤 Buyer Name:", ["PLEASE SELECT", "Luqman", "Alisya", "Tan"])
    with col_ui2:
        code_category = st.selectbox("🗂️ Category Type:", ["PLEASE SELECT", "Controllable ", "Comparison ", "Chemical", "Non-Controllable "])

    st.markdown("### 📝 Upload & Clean Request Rows")
    user_request_file = st.file_uploader("Upload your request spreadsheet (.xlsx)", type=["xlsx", "xls", "xlsm"])

    if user_request_file and operator != "PLEASE SELECT" and code_category != "PLEASE SELECT":
        # Process the uploaded template target rows
        wb_upload = load_workbook(user_request_file, data_only=True)
        target_sheets = [s for s in wb_upload.sheetnames if "PivVendorMaster" not in s]
        sheet = wb_upload[target_sheets[0]]

        col_map = {}
        for col in range(1, sheet.max_column + 1):
            header_val = str(sheet.cell(row=13, column=col).value or "").upper()
            if "PARTNER NAME" in header_val: col_map["partner"] = col
            elif "CATEGORY" in header_val: col_map["category"] = col
            elif "ITEM DESCRIPTION" in header_val and "46" in header_val: col_map["desc"] = col
            elif "PURCHASE UMSR" in header_val: col_map["umsr"] = col
            elif "VALIDITY" in header_val: col_map["validity"] = col
            elif "FINAL U/PRICE" in header_val: col_map["final_price"] = col
            elif "CURRENCY" in header_val: col_map["currency"] = col
            elif "INITIAL" in header_val: col_map["init_price"] = col
            elif "MOQ" in header_val: col_map["moq"] = col
            elif "LEAD TIME" in header_val: col_map["lead"] = col
            elif "SHIPPING TERM" in header_val: col_map["ship_term"] = col

        # Extract data into list for Streamlit's Interactive Data Editor
        staged_rows = []
        curr_row = 14
        while True:
            test_val = sheet.cell(row=curr_row, column=col_map.get("desc", 3)).value
            if test_val is None or str(test_val).strip() == "":
                break

            def get_row_val(key, default=""):
                if key in col_map:
                    v = sheet.cell(row=curr_row, column=col_map[key]).value
                    return default if v is None else str(v).strip()
                return default

            ext_partner = get_row_val("partner").upper()
            p_details = partner_db.get(ext_partner, {"partner_code": "", "buyer_code": "ZN", "currency": "MYR", "shipping_term": "PLEASE SELECT"})

            staged_rows.append({
                "Source Row": f"Row {curr_row}",
                "Partner Name": ext_partner if ext_partner in sorted_partners else sorted_partners[0],
                "Category": get_row_val("category", "Others"),
                "Description": get_row_val("desc").upper()[:46],
                "Partner Code": p_details["partner_code"],
                "Buyer Code": p_details["buyer_code"],
                "Purchase UMSR": get_row_val("umsr", "PCE").upper(),
                "Validity (YYYYMMDD)": get_row_val("validity").split('.')[0] if "none" not in get_row_val("validity").lower() else "",
                "Final Price": get_row_val("final_price"),
                "Currency": get_row_val("currency", p_details["currency"]).upper(),
                "Initial Price": get_row_val("init_price"),
                "MOQ": "",
                "Lead Time": "",
                "Shipping Term": get_row_val("ship_term", p_details["shipping_term"])
            })
            curr_row += 1
        wb_upload.close()

        if staged_rows:
            # Display spreadsheet using Streamlit Data Editor instead of drawing dozens of individual boxes
            df_staged = pd.DataFrame(staged_rows)
            
            st.info("💡 You can double-click any cell below to fix typo manual overrides directly inside the live grid!")
            edited_df = st.data_editor(
                df_staged,
                column_config={
                    "Partner Name": st.column_config.SelectboxColumn(options=sorted_partners),
                    "Purchase UMSR": st.column_config.SelectboxColumn(options=umsr_options),
                    "Shipping Term": st.column_config.SelectboxColumn(options=shipping_options),
                },
                hide_index=True,
                use_container_width=True
            )

            # --- 4. EXCEL GENERATION ENGINE ---
            if st.button("🚀 Confirm & Process All Rows", type="primary", use_container_width=True):
                try:
                    # Target prefix structures
                    suffix_char = "R" if operator == "Luqman" else ("P" if operator == "Alisya" else "D")
                    prefix_letter = "C" if code_category == "Controllable " else ("CC" if code_category == "Comparison " else ("X" if code_category == "Chemical" else "N"))
                    full_prefix = f"{prefix_letter}{suffix_char}"

                    sheet_mapping = {"Controllable ": "Controllable", "Comparison ": "Comparison", "Chemical": "Chemical (X)", "Non-Controllable ": "Non Controllable"}
                    target_sheet_name = sheet_mapping.get(code_category, "Non Controllable")

                    # Load template configurations safely
                    src_buyer = io.BytesIO(BUYER_FILE.getvalue()) if hasattr(BUYER_FILE, 'getvalue') else BUYER_FILE
                    wb_b_check = load_workbook(src_buyer, data_only=True)
                    highest_seq = 0
                    if target_sheet_name in wb_b_check.sheetnames:
                        highest_seq = get_highest_sequence_number(wb_b_check[target_sheet_name], full_prefix)
                    wb_b_check.close()
                    
                    if highest_seq == 0: 
                        highest_seq = get_base_fallback(full_prefix)

                    # Open read/write workbook buffers
                    src_master_rw = io.BytesIO(MASTER_FILE.getvalue()) if hasattr(MASTER_FILE, 'getvalue') else MASTER_FILE
                    wb_m = load_workbook(src_master_rw)
                    m_sheets = [s for s in wb_m.sheetnames if s != "PivVendorMaster"]
                    sheet_m = wb_m[m_sheets[0]]
                    m_row = sheet_m.max_row + 1

                    wb_b = load_workbook(src_buyer)
                    sheet_b = wb_b[target_sheet_name] if target_sheet_name in wb_b.sheetnames else wb_b.create_sheet(title=target_sheet_name)
                    b_row = sheet_b.max_row + 1

                    # Processing updates
                    for index, row_data in edited_df.iterrows():
                        highest_seq += 1
                        generated_code = f"{full_prefix}{highest_seq}"
                        
                        def clean_int(val): return int(float(val)) if str(val).strip().replace('.0','').isdigit() else None
                        def clean_float(val):
                            try: return float(val)
                            except: return None

                        # Write to Catalog Master
                        sheet_m.cell(row=m_row, column=1, value=max(1, m_row - 13))
                        sheet_m.cell(row=m_row, column=2, value=str(row_data["Partner Name"]).upper())
                        sheet_m.cell(row=m_row, column=3, value=row_data["Category"])
                        sheet_m.cell(row=m_row, column=4, value=str(row_data["Description"]).upper()[:46])
                        sheet_m.cell(row=m_row, column=5, value=generated_code)
                        sheet_m.cell(row=m_row, column=6, value="RST")
                        sheet_m.cell(row=m_row, column=7, value=str(row_data["Description"]).upper()[:46])
                        sheet_m.cell(row=m_row, column=8, value=str(row_data["Partner Code"]).upper())
                        sheet_m.cell(row=m_row, column=9, value=str(row_data["Buyer Code"]).upper())
                        sheet_m.cell(row=m_row, column=10, value=row_data["Purchase UMSR"])
                        sheet_m.cell(row=m_row, column=11, value=0)
                        sheet_m.cell(row=m_row, column=13, value=row_data["Validity (YYYYMMDD)"])
                        sheet_m.cell(row=m_row, column=14, value=clean_float(row_data["Final Price"]))
                        sheet_m.cell(row=m_row, column=15, value=str(row_data["Currency"]).upper())
                        sheet_m.cell(row=m_row, column=16, value=clean_float(row_data["Initial Price"]))
                        sheet_m.cell(row=m_row, column=17, value=clean_int(row_data["MOQ"]))
                        sheet_m.cell(row=m_row, column=19, value=clean_int(row_data["Lead Time"]))
                        sheet_m.cell(row=m_row, column=20, value=row_data["Shipping Term"])

                        # Write to Buyer File
                        sheet_b.cell(row=b_row, column=1, value="")
                        sheet_b.cell(row=b_row, column=2, value=generated_code)
                        sheet_b.cell(row=b_row, column=3, value=str(row_data["Description"]).upper()[:46])
                        sheet_b.cell(row=b_row, column=4, value=str(row_data["Partner Name"]).upper())

                        m_row += 1
                        b_row += 1

                    # Save both file engines into memory buffers for explicit web downloads
                    buffer_m = io.BytesIO()
                    wb_m.save(buffer_m)
                    wb_m.close()

                    buffer_b = io.BytesIO()
                    wb_b.save(buffer_b)
                    wb_b.close()

                    st.success("🎉 Processing complete! Download your updated files below:")
                    
                    # Generate download layouts side-by-side
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        st.download_button(
                            label="📥 Download Updated Catalog Master",
                            data=buffer_m.getvalue(),
                            file_name="Catalog Master_Updated.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with col_dl2:
                        st.download_button(
                            label="📥 Download Updated Buyer File",
                            data=buffer_b.getvalue(),
                            file_name="Buyer File_Updated.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as ex:
                    st.error(f"❌ Processing Fault: {ex}")
        else:
            st.warning("⚠️ No valid data rows detected in row layout.")
    elif user_request_file:
        st.warning("👉 Please configure both a specific 'Buyer Name' and 'Category Type' option fields above first.")
else:
    st.info("ℹ️ To begin, please upload the primary configurations ('Catalog Master.xlsx' and 'Buyer File.xlsx') via the sidebar menu.")
