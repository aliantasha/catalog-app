import os
import re
import time
import shutil
import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# --- CONFIG & PATHS ---
st.set_page_config(layout="wide", page_title="Catalog Hub")
st.title("⚙️ Dynamic Multi-Row Catalog Hub")
st.caption("Manual Correction and Live Sync Mode")

# Adjust these paths to your server or local directory structure
PROJECT_FOLDER = "./" 
MASTER_CATALOG_FILENAME = "Catalog Master.xlsx"
MASTER_CATALOG_PATH = os.path.join(PROJECT_FOLDER, MASTER_CATALOG_FILENAME)
BUYER_FILE_NAME = "Buyer File.xlsx"
BUYER_FILE_PATH = os.path.join(PROJECT_FOLDER, BUYER_FILE_NAME)
LOCK_FILE_PATH = os.path.join(PROJECT_FOLDER, "drive_write.lock")

# Baseline files check
if not os.path.exists(MASTER_CATALOG_PATH) or not os.path.exists(BUYER_FILE_PATH):
    st.error("❌ Missing baseline master dependencies in project folder.")
    st.stop()

# --- 1. CORE DATA CACHING & PARSING FUNCTIONS ---
@st.cache_data(ttl=60) # Caches the data for 60s so it doesn't re-read from Excel on every tiny button click
def load_partner_db():
    wb_master = load_workbook(MASTER_CATALOG_PATH, data_only=True)
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

partner_db = load_partner_db()
sorted_partners = sorted(list(set([v["original_name"] for v in partner_db.values()])))

umsr_options = ["BT", "BUL", "BX", "CT", "DAY", "DRM", "GRM", "HR", "JOB", "KGM", "LOT", "LTR", "M3", "MAN", "MLT", "MM", "MTH", "NRL", "PCE", "PKT", "PR", "PT", "SET", "SHT", "TON", "TRP", "UT", "PLEASE SELECT"]
shipping_options = ["CIF", "CIP", "DAP", "DAT", "DDP", "DDU", "EXW", "FCA", "FOB", "OTHERS", "PLEASE SELECT"]

def get_highest_sequence_number_from_buyer_file(sheet, prefix):
    highest_num = 0
    r = 3
    empty_row_allowance = 0
    while True:
        cell_val = sheet.cell(row=r, column=2).value
        if cell_val is None or str(cell_val).strip() == "" or "PLEASE SELECT" in str(cell_val).upper():
            empty_row_allowance += 1
            if empty_row_allowance > 15: break
        else:
            empty_row_allowance = 0
            cell_str = str(cell_val).strip().upper()
            match = re.match(rf"^{re.escape(prefix)}\s*(\d+)", cell_str)
            if match:
                val_num = int(match.group(1))
                if val_num > highest_num: highest_num = val_num
        r += 1
    return highest_num

def get_base_fallback(full_prefix):
    fallbacks = {"CP": 239, "CR": 8974, "NP": 12514, "NR": 16802, "CCP": 364, "CCR": 1575, "XP": 58, "XR": 110}
    return fallbacks.get(full_prefix, 2600)

def safe_w(sheet, r, c, val):
    if type(sheet.cell(row=r, column=c)).__name__ != 'MergedCell':
        sheet.cell(row=r, column=c).value = val

# --- 2. SIDEBAR / META CHOICES ---
st.sidebar.subheader("👤 User Selection Profiles")
operator = st.sidebar.selectbox("Select Buyer Profile:", ["PLEASE SELECT", "Luqman", "Alisya", "Tan"])
category_type = st.sidebar.selectbox("Select Category Type:", ["PLEASE SELECT", "Controllable ", "Comparison ", "Chemical", "Non-Controllable "])

# --- 3. DYNAMIC LOGIC FOR PREFIX GENERATION ---
ready_for_preview = operator != "PLEASE SELECT" and category_type != "PLEASE SELECT"
full_prefix = ""
target_sheet_name = ""

if ready_for_preview:
    suffix_char = "R" if operator == "Luqman" else ("P" if operator == "Alisya" else "D")
    prefix_letter = "C" if category_type == "Controllable " else ("CC" if category_type == "Comparison " else ("X" if category_type == "Chemical" else "N"))
    full_prefix = f"{prefix_letter}{suffix_char}"
    
    sheet_mapping = {"Controllable ": "Controllable", "Comparison ": "Comparison", "Chemical": "Chemical (X)", "Non-Controllable ": "Non Controllable"}
    target_sheet_name = sheet_mapping.get(category_type, "Non Controllable")

# --- 4. DATA HANDLING AND UPLOAD ENGINE ---
uploaded_file = st.file_uploader("📂 Step 1: Upload Request File (.xlsx)", type=["xlsx", "xlsm"])

if uploaded_file and ready_for_preview:
    # We parse the file into a clean data editor view
    wb_upload = load_workbook(io.BytesIO(uploaded_file.read()), data_only=True)
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

    rows_data = []
    curr_row = 14
    
    # Get the initial sequence counter setup ahead of matrix generation loop
    wb_b_check = load_workbook(BUYER_FILE_PATH, data_only=True)
    highest_seq = 0
    if target_sheet_name in wb_b_check.sheetnames:
        highest_seq = get_highest_sequence_number_from_buyer_file(wb_b_check[target_sheet_name], full_prefix)
    wb_b_check.close()
    if highest_seq == 0: highest_seq = get_base_fallback(full_prefix)
    
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
        
        highest_seq += 1
        
        rows_data.append({
            "Original Row": curr_row,
            "Catalog Code": f"{full_prefix}{highest_seq}",
            "Partner Name": ext_partner if ext_partner in sorted_partners else sorted_partners[0],
            "Category": get_row_val("category", "Others"),
            "Item Description": get_row_val("desc").upper()[:46],
            "Partner Code": p_details["partner_code"],
            "Buyer Code": p_details["buyer_code"],
            "UMSR": get_row_val("umsr", "PCE").upper(),
            "Validity": get_row_val("validity").split('.')[0] if "none" not in get_row_val("validity").lower() else "",
            "Final Price": get_row_val("final_price"),
            "Currency": get_row_val("currency", p_details["currency"]).upper(),
            "Initial Price": get_row_val("init_price"),
            "MOQ": "",
            "Lead Time": "",
            "Shipping Term": get_row_val("ship_term", p_details["shipping_term"])
        })
        curr_row += 1
    wb_upload.close()

    df = pd.DataFrame(rows_data)
    
    st.subheader("📝 Step 2: Correct Staged Data Inline")
    # Streamlit's awesome interactable data matrix 
    edited_df = st.data_editor(
        df,
        column_config={
            "Partner Name": st.column_config.SelectboxColumn(options=sorted_partners),
            "Category": st.column_config.SelectboxColumn(options=["Goods/Item", "Others", "blank"]),
            "UMSR": st.column_config.SelectboxColumn(options=umsr_options),
            "Shipping Term": st.column_config.SelectboxColumn(options=shipping_options),
            "Catalog Code": st.column_config.TextColumn(disabled=True)
        },
        use_container_width=True,
        hide_index=True
    )
    
    # --- 5. DATA SYNCHRONIZATION AND SAVE ENGINE ---
    st.markdown("---")
    if st.button("🚀 Confirm & Sync All Rows to Drive Source Files", type="primary", use_container_width=True):
        
        # Concurrency Lock Protection
        retry_count = 0
        lock_acquired = False
        while os.path.exists(LOCK_FILE_PATH):
            st.warning("⏳ Another buyer is sync saving data right now. Retrying queue clear shortly...")
            time.sleep(2)
            retry_count += 1
            if retry_count > 5:
                st.error("⚠️ Sync Error: Database access channel is heavily congested. Please retry saving data.")
                st.stop()
                
        with open(LOCK_FILE_PATH, 'w') as lock_f:
            lock_f.write("locked")
            
        try:
            # Fresh sequence count reread post lock
            wb_b_check = load_workbook(BUYER_FILE_PATH, data_only=True)
            highest_seq = 0
            if target_sheet_name in wb_b_check.sheetnames:
                highest_seq = get_highest_sequence_number_from_buyer_file(wb_b_check[target_sheet_name], full_prefix)
            wb_b_check.close()
            if highest_seq == 0: highest_seq = get_base_fallback(full_prefix)
            
            wb_m = load_workbook(MASTER_CATALOG_PATH)
            m_sheets = [s for s in wb_m.sheetnames if s != "PivVendorMaster"]
            sheet_m = wb_m[m_sheets[0]]
            m_row = sheet_m.max_row + 1
            
            wb_b = load_workbook(BUYER_FILE_PATH)
            sheet_b = wb_b[target_sheet_name] if target_sheet_name in wb_b.sheetnames else wb_b.create_sheet(title=target_sheet_name)
            b_row = sheet_b.max_row + 1
            
            progress_bar = st.progress(0)
            total_items = len(edited_df)
            
            for idx, row_data in edited_df.iterrows():
                highest_seq += 1
                generated_code = f"{full_prefix}{highest_seq}"
                
                # Check live changes made by partner drop downs to update structural mapping fields automatically
                p_sel = row_data["Partner Name"].upper()
                live_pcode = partner_db.get(p_sel, {}).get("partner_code", row_data["Partner Code"])
                live_bcode = partner_db.get(p_sel, {}).get("buyer_code", row_data["Buyer Code"])
                live_curr = partner_db.get(p_sel, {}).get("currency", row_data["Currency"])
                
                # Write to Catalog Master
                safe_w(sheet_m, m_row, 1, max(1, m_row - 13))
                safe_w(sheet_m, m_row, 2, p_sel)
                safe_w(sheet_m, m_row, 3, row_data["Category"])
                safe_w(sheet_m, m_row, 4, str(row_data["Item Description"]).strip().upper())
                safe_w(sheet_m, m_row, 5, generated_code)
                safe_w(sheet_m, m_row, 6, "RST")
                safe_w(sheet_m, m_row, 7, str(row_data["Item Description"]).strip().upper())
                safe_w(sheet_m, m_row, 8, str(live_pcode).upper())
                safe_w(sheet_m, m_row, 9, str(live_bcode).upper())
                safe_w(sheet_m, m_row, 10, row_data["UMSR"])
                safe_w(sheet_m, m_row, 11, 0)
                safe_w(sheet_m, m_row, 12, "")
                safe_w(sheet_m, m_row, 13, row_data["Validity"])
                
                try: safe_w(sheet_m, m_row, 14, float(row_data["Final Price"]))
                except: safe_w(sheet_m, m_row, 14, None)
                
                safe_w(sheet_m, m_row, 15, str(live_curr).upper())
                
                try: safe_w(sheet_m, m_row, 16, float(row_data["Initial Price"]))
                except: safe_w(sheet_m, m_row, 16, None)
                
                try: safe_w(sheet_m, m_row, 17, int(row_data["MOQ"]))
                except: safe_w(sheet_m, m_row, 17, None)
                
                safe_w(sheet_m, m_row, 18, "")
                
                try: safe_w(sheet_m, m_row, 19, int(row_data["Lead Time"]))
                except: safe_w(sheet_m, m_row, 19, None)
                
                safe_w(sheet_m, m_row, 20, row_data["Shipping Term"])
                
                # Write to Buyer File
                safe_w(sheet_b, b_row, 1, "")
                safe_w(sheet_b, b_row, 2, generated_code)
                safe_w(sheet_b, b_row, 3, str(row_data["Item Description"]).strip().upper())
                safe_w(sheet_b, b_row, 4, p_sel)
                
                m_row += 1
                b_row += 1
                progress_bar.progress((idx + 1) / total_items)
            
            # Atomic save routines
            temp_catalog_path = MASTER_CATALOG_PATH + ".tmp"
            temp_buyer_path = BUYER_FILE_PATH + ".tmp"
            wb_m.save(temp_catalog_path)
            wb_b.save(temp_buyer_path)
            wb_m.close()
            wb_b.close()
            
            shutil.move(temp_catalog_path, MASTER_CATALOG_PATH)
            shutil.move(temp_buyer_path, BUYER_FILE_PATH)
            
            st.success("🎉 All rows safely committed to system records database file successfully!")
            st.balloons()
            
        except Exception as ex:
            st.error(f"❌ Database Transaction Error: {ex}")
        finally:
            if os.path.exists(LOCK_FILE_PATH):
                os.remove(LOCK_FILE_PATH)

else:
    if not ready_for_preview:
        st.info("💡 Please specify your assigned **Buyer Profile** and **Category Type** on the left side-panel options block to begin workflow staging arrays.")
