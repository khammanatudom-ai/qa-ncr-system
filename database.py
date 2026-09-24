import os
import re
import pandas as pd
from datetime import datetime, timedelta

TRACKING_FILE = "claim_tracking.xlsx"

def extract_clean_email(raw_text):
    if not raw_text or pd.isna(raw_text): return ""
    text_str = str(raw_text).strip()
    match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text_str)
    if match: return match.group(0)
    return text_str

def lookup_part_master(part_no, excel_file="part_master.xlsx"):
    if not os.path.exists(excel_file): return None
    try:
        df = pd.read_excel(excel_file)
        clean_target = str(part_no).strip().upper().replace(" ", "").replace("-", "")
        df['clean_part_no'] = df['Part no'].astype(str).str.strip().str.upper().str.replace(" ", "").str.replace("-", "")
        match = df[df['clean_part_no'] == clean_target]
        if not match.empty:
            row = match.iloc[0]
            raw_email = row.get('Supplier Email', '')
            clean_email = extract_clean_email(raw_email)
            return {
                "part_name": str(row['Part name']).strip(),
                "supplier": str(row['Supplier Name']).strip(),
                "supplier_email": clean_email,
                "found": True
            }
    except Exception as e:
        print(f"Error reading master excel: {e}")
    return None

def init_tracking_db():
    if not os.path.exists(TRACKING_FILE):
        df = pd.DataFrame(columns=[
            "NCR_No", "Issue_Date", "Customer", "Part_No", "Part_Name", "Supplier", 
            "Problem_Summary", "Status", "SLA_Due_Date", 
            "Root_Cause", "Countermeasure", "Supplier_File", "WhyWhy_File"
        ])
        df.to_excel(TRACKING_FILE, index=False)
    else:
        df = pd.read_excel(TRACKING_FILE)
        changed = False
        if "Customer" not in df.columns:
            df["Customer"] = "General"
            changed = True
        if "Supplier_File" not in df.columns:
            df["Supplier_File"] = df.get("Reply_File", "-")
            changed = True
        if "WhyWhy_File" not in df.columns:
            df["WhyWhy_File"] = "-"
            changed = True
        if changed:
            df.to_excel(TRACKING_FILE, index=False)

def generate_ncr_no():
    init_tracking_db()
    df = pd.read_excel(TRACKING_FILE)
    now = datetime.now()
    prefix = f"NCR{now.strftime('%y%m')}-"
    current_month_ncrs = df[df['NCR_No'].str.startswith(prefix, na=False)]
    if current_month_ncrs.empty:
        return f"{prefix}001"
    else:
        nums = current_month_ncrs['NCR_No'].str.replace(prefix, "").astype(int)
        next_num = nums.max() + 1
        return f"{prefix}{next_num:03d}"

def save_new_claim(ncr_no, part_no, part_name, supplier, problem_summary, customer="General"):
    df = pd.read_excel(TRACKING_FILE)
    now = datetime.now()
    sla_due = now + timedelta(days=10)
    
    new_row = {
        "NCR_No": ncr_no,
        "Issue_Date": now.strftime("%Y-%m-%d"),
        "Customer": customer,
        "Part_No": part_no,
        "Part_Name": part_name,
        "Supplier": supplier,
        "Problem_Summary": str(problem_summary)[:150],
        "Status": "Waiting for Supplier",
        "SLA_Due_Date": sla_due.strftime("%Y-%m-%d"),
        "Root_Cause": "-",
        "Countermeasure": "-",
        "Supplier_File": "-",
        "WhyWhy_File": "-"
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(TRACKING_FILE, index=False)

def load_tracking_data():
    init_tracking_db()
    return pd.read_excel(TRACKING_FILE)

def update_claim(ncr_no, status, root_cause, countermeasure, supplier_file="-", whywhy_file="-"):
    df = pd.read_excel(TRACKING_FILE)
    idx = df[df['NCR_No'] == ncr_no].index
    if not idx.empty:
        df.loc[idx, 'Status'] = status
        if root_cause: df.loc[idx, 'Root_Cause'] = root_cause
        if countermeasure: df.loc[idx, 'Countermeasure'] = countermeasure
        if supplier_file and supplier_file != "-": df.loc[idx, 'Supplier_File'] = supplier_file
        if whywhy_file and whywhy_file != "-": df.loc[idx, 'WhyWhy_File'] = whywhy_file
        df.to_excel(TRACKING_FILE, index=False)