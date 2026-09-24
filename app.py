import streamlit as st
import importlib
import os
import re
import ai_engine
import ncr_exporter
import database
from datetime import datetime
import pandas as pd
from gmail_notifier import send_ncr_alert

# Reload modules สำหรับการอัปเดตแบบ Live
importlib.reload(ai_engine)
importlib.reload(ncr_exporter)
importlib.reload(database)

# ฟังก์ชันช่วยสร้างโฟลเดอร์ตามโครงสร้างลูกค้ารายเจ้า
def init_customer_folders():
    # 1. ลูกค้า SKC
    os.makedirs("SKC/NCR_Issued", exist_ok=True)
    os.makedirs("SKC/Supplier_Report", exist_ok=True)
    os.makedirs("SKC/WhyWhy_Approved", exist_ok=True)
    
    # 2. ลูกค้า VCS
    os.makedirs("VCS/Customer_Claims", exist_ok=True)
    os.makedirs("VCS/Supplier_Report", exist_ok=True)
    os.makedirs("VCS/Completed_Forms_For_Review", exist_ok=True)
    
    # 3. ลูกค้า Yanmar
    os.makedirs("Yanmar/Customer_Claims", exist_ok=True)
    os.makedirs("Yanmar/Supplier_Report", exist_ok=True)
    os.makedirs("Yanmar/Completed_Forms_For_Review", exist_ok=True)

init_customer_folders()

st.set_page_config(page_title="QA Claim System", layout="wide")
st.title("🛠️ QA Overseas Claim & NCR Management System")

# รายชื่อลูกค้าตามระบบใหม่
CUSTOMER_LIST = [
    "1. SKC (General PPTX Form)",
    "2. VCS (U-SHIN Excel Form)",
    "3. Yanmar (YSP Meter Excel Form)"
]

def get_customer_key(cust_str):
    """ฟังก์ชันระบุโฟลเดอร์หลักจากชื่อลูกค้า"""
    if "SKC" in str(cust_str):
        return "SKC"
    elif "VCS" in str(cust_str):
        return "VCS"
    elif "Yanmar" in str(cust_str):
        return "Yanmar"
    return "SKC"

tab1, tab2, tab3 = st.tabs(["🆕 1. รับเคลมลูกค้า & ร่างแจ้ง Supplier", "📋 2. สรุปรายงาน & เติมแบบฟอร์มส่งลูกค้า", "📈 3. รายงานสรุป Management"])

# ==========================================
# TAB 1: RECEIVE CLAIM & DRAFT SUPPLIER EMAIL
# ==========================================
with tab1:
    st.subheader("1. บันทึกเคสจากลูกค้า & ออกเอกสารอ้างอิง")
    col1, col2 = st.columns([1, 1])
    
    with col1:
        selected_customer = st.selectbox("🏢 เลือกลูกค้าเจ้าของเคส (Customer):", CUSTOMER_LIST)
        cust_folder = get_customer_key(selected_customer)
        
        # เช็กเงื่อนไขการออกเลขเอกสาร
        if cust_folder == "SKC":
            st.info("ℹ️ ลูกค้า SKC: ระบบจะรันหมายเลข NCR No. อัตโนมัติ (เช่น NCR2609-001)")
            customer_doc_no = None
        else:
            customer_doc_no = st.text_input(
                "📌 ระบุหมายเลขเอกสารเคลมของลูกค้า (Customer Claim Ref / Doc No.):",
                value=f"CLAIM-{datetime.now().strftime('%Y%m%d-%H%M')}"
            )
            st.info(f"ℹ️ ลูกค้า {selected_customer}: ไม่รันเลข NCR (ใช้อ้างอิงรหัส: **{customer_doc_no}**)")

        raw_text = st.text_area("วางข้อความแจ้งปัญหาจากลูกค้า (Raw Data):", height=180,
                                value="""#แจ้งปัญหาคุณภาพ #\nวันที่พบปัญหา: 28/08/2026\nPart No: W95J2-1338-1\nปัญหา: Mold ปลั๊กสายไฟไม่เต็ม""")
        
    with col2:
        customer_file = st.file_uploader("แนบไฟล์เคลมต้นฉบับจากลูกค้า (Excel/PDF/Image):", type=['xlsx', 'xls', 'pdf', 'png', 'jpg', 'jpeg'])

    if st.button("🤖 1. บันทึกเคส & ร่างอีเมลแจ้ง Supplier"):
        if raw_text:
            with st.spinner("กำลังประมวลผลข้อมูล..."):
                try:
                    result = ai_engine.process_ncr_data(raw_text, customer_file)
                    ncr = result.get('ncr_data', {})
                    email = result.get('email_draft', {})
                    
                    # ตัดสินใจเรื่องเลขเอกสารตามประเภทลูกค้า
                    if cust_folder == "SKC":
                        doc_no = database.generate_ncr_no()
                        st.success(f"🎯 Matched Master Data! | Auto NCR No: **{doc_no}**")
                    else:
                        doc_no = customer_doc_no if customer_doc_no else f"REF-{datetime.now().strftime('%Y%m%d-%H%M')}"
                        st.success(f"🎯 Matched Master Data! | Ref Doc No: **{doc_no}**")

                    ncr['doc_no'] = doc_no

                    # บันทึกไฟล์เคลมต้นฉบับลงโฟลเดอร์เฉพาะของลูกค้ารายนั้น
                    if customer_file:
                        if cust_folder == "SKC":
                            save_dir = f"{cust_folder}/NCR_Issued"
                        else:
                            save_dir = f"{cust_folder}/Customer_Claims"
                            
                        os.makedirs(save_dir, exist_ok=True)
                        cust_save_path = f"{save_dir}/{doc_no}_CustomerRaw_{customer_file.name}"
                        with open(cust_save_path, "wb") as f:
                            f.write(customer_file.getbuffer())

                    master_info = database.lookup_part_master(ncr.get('part_no', ''))
                    if master_info and master_info['found']:
                        ncr['part_name'] = master_info['part_name'].strip()
                        ncr['supplier'] = master_info['supplier'].strip()
                        ncr['supplier_email'] = master_info.get('supplier_email', '')
                    else:
                        st.warning("⚠️ ไม่พบ Part No ใน Master Data ระบบใช้ข้อมูลสกัดจาก AI")

                    database.save_new_claim(
                        doc_no, ncr.get('part_no'), ncr.get('part_name'), 
                        ncr.get('supplier'), ncr.get('problem'), customer=selected_customer
                    )
                    
                    st.session_state['active_ncr_no'] = doc_no
                    st.session_state['draft_to'] = ncr.get('supplier_email', '')
                    st.session_state['draft_sub'] = f"[{doc_no}] " + email.get('subject', '')
                    st.session_state['draft_body'] = email.get('body', '')
                    
                    st.success(f"✅ บันทึกข้อมูลเข้าโฟลเดอร์ [{cust_folder}] สำเร็จ! โปรดตรวจสอบร่างอีเมลด้านล่าง")
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")

    # แสดงร่างอีเมลสำหรับส่งแจ้ง Supplier
    if 'draft_sub' in st.session_state:
        st.divider()
        st.subheader("✉️ ร่างอีเมลเปิดเคสแจ้ง Supplier")
        to_email = st.text_input("To (Supplier Email):", value=st.session_state.get('draft_to', ''))
        sub_email = st.text_input("Subject:", value=st.session_state.get('draft_sub', ''))
        body_email = st.text_area("Body:", value=st.session_state.get('draft_body', ''), height=200)
        
        if st.button("📤 สั่งส่งอีเมลหา Supplier ทันที"):
            if to_email:
                success, msg = send_ncr_alert(to_email, sub_email, body_email)
                if success:
                    st.success("📧 ส่งอีเมลแจ้ง Supplier เรียบร้อยแล้ว!")
                else:
                    st.error(f"❌ ส่งอีเมลไม่สำเร็จ: {msg}")

# ==========================================
# TAB 2: PROCESS SUPPLIER REPLY & AUTO FILL
# ==========================================
with tab2:
    st.subheader("📋 2. อ่านรายงาน Supplier & กรอกแบบฟอร์มส่งลูกค้า")
    
    df = database.load_tracking_data()
    if not df.empty:
        active_ncrs = df[df['Status'] != 'Closed']['NCR_No'].tolist()
        selected_ncr = st.selectbox("เลือกเอกสารอ้างอิง (NCR / Claim Ref No.) ที่ต้องการอัปเดต:", active_ncrs if active_ncrs else ["ไม่มีรายการค้าง"])
        
        if selected_ncr != "ไม่มีรายการค้าง":
            current_row = df[df['NCR_No'] == selected_ncr].iloc[0]
            cust_type = current_row.get('Customer', '1. SKC (General PPTX Form)')
            cust_folder = get_customer_key(cust_type)
            
            st.info(f"📌 เคส: **{selected_ncr}** | ลูกค้า: **{cust_type}** (บันทึกลงโฟลเดอร์ `{cust_folder}/`) | Part: **{current_row.get('Part_No')}**")
            
            c1, c2 = st.columns(2)
            with c1:
                sup_reply_file = st.file_uploader("1. แนบเอกสารตอบกลับจาก Supplier (8D / Report):", type=['pdf', 'pptx', 'xlsx'], key="sup_reply")
            with c2:
                cust_template_file = st.file_uploader("2. แนบแบบฟอร์ม Excel/PPTX ของลูกค้า (เพื่อเติมข้อมูล):", type=['xlsx', 'xls', 'pptx'], key="cust_temp")
                
            if sup_reply_file and cust_template_file and st.button("🤖 ให้ AI สรุปรายงาน & หยอดใส่แบบฟอร์มลูกค้า"):
                with st.spinner(f"AI กำลังวิเคราะห์และหยอดข้อความลงในช่องของลูกค้า {cust_folder}..."):
                    try:
                        # 1. บันทึกไฟล์รายงาน Supplier ลงโฟลเดอร์ของลูกค้ารายนั้น
                        sup_dir = f"{cust_folder}/Supplier_Report"
                        os.makedirs(sup_dir, exist_ok=True)
                        sup_save_name = f"{selected_ncr}_Supplier_{sup_reply_file.name}"
                        with open(f"{sup_dir}/{sup_save_name}", "wb") as f:
                            f.write(sup_reply_file.getbuffer())
                        
                        filled_file_bytes = None
                        file_ext = ".xlsx"
                        
                        # 2. ประมวลผลและหยอดข้อมูลตามรายลูกค้า
                        if cust_folder == "VCS":
                            summary_data = ai_engine.summarize_for_customer1(sup_reply_file.getvalue(), sup_reply_file.name)
                            filled_file_bytes = ncr_exporter.fill_customer1_excel(cust_template_file.getvalue(), summary_data)
                            file_ext = ".xlsx"
                            st.subheader("💡 สรุปข้อความที่ AI หยอดลงช่องสีเขียว (VCS):")
                            st.json(summary_data)
                            review_dir = f"{cust_folder}/Completed_Forms_For_Review"
                            
                        elif cust_folder == "Yanmar":
                            summary_data = ai_engine.summarize_for_customer3(sup_reply_file.getvalue(), sup_reply_file.name)
                            filled_file_bytes = ncr_exporter.fill_customer3_excel(cust_template_file.getvalue(), summary_data)
                            file_ext = ".xlsx"
                            st.subheader("💡 สรุปข้อความที่ AI หยอดลง 3 ช่องเป้าหมาย (Yanmar):")
                            st.json(summary_data)
                            review_dir = f"{cust_folder}/Completed_Forms_For_Review"
                            
                        else: # SKC (General PPTX)
                            summary_data = ai_engine.summarize_supplier_reply_th(sup_reply_file.getvalue(), sup_reply_file.name)
                            filled_file_bytes = ncr_exporter.fill_ncr_pptx_template(cust_template_file.getvalue(), ncr_data={"doc_no": selected_ncr})
                            file_ext = ".pptx"
                            review_dir = f"{cust_folder}/WhyWhy_Approved"
                        
                        # 3. บันทึกเอกสารฉบับสมบูรณ์ลงโฟลเดอร์ทบทวนของลูกค้ารายนั้น
                        os.makedirs(review_dir, exist_ok=True)
                        completed_filename = f"{selected_ncr}_For_Customer_Review{file_ext}"
                        review_path = f"{review_dir}/{completed_filename}"
                        with open(review_path, "wb") as f:
                            f.write(filled_file_bytes.getvalue())
                            
                        st.success(f"✅ AI เติมข้อมูลและบันทึกไฟล์ไปที่ `{review_path}` เรียบร้อยแล้ว!")
                        
                        st.download_button(
                            label=f"📥 ดาวน์โหลดเอกสารสำหรับ QA ตรวจเช็กก่อนส่งลูกค้า ({completed_filename})",
                            data=filled_file_bytes.getvalue(),
                            file_name=completed_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_ext == ".xlsx" else "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                        
                        # อัปเดตสถานะในฐานข้อมูล
                        database.update_claim(selected_ncr, "Supplier Replied", "AI Analyzed", "AI Analyzed", sup_save_name, completed_filename)
                        
                    except Exception as e:
                        st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")

        st.divider()
        st.subheader("📊 ตารางติดตามสถานะทั้งหมด (Tracking Dashboard)")
        st.dataframe(df, use_container_width=True)

# ==========================================
# TAB 3: MANAGEMENT REPORTS
# ==========================================
with tab3:
    st.subheader("📊 สรุปประวัติการเคลมภาพรวม")
    df_all = database.load_tracking_data()
    if not df_all.empty:
        col_a, col_b = st.columns(2)
        with col_a:
            st.write("**จำนวน Claim แยกตาม Supplier**")
            st.bar_chart(df_all['Supplier'].value_counts())
        with col_b:
            st.write("**จำนวน Claim แยกตาม Customer**")
            if 'Customer' in df_all.columns:
                st.bar_chart(df_all['Customer'].value_counts())