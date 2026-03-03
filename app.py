import streamlit as st
import pandas as pd
from excel_mapper import process_mapping_execution, create_zip_archive, generate_mock_mapping_file

st.set_page_config(page_title="Excel Report Automator", layout="wide")

st.title("📊 Excel Report Automator")
st.markdown("ดึงค่าจากหลายไฟล์มารวมใส่ Template ที่คุณออกแบบไว้ง่ายๆ ในคลิกเดียว")

# Provide an example Mapping Template
st.sidebar.header("เครื่องมือช่วยเหลือ")
st.sidebar.download_button(
    "📥 ดาวน์โหลดไฟล์ตัวอย่าง Mapping Config",
    data=generate_mock_mapping_file(),
    file_name="Mapping_Template_Example.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. ไฟล์ตั้งค่า Mapping")
    mapping_file = st.file_uploader("อัปโหลดไฟล์ Mapping (Excel/CSV)", type=["xlsx", "csv"])

with col2:
    st.subheader("2. ไฟล์ข้อมูลต้นทาง (Source)")
    source_files = st.file_uploader("อัปโหลดไฟล์ Excel ต้นทางทั้งหมด", type=["xlsx"], accept_multiple_files=True)

with col3:
    st.subheader("3. ไฟล์ปลายทาง (Templates)")
    template_files = st.file_uploader("อัปโหลดไฟล์ Template ทั้งหมด", type=["xlsx"], accept_multiple_files=True)

st.divider()

if st.button("🚀 Generate Reports (เริ่มดึงข้อมูลใส่ Template)", type="primary"):
    if not mapping_file or not source_files or not template_files:
        st.error("กรุณาอัปโหลดไฟล์ให้ครบทั้ง 3 ช่อง ก่อนกดเริ่มประมวลผลครับ")
    else:
        with st.spinner("กำลังประมวลผลการแมปข้อมูล... กรุณารอสักครู่"):
            try:
                # Load mapping
                if mapping_file.name.endswith(".csv"):
                    mapping_df = pd.read_csv(mapping_file)
                else:
                    mapping_df = pd.read_excel(mapping_file)

                # Process Source files into dict format required by backend
                source_dict = {f.name: f.read() for f in source_files}
                # Process Template files
                template_dict = {f.name: f.read() for f in template_files}

                # Core execution
                generated_reports = process_mapping_execution(mapping_df, source_dict, template_dict)
                
                if not generated_reports:
                    st.warning("ประมวลผลสำเร็จ แต่ไม่มีข้อมูลถูกสร้าง กรุณาตรวจสอบว่าชื่อไฟล์ใน Mapping ถูกต้องตรงตามไฟล์ที่อัปโหลด")
                else:
                    st.success(f"✅ สร้าง Report สำเร็จแล้วจำนวน {len(generated_reports)} ไฟล์!")
                    
                    if len(generated_reports) == 1:
                        # Single file download
                        filename, file_bytes = list(generated_reports.items())[0]
                        st.download_button(
                            label=f"💾 ดาวน์โหลด {filename}",
                            data=file_bytes,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        # Zip download for multiple reports
                        zip_bytes = create_zip_archive(generated_reports)
                        st.download_button(
                            label="🗂️ ดาวน์โหลดผลลัพธ์ทั้งหมดเป็น ZIP",
                            data=zip_bytes,
                            file_name="Generated_Reports.zip",
                            mime="application/zip"
                        )

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดระหว่างประมวลผล: {e}")
                st.info("โปรดตรวจสอบว่าชื่อไฟล์ Sheet และ Cell ในไฟล์ Mapping ตรงกับความเป็นจริงทุกประการ")
