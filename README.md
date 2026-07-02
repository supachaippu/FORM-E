# FORM E - Document Comparison Tool

ระบบเปรียบเทียบข้อมูลรายการสินค้าระหว่างเอกสาร **Excel (Invoice/Packing List)** และ **PDF (FORM E)** โดยแยกการตรวจจับเป็น 4 หัวข้อสำคัญ ได้แก่ **English Name, Part Number, Qty, และ HS Code** พร้อมไฮไลต์สีแดง-เขียวแสดงความผิดพลาดและจับคู่ได้อย่างชัดเจน

## 🚀 ช่องทางการเข้าใช้งานออนไลน์ (Streamlit Cloud)
ท่านสามารถติดตั้งและรันหน้าเว็บแอปพลิเคชันออนไลน์นี้ได้ง่ายๆ 24 ชั่วโมงฟรี โดยการกดปุ่มด้านล่างนี้เพื่อเชื่อมโยงและรันบน Streamlit Community Cloud:

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/deploy?repository=supachaippu/FORM-E&branch=main&mainModule=compare_streamlit.py)

## 💻 วิธีการติดตั้งและรันใช้งานเองในเครื่อง (Local Setup)

### 1. การติดตั้งไลบรารีที่จำเป็น (Dependencies)
```bash
pip install -r requirements.txt
```

### 2. การรันใช้งานโปรแกรมหน้าเว็บ (Streamlit)
```bash
streamlit run compare_streamlit.py
```

### 3. การรันใช้งานโปรแกรมหน้าจอคอมพิวเตอร์ (Desktop GUI)
* **สำหรับ Windows:** ดับเบิ้ลคลิกไฟล์ `run_compare.bat` ได้ทันที (ระบบจะตรวจสอบการติดตั้ง Python ให้โดยอัตโนมัติ)
* **สำหรับ macOS / Linux:** รันผ่าน Terminal:
  ```bash
  python3 compare_gui.py
  ```

### 4. การรันสคริปต์คอมมานด์ไลน์แบบธรรมดา (CLI)
```bash
python3 compare_form_e.py
```

## 📂 โครงสร้างโฟลเดอร์โครงการ
* `compare_streamlit.py` - เว็บแอปพลิเคชันผ่านระบบ Streamlit
* `compare_gui.py` - โปรแกรมหน้าจอ Desktop GUI
* `compare_form_e.py` - สคริปต์เปรียบเทียบแบบคอมมานด์ไลน์หลัก
* `run_compare.bat` - ปุ่มลัดด่วนสำหรับดับเบิ้ลคลิกเปิดโปรแกรมบน Windows
* `requirements.txt` - ไฟล์ระบุไลบรารีที่ต้องใช้ในการติดตั้ง
* `FE-T1.pdf` - ไฟล์ตัวอย่างเอกสาร PDF (FORM E)
* `Invoice and Packing list 326T1.xlsx` - ไฟล์ตัวอย่างตาราง Excel (Invoice)
* `form_e_comparison_report.xlsx` - ตัวอย่างรายงานผลลัพธ์การตรวจสอบแบบ Excel ที่จัดสีกรีน-เรดเรียบร้อย
