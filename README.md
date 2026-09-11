# 📄 PDF Digger

ระบบแกะข้อความภาษาไทยและอังกฤษจากเอกสาร PDF และไฟล์รูปภาพ (ทั้งแบบดิจิทัลที่แปลงมาจาก Word, ภาพสแกน และรูปถ่ายจากมือถือ) พร้อมระบบ**แปลภาษาในรอบเดียว (One-Pass Translation)**, ถอดแบบตาราง, สกัดรูปภาพประกอบ, และปรับแต่งความคมชัดภาพ แปลงออกมาเป็นไฟล์ Microsoft Word (`.docx`) คุณภาพสูง โดยใช้ **Google Gemini AI**

---

## ✨ จุดเด่นและความสามารถ (Features)

1. **🌐 ระบบแปลภาษาในรอบเดียว (One-Pass Translation):**
   - รองรับการแปล: **ไทย (Thai)**, **อังกฤษ (English)**, **จีน (Chinese)**, **ญี่ปุ่น (Japanese)** หรือคงภาษาตามต้นฉบับ
   - ทำงานในรอบเดียว ไม่เสียเวลาเรียก API ซ้ำ ไม่เปลืองโควตา
   - โครงสร้างตาราง หัวข้อ รายการ ลำดับตัวเลข และตำแหน่งรูปภาพจะยังคงเดิม 100%

2. **🖼️ รองรับไฟล์รูปภาพโดยตรง (Direct Image Input):**
   - ลากไฟล์ภาพเดี่ยวเข้ามาแปลงเป็น Word ได้ทันที: รองรับ `.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.tiff`
   - เหมาะอย่างยิ่งสำหรับภาพถ่ายเอกสารจากสมาร์ตโฟนหรือใบเสร็จ

3. **✨ โหมดเพิ่มความคมชัดพิเศษสำหรับภาพสแกน (Image Enhancement Mode):**
   - **Auto-Contrast & White Balance:** ปรับจูนแสงกระดาษที่เหลือง/หมองให้ขาวขึ้น และขับตัวหนังสือจางให้เข้มขึ้น
   - **Edge Sharpening:** เร่งความคมชัดของขอบตัวอักษรและสระภาษาไทย ป้องกันสระและวรรณยุกต์กลืนหายไป

4. **⚠️ ความถูกต้องของข้อมูลและการจัดการจุดที่อ่านไม่ออก (Strict Anti-Hallucination):**
   - AI จะไม่เดาหรือแต่งเติมคำ/ตัวเลขขึ้นมาเองเด็ดขาด
   - หากจุดใดเบลอมากจนไม่มั่นใจ จะติดแท็กแจ้งเตือนชัดเจน เช่น `[ข้อความไม่ชัดเจน]` หรือ `[ตัวเลขไม่ชัดเจน]` เพื่อให้ผู้ใช้ตรวจสอบได้ทันที
   - หากภาพเสียหายทั้งหน้า จะมีข้อความแจ้งเตือนพร้อมแนบภาพต้นฉบับลงใน Word ให้ดูเปรียบเทียบ

5. **🔄 ระบบสลับโมเดลอัตโนมัติอย่างไร้รอยต่อ (Seamless Model Fallback):**
   - รองรับโมเดล Gemini ล่าสุด: `gemini-3.6-flash` ➔ `gemini-3.7-flash` ➔ `gemini-flash-latest` ➔ `gemini-3-flash-preview`
   - หากโมเดลหลักโควตาเต็ม (HTTP 429 Quota Exceeded) ระบบจะสลับไปใช้โมเดลสำรองถัดไปทันที งานไม่สะดุด

6. **📊 ถอดแบบตารางเหมือนต้นฉบับ (High-Fidelity Tables):**
   - แปลงตารางในเอกสารออกมาเป็น **Native Word Table**
   - ตีเส้นตาราง (Table Grid), จัดสีพื้นหลังหัวตาราง, รักษาระยะคอลัมน์ และจัดแนวตัวเลขชิดขวาอัตโนมัติ

7. **🔍 ตรวจจับแบบอักษรต้นฉบับอัตโนมัติ (Auto-Detect Font):**
   - ตรวจจับฟอนต์จริงจาก PDF (เช่น `TH Sarabun New`, `Cordia New`, `Angsana New`, `Tahoma`, `Calibri`, `Arial`)
   - มีฟอนต์ `Cordia New` (OpenXML Complex Script `w:cs`) เป็นค่าเริ่มต้นที่ปลอดภัย เปิดได้ถูกต้องบนทุกเครื่อง

8. **รูปแบบการใช้งานที่หลากหลาย:**
   - 🖥️ **Desktop App:** หน้าต่างโปรแกรมสวยงาม (CustomTkinter) สำหรับ macOS / Windows
   - 🌐 **Web UI:** หน้าเว็บ Streamlit สำหรับรันบน Cloud ออนไลน์
   - 💻 **CLI:** คำสั่งผ่าน Terminal สำหรับการประมวลผลไฟล์เดี่ยวหรือทั้งโฟลเดอร์ (Batch)

---

## 🚀 การติดตั้ง (Installation)

### 1. Clone Repository
```bash
git clone https://github.com/cmnhlabit01/PDF-digger.git
cd PDF-digger
```

### 2. สร้างและเปิดใช้งาน Virtual Environment
**บน macOS / Linux:**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

**บน Windows:**
```powershell
python -m venv .venv
.venv\Scripts\activate
```

### 3. ติดตั้งไลบรารีที่จำเป็น
```bash
pip install -r requirements.txt
```

### 4. ตั้งค่า Gemini API Key
คัดลอกไฟล์ตัวอย่างคอนฟิก:
```bash
cp .env.example .env
```
เปิดไฟล์ `.env` แล้วใส่ API Key ที่ได้รับจาก [Google AI Studio](https://aistudio.google.com/):
```ini
GEMINI_API_KEY=AIzaSy...
```

---

## 💻 วิธีการใช้งาน (Usage)

### วิธีที่ 1: ใช้งานผ่าน Desktop Application (GUI)
```bash
python gui.py
```
- บน macOS สามารถดับเบิ้ลคลิกไฟล์ `เปิดโปรแกรม_PDF_Digger.command` หรือแอป `PDF Digger.app`
- เลือกไฟล์ (PDF หรือ รูปภาพ), เลือกภาษาเป้าหมายที่ต้องการแปล, ติ๊กเลือกโหมดปรับความคมชัดภาพ แล้วกด **"🚀 เริ่มแปลงเอกสารเป็น Word (.docx)"**

### วิธีที่ 2: ใช้งานผ่านหน้าเว็บ (Streamlit Web UI - Online)
```bash
streamlit run app.py
```
- สามารถนำไป Deploy ขึ้น **Streamlit Community Cloud** เพื่อใช้งานแบบ Online ผ่านเบราว์เซอร์ได้ฟรีตลอดชีพ

### วิธีที่ 3: ใช้งานผ่าน Command Line (CLI)

```bash
# แปลงไฟล์เดี่ยว (ภาษาตามต้นฉบับ)
python main.py sample.pdf

# แปลงเอกสารเป็นภาษาไทย พร้อมเปิดโหมดเพิ่มความคมชัดภาพ
python main.py english_paper.pdf -t th --enhance

# แปลงไฟล์รูปภาพเดี่ยว (.png / .jpg) เป็น Word
python main.py receipt.jpg -o receipt.docx

# แปลงเอกสารเป็นภาษาอังกฤษ และเลือกฟอนต์
python main.py thai_doc.pdf -t en --font "Calibri"

# แปลงทุกไฟล์ PDF และรูปภาพในโฟลเดอร์พร้อมกัน (Batch Processing)
python main.py /path/to/folder/ --batch -t th
```

---

## 🧪 การรันชุดทดสอบ (Running Tests)

ทดสอบการทำงานของทุกโมดูล (DocxBuilder, Font Detector, Fallback, Direct Image, Enhancement, Translation):
```bash
python -m unittest tests/test_pdf_digger.py
```

---

## 📂 โครงสร้างโปรเจกต์ (Project Structure)

```
PDF-digger/
├── .env.example              # ตัวอย่างไฟล์คอนฟิก
├── .gitignore                # ป้องกันการอัปโหลด API Key และไฟล์ชั่วคราว
├── requirements.txt          # รายการไลบรารี
├── README.md                 # คู่มือการใช้งาน
├── gui.py                    # โปรแกรม Desktop App (CustomTkinter)
├── app.py                    # แอปพลิเคชันหน้าเว็บ (Streamlit)
├── main.py                   # อินเทอร์เฟซคำสั่ง CLI
├── tests/
│   └── test_pdf_digger.py    # ชุดทดสอบระบบอัตโนมัติ (6 unit tests)
└── src/
    ├── __init__.py
    ├── config.py             # จัดการคอนฟิก และ Fallback Models
    ├── font_detector.py      # ตรวจจับและแมปชื่อฟอนต์จาก PDF
    ├── pdf_processor.py      # เรนเดอร์ PDF/รูปภาพ และปรับแต่งภาพ (Image Enhancement)
    ├── gemini_extractor.py   # Multilingual OCR & Translation ด้วย Gemini AI
    ├── docx_builder.py       # สร้างไฟล์ Word (OpenXML, ตาราง, สไตล์)
    └── pipeline.py           # ควบคุมขั้นตอน Pipeline ทั้งหมด
```

---

## 📄 License
MIT License
