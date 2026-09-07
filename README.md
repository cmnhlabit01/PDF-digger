# 📄 PDF Digger

ระบบแกะข้อความภาษาไทยและอังกฤษจากเอกสาร PDF (ทั้งแบบดิจิทัลที่แปลงมาจาก Word และภาพสแกน/รูปถ่าย) พร้อมถอดแบบตารางและรูปภาพประกอบ แปลงออกมาเป็นไฟล์ Microsoft Word (`.docx`) คุณภาพสูง โดยใช้ **Google Gemini AI**

---

## ✨ จุดเด่นและความสามารถ (Features)

1. **ถอดข้อความภาษาไทยและอังกฤษได้อย่างแม่นยำสูง (Verbatim Bilingual OCR):**
   - รองรับทั้งเอกสารดิจิทัลและเอกสารภาพสแกน/เอกสารเก่า
   - รักษาความถูกต้องของสระและวรรณยุกต์ภาษาไทยครบ 100% (สระลอย, วรรณยุกต์ซ้อน, สระอำ, การันต์)
   - เว้นวรรคคำภาษาไทยอย่างถูกต้องเป็นธรรมชาติ ไม่ตัดคำมั่ว

2. **ระบบสลับโมเดลอัตโนมัติอย่างไร้รอยต่อ (Seamless Model Fallback):**
   - รองรับสถาปัตยกรรมโมเดล Gemini ล่าสุด: `gemini-3.6-flash` ➔ `gemini-3.7-flash` ➔ `gemini-flash-latest` ➔ `gemini-3-flash-preview`
   - หากโมเดลหลักโควตาเต็มหรือติดขัด (HTTP 429 Quota Exceeded) ระบบจะสลับไปใช้โมเดลสำรองถัดไปทันที งานไม่สะดุด

3. **ถอดแบบตารางเหมือนต้นฉบับ (High-Fidelity Tables):**
   - แปลงตารางใน PDF ออกมาเป็น **Native Word Table**
   - ตีเส้นตาราง (Table Grid), จัดสีพื้นหลังหัวตาราง, รักษาระยะคอลัมน์ และจัดแนวตัวเลขชิดขวาอัตโนมัติ

4. **ตรวจจับแบบอักษรต้นฉบับอัตโนมัติ (Auto-Detect Font):**
   - ตรวจจับฟอนต์จริงจาก PDF (เช่น `TH Sarabun New`, `Cordia New`, `Angsana New`, `Tahoma`, `Calibri`, `Arial`)
   - นำฟอนต์ที่ตรวจพบไปจัดรูปแบบใน Word ให้อัตโนมัติ เพื่อให้หน้าตาเหมือนต้นฉบับมากที่สุด
   - มีฟอนต์ `Cordia New` เป็นค่าเริ่มต้นที่ปลอดภัย เปิดได้ถูกต้องบนทุกเครื่อง

5. **รองรับการแสดงผลภาษาไทยบน Word ทุกแพลตฟอร์ม:**
   - ฝังแท็ก OpenXML Complex Script (`w:cs`, `w:ascii`) เปิดบน Windows, Mac, iOS, Android ฟอนต์ไม่เพี้ยน
   - ปรับระยะบรรทัด 1.15x ป้องกันสระบน-ล่างและวรรณยุกต์ถูกตัดขอบ

6. **ระบบประมวลผลความเร็วสูง (High-Speed Processing):**
   - ประมวลผลแบบคู่ขนานหลายหน้าพร้อมกัน (Multi-threading Parallel Processing)
   - เรนเดอร์ 200 DPI High-Quality JPEG พร้อมระบบกรองภาพสแกนพื้นหลัง

7. **รูปแบบการใช้งานที่หลากหลาย:**
   - 🖥️ **Desktop App:** หน้าต่างโปรแกรมสวยงาม (CustomTkinter) เข้ากับธีมระบบ
   - 🌐 **Web UI:** หน้าเว็บอินเทอร์แอคทีฟด้วย Streamlit
   - 💻 **CLI:** คำสั่งผ่าน Terminal สำหรับการประมวลผลไฟล์เดี่ยวหรือทั้งโฟลเดอร์ (Batch)

---

## 🚀 การติดตั้ง (Installation)

### 1. Clone Repository
```bash
git clone https://github.com/<username>/pdf-digger.git
cd pdf-digger
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
- บน macOS สามารถดับเบิ้ลคลิกไฟล์ `เปิดโปรแกรม_PDF_Digger.command` ได้ทันที
- เลือกไฟล์ PDF, ตรวจสอบ API Key แล้วกด **"🚀 เริ่มแปลงเอกสารเป็น Word (.docx)"**

### วิธีที่ 2: ใช้งานผ่านหน้าเว็บ (Streamlit Web UI)
```bash
streamlit run app.py
```
เปิดเบราว์เซอร์แล้วลากไฟล์ PDF มาวาง พร้อมกดเริ่มแปลงและดาวน์โหลดไฟล์ Word ได้ทันที

### วิธีที่ 3: ใช้งานผ่าน Command Line (CLI)

```bash
# แปลงไฟล์เดี่ยว (ตรวจจับฟอนต์ต้นฉบับอัตโนมัติ)
python main.py sample.pdf

# กำหนดชื่อไฟล์ปลายทางและเลือกฟอนต์เอง
python main.py sample.pdf -o output.docx --font "TH Sarabun New"

# แปลงทุกไฟล์ PDF ในโฟลเดอร์พร้อมกัน (Batch Processing)
python main.py /path/to/pdf_folder/ --batch
```

---

## 🧪 การรันชุดทดสอบ (Running Tests)

ทดสอบการทำงานของโมดูลทั้งหมด (DocxBuilder, Font Detector, Model Fallback, PDF Processor):
```bash
python -m unittest tests/test_pdf_digger.py
```

---

## 📂 โครงสร้างโปรเจกต์ (Project Structure)

```
pdf-digger/
├── .env.example              # ตัวอย่างไฟล์คอนฟิก
├── .gitignore                # ป้องกันการอัปโหลด API Key และไฟล์ชั่วคราว
├── requirements.txt          # รายการไลบรารี
├── README.md                 # คู่มือการใช้งาน
├── gui.py                    # โปรแกรม Desktop App (CustomTkinter)
├── app.py                    # แอปพลิเคชันหน้าเว็บ (Streamlit)
├── main.py                   # อินเทอร์เฟซคำสั่ง CLI
├── tests/
│   └── test_pdf_digger.py    # ชุดทดสอบระบบอัตโนมัติ
└── src/
    ├── __init__.py
    ├── config.py             # จัดการคอนฟิก และ Fallback Models
    ├── font_detector.py      # ตรวจจับและแมปชื่อฟอนต์จาก PDF
    ├── pdf_processor.py      # เรนเดอร์หน้า PDF และสกัดรูปภาพ
    ├── gemini_extractor.py   # OCR & โครงสร้างเอกสารด้วย Gemini AI
    ├── docx_builder.py       # สร้างไฟล์ Word (OpenXML, ตาราง, สไตล์)
    └── pipeline.py           # ควบคุมขั้นตอน Pipeline ทั้งหมด
```

---

## 📄 License
MIT License
