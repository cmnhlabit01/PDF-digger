import os
from pathlib import Path
from dotenv import load_dotenv

# โหลดตัวแปรจาก .env ถ้ามี
load_dotenv()

# รายชื่อโมเดลสำหรับการทำ Seamless Fallback (หากตัวแรกติด 429 Quota Exceeded จะสลับไปตัวถัดไปทันที)
FALLBACK_MODELS = [
    "gemini-3.6-flash",
    "gemini-3.7-flash",
    "gemini-flash-latest",
    "gemini-3-flash-preview",
]

# ฟอนต์มาตรฐานสำหรับไฟล์ Word
DEFAULT_FONT = os.getenv("DEFAULT_FONT", "Cordia New")
FONT_SIZE_BODY = 16
FONT_SIZE_H3 = 17
FONT_SIZE_H2 = 18
FONT_SIZE_H1 = 22

# การเรนเดอร์ภาพหน้า PDF สำหรับส่งให้ Gemini อ่าน (200 DPI คมชัดระดับสิ่งพิมพ์และประมวลผลเร็วกว่า 300 DPI ถึง 2.25 เท่า)
RENDER_DPI = 200

# จำนวนหน้าที่จะประมวลผลพร้อมกันในเวลาเดียวกัน (Parallel Workers) ช่วยลดเวลาทำงานลง 3-4 เท่า
MAX_PARALLEL_WORKERS = 3

# ขนาดรูปภาพสูงสุดสำหรับแทรกลงในไฟล์ Word (นิ้ว)
MAX_DOCX_IMAGE_WIDTH_INCHES = 6.0

def get_api_key(override_key: str | None = None) -> str:
    """ดึง Gemini API Key จาก parameter หรือ environment variable"""
    key = override_key or os.getenv("GEMINI_API_KEY")
    if not key or key.strip() == "" or key.strip() == "your_gemini_api_key_here":
        raise ValueError(
            "ไม่พบ GEMINI_API_KEY! กรุณาระบุในไฟล์ .env หรือส่งผ่านพารามิเตอร์ "
            "(สามารถขอรับ API Key ฟรีได้จาก https://aistudio.google.com/)"
        )
    return key.strip()
