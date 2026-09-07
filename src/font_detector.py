import re
from collections import Counter
from pathlib import Path
from typing import Optional
import pymupdf as fitz

from .config import DEFAULT_FONT

# พจนานุกรมจับคู่ชื่อฟอนต์ใน PDF ไปเป็นชื่อฟอนต์ทางการใน Word
FONT_MAPPINGS = [
    (r"sarabun", "TH Sarabun New"),
    (r"cordia", "Cordia New"),
    (r"angsana", "Angsana New"),
    (r"browallia", "Browallia New"),
    (r"tahoma", "Tahoma"),
    (r"calibri", "Calibri"),
    (r"arial", "Arial"),
    (r"times(newroman)?", "Times New Roman"),
    (r"segoe", "Segoe UI"),
    (r"garuda", "Garuda"),
    (r"kinnari", "Kinnari"),
    (r"sukhumvit", "Sukhumvit Set"),
    (r"helvetica", "Arial"),
    (r"cambria", "Cambria"),
    (r"georgia", "Georgia"),
    (r"verdana", "Verdana"),
]


def clean_font_name(raw_name: str) -> str:
    """
    ทำความสะอาดชื่อฟอนต์ที่ดึงมาจาก PDF เช่น:
    - 'ABCDEF+THSarabunPSK-Bold' -> 'TH Sarabun New'
    - 'CordiaNew-Regular' -> 'Cordia New'
    - 'AngsanaUPC,Bold' -> 'Angsana New'
    """
    if not raw_name:
        return DEFAULT_FONT

    # ตัด subset prefix (เช่น 'ABCDEF+')
    name = re.sub(r"^[A-Z]{6}\+", "", raw_name)

    # ตัดสไตล์ส่วนท้าย (Bold, Italic, Regular, PS, MT)
    name_clean = re.sub(
        r"[-_, ]*(Bold|Italic|Regular|BoldItalic|Bd|It|MT|PS|Pro|Uni|PSK|UPC)\b",
        "",
        name,
        flags=re.IGNORECASE,
    ).strip()

    # ตรวจสอบกับรายการ Mapping ฟอนต์มาตรฐาน
    name_lower = name.lower()
    for pattern, official_name in FONT_MAPPINGS:
        if re.search(pattern, name_lower):
            return official_name

    # หากไม่ตรงกับรายการมาตรฐาน ให้คืนค่าชื่อที่ตัดแต่งแล้ว หรือค่าเริ่มต้น
    return name_clean if len(name_clean) >= 3 else DEFAULT_FONT


def detect_dominant_font(pdf_path: str | Path, max_pages_to_check: int = 5) -> Optional[str]:
    """
    ตรวจจับฟอนต์หลักที่ใช้ในเอกสาร PDF โดยวิเคราะห์จากความถี่ของตัวอักษรในแต่ละฟอนต์
    
    pdf_path: ที่อยู่ไฟล์ PDF
    max_pages_to_check: จำนวนหน้าที่ใช้วิเคราะห์ (ค่าเริ่มต้น 5 หน้าแรก)
    ส่งคืน: ชื่อฟอนต์ที่ตรวจพบ (เช่น 'TH Sarabun New', 'Cordia New') หรือ None หากเป็นภาพสแกน
    """
    path = Path(pdf_path)
    if not path.exists():
        return None

    font_char_counts = Counter()

    try:
        with fitz.open(path) as doc:
            num_pages = min(len(doc), max_pages_to_check)
            for page_idx in range(num_pages):
                page = doc[page_idx]
                page_dict = page.get_text("dict")

                for block in page_dict.get("blocks", []):
                    # เฉพาะ text block (type == 0)
                    if block.get("type") == 0:
                        for line in block.get("lines", []):
                            for span in line.get("spans", []):
                                font_name = span.get("font", "")
                                text = span.get("text", "")
                                char_count = len(text.strip())
                                if font_name and char_count > 0:
                                    font_char_counts[font_name] += char_count

            # หากไม่พบ text span เลย (เช่น เป็นไฟล์ภาพสแกน) ให้ลองตรวจจับจาก get_fonts
            if not font_char_counts:
                for page_idx in range(num_pages):
                    for font_info in doc[page_idx].get_fonts():
                        # font_info = (xref, ext, type, basefont, name, encoding)
                        basefont = font_info[3] if len(font_info) > 3 else ""
                        if basefont:
                            font_char_counts[basefont] += 1

    except Exception:
        return None

    if not font_char_counts:
        # เป็น PDF ภาพสแกน ไม่มี metadata ฟอนต์ดิจิทัล
        return None

    # ดึงฟอนต์ที่มีจำนวนตัวอักษรมากที่สุด (Dominant Font)
    most_common_raw_font, _ = font_char_counts.most_common(1)[0]
    detected_font = clean_font_name(most_common_raw_font)
    return detected_font
