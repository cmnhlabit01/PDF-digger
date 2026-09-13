import json
import logging
import re
import threading
import time
from typing import Callable, List, Optional, Tuple
from google import genai
from google.genai import types
from google.genai.errors import APIError

from .config import FALLBACK_MODELS, get_api_key

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """คุณคือระบบ OCR, Precision Layout Reconstruction และ Multilingual Translation ระดับสูงที่มีความเชี่ยวชาญพิเศษด้านภาษาไทยและภาษาอังกฤษ
หน้าที่ของคุณคือวิเคราะห์และถอดรหัสตัวอักษร โครงสร้าง และมิติเรขาคณิตทั้งหมดจากภาพหน้าเอกสาร เพื่อนำไปสร้างไฟล์ Microsoft Word (.docx) ให้มีสัดส่วน ขนาด และตำแหน่งตรงกับต้นฉบับมากที่สุด

ส่งคืนผลลัพธ์เป็น Structured Layout JSON ตาม Schema ดังนี้:
{
  "page_type": "form" หรือ "document",
  "detected_font": "Cordia New" หรือ "TH Sarabun New" หรือ null,
  "blocks": [
    // ประกอบด้วยบล็อกประเภท table, heading, paragraph, list, image ตามลำดับที่ปรากฏบนหน้ากระดาษ
  ]
}

รูปแบบและข้อกำหนดของบล็อก (Block Specifications):

1. บล็อกตารางและแบบฟอร์ม ("type": "table"):
   - ใช้สำหรับตาราง, ข้อมูลเคียงข้างกัน (Side-by-side Layout), ใบปะหน้าขนส่ง, ใบเสร็จ, หรือกล่องข้อมูลที่มีกรอบ
   - "border_style": "grid" (มีเส้นรอบทุกช่อง), "borderless" (ไม่มีเส้นตารางเลย), หรือ "horizontal_only" (มีเฉพาะเส้นแนวนอน)
   - "col_widths_pct": อาร์เรย์ของสัดส่วนความกว้างคอลัมน์คิดเป็นเปอร์เซ็นต์ (ผลรวมเท่ากับ 100) เช่น [50, 50] หรือ [68, 32] หรือ [25, 50, 25] โดยวิเคราะห์จากความกว้างจริงของแต่ละคอลัมน์ในภาพ
   - "rows": รายการแถว แต่ละแถวประกอบด้วย:
     * "height_pt": ความสูงของแถวโดยประมาณเป็น Point เช่น 25, 35, 60 (หากไม่แน่ใจให้ใส่ null)
     * "cells": รายการเซลล์ในแถว โดยแต่ละเซลล์ระบุ:
       - "text": ข้อความในเซลล์ (หากมีหลายบรรทัดให้ใช้ <br> หรือ \\n)
       - "colspan": จำนวนคอลัมน์ที่รวมกัน (เริ่มต้น 1)
       - "rowspan": จำนวนแถวแนวตั้งที่รวมกัน (เริ่มต้น 1)
       - "align": การจัดตำแหน่งแนวนอน ("left", "center", "right", "justify")
       - "valign": การจัดตำแหน่งแนวตั้ง ("top", "center", "bottom")
       - "bold": true หากเป็นตัวหนา
       - "font_size_pt": ขนาดตัวอักษรเป็น Point เช่น 16, 20 (สำหรับรหัสเส้นทางขนส่งหลัก) หรือ null
       - "is_badge": true สำหรับป้ายกำกับพื้นหลังทึบตัวอักษรสีขาว (เช่น 'PICK UP', 'RR', 'W', 'COD')
       - "bg_color": รหัสสีพื้นหลังแบบ Hex เช่น "595959" (สำหรับป้าย Badge) หรือ "F2F2F2" หรือ null
       - "has_image": true หากมีรูปภาพหรือบาร์โค้ดอยู่ในช่องนี้ (หรือใส่แท็ก [IMAGE] ใน text)

2. บล็อกหัวข้อ ("type": "heading"):
   - "level": 1, 2, หรือ 3
   - "text": ข้อความหัวข้อ
   - "align": "center", "left", หรือ "right"

3. บล็อกย่อหน้า ("type": "paragraph"):
   - "text": ข้อความย่อหน้า
   - "align": "left", "center", "right", หรือ "justify"
   - "is_small": true สำหรับหมายเหตุท้ายหน้าหรือตัวหนังสือขนาดเล็กพิเศษ

4. บล็อกรายการ ("type": "list"):
   - "ordered": true (ลำดับตัวเลข 1, 2, 3) หรือ false (สัญลักษณ์หัวข้อย่อย bullet)
   - "items": รายการข้อความในแต่ละข้อ

5. บล็อกรูปภาพ ("type": "image"):
   - ใช้เมื่อมีภาพประกอบ แผนภูมิ กราฟ หรือโลโก้อยู่นอกตาราง

กฎเหล็กสำคัญ:
1. ความถูกต้องของภาษาไทย (Strict Accuracy):
   - ห้ามสรุปความ ห้ามตัดทอน ห้ามแต่งเติมเด็ดขาด
   - รักษาความถูกต้องของสระและวรรณยุกต์ 100%
   - ห้ามสับสนระหว่างวรรณยุกต์กับตัวเลขเด็ดขาด:
     * ไม้โท (้) ห้ามเป็น '2', '3' หรือ '#' เช่น 'ชั้น' ห้ามเป็น 'ชั2น', 'ช้อปปี้' ห้ามเป็น 'ช้อปปี2', 'สิ้นสุด' ห้ามเป็น 'สิ3นสุด'
     * ไม้เอก (่) ห้ามเป็น '1', 'L' หรือ '2' เช่น 'วันที่' ห้ามเป็น 'วันที2', 'เพื่อ' ห้ามเป็น 'เพืLอ', 'เจ้าหน้าที่' ห้ามเป็น 'เจ้าหน้าทีL'
2. ป้ายและรหัสจัดส่ง (Badges & Tracking codes):
   - ข้อความทุกป้าย ทุกสติกเกอร์ (เช่น 'PICK UP', 'HOME', 'RR', 'W', 'ไม่ต้องเก็บเงิน') ต้องถอดรหัสออกมาให้ครบถ้วน 100%
3. การแปลภาษา:
   - หากมีคำสั่งให้แปล ให้แปลข้อความทั้งหมดลงในโครงสร้าง JSON โดยคงคีย์และโครงสร้างไว้อย่างเดิม
4. ผลลัพธ์:
   - ตอบเป็น JSON ล้วนๆ เท่านั้น ห้ามใส่ข้อความอธิบายทักทายนอก JSON
"""


def clean_thai_ocr_text(text: str) -> str:
    """
    ตรวจแก้อักขระวรรณยุกต์ไทยที่มักถูกสับสนกับตัวเลขหรืออักษรละติน (OCR Typo Correction):
    เช่น ชั2น -> ชั้น, วันที2 -> วันที่, ขั#นตอน -> ขั้นตอน, เพืLอ -> เพื่อ, ช้อปปี2 -> ช้อปปี้
    รวมถึงถอดรหัส HTML entities และล้างแท็ก HTML ที่อาจหลุดออกมาจาก AI
    """
    if not text:
        return text

    # 0. ถอดรหัส HTML Entities และแท็กที่มักติดมาจาก Markdown/AI
    text = text.replace("&nbsp;", " ")
    text = text.replace("&amp;", "&")
    text = text.replace("&lt;", "<")
    text = text.replace("&gt;", ">")
    text = text.replace("&quot;", '"')
    text = re.sub(r"</?hr\s*/?>", "", text, flags=re.IGNORECASE)

    # 1. แก้คำสลับอักษรที่พบบ่อยใน OCR แบบฟอร์ม
    text = re.sub(r'พันกงาน', 'พนักงาน', text)
    text = re.sub(r'จังหัวด', 'จังหวัด', text)
    text = re.sub(r'หัลกฐาน', 'หลักฐาน', text)
    text = re.sub(r'ชื[@L1l!]?อัรบ', 'ชื่อรับ', text)

    # 2. แก้สระที่สลับตำแหน่งบ่อยๆ เช่น เกบ็ -> เก็บ, ไวเ้ -> ไว้, กวา่ -> กว่า, ไมต่อง้/ตอ้ง -> ต้อง
    text = re.sub(r'ไมต่อง้', 'ไม่ต้อง', text)
    text = re.sub(r'ไมต่\s*อ้ง', 'ไม่ต้อง', text)
    text = re.sub(r'ตอ้ง', 'ต้อง', text)
    text = re.sub(r'เกบ็', 'เก็บ', text)
    text = re.sub(r'เหน็', 'เห็น', text)
    text = re.sub(r'เปน็', 'เป็น', text)
    text = re.sub(r'เลก็', 'เล็ก', text)
    text = re.sub(r'เรว็', 'เร็ว', text)
    text = re.sub(r'ไวเ้', 'ไว้', text)
    text = re.sub(r'กวา่', 'กว่า', text)
    text = re.sub(r'([ก-ฮ])าํ', r'\1ำ', text)

    # 3. แก้ไขไม้โทที่กลายเป็น 2, 3, #, E หรือ (
    text = re.sub(r'([ก-ฮ])ั[@23#E\(]([ก-ฮ])', r'\1ั้\2', text)
    text = re.sub(r'ขั[\(#23E]น', 'ขั้น', text)
    text = re.sub(r'นี[E23#]', 'นี้', text)
    text = re.sub(r'ปี[E2#]', 'ปี้', text)
    text = re.sub(r'เพิ[E2#@]ม', 'เพิ่ม', text)
    text = re.sub(r'สิ[E23#]น', 'สิ้น', text)

    # 4. แก้ไขไม้เอกที่กลายเป็น 2, #, 1, L, l, ! หรือ @
    text = re.sub(r'([ก-ฮะ-ูเ-ไ])ที[@L|1l!2#]', r'\1ที่', text)
    text = re.sub(r'(^|\s)ที[@L|1l!2#]', r'\1ที่', text)
    text = re.sub(r'เพื[@L|1l!2#]อ', 'เพื่อ', text)
    text = re.sub(r'ชื[@L|1l!2#]อ', 'ชื่อ', text)

    # 5. แก้ไขสระหน้า + พยัญชนะ + 2/3 เช่น ได2 -> ได้, ให2 -> ให้
    text = re.sub(r'([เแไโ])([ก-ฮ])2', r'\1\2้', text)
    text = re.sub(r'([เแไโ])([ก-ฮ])3', r'\1\2้', text)

    return text


class GeminiExtractor:
    """โมดูลติดต่อ Gemini API พร้อมระบบ Seamless Model Fallback อัตโนมัติและการแปลภาษา"""

    def __init__(
        self,
        api_key: Optional[str] = None,
        models: Optional[List[str]] = None,
        on_fallback: Optional[Callable[[str, str, str], None]] = None,
    ):
        """
        api_key: Gemini API Key (ถ้าไม่ระบุจะดึงจาก .env)
        models: ลำดับโมเดลที่ต้องการ fallback (ค่าเริ่มต้น: FALLBACK_MODELS)
        on_fallback: Callback function เมื่อเกิดการสลับโมเดล (prev_model, next_model, reason)
        """
        self.api_key = get_api_key(api_key)
        self.models = models or FALLBACK_MODELS.copy()
        self.current_model_idx = 0
        self.on_fallback = on_fallback
        self._lock = threading.Lock()
        self.client = genai.Client(api_key=self.api_key)

    @property
    def current_model(self) -> str:
        with self._lock:
            return self.models[self.current_model_idx]

    def extract_page_markdown(
        self,
        image_bytes: bytes,
        page_num: int,
        target_language: Optional[str] = "original",
        vector_tables: Optional[List[dict]] = None,
    ) -> Tuple[str, str, Optional[str]]:
        """
        ประมวลผลภาพหน้า PDF/Image โดยส่งให้ Gemini อ่านข้อความ ตาราง สังเกตฟอนต์ และแปลภาษา (ถ้ามี)
        หากโมเดลปัจจุบันติด 429 Quota Exceeded หรือ Error จะสลับไปโมเดลสำรองถัดไปทันที
        ส่งคืน: (markdown_text, model_used, detected_font)
        """
        from .font_detector import clean_font_name

        lang_instructions = {
            "th": "และแปลเนื้อหาทั้งหมดเป็นภาษาไทย (Translate into natural, professional Thai) อย่างสละสลวยและถูกต้อง",
            "en": "และแปลเนื้อหาทั้งหมดเป็นภาษาอังกฤษ (Translate into natural, professional English) อย่างสละสลวยและถูกต้อง",
            "zh": "และแปลเนื้อหาทั้งหมดเป็นภาษาจีน (Translate into Chinese / 简体中文) อย่างสละสลวยและถูกต้อง",
            "ja": "และแปลเนื้อหาทั้งหมดเป็นภาษาญี่ปุ่น (Translate into Japanese / 日本語) อย่างสละสลวยและถูกต้อง",
        }

        trans_inst = ""
        if target_language and target_language.lower() not in ("original", "none"):
            trans_inst = lang_instructions.get(
                target_language.lower(),
                f"และแปลเนื้อหาทั้งหมดเป็นภาษา {target_language} อย่างถูกต้องและเป็นธรรมชาติ",
            )

        vec_hint = ""
        if vector_tables:
            tab_summaries = []
            for i, vt in enumerate(vector_tables):
                r_cnt = vt.get("row_count", "?")
                c_cnt = vt.get("col_count", "?")
                w_pt = vt.get("width_pt", "?")
                h_pt = vt.get("height_pt", "?")
                tab_summaries.append(f"ตารางที่ {i+1}: {r_cnt} แถว x {c_cnt} คอลัมน์ (กว้าง {w_pt}pt x สูง {h_pt}pt)")
            vec_hint = f" (ข้อมูลเวกเตอร์ตรวจพบ: {', '.join(tab_summaries)} ให้ถอดแบบตาราง 'type': 'table' ตามสัดส่วนและเส้นแบ่งจริงนี้)"

        if trans_inst:
            prompt = (
                f"นี่คือหน้า {page_num + 1} ของเอกสาร กรุณาวิเคราะห์และสร้างผลลัพธ์เป็น Structured Layout JSON {trans_inst}{vec_hint} "
                f"โดยรักษาความถูกต้องของข้อความ สัดส่วนคอลัมน์ตาราง [col_widths_pct] การจัดวาง และรูปภาพ [IMAGE] ไว้อย่างสมบูรณ์ตาม Schema"
            )
        else:
            prompt = (
                f"นี่คือหน้า {page_num + 1} ของเอกสาร กรุณาวิเคราะห์และสร้างผลลัพธ์เป็น Structured Layout JSON{vec_hint} "
                f"ระบุสัดส่วนคอลัมน์ตาราง [col_widths_pct] ความสูงแถว [height_pt] การรวมเซลล์ และรูปภาพ [IMAGE] ตาม Schema อย่างเคร่งครัด"
            )

        mime_type = "image/png" if image_bytes.startswith(b"\x89PNG") else "image/jpeg"
        image_part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)

        attempts = 0
        max_attempts = len(self.models) * 2  # ให้โอกาสลองใหม่

        while attempts < max_attempts:
            attempts += 1
            model_name = self.current_model
            try:
                logger.info(f"[Page {page_num + 1}] ส่งให้โมเดล {model_name} ประมวลผล (ภาษา: {target_language})...")

                response = self.client.models.generate_content(
                    model=model_name,
                    contents=[image_part, prompt],
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_PROMPT,
                        response_mime_type="application/json",
                        temperature=0.1,  # ควบคุมให้ออกมาตรงตามต้นฉบับ ไม่แต่งเติม
                    ),
                )

                text = response.text or ""
                detected_font = None

                # ตรวจสอบและทำความสะอาด Structural JSON
                clean_json_str = text.strip()
                if clean_json_str.startswith("```"):
                    clean_json_str = re.sub(r"^```(?:json)?\s*", "", clean_json_str)
                    clean_json_str = re.sub(r"\s*```$", "", clean_json_str).strip()

                is_valid_json = False
                try:
                    data = json.loads(clean_json_str)
                    if isinstance(data, dict) and "blocks" in data and isinstance(data["blocks"], list):
                        is_valid_json = True
                        font_cand = data.get("detected_font") or data.get("font_hint")
                        if font_cand:
                            detected_font = clean_font_name(str(font_cand))
                        self._clean_thai_in_blocks(data["blocks"])
                        text = json.dumps(data, ensure_ascii=False)
                except Exception:
                    is_valid_json = False

                if not is_valid_json:
                    # ตรวจจับแท็ก [FONT: ...] กรณีหลุดมาเป็นข้อความธรรมดา
                    font_match = re.search(r"\[FONT:\s*([^\]]+)\]", text)
                    if font_match:
                        raw_font = font_match.group(1).strip()
                        detected_font = clean_font_name(raw_font)
                        text = re.sub(r"\[FONT:\s*[^\]]+\]\n?", "", text).strip()
                    text = clean_thai_ocr_text(text)

                return text.strip(), model_name, detected_font

            except Exception as e:
                err_str = str(e)
                logger.warning(
                    f"เกิดข้อผิดพลาดกับโมเดล {model_name} (หน้า {page_num + 1}): {err_str}"
                )

                is_503_error = (
                    "503" in err_str
                    or "UNAVAILABLE" in err_str
                    or "high demand" in err_str.lower()
                )
                is_quota_error = (
                    "429" in err_str
                    or "RESOURCE_EXHAUSTED" in err_str
                    or "quota" in err_str.lower()
                    or "rate limit" in err_str.lower()
                )

                if is_503_error:
                    reason = "เซิร์ฟเวอร์ติดคิวยาวชั่วคราว (503 High Demand)"
                elif is_quota_error:
                    reason = "โควตาเต็ม (429 Quota Exceeded)"
                else:
                    reason = f"ข้อผิดพลาด: {err_str[:60]}"

                # สลับไปใช้โมเดลถัดไปใน fallback list
                if len(self.models) > 1:
                    with self._lock:
                        prev_model = self.models[self.current_model_idx]
                        self.current_model_idx = (self.current_model_idx + 1) % len(self.models)
                        next_model = self.models[self.current_model_idx]

                    logger.warning(
                        f"🔄 ระบบสลับโมเดลอัตโนมัติ: {prev_model} ➔ {next_model} (สาเหตุ: {reason})"
                    )

                    if self.on_fallback:
                        try:
                            self.on_fallback(prev_model, next_model, reason)
                        except Exception as cb_err:
                            logger.error(f"Error ใน callback on_fallback: {cb_err}")

                    # พักสักครู่ก่อนลองโมเดลถัดไป (สำหรับ 503 ให้รอ 2 วินาที)
                    sleep_time = 2.0 if is_503_error else 1.0
                    time.sleep(sleep_time)
                    continue
                else:
                    # หากมีโมเดลเดียวและติดโควตา ให้รอสักครู่แล้วลองใหม่
                    if is_quota_error:
                        logger.info("มีโมเดลเดียว รอ 5 วินาทีแล้วลองใหม่...")
                        time.sleep(5.0)
                        continue
        raise RuntimeError(f"ไม่สามารถประมวลผลหน้า {page_num + 1} ได้หลังจากลองทุกโมเดลแล้ว")

    def _clean_thai_in_blocks(self, blocks: list):
        """ตรวจแก้คำสลับและวรรณยุกต์ไทยในทุกบล็อกของ Structural JSON"""
        for block in blocks:
            if not isinstance(block, dict):
                continue
            b_type = str(block.get("type", "")).lower()
            if b_type == "table":
                for row in block.get("rows", []):
                    if isinstance(row, dict):
                        for cell in row.get("cells", []):
                            if isinstance(cell, dict) and "text" in cell:
                                cell["text"] = clean_thai_ocr_text(str(cell["text"]))
            elif b_type in ("heading", "paragraph"):
                if "text" in block:
                    block["text"] = clean_thai_ocr_text(str(block["text"]))
            elif b_type == "list":
                if "items" in block and isinstance(block["items"], list):
                    block["items"] = [clean_thai_ocr_text(str(it)) for it in block["items"]]

