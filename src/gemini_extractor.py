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

6. บล็อกหัวกระดาษ ("type": "header"):
   - ใช้เฉพาะเมื่อเอกสารมีหัวกระดาษ (Running Header) ด้านบนสุดของหน้ากระดาษ (เช่น ชื่อเอกสาร, รหัสโครงการ, โลโก้บนหัวกระดาษ)
   - "text": ข้อความหัวกระดาษ
   - "align": "left", "center", หรือ "right"

7. บล็อกท้ายกระดาษ ("type": "footer"):
   - ใช้เฉพาะเมื่อเอกสารมีท้ายกระดาษ (Running Footer) ด้านล่างสุดของหน้ากระดาษ (เช่น หมายเลขหน้า "หน้า 1", วันที่พิมพ์, ข้อความสงวนลิขสิทธิ์, รหัสเอกสาร)
   - "text": ข้อความท้ายกระดาษ
   - "align": "left", "center", หรือ "right"
   - "is_page_number": true หากมีข้อความระบุหมายเลขหน้า

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


def is_raw_json_code(text: str) -> bool:
    """ตรวจสอบว่าข้อความดูเหมือนเป็นโค้ด JSON หรือไม่ เพื่อป้องกันการรั่วไหลลงเอกสาร"""
    if not text:
        return False
    s = text.strip()
    if (s.startswith("{") or s.startswith("[")) and (
        '"type":' in s or '"blocks":' in s or '"cells":' in s or '"rows":' in s
    ):
        return True
    if '"type": "table"' in s or '"type": "paragraph"' in s or '"col_widths_pct"' in s:
        return True
    return False


def fallback_extract_text_from_broken_json(json_str: str) -> List[str]:
    """สกัดเฉพาะข้อความที่อ่านได้จาก JSON ที่เสีย แทนที่จะปล่อยให้แท็กและโค้ด JSON หลุดลงหน้าเอกสาร"""
    extracted_lines = []
    # ค้นหาค่าใน "text": "..."
    matches = re.findall(r'"text"\s*:\s*"((?:\\.|[^"\\])*)"', json_str)
    for m in matches:
        val = m.replace('\\"', '"').replace('\\n', '\n').replace('\\t', ' ').strip()
        val = clean_thai_ocr_text(val)
        if val and not is_raw_json_code(val):
            extracted_lines.append(val)
    return extracted_lines


def unwrap_gemini_json(text: str) -> Tuple[bool, Optional[dict], Optional[str]]:
    """
    ถอดรหัสและ Normalize ผลลัพธ์ JSON จาก Gemini ไม่ว่าจะมาในรูปแบบ:
    1. Single Object: {"page_type": "...", "blocks": [...]}
    2. Array of Single Object: [{"page_type": "...", "blocks": [...]}]
    3. Array of Blocks: [{"type": "table", ...}, {"type": "paragraph", ...}]
    4. มี Markdown code fence: ```json ... ```
    """
    from .font_detector import clean_font_name
    if not text:
        return False, None, None

    clean_str = text.strip()
    if clean_str.startswith("```"):
        clean_str = re.sub(r"^```(?:json)?\s*", "", clean_str)
        clean_str = re.sub(r"\s*```$", "", clean_str).strip()

    data = None
    try:
        data = json.loads(clean_str)
    except Exception:
        for open_ch, close_ch in [("{", "}"), ("[", "]")]:
            start_i = clean_str.find(open_ch)
            end_i = clean_str.rfind(close_ch)
            if start_i != -1 and end_i != -1 and end_i > start_i:
                sub = clean_str[start_i : end_i + 1]
                try:
                    data = json.loads(sub)
                    break
                except Exception:
                    sub_fixed = re.sub(r",\s*([\}\]])", r"\1", sub)
                    try:
                        data = json.loads(sub_fixed)
                        break
                    except Exception:
                        pass

    if data is None:
        return False, None, None

    blocks = []
    page_type = "document"
    detected_font = None

    if isinstance(data, list):
        if not data:
            return False, None, None
        if isinstance(data[0], dict) and "blocks" in data[0] and isinstance(data[0]["blocks"], list):
            blocks = data[0]["blocks"]
            page_type = data[0].get("page_type", "document")
            font_cand = data[0].get("detected_font") or data[0].get("font_hint")
            if font_cand:
                detected_font = clean_font_name(str(font_cand))
        elif any(isinstance(x, dict) and "type" in x for x in data):
            blocks = [x for x in data if isinstance(x, dict)]
    elif isinstance(data, dict):
        if "blocks" in data and isinstance(data["blocks"], list):
            blocks = data["blocks"]
            page_type = data.get("page_type", "document")
            font_cand = data.get("detected_font") or data.get("font_hint")
            if font_cand:
                detected_font = clean_font_name(str(font_cand))
        else:
            for val in data.values():
                if isinstance(val, dict) and "blocks" in val and isinstance(val["blocks"], list):
                    blocks = val["blocks"]
                    page_type = val.get("page_type", "document")
                    break

    if blocks:
        normalized = {
            "page_type": page_type,
            "detected_font": detected_font,
            "blocks": blocks,
        }
        return True, normalized, detected_font

    return False, None, None


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
        self._model_cooldowns: dict[str, float] = {}  # model_name -> unblock timestamp
        self.client = genai.Client(api_key=self.api_key)

    @property
    def current_model(self) -> str:
        with self._lock:
            return self.models[self.current_model_idx]

    def _is_model_cooling_down(self, model_name: str) -> bool:
        with self._lock:
            until = self._model_cooldowns.get(model_name, 0.0)
            return time.time() < until

    def _set_model_cooldown(self, model_name: str, duration_sec: float):
        with self._lock:
            self._model_cooldowns[model_name] = time.time() + duration_sec

    def _clear_model_cooldown(self, model_name: str):
        with self._lock:
            self._model_cooldowns.pop(model_name, None)

    def extract_page_markdown(
        self,
        image_bytes: bytes,
        page_num: int,
        target_language: Optional[str] = "original",
        vector_tables: Optional[List[dict]] = None,
        header_hint: Optional[str] = None,
        footer_hint: Optional[str] = None,
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

        hf_hint = ""
        if header_hint:
            hf_hint += f" (ตรวจพบข้อความส่วนหัวกระดาษ: {header_hint} หากเป็น Running Header ให้ใช้ 'type': 'header')"
        if footer_hint:
            hf_hint += f" (ตรวจพบข้อความส่วนท้ายกระดาษ: {footer_hint} ให้ใช้ 'type': 'footer')"

        if trans_inst:
            prompt = (
                f"นี่คือหน้า {page_num + 1} ของเอกสาร กรุณาวิเคราะห์และสร้างผลลัพธ์เป็น Structured Layout JSON {trans_inst}{vec_hint}{hf_hint} "
                f"โดยรักษาความถูกต้องของข้อความ สัดส่วนคอลัมน์ตาราง [col_widths_pct] การจัดวาง และรูปภาพ [IMAGE] ไว้อย่างสมบูรณ์ตาม Schema"
            )
        else:
            prompt = (
                f"นี่คือหน้า {page_num + 1} ของเอกสาร กรุณาวิเคราะห์และสร้างผลลัพธ์เป็น Structured Layout JSON{vec_hint}{hf_hint} "
                f"ระบุสัดส่วนคอลัมน์ตาราง [col_widths_pct] ความสูงแถว [height_pt] การรวมเซลล์ หัวกระดาษ [header] ท้ายกระดาษ [footer] และรูปภาพ [IMAGE] ตาม Schema อย่างเคร่งครัด"
            )

        mime_type = "image/png" if image_bytes.startswith(b"\x89PNG") else "image/jpeg"
        image_part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)

        max_rounds = 2
        for round_idx in range(max_rounds):
            with self._lock:
                start_idx = self.current_model_idx
                candidate_models = [
                    self.models[(start_idx + i) % len(self.models)]
                    for i in range(len(self.models))
                ]

            # ตรวจสอบว่ามีโมเดลที่ไม่อยู่ใน cooldown หรือไม่
            non_cooling = [m for m in candidate_models if not self._is_model_cooling_down(m)]
            models_to_try = non_cooling if non_cooling else candidate_models

            # หากทุกโมเดลติด cooldown และเวลารอไม่นาน ให้รอก่อนลอง
            if not non_cooling and len(self.models) > 1:
                with self._lock:
                    now = time.time()
                    waits = [max(0.0, self._model_cooldowns.get(m, 0.0) - now) for m in self.models]
                min_wait = min(waits) if waits else 0.0
                if 0 < min_wait <= 35:
                    logger.info(f"[Page {page_num + 1}] โมเดลทั้งหมดติด Cooldown ชั่วคราว รอ {min_wait:.1f} วินาที...")
                    time.sleep(min_wait + 0.5)

            for model_name in models_to_try:
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

                    # ตรวจสอบและ Normalize ผลลัพธ์ Structural JSON ทุกรูปแบบ (Object, Array, Nested)
                    is_valid_json, normalized_data, detected_font_cand = unwrap_gemini_json(text)
                    if is_valid_json and normalized_data:
                        if detected_font_cand:
                            detected_font = detected_font_cand
                        self._clean_thai_in_blocks(normalized_data["blocks"])
                        text = json.dumps(normalized_data, ensure_ascii=False)
                    else:
                        # หากไม่ใช่ JSON ที่สมบูรณ์ ให้ตรวจสอบว่ามีรหัส JSON ปนเปื้อนหรือไม่ (Anti-leak guard)
                        if is_raw_json_code(text):
                            logger.warning(f"[Page {page_num + 1}] ตรวจพบโค้ด JSON หลุดจากโมเดล ทำการสกัดเฉพาะข้อความจริง...")
                            clean_lines = fallback_extract_text_from_broken_json(text)
                            text = "\n\n".join(clean_lines)
                        else:
                            font_match = re.search(r"\[FONT:\s*([^\]]+)\]", text)
                            if font_match:
                                raw_font = font_match.group(1).strip()
                                detected_font = clean_font_name(raw_font)
                                text = re.sub(r"\[FONT:\s*[^\]]+\]\n?", "", text).strip()
                            text = clean_thai_ocr_text(text)

                    # สำเร็จ: ปรับให้หน้านี้และหน้าถัดไปใช้โมเดลนี้เป็นโมเดลหลัก
                    with self._lock:
                        if model_name in self.models:
                            self.current_model_idx = self.models.index(model_name)
                    self._clear_model_cooldown(model_name)

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
                        cooldown_sec = 3.0
                        sleep_time = 1.5
                    elif is_quota_error:
                        reason = "โควตาเต็ม (429 Quota Exceeded)"
                        if "perday" in err_str.lower() or "free_tier_requests" in err_str.lower():
                            cooldown_sec = 3600.0  # โควตารายวันเต็ม พัก 1 ชั่วโมง
                        else:
                            delay_match = re.search(r"retryDelay['\"]?\s*:\s*['\"]?(\d+)s?", err_str)
                            if delay_match:
                                cooldown_sec = float(delay_match.group(1)) + 2.0
                            else:
                                cooldown_sec = 30.0
                        sleep_time = 1.0
                    else:
                        reason = f"ข้อผิดพลาด: {err_str[:60]}"
                        cooldown_sec = 15.0
                        sleep_time = 1.0

                    self._set_model_cooldown(model_name, cooldown_sec)

                    # สลับไปใช้โมเดลถัดไปใน fallback list
                    if len(self.models) > 1:
                        next_idx = (self.models.index(model_name) + 1) % len(self.models)
                        next_model = self.models[next_idx]
                        with self._lock:
                            self.current_model_idx = next_idx

                        logger.warning(
                            f"🔄 ระบบสลับโมเดลอัตโนมัติ: {model_name} ➔ {next_model} (สาเหตุ: {reason})"
                        )

                        if self.on_fallback:
                            try:
                                self.on_fallback(model_name, next_model, reason)
                            except Exception as cb_err:
                                logger.error(f"Error ใน callback on_fallback: {cb_err}")

                        time.sleep(sleep_time)
                    else:
                        if is_quota_error:
                            logger.info("มีโมเดลเดียว รอ 5 วินาทีแล้วลองใหม่...")
                            time.sleep(5.0)

            # หากจบ Candidate Models ในรอบแรกแล้ว ให้ดูว่ามี cooldown สั้นที่รอได้หรือไม่
            if round_idx < max_rounds - 1:
                with self._lock:
                    now = time.time()
                    waits = [max(0.0, self._model_cooldowns.get(m, 0.0) - now) for m in self.models]
                short_waits = [cd for cd in waits if 0 < cd <= 35]
                if short_waits:
                    wait_sec = min(short_waits) + 0.5
                    logger.info(f"[Page {page_num + 1}] โมเดลทั้งหมดติดสถานะชั่วคราว รอ {wait_sec:.1f} วินาทีเพื่อลองใหม่...")
                    time.sleep(wait_sec)

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
            elif b_type in ("heading", "paragraph", "header", "footer"):
                if "text" in block:
                    block["text"] = clean_thai_ocr_text(str(block["text"]))
            elif b_type == "list":
                if "items" in block and isinstance(block["items"], list):
                    block["items"] = [clean_thai_ocr_text(str(it)) for it in block["items"]]

