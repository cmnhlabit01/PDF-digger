import logging
import threading
import time
from typing import Callable, List, Optional, Tuple
from google import genai
from google.genai import types
from google.genai.errors import APIError

from .config import FALLBACK_MODELS, get_api_key

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """คุณคือระบบ OCR และ Document Structure Extractor ระดับสูงที่มีความเชี่ยวชาญพิเศษด้านภาษาไทยและภาษาอังกฤษ
หน้าที่ของคุณคือแกะตัวอักษรและโครงสร้างทั้งหมดจากภาพหน้าเอกสารที่ได้รับ และแปลงเป็น Markdown โดยมีกฎเหล็กดังต่อไปนี้:

1. ความถูกต้องของภาษาไทยและอังกฤษ (Verbatim & Accurate):
   - แกะข้อความตามต้นฉบับคำต่อคำ ห้ามสรุปความ ห้ามตัดทอน ห้ามแต่งเติมเด็ดขาด
   - รักษาความถูกต้องของสระและวรรณยุกต์ภาษาไทย 100% เช่น สระอิ สระอี ไม้เอก ไม้โท ไม้ตรี ไม้จัตวา สระอำ (ำ) การันต์ (์) รวมถึงสระลอยและวรรณยุกต์ซ้อน
   - เว้นวรรคคำภาษาไทยให้ถูกต้องเป็นธรรมชาติ ไม่เว้นวรรคมั่วกลางคำ

2. โครงสร้างและลำดับเอกสาร (Document Hierarchy):
   - ใช้ `#` สำหรับชื่อเรื่องหลัก (Title / Heading 1)
   - ใช้ `##` สำหรับหัวข้อย่อยระดับ 1 (Heading 2)
   - ใช้ `###` สำหรับหัวข้อย่อยระดับ 2 (Heading 3)
   - ใช้ `- ` หรือ `* ` สำหรับรายการหัวข้อย่อย (Bullet lists)
   - ใช้ `1. `, `2. ` สำหรับรายการแบบเรียงลำดับตัวเลข
   - ใช้ `**ข้อความ**` สำหรับตัวอักษรตัวหนา

3. การถอดแบบตารางอย่างสมบูรณ์แบบ (High-Fidelity Tables):
   - หากมีตารางในเอกสาร ให้แปลงเป็น Markdown Table (`| Header 1 | Header 2 |`) ให้ครบทุกแถวและทุกคอลัมน์
   - ต้องคงจำนวนคอลัมน์และแถวให้เหมือนต้นฉบับ ห้ามข้ามเซลล์ที่ว่าง (ให้เว้นเป็นช่องว่าง `| |`)
   - หากในเซลล์มีตัวเลข ข้อมูลทางการเงิน หรือเปอร์เซ็นต์ ให้แกะอย่างแม่นยำที่สุด
   - ตารางต้องเริ่มต้นด้วยหัวตารางและเส้นขีดคั่น `|---|---|` เสมอ

4. ตำแหน่งรูปภาพและแผนภาพ (Images & Illustrations):
   - ห้ามใส่แท็ก [IMAGE] แทนหน้าเอกสารสแกน แบบฟอร์ม สไลด์ ใบเสร็จ หรือข้อความเด็ดขาด!
   - หากหน้านี้เป็นเอกสารสแกนหรือภาพถ่ายหน้ากระดาษ ต้องถอดรหัสและแกะข้อความทุกตัวอักษรและตารางออกมาเป็นตัวหนังสือ (Editable Text) คำต่อคำ
   - ใส่แท็ก `[IMAGE]` เฉพาะเมื่อมี "รูปภาพประกอบจริงๆ" อยู่ในเนื้อหาเท่านั้น เช่น ภาพถ่ายบุคคล/สินค้า, ตราสัญลักษณ์ (Logo), แผนภูมิ (Diagram) หรือกราฟ ในบรรทัดแยกต่างหาก

5. การสังเกตแบบอักษร (Font Style Analysis):
   - หากสังเกตเห็นฟอนต์หลักที่ใช้ในหน้านี้อย่างชัดเจน (เช่น TH Sarabun, Cordia New, Angsana New, Tahoma, Calibri) สามารถระบุไว้บรรทัดแรกสุด เช่น:
     `[FONT: TH Sarabun New]`
     หากไม่แน่ใจ ไม่จำเป็นต้องใส่แท็กนี้

6. ผลลัพธ์:
   - ส่งออกผลลัพธ์เป็นข้อความ Markdown ล้วนๆ ไม่ต้องใส่ข้อความอธิบายทักทายใดๆ นอกเหนือจากเนื้อหาของเอกสาร
"""


class GeminiExtractor:
    """โมดูลติดต่อ Gemini API พร้อมระบบ Seamless Model Fallback อัตโนมัติ"""

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
        self, image_bytes: bytes, page_num: int
    ) -> Tuple[str, str, Optional[str]]:
        """
        ประมวลผลภาพหน้า PDF โดยส่งให้ Gemini อ่านข้อความ ตาราง และสังเกตฟอนต์
        หากโมเดลปัจจุบันติด 429 Quota Exceeded หรือ Error จะสลับไปโมเดลสำรองถัดไปทันที
        ส่งคืน: (markdown_text, model_used, detected_font)
        """
        import re
        from .font_detector import clean_font_name

        prompt = (
            f"นี่คือหน้า {page_num + 1} ของเอกสาร กรุณาแกะข้อความ หัวข้อ รายการ ตาราง "
            f"ระบุตำแหน่งรูปภาพ [IMAGE] และสังเกตฟอนต์ [FONT: ...] ตามกฎที่กำหนดอย่างเคร่งครัด"
        )

        mime_type = "image/png" if image_bytes.startswith(b"\x89PNG") else "image/jpeg"
        image_part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)

        attempts = 0
        max_attempts = len(self.models) * 2  # ให้โอกาสลองใหม่

        while attempts < max_attempts:
            attempts += 1
            model_name = self.current_model
            try:
                logger.info(f"[Page {page_num + 1}] ส่งให้โมเดล {model_name} ประมวลผล...")

                response = self.client.models.generate_content(
                    model=model_name,
                    contents=[image_part, prompt],
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_PROMPT,
                        temperature=0.1,  # ควบคุมให้ออกมาตรงตามต้นฉบับ ไม่แต่งเติม
                    ),
                )

                text = response.text or ""

                # ตรวจจับแท็ก [FONT: ...]
                detected_font = None
                font_match = re.search(r"\[FONT:\s*([^\]]+)\]", text)
                if font_match:
                    raw_font = font_match.group(1).strip()
                    detected_font = clean_font_name(raw_font)
                    # ตัดแท็กฟอนต์ออกจากเนื้อหา
                    text = re.sub(r"\[FONT:\s*[^\]]+\]\n?", "", text).strip()

                return text.strip(), model_name, detected_font

            except Exception as e:
                err_str = str(e)
                logger.warning(
                    f"เกิดข้อผิดพลาดกับโมเดล {model_name} (หน้า {page_num + 1}): {err_str}"
                )

                # ตรวจสอบว่าเป็นข้อผิดพลาดด้านโควตา (429 / RESOURCE_EXHAUSTED) หรือ Rate limit
                is_quota_error = (
                    "429" in err_str
                    or "RESOURCE_EXHAUSTED" in err_str
                    or "quota" in err_str.lower()
                    or "rate limit" in err_str.lower()
                )

                # สลับไปใช้โมเดลถัดไปใน fallback list
                if len(self.models) > 1:
                    with self._lock:
                        prev_model = self.models[self.current_model_idx]
                        self.current_model_idx = (self.current_model_idx + 1) % len(self.models)
                        next_model = self.models[self.current_model_idx]

                    reason = "โควตาเต็ม (429 Quota Exceeded)" if is_quota_error else f"ข้อผิดพลาด: {err_str[:80]}"

                    logger.warning(
                        f"🔄 ระบบสลับโมเดลอัตโนมัติ: {prev_model} ➔ {next_model} (สาเหตุ: {reason})"
                    )

                    if self.on_fallback:
                        try:
                            self.on_fallback(prev_model, next_model, reason)
                        except Exception as cb_err:
                            logger.error(f"Error ใน callback on_fallback: {cb_err}")

                    # พักสักครู่ก่อนลองโมเดลถัดไป
                    time.sleep(1.0)
                    continue
                else:
                    # หากมีโมเดลเดียวและติดโควตา ให้รอสักครู่แล้วลองใหม่
                    if is_quota_error:
                        logger.info("มีโมเดลเดียว รอ 5 วินาทีแล้วลองใหม่...")
                        time.sleep(5.0)
                        continue
                    raise e

        raise RuntimeError(f"ไม่สามารถประมวลผลหน้า {page_num + 1} ได้หลังจากลองทุกโมเดลแล้ว")
