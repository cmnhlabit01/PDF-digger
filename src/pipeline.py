import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, List, Optional
from .config import DEFAULT_FONT, MAX_PARALLEL_WORKERS
from .docx_builder import DocxBuilder
from .gemini_extractor import GeminiExtractor
from .pdf_processor import PDFProcessor

logger = logging.getLogger(__name__)


@dataclass
class PageConversionResult:
    """ผลลัพธ์การแปลงของแต่ละหน้า"""
    page_num: int
    markdown_text: str
    model_used: str
    images_count: int


@dataclass
class ConversionReport:
    """รายงานสรุปผลการแปลงไฟล์ PDF ทั้งฉบับ"""
    pdf_path: str
    output_docx_path: str
    total_pages: int
    successful_pages: int
    total_images_extracted: int
    models_used: List[str]
    detected_font: Optional[str] = None
    applied_font: str = DEFAULT_FONT
    target_language: str = "original"
    enhanced: bool = False
    fallback_events: List[str] = field(default_factory=list)
    page_results: List[PageConversionResult] = field(default_factory=list)


class PDFToWordPipeline:
    """คลาสประสานการทำงานตั้งแต่เปิด PDF/รูปภาพ -> ปรับแต่งภาพ -> แกะ/แปลผ่าน Gemini -> สร้างไฟล์ Word"""

    def __init__(
        self,
        api_key: Optional[str] = None,
        font_name: Optional[str] = None,
        auto_detect_font: bool = True,
        models: Optional[List[str]] = None,
        max_workers: int = MAX_PARALLEL_WORKERS,
        target_language: str = "original",
        enhance_image: bool = False,
    ):
        self.api_key = api_key
        self.font_name = font_name
        self.auto_detect_font = auto_detect_font
        self.models = models
        self.max_workers = max_workers
        self.target_language = target_language
        self.enhance_image = enhance_image
        self.fallback_events: List[str] = []

    def _handle_fallback(self, prev_model: str, next_model: str, reason: str):
        event_msg = f"สลับจาก {prev_model} ➔ {next_model} (สาเหตุ: {reason})"
        self.fallback_events.append(event_msg)
        logger.warning(f"🚨 Pipeline Fallback: {event_msg}")

    def convert(
        self,
        pdf_path: str | Path,
        output_docx_path: Optional[str | Path] = None,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        target_language: Optional[str] = None,
        enhance_image: Optional[bool] = None,
    ) -> ConversionReport:
        """
        แปลงไฟล์ PDF หรือไฟล์รูปภาพเป็น Word (.docx)
        
        pdf_path: ที่อยู่ไฟล์ต้นทาง (PDF หรือ รูปภาพ)
        output_docx_path: ที่อยู่ไฟล์ Word ปลายทาง (ถ้าไม่ระบุจะบันทึกชื่อเดียวกันแต่นามสกุล .docx)
        progress_callback: Callback สำหรับรายงานความคืบหน้า (current_page, total_pages, status_message)
        target_language: ภาษาเป้าหมาย ('original', 'th', 'en', 'zh', 'ja')
        enhance_image: เปิดโหมดปรับความคมชัดภาพสแกน (True/False)
        """
        input_file = Path(pdf_path)
        if not input_file.exists():
            raise FileNotFoundError(f"ไม่พบไฟล์ต้นทาง: {input_file}")

        if output_docx_path is None:
            output_docx_path = input_file.with_suffix(".docx")
        output_file = Path(output_docx_path)

        eff_lang = target_language or self.target_language or "original"
        eff_enhance = enhance_image if enhance_image is not None else self.enhance_image

        self.fallback_events = []

        # 1. เตรียมโมดูลต่างๆ
        processor = PDFProcessor(input_file)
        extractor = GeminiExtractor(
            api_key=self.api_key,
            models=self.models,
            on_fallback=self._handle_fallback,
        )

        # 2. ตรวจจับฟอนต์จากเอกสารต้นฉบับ
        detected_font = None
        if self.auto_detect_font and not processor.is_image:
            detected_font = processor.detect_font()
            if detected_font:
                logger.info(f"🔍 ตรวจพบฟอนต์ต้นฉบับจาก PDF: {detected_font}")

        # ตัดสินใจเลือกฟอนต์ที่จะใช้
        if self.font_name and self.font_name.lower() != "auto":
            applied_font = self.font_name
        elif detected_font:
            applied_font = detected_font
        else:
            applied_font = DEFAULT_FONT

        builder = DocxBuilder(font_name=applied_font)

        total_pages = processor.get_page_count()
        total_images = 0
        models_used_set = set()
        page_results: List[PageConversionResult] = []

        lang_label = f" (แปลภาษา ➔ {eff_lang})" if eff_lang != "original" else ""
        enhance_label = " + เพิ่มความคมชัดภาพ" if eff_enhance else ""
        logger.info(f"เริ่มแปลงไฟล์: {input_file.name} (จำนวน {total_pages} หน้า, ฟอนต์: {applied_font}{lang_label}{enhance_label})")

        def process_page_task(p_idx: int):
            p_data = processor.process_page(p_idx, enhance=eff_enhance)
            md_text, m_used, p_font = extractor.extract_page_markdown(
                p_data.rendered_image_bytes, p_idx, target_language=eff_lang
            )
            return (p_idx, p_data, md_text, m_used, p_font)

        # 3. ประมวลผลหน้าเอกสารแบบคู่ขนาน (Parallel Concurrency) เพื่อเพิ่มความเร็วสูงสุด
        effective_workers = min(self.max_workers, total_pages)
        raw_results = []
        completed_count = 0

        if effective_workers > 1:
            logger.info(f"⚡ เริ่มประมวลผลแบบคู่ขนาน ({effective_workers} หน้าพร้อมกัน)...")
            with ThreadPoolExecutor(max_workers=effective_workers) as executor:
                future_to_idx = {
                    executor.submit(process_page_task, p_idx): p_idx
                    for p_idx in range(total_pages)
                }
                for future in as_completed(future_to_idx):
                    completed_count += 1
                    res = future.result()
                    raw_results.append(res)
                    if progress_callback:
                        progress_callback(
                            completed_count,
                            total_pages,
                            f"อ่านและสกัดข้อมูลเสร็จแล้ว {completed_count}/{total_pages} หน้า...",
                        )
        else:
            for p_idx in range(total_pages):
                if progress_callback:
                    progress_callback(
                        p_idx + 1,
                        total_pages,
                        f"กำลังอ่านและสกัดข้อมูลหน้า {p_idx + 1}/{total_pages} ด้วย {extractor.current_model}...",
                    )
                res = process_page_task(p_idx)
                raw_results.append(res)

        # เรียงผลลัพธ์กลับตามลำดับหน้า 0, 1, 2, ... ให้ตรงตามต้นฉบับเสมอ
        raw_results.sort(key=lambda x: x[0])

        # ตรวจสอบฟอนต์จากหน้าแรกหากตรวจไม่พบจาก PDF ดิจิทัล
        if raw_results:
            page_0_font = raw_results[0][4]
            if not detected_font and page_0_font and (not self.font_name or self.font_name.lower() == "auto"):
                detected_font = page_0_font
                applied_font = page_0_font
                builder.set_font(applied_font)
                logger.info(f"🔍 Gemini ตรวจพบฟอนต์จากภาพสแกน: {applied_font}")

        # 4. ประกอบข้อมูลลงในเอกสาร Word ตามลำดับหน้าที่ถูกต้อง
        for p_idx, p_data, md_text, m_used, _ in raw_results:
            p_num = p_idx + 1
            total_images += len(p_data.embedded_images)
            models_used_set.add(m_used)

            builder.add_page_content(
                markdown_text=md_text,
                page_num=p_num,
                images=p_data.embedded_images,
                is_first_page=(p_idx == 0),
            )

            page_results.append(
                PageConversionResult(
                    page_num=p_num,
                    markdown_text=md_text,
                    model_used=m_used,
                    images_count=len(p_data.embedded_images),
                )
            )

        # 5. บันทึกไฟล์ Word
        if progress_callback:
            progress_callback(total_pages, total_pages, "กำลังบันทึกไฟล์ Word (.docx)...")

        builder.save(output_file)
        logger.info(f"สร้างไฟล์ Word สำเร็จที่: {output_file} (ใช้ฟอนต์: {applied_font})")

        return ConversionReport(
            pdf_path=str(input_file),
            output_docx_path=str(output_file),
            total_pages=total_pages,
            successful_pages=len(page_results),
            total_images_extracted=total_images,
            models_used=sorted(list(models_used_set)),
            detected_font=detected_font,
            applied_font=applied_font,
            target_language=eff_lang,
            enhanced=eff_enhance,
            fallback_events=self.fallback_events,
            page_results=page_results,
        )
