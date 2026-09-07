import io
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple
import pymupdf as fitz
from PIL import Image

from .config import RENDER_DPI


@dataclass
class ExtractedImage:
    """ข้อมูลรูปภาพที่สกัดมาจากหน้า PDF"""
    page_num: int
    image_index: int
    image_bytes: bytes
    ext: str
    width: int
    height: int
    bbox: Optional[Tuple[float, float, float, float]] = None


@dataclass
class PDFPageData:
    """ข้อมูลของแต่ละหน้าใน PDF สำหรับนำไปประมวลผลต่อ"""
    page_num: int
    rendered_image_bytes: bytes
    embedded_images: List[ExtractedImage]
    has_text: bool


class PDFProcessor:
    """คลาสสำหรับจัดการไฟล์ PDF: เรนเดอร์หน้าเป็นภาพความละเอียดสูง และสกัดรูปภาพประกอบ"""

    def __init__(self, pdf_path: str | Path, dpi: int = RENDER_DPI):
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"ไม่พบไฟล์ PDF ที่: {self.pdf_path}")
        self.dpi = dpi

    def get_page_count(self) -> int:
        """นับจำนวนหน้าทั้งหมดของ PDF"""
        with fitz.open(self.pdf_path) as doc:
            return len(doc)

    def detect_font(self) -> Optional[str]:
        """ตรวจจับฟอนต์หลักที่ใช้ในเอกสาร PDF จาก metadata หรือ text spans"""
        from .font_detector import detect_dominant_font
        return detect_dominant_font(self.pdf_path)

    def process_page(self, page_num: int) -> PDFPageData:
        """
        ประมวลผลหน้า PDF ตามหมายเลขหน้าที่ระบุ (0-indexed)
        - เรนเดอร์หน้าทั้งหมดเป็นภาพความละเอียดสูงสำหรับ Gemini
        - สกัดรูปภาพประกอบที่ฝังอยู่ในหน้านี้
        """
        with fitz.open(self.pdf_path) as doc:
            if page_num < 0 or page_num >= len(doc):
                raise IndexError(f"หมายเลขหน้า {page_num} อยู่นอกช่วง (มีทั้งหมด {len(doc)} หน้า)")

            page = doc[page_num]

            # 1. ตรวจสอบว่าหน้านี้มีข้อความแบบ digital หรือไม่
            raw_text = page.get_text()
            has_text = len(raw_text.strip()) > 0

            # 2. เรนเดอร์หน้าเป็นภาพความละเอียดสูง 200 DPI (JPEG 92% คมชัดสูง ไฟล์เล็ก อัปโหลดเร็วขึ้น 10 เท่า)
            zoom = self.dpi / 72.0
            matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            rendered_image_bytes = pix.tobytes(output="jpeg", jpg_quality=92)

            # 3. สกัดรูปภาพที่ฝังอยู่ในหน้า
            embedded_images = self._extract_images_from_page(doc, page, page_num)

            return PDFPageData(
                page_num=page_num,
                rendered_image_bytes=rendered_image_bytes,
                embedded_images=embedded_images,
                has_text=has_text,
            )

    def _extract_images_from_page(
        self, doc: fitz.Document, page: fitz.Page, page_num: int
    ) -> List[ExtractedImage]:
        """สกัดรูปภาพประกอบจากหน้า PDF โดยกรองรูปขนาดเล็กหรือไอคอนออก"""
        extracted = []
        image_list = page.get_images(full=True)

        for img_idx, img_info in enumerate(image_list):
            xref = img_info[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image.get("image")
            image_ext = base_image.get("ext", "png")
            width = base_image.get("width", 0)
            height = base_image.get("height", 0)

            # กรองภาพขนาดเล็กมากๆ เช่น เส้นประ, ไอคอนตกแต่ง, หรือ 1x1 spacer
            if width < 60 or height < 60:
                continue

            # ค้นหาตำแหน่ง Bounding Box ของรูปภาพในหน้าถ้ามี
            bbox = None
            is_full_page_scan = False
            page_area = page.rect.width * page.rect.height

            rects = page.get_image_rects(xref)
            for rect in rects:
                bbox = (rect.x0, rect.y0, rect.x1, rect.y1)
                img_area = rect.width * rect.height
                # หากภาพครอบคลุมพื้นที่หน้ากระดาษเกิน 60% แสดงว่าเป็นภาพสแกนเอกสารทั้งหน้า
                if page_area > 0 and (img_area / page_area) > 0.60:
                    is_full_page_scan = True
                break

            # กรณีไม่มี rect ชัดเจน หรือภาพครอบคลุมทั้งหน้า
            if is_full_page_scan:
                continue

            # หากทั้งหน้านี้มีรูปเดียวและครอบคลุมทั้งหน้ากระดาษ (ภาพสแกนทั้งหน้า) ให้ข้าม
            if len(image_list) == 1 and width >= 700 and height >= 700:
                continue

            extracted.append(
                ExtractedImage(
                    page_num=page_num,
                    image_index=img_idx + 1,
                    image_bytes=image_bytes,
                    ext=image_ext,
                    width=width,
                    height=height,
                    bbox=bbox,
                )
            )

        return extracted
