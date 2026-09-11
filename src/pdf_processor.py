import io
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple
from PIL import Image, ImageEnhance, ImageOps
import pymupdf as fitz

from .config import RENDER_DPI

SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff", ".tif"}


def enhance_document_image(image_bytes: bytes) -> bytes:
    """
    ปรับแต่งภาพหน้าเอกสาร/ภาพสแกนให้คมชัดยิ่งขึ้นก่อนส่งให้ OCR:
    - ปรับ Auto-contrast (ตัดขอบสี 0.5% เพื่อลดรอยกระดาษเหลือง/หมอง ให้พื้นหลังขาวขึ้น)
    - เร่งความคมชัดของขอบตัวอักษรและสระภาษาไทย (Edge Sharpening)
    - ปรับสมดุล Contrast ให้หมึกตัวหนังสือเข้มเด่นชัด
    """
    try:
        img = Image.open(io.BytesIO(image_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        elif img.mode != "RGB":
            img = img.convert("RGB")

        # 1. Auto-contrast ตัดฝุ่นและปรับแสงกระดาษ
        img = ImageOps.autocontrast(img, cutoff=0.5)

        # 2. เร่งความคมชัดขอบตัวอักษร (Sharpness)
        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(1.4)

        # 3. เร่ง Contrast อีกเล็กน้อยให้ตัวหนังสือเข้มขึ้น
        c_enhancer = ImageEnhance.Contrast(img)
        img = c_enhancer.enhance(1.15)

        out_buf = io.BytesIO()
        img.save(out_buf, format="JPEG", quality=92)
        return out_buf.getvalue()
    except Exception:
        return image_bytes


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
    """ข้อมูลของแต่ละหน้าใน PDF/Image สำหรับนำไปประมวลผลต่อ"""
    page_num: int
    rendered_image_bytes: bytes
    embedded_images: List[ExtractedImage]
    has_text: bool


class PDFProcessor:
    """คลาสสำหรับจัดการไฟล์ PDF หรือไฟล์รูปภาพ: เรนเดอร์หน้าเป็นภาพความละเอียดสูง สกัดรูปประกอบ และปรับแต่งภาพ"""

    def __init__(self, file_path: str | Path, dpi: int = RENDER_DPI):
        self.file_path = Path(file_path)
        self.pdf_path = self.file_path  # เพื่อความเข้ากันได้ย้อนหลัง
        if not self.file_path.exists():
            raise FileNotFoundError(f"ไม่พบไฟล์ต้นทางที่: {self.file_path}")
        self.dpi = dpi
        self.is_image = self.file_path.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS

    def get_page_count(self) -> int:
        """นับจำนวนหน้าทั้งหมด (ถ้าเป็นไฟล์รูปภาพเดี่ยว จะคืนค่า 1)"""
        if self.is_image:
            return 1

        with fitz.open(self.file_path) as doc:
            return len(doc)

    def detect_font(self) -> Optional[str]:
        """ตรวจจับฟอนต์หลักที่ใช้ในเอกสาร PDF จาก metadata หรือ text spans (หากเป็นรูปภาพจะคืนค่า None)"""
        if self.is_image:
            return None
        from .font_detector import detect_dominant_font
        return detect_dominant_font(self.file_path)

    def process_page(self, page_num: int, enhance: bool = False) -> PDFPageData:
        """
        ประมวลผลหน้า PDF หรือรูปภาพตามหมายเลขหน้าที่ระบุ (0-indexed)
        - เรนเดอร์หน้าทั้งหมดเป็นภาพความละเอียดสูงสำหรับ Gemini
        - หาก enhance=True จะทำการปรับ Contrast & Sharpening เพิ่มเติม
        - สกัดรูปภาพประกอบที่ฝังอยู่ในหน้านี้
        """
        if self.is_image:
            # กรณีไฟล์ต้นทางเป็นรูปภาพโดยตรง (.png, .jpg, .webp ฯลฯ)
            with Image.open(self.file_path) as img:
                if img.mode in ("RGBA", "P"):
                    img_rgb = img.convert("RGB")
                elif img.mode != "RGB":
                    img_rgb = img.convert("RGB")
                else:
                    img_rgb = img.copy()

                # ปรับขนาดภาพหากใหญ่เกินไปเพื่อความเร็วในการส่ง API (สูงสุด 2400px)
                max_dim = 2400
                if max(img_rgb.size) > max_dim:
                    scale = max_dim / float(max(img_rgb.size))
                    new_size = (int(img_rgb.size[0] * scale), int(img_rgb.size[1] * scale))
                    img_rgb = img_rgb.resize(new_size, Image.Resampling.LANCZOS)

                out_buf = io.BytesIO()
                img_rgb.save(out_buf, format="JPEG", quality=92)
                rendered_bytes = out_buf.getvalue()

                if enhance:
                    rendered_bytes = enhance_document_image(rendered_bytes)

                return PDFPageData(
                    page_num=page_num,
                    rendered_image_bytes=rendered_bytes,
                    embedded_images=[],
                    has_text=False,
                )

        # กรณีไฟล์ต้นทางเป็น PDF
        with fitz.open(self.file_path) as doc:
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

            if enhance:
                rendered_image_bytes = enhance_document_image(rendered_image_bytes)

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
        """
        สกัดรูปภาพประกอบจากหน้า PDF โดย:
        - เรนเดอร์ด้วย page.get_pixmap(clip=rect) เพื่อให้ได้สีและ Alpha/Soft Mask ที่ถูกต้อง ป้องกันปัญหากล่องดำ QR Code
        - กรองรูปสแกนทั้งหน้า (> 60% ของพื้นที่) และสเปเซอร์ขนาดจิ๋วออก
        - จัดเรียงรูปภาพตามพิกัดสายตาจริงจากบนลงล่าง และจากซ้ายไปขวา (y0, x0)
        """
        extracted = []
        image_list = page.get_images(full=True)
        seen_rects = set()
        page_area = page.rect.width * page.rect.height

        for img_info in image_list:
            xref = img_info[0]
            rects = page.get_image_rects(xref)

            if rects:
                for rect in rects:
                    # ป้องกันการบันทึกภาพซ้ำซ้อนที่พิกัดเดียวกัน
                    r_key = (round(rect.x0, 1), round(rect.y0, 1), round(rect.x1, 1), round(rect.y1, 1))
                    if r_key in seen_rects:
                        continue
                    seen_rects.add(r_key)

                    # กรองภาพขนาดเล็กมากๆ เช่น เส้นประ, ไอคอนตกแต่ง 1x1 หรือ spacer
                    if rect.width < 15 or rect.height < 15:
                        continue

                    # กรองภาพสแกนเอกสารทั้งหน้า (> 60% ของพื้นที่หน้ากระดาษ)
                    img_area = rect.width * rect.height
                    if page_area > 0 and (img_area / page_area) > 0.60:
                        continue

                    try:
                        # เรนเดอร์ภาพจากหน้าเอกสารโดยตรงเพื่อรักษา Alpha Channel, Soft Mask และสีที่ถูกต้อง 100%
                        # (แก้ปัญหารหัส QR Code หรือภาพโปร่งใสกลายเป็นกล่องดำทึบ)
                        pix = page.get_pixmap(clip=rect, dpi=200)
                        img_bytes = pix.tobytes(output="png")
                        bbox = (rect.x0, rect.y0, rect.x1, rect.y1)

                        extracted.append(
                            ExtractedImage(
                                page_num=page_num,
                                image_index=len(extracted) + 1,
                                image_bytes=img_bytes,
                                ext="png",
                                width=pix.width,
                                height=pix.height,
                                bbox=bbox,
                            )
                        )
                    except Exception:
                        # Fallback กรณีเรนเดอร์ล้มเหลว
                        base_image = doc.extract_image(xref)
                        extracted.append(
                            ExtractedImage(
                                page_num=page_num,
                                image_index=len(extracted) + 1,
                                image_bytes=base_image.get("image"),
                                ext=base_image.get("ext", "png"),
                                width=base_image.get("width", 0),
                                height=base_image.get("height", 0),
                                bbox=(rect.x0, rect.y0, rect.x1, rect.y1),
                            )
                        )
            else:
                # กรณีไม่มี rect บนหน้า (ภาพที่ไม่ถูกวาดตรงๆ บนเลเยอร์หน้า)
                base_image = doc.extract_image(xref)
                width = base_image.get("width", 0)
                height = base_image.get("height", 0)
                if width < 40 or height < 40:
                    continue
                if len(image_list) == 1 and width >= 700 and height >= 700:
                    continue

                extracted.append(
                    ExtractedImage(
                        page_num=page_num,
                        image_index=len(extracted) + 1,
                        image_bytes=base_image.get("image"),
                        ext=base_image.get("ext", "png"),
                        width=width,
                        height=height,
                        bbox=None,
                    )
                )

        # จัดเรียงรูปภาพตามลำดับการอ่านจริงจากบนลงล่าง และจากซ้ายไปขวา (y0, x0)
        extracted.sort(
            key=lambda img: (
                img.bbox[1] if img.bbox else 9999.0,
                img.bbox[0] if img.bbox else 9999.0,
            )
        )

        # รันหมายเลขลำดับรูปภาพใหม่ให้ตรงกับลำดับบนลงล่าง
        for idx, img in enumerate(extracted, start=1):
            img.image_index = idx

        return extracted
