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
    vector_tables: Optional[List[dict]] = None
    width_pt: Optional[float] = None
    height_pt: Optional[float] = None
    header_hint: Optional[str] = None
    footer_hint: Optional[str] = None


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
                orig_w, orig_h = img.size
                width_pt = round(orig_w * 72.0 / self.dpi, 1)
                height_pt = round(orig_h * 72.0 / self.dpi, 1)

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
                    width_pt=width_pt,
                    height_pt=height_pt,
                )

        # กรณีไฟล์ต้นทางเป็น PDF
        with fitz.open(self.file_path) as doc:
            if page_num < 0 or page_num >= len(doc):
                raise IndexError(f"หมายเลขหน้า {page_num} อยู่นอกช่วง (มีทั้งหมด {len(doc)} หน้า)")

            page = doc[page_num]
            p_width_pt = round(page.rect.width, 1)
            p_height_pt = round(page.rect.height, 1)

            # 1. ตรวจสอบว่าหน้านี้มีข้อความแบบ digital หรือไม่ พร้อมตรวจจับ Header/Footer พื้นที่บน-ล่าง
            raw_text = page.get_text()
            has_text = len(raw_text.strip()) > 0
            header_hint = None
            footer_hint = None

            try:
                blocks = page.get_text("blocks")
                header_candidates = []
                footer_candidates = []
                for b in blocks:
                    # b format: (x0, y0, x1, y1, text, block_no, block_type)
                    if len(b) >= 7 and b[6] == 0:  # text block
                        txt = b[4].strip()
                        if not txt:
                            continue
                        y0, y1 = b[1], b[3]
                        # Top zone: y0 < 54 pt (0.75 นิ้วจากขอบบน)
                        if y0 < 54:
                            header_candidates.append(txt.replace("\n", " "))
                        # Bottom zone: y1 > p_height_pt - 54 pt (0.75 นิ้วจากขอบล่าง)
                        elif y1 > p_height_pt - 54:
                            footer_candidates.append(txt.replace("\n", " "))

                if header_candidates:
                    header_hint = " | ".join(header_candidates)
                if footer_candidates:
                    footer_hint = " | ".join(footer_candidates)
            except Exception:
                pass

            # 2. เรนเดอร์หน้าเป็นภาพความละเอียดสูง 200 DPI (JPEG 92% คมชัดสูง ไฟล์เล็ก อัปโหลดเร็วขึ้น 10 เท่า)
            zoom = self.dpi / 72.0
            matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            rendered_image_bytes = pix.tobytes(output="jpeg", jpg_quality=92)

            if enhance:
                rendered_image_bytes = enhance_document_image(rendered_image_bytes)

            # 3. สกัดรูปภาพที่ฝังอยู่ในหน้า
            embedded_images = self._extract_images_from_page(doc, page, page_num)

            # 4. สกัดพิกัดตารางจริงจากเวกเตอร์ (Vector Table Geometry)
            vector_tables = []
            try:
                tabs = page.find_tables()
                for t in tabs.tables:
                    bbox = t.bbox
                    t_w = bbox[2] - bbox[0]
                    t_h = bbox[3] - bbox[1]
                    if t_w > 40 and t_h > 15 and t.col_count >= 1:
                        vector_tables.append({
                            "col_count": t.col_count,
                            "row_count": t.row_count,
                            "bbox": bbox,
                            "width_pt": t_w,
                            "height_pt": t_h,
                        })
            except Exception:
                vector_tables = []

            return PDFPageData(
                page_num=page_num,
                rendered_image_bytes=rendered_image_bytes,
                embedded_images=embedded_images,
                has_text=has_text,
                vector_tables=vector_tables,
                width_pt=p_width_pt,
                height_pt=p_height_pt,
                header_hint=header_hint,
                footer_hint=footer_hint,
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

        # 1. รวบรวม Bounding Box ของภาพทั้งหมดบนหน้า
        raw_rects = []
        for img_info in image_list:
            xref = img_info[0]
            rects = page.get_image_rects(xref)
            for r in rects:
                # กรองภาพขนาดเล็กมากๆ เช่น เส้นประ, spacer
                if r.width < 12 or r.height < 12:
                    continue
                # กรองภาพสแกนเอกสารทั้งหน้า (> 60% ของพื้นที่หน้ากระดาษ)
                if page_area > 0 and ((r.width * r.height) / page_area) > 0.60:
                    continue
                raw_rects.append(r)

        # 2. รวมพิกัดรูปภาพที่ทับซ้อนหรือเป็นกราฟิกชิ้นเดียวกัน (Overlapping Graphic Merging)
        # เช่น ไอคอนวงกลมตัวเลขสีแดงที่วางทับบนภาพกล่องพัสดุขั้นตอนที่ 1, 2, 3
        merged_rects = []
        for r in raw_rects:
            found = False
            for i, m in enumerate(merged_rects):
                if m.intersects(r) or (abs(m.y1 - r.y0) < 10 and abs(m.x0 - r.x0) < 30):
                    merged_rects[i] = m | r
                    found = True
                    break
            if not found:
                merged_rects.append(r)

        # 3. จัดเรียงรูปภาพตามระดับสายตาจริง (Row Band Clustering) จากบนลงล่าง และจากซ้ายไปขวา
        merged_rects.sort(key=lambda r: (round(r.y0 / 25.0), r.x0))

        # 4. สกัดภาพด้วย Pixmap ตามพิกัดที่รวมแล้ว เพื่อรักษา Alpha Mask และความคมชัด 100%
        for idx, rect in enumerate(merged_rects, start=1):
            try:
                pix = page.get_pixmap(clip=rect, dpi=200)
                img_bytes = pix.tobytes(output="png")
                extracted.append(
                    ExtractedImage(
                        page_num=page_num,
                        image_index=idx,
                        image_bytes=img_bytes,
                        ext="png",
                        width=pix.width,
                        height=pix.height,
                        bbox=(rect.x0, rect.y0, rect.x1, rect.y1),
                    )
                )
            except Exception:
                pass

        # กรณีภาพที่ไม่ถูกวาดตรงๆ บนเลเยอร์หน้า (Fallback images without rects)
        if not extracted:
            for img_info in image_list:
                xref = img_info[0]
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

        return extracted
