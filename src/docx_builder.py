import io
import re
from pathlib import Path
from typing import List, Optional, Tuple
import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

from .config import (
    DEFAULT_FONT,
    FONT_SIZE_BODY,
    FONT_SIZE_H1,
    FONT_SIZE_H2,
    FONT_SIZE_H3,
    MAX_DOCX_IMAGE_WIDTH_INCHES,
    get_font_sizes,
)
from .pdf_processor import ExtractedImage


def set_run_font(run, font_name: str = DEFAULT_FONT, size_pt: float = FONT_SIZE_BODY, bold: bool = False, italic: bool = False):
    """
    กำหนดแบบอักษรให้รองรับภาษาไทยใน OpenXML อย่างสมบูรณ์
    โดยตั้งค่าทั้ง ascii, hAnsi และ cs (Complex Script) เพื่อให้เปิดบน Word ทุกเครื่องได้ฟอนต์ถูกต้อง
    """
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic

    # กำหนดค่า XML element โดยตรงสำหรับ Complex Script (ภาษาไทย)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(
            f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>'
        )
        rPr.append(rFonts)
    else:
        rFonts.set(qn("w:ascii"), font_name)
        rFonts.set(qn("w:hAnsi"), font_name)
        rFonts.set(qn("w:cs"), font_name)


def set_cell_background(cell, fill_hex: str):
    """ตั้งสีพื้นหลังของเซลล์ตาราง (เช่น 'F2F2F2')"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_table_margins(table, top: int = 100, bottom: int = 100, left: int = 150, right: int = 150):
    """ตั้งระยะขอบด้านในของเซลล์ตาราง (Cell Margins/Padding)"""
    tblPr = table._tbl.tblPr
    tblCellMar = parse_xml(
        f'<w:tblCellMar {nsdecls("w")}>'
        f'  <w:top w:w="{top}" w:type="dxa"/>'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'  <w:left w:w="{left}" w:type="dxa"/>'
        f'  <w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tblCellMar>'
    )
    tblPr.append(tblCellMar)


class DocxBuilder:
    """คลาสสร้างเอกสาร Word (.docx) จาก Markdown ที่ได้จาก Gemini พร้อมแทรกตาราง รูปภาพ และการจัดหน้าตามต้นฉบับ"""

    def __init__(self, font_name: str = DEFAULT_FONT):
        self.font_name = font_name
        self.font_sizes = get_font_sizes(self.font_name)
        self.doc = docx.Document()
        self._setup_document_styles()

    def set_font(self, font_name: str):
        """เปลี่ยนหรืออัปเดตฟอนต์ของเอกสาร (เช่น เมื่อตรวจพบฟอนต์จากต้นฉบับ)"""
        if font_name and font_name.strip():
            self.font_name = font_name.strip()
            self.font_sizes = get_font_sizes(self.font_name)
            self._setup_document_styles()

    def _setup_document_styles(self):
        """ตั้งค่าสไตล์เริ่มต้นของเอกสาร ขอบกระดาษ 0.5 นิ้ว และระยะบรรทัดภาษาไทยให้กระชับลงตัว"""
        # ปรับระยะขอบหน้ากระดาษเป็น 0.5 นิ้ว เพื่อให้พื้นที่พิมพ์กว้างขึ้น เหมาะกับฟอร์มและเอกสารทั่วไป
        for section in self.doc.sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        # สไตล์ Normal
        style_normal = self.doc.styles["Normal"]
        style_normal.font.name = self.font_name
        style_normal.font.size = Pt(self.font_sizes["body"])
        style_normal.paragraph_format.line_spacing = 1.15
        style_normal.paragraph_format.space_after = Pt(2)

    def parse_line_formatting(self, raw_line: str) -> Tuple[str, Optional[WD_ALIGN_PARAGRAPH], bool]:
        """
        ตรวจจับแท็กการจัดตำแหน่ง [CENTER], [RIGHT], [JUSTIFY] และแท็กขนาด [SMALL]
        คืนค่า: (cleaned_line, alignment, is_small)
        """
        alignment = None
        is_small = False
        cleaned = raw_line.strip()

        # ตรวจสอบแท็กขนาดเล็ก
        if "[SMALL]" in cleaned.upper() or "[/SMALL]" in cleaned.upper():
            is_small = True
            cleaned = re.sub(r"\[/?SMALL\]", "", cleaned, flags=re.IGNORECASE).strip()

        # ตรวจสอบแท็กจัดตำแหน่ง
        if "[CENTER]" in cleaned.upper() or "[/CENTER]" in cleaned.upper():
            alignment = WD_ALIGN_PARAGRAPH.CENTER
            cleaned = re.sub(r"\[/?CENTER\]", "", cleaned, flags=re.IGNORECASE).strip()
        elif "[RIGHT]" in cleaned.upper() or "[/RIGHT]" in cleaned.upper():
            alignment = WD_ALIGN_PARAGRAPH.RIGHT
            cleaned = re.sub(r"\[/?RIGHT\]", "", cleaned, flags=re.IGNORECASE).strip()
        elif "[JUSTIFY]" in cleaned.upper() or "[/JUSTIFY]" in cleaned.upper():
            alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            cleaned = re.sub(r"\[/?JUSTIFY\]", "", cleaned, flags=re.IGNORECASE).strip()

        return cleaned, alignment, is_small

    def add_page_content(
        self,
        markdown_text: str,
        page_num: int,
        images: Optional[List[ExtractedImage]] = None,
        is_first_page: bool = False,
    ):
        """
        แปลงเนื้อหา Markdown ของหนึ่งหน้าลงในเอกสาร Word
        พร้อมแทรกรูปภาพที่สกัดได้จากหน้านั้นๆ และจัดรูปแบบตำแหน่งตามต้นฉบับ
        """
        if not is_first_page:
            # เพิ่มการขึ้นหน้าใหม่ตามต้นฉบับ PDF
            self.doc.add_page_break()

        images_queue = list(images) if images else []
        lines = markdown_text.splitlines()

        i = 0
        while i < len(lines):
            raw_line = lines[i].strip()

            # 1. บรรทัดว่าง
            if not raw_line:
                i += 1
                continue

            # 2. ตรวจสอบแท็กรูปภาพ [IMAGE]
            if "[IMAGE]" in raw_line.upper():
                if images_queue:
                    img = images_queue.pop(0)
                    self._insert_image(img)
                i += 1
                continue

            # 3. ตรวจสอบตาราง (Markdown Table)
            if raw_line.startswith("|") and ("|" in raw_line[1:]):
                table_lines = []
                while i < len(lines) and lines[i].strip().startswith("|") and ("|" in lines[i].strip()[1:]):
                    table_lines.append(lines[i].strip())
                    i += 1
                self._create_word_table(table_lines, images_queue=images_queue)
                continue

            # ตรวจสอบและแยกแท็กจัดตำแหน่งและขนาดก่อน
            line, alignment, is_small = self.parse_line_formatting(raw_line)
            if not line:
                i += 1
                continue

            # 4. ตรวจสอบหัวข้อ (Headings)
            if line.startswith("### "):
                p = self.doc.add_paragraph()
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.keep_with_next = True
                if alignment:
                    p.alignment = alignment
                font_sz = self.font_sizes["h3"] if not is_small else self.font_sizes["small"]
                self._add_formatted_text_to_paragraph(p, line[4:].strip(), size_pt=font_sz, bold=True)
                i += 1
                continue
            elif line.startswith("## "):
                p = self.doc.add_paragraph()
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.keep_with_next = True
                if alignment:
                    p.alignment = alignment
                font_sz = self.font_sizes["h2"] if not is_small else self.font_sizes["small"]
                self._add_formatted_text_to_paragraph(p, line[3:].strip(), size_pt=font_sz, bold=True)
                i += 1
                continue
            elif line.startswith("# "):
                p = self.doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.keep_with_next = True
                if alignment:
                    p.alignment = alignment
                font_sz = self.font_sizes["h1"] if not is_small else self.font_sizes["small"]
                self._add_formatted_text_to_paragraph(p, line[2:].strip(), size_pt=font_sz, bold=True)
                i += 1
                continue

            # 5. ตรวจสอบรายการหัวข้อย่อย (Bullet Lists)
            if line.startswith(("- ", "* ", "• ")):
                p = self.doc.add_paragraph(style="List Bullet")
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.15
                if alignment:
                    p.alignment = alignment
                font_sz = self.font_sizes["body"] if not is_small else self.font_sizes["small"]
                self._add_formatted_text_to_paragraph(p, line[2:].strip(), size_pt=font_sz)
                i += 1
                continue

            # 6. ตรวจสอบรายการตัวเลข (Numbered Lists)
            num_match = re.match(r"^(\d+[\.\)])\s*(.*)$", line)
            if num_match:
                p = self.doc.add_paragraph(style="List Number")
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.15
                if alignment:
                    p.alignment = alignment
                font_sz = self.font_sizes["body"] if not is_small else self.font_sizes["small"]
                self._add_formatted_text_to_paragraph(p, num_match.group(2).strip(), size_pt=font_sz)
                i += 1
                continue

            # 7. ย่อหน้าปกติ (Paragraph)
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            if alignment:
                p.alignment = alignment
                if alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                    p.paragraph_format.first_line_indent = Inches(0.4)
            elif len(line) > 80 and not any(line.startswith(pfx) for pfx in ["ที่ ", "เรื่อง ", "เรียน ", "ถึง ", "จาก ", "วันที่ ", "เอกสาร "]):
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Inches(0.4)

            font_sz = self.font_sizes["body"] if not is_small else self.font_sizes["small"]
            self._add_formatted_text_to_paragraph(p, line, size_pt=font_sz)
            i += 1

    def _add_formatted_text_to_paragraph(
        self,
        paragraph,
        text: str,
        size_pt: Optional[float] = None,
        bold: bool = False,
        italic: bool = False,
    ):
        """แยกแท็กตัวหนา (**ข้อความ**) และตัวเอียง (*ข้อความ*) ออกมาใส่ใน Run พร้อมกำหนดขนาด"""
        from .gemini_extractor import clean_thai_ocr_text
        text = clean_thai_ocr_text(text)
        # ปรับทอนเส้นใต้ลายเซ็นที่ยาวเกินไป ไม่ให้ล้นและตกบรรทัด
        text = re.sub(r"_{20,}", "____________________", text)
        if "✂" in text:
            text = re.sub(r"[-–—_]{30,}", "--------------------------------------------------", text)

        actual_size = size_pt if size_pt is not None else self.font_sizes["body"]
        tokens = re.split(r"(\*\*.*?\*\*|\*.*?\*)", text)
        for token in tokens:
            if not token:
                continue
            if token.startswith("**") and token.endswith("**") and len(token) >= 4:
                run = paragraph.add_run(token[2:-2])
                set_run_font(run, self.font_name, actual_size, bold=True, italic=italic)
            elif token.startswith("*") and token.endswith("*") and len(token) >= 2:
                run = paragraph.add_run(token[1:-1])
                set_run_font(run, self.font_name, actual_size, bold=bold, italic=True)
            else:
                run = paragraph.add_run(token)
                set_run_font(run, self.font_name, actual_size, bold=bold, italic=italic)

    def _create_word_table(self, table_lines: List[str], images_queue: Optional[List[ExtractedImage]] = None):
        """แปลงตาราง Markdown เป็น Table Object ของ Word พร้อมจัดขอบตาราง ถอดรหัสแท็ก และแทรกรูปในเซลล์"""
        if not table_lines:
            return

        parsed_rows = []
        for line in table_lines:
            # ตัด | ตัวแรกและตัวสุดท้ายออก
            content = line.strip()
            if content.startswith("|"):
                content = content[1:]
            if content.endswith("|"):
                content = content[:-1]

            cols = [c.strip() for c in content.split("|")]

            # ตรวจสอบว่าเป็นบรรทัดขีดคั่นหัวตารางหรือไม่ เช่น |---|---|
            is_separator = all(re.match(r"^:?-+:?$", c) for c in cols if c)
            if is_separator:
                continue

            parsed_rows.append(cols)

        if not parsed_rows:
            return

        # คำนวณจำนวนคอลัมน์สูงสุด
        num_cols = max(len(r) for r in parsed_rows)
        num_rows = len(parsed_rows)

        table = self.doc.add_table(rows=num_rows, cols=num_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        # ขอบเซลล์แบบกะทัดรัด (Compact cell padding)
        set_table_margins(table, top=60, bottom=60, left=100, right=100)

        # คำนวณขนาดตัวอักษรสำหรับตารางให้กะทัดรัด พอดีกับช่องเอกสาร
        table_font_size = max(9.0, self.font_sizes["body"] - 3.0)
        header_font_size = max(10.0, self.font_sizes["body"] - 2.0)

        for row_idx, row_data in enumerate(parsed_rows):
            is_header = (row_idx == 0 and len(parsed_rows) > 1)
            row = table.rows[row_idx]

            # กำหนดคุณสมบัติไม่ให้แถวแตกข้ามหน้าถ้าไม่จำเป็น
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            # ถ้าเป็นหัวตาราง ให้วนซ้ำหัวตารางเมื่อขึ้นหน้าใหม่ (Repeat Header Row)
            if is_header:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

            for col_idx in range(num_cols):
                cell = row.cells[col_idx]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                raw_cell_content = row_data[col_idx] if col_idx < len(row_data) else ""
                cell.text = ""  # เคลียร์พารากราฟเริ่มต้น

                # แยกหลายบรรทัดในเซลล์ด้วย <br> หรือ \n
                cell_lines = re.split(r"<br\s*/?>|\n", raw_cell_content, flags=re.IGNORECASE)
                while cell_lines and not cell_lines[0].strip():
                    cell_lines.pop(0)
                while cell_lines and not cell_lines[-1].strip():
                    cell_lines.pop()

                if not cell_lines:
                    cell_lines = [""]

                for line_idx, cell_line in enumerate(cell_lines):
                    line_str = cell_line.strip()
                    if line_idx == 0:
                        p = cell.paragraphs[0]
                    else:
                        p = cell.add_paragraph()

                    p.paragraph_format.space_before = Pt(1)
                    p.paragraph_format.space_after = Pt(1)
                    p.paragraph_format.line_spacing = 1.05

                    # ตรวจสอบแท็กรูปภาพ [IMAGE] ในเซลล์ตาราง (เช่น บาร์โค้ดหรือโลโก้)
                    if "[IMAGE]" in line_str.upper():
                        if images_queue:
                            img = images_queue.pop(0)
                            self._insert_image_to_paragraph(p, img, max_width_inches=3.0)
                        continue

                    # แยกแท็กจัดตำแหน่ง [CENTER], [RIGHT], [SMALL]
                    clean_text, align, is_small = self.parse_line_formatting(line_str)

                    if is_small:
                        curr_size = self.font_sizes["small"]
                    elif is_header:
                        curr_size = header_font_size
                    else:
                        curr_size = table_font_size

                    # กำหนดการจัดตำแหน่งในเซลล์
                    if align:
                        p.alignment = align
                    elif is_header:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif re.match(r"^[\$฿€¥]?\s*[\d,]+(\.\d+)?%?$", clean_text):
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    if is_header:
                        set_cell_background(cell, "F2F2F2")
                        self._add_formatted_text_to_paragraph(p, clean_text, size_pt=curr_size, bold=True)
                    else:
                        self._add_formatted_text_to_paragraph(p, clean_text, size_pt=curr_size)


    def _insert_image_to_paragraph(self, p, img: ExtractedImage, max_width_inches: float = 6.0):
        """แทรกรูปภาพลงใน Paragraph ที่ระบุ พร้อมจำกัดความกว้างสูงสุด"""
        try:
            image_stream = io.BytesIO(img.image_bytes)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            if img.bbox and len(img.bbox) == 4:
                bbox_w_pt = abs(img.bbox[2] - img.bbox[0])
                width_in_inches = min(bbox_w_pt / 72.0, max_width_inches)
                width_in_inches = max(0.4, width_in_inches)
            else:
                width_in_inches = min(img.width / 150.0, max_width_inches)
                width_in_inches = max(0.4, width_in_inches)
            run.add_picture(image_stream, width=Inches(width_in_inches))
        except Exception:
            run = p.add_run(f"[ภาพประกอบ {img.image_index}]")
            set_run_font(run, self.font_name, self.font_sizes["body"], italic=True)

    def _insert_image(self, img: ExtractedImage):
        """แทรกรูปภาพลงในเอกสาร Word โดยรักษาขนาดจริง 1:1 และตำแหน่งชิดซ้าย/กลาง/ขวา"""
        try:
            image_stream = io.BytesIO(img.image_bytes)
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            # 1. คำนวณขนาดจริงและความกว้าง (Physical sizing 1:1)
            if img.bbox and len(img.bbox) == 4:
                # PDF points: 72 points = 1 inch
                bbox_w_pt = abs(img.bbox[2] - img.bbox[0])
                width_in_inches = bbox_w_pt / 72.0
                # จำกัดขนาดความกว้างระหว่าง 0.4 นิ้ว ถึงค่าสูงสุดหน้ากระดาษ
                width_in_inches = max(0.4, min(width_in_inches, MAX_DOCX_IMAGE_WIDTH_INCHES))

                # 2. คำนวณการจัดตำแหน่งแนวนอนจาก Bounding Box
                # หน้า A4 กว้าง ~595 pt (ขอบซ้าย 72 pt, ขอบขวา 523 pt)
                center_x = (img.bbox[0] + img.bbox[2]) / 2.0
                if center_x < 220.0:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                elif center_x > 375.0:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                # กรณีไม่มี bbox (เช่น ไฟล์รูปภาพเดี่ยว หรือ fallback)
                width_in_inches = min(img.width / 150.0, MAX_DOCX_IMAGE_WIDTH_INCHES)
                width_in_inches = max(0.6, width_in_inches)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            run = p.add_run()
            run.add_picture(image_stream, width=Inches(width_in_inches))
        except Exception as e:
            # หากแทรกรูปภาพล้มเหลว ให้ใส่ข้อความเตือนแทนโดยไม่ให้กระบวนการหลักหยุดชะงัก
            p = self.doc.add_paragraph()
            run = p.add_run(f"[ภาพประกอบ {img.image_index}]")
            set_run_font(run, self.font_name, self.font_sizes["body"], italic=True)

    def save(self, output_path: str | Path):
        """บันทึกเอกสาร Word ลงในไฟล์เป้าหมาย"""
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(path))
