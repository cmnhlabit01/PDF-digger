import io
import json
import logging
import re
from pathlib import Path
from typing import List, Optional, Tuple
import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

logger = logging.getLogger(__name__)

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


def set_run_font(run, font_name: str = DEFAULT_FONT, size_pt: float = FONT_SIZE_BODY, bold: bool = False, italic: bool = False, color_rgb: Optional[RGBColor] = None):
    """
    กำหนดแบบอักษรให้รองรับภาษาไทยใน OpenXML อย่างสมบูรณ์
    โดยตั้งค่าทั้ง ascii, hAnsi และ cs (Complex Script) เพื่อให้เปิดบน Word ทุกเครื่องได้ฟอนต์ถูกต้อง
    """
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb is not None:
        run.font.color.rgb = color_rgb

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


def set_table_borderless(table):
    """ลบเส้นขอบตารางออกทั้งหมด สำหรับตารางจัดเค้าโครงหน้า (Layout Table ไร้ขอบ)"""
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="none"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:bottom w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideH w:val="none"/>'
        f'  <w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)


def set_table_horizontal_borders(table):
    """กำหนดเส้นขอบตารางเฉพาะแนวนอน (ไม่มีเส้นแนวตั้ง) สไตล์รายงานผลแล็บ/วิจัย/งบการเงิน"""
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        f'  <w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)


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
        """ตั้งค่าหน้ากระดาษ A4 เริ่มต้น ขอบกระดาษ 0.5 นิ้ว และระยะบรรทัดภาษาไทยให้กระชับลงตัว"""
        for section in self.doc.sections:
            self._apply_page_dimensions(section, 595.3, 841.9)

    def _apply_page_dimensions(self, section, width_pt: float, height_pt: float):
        """กำหนดขนาดหน้ากระดาษ ทิศทาง และขอบกระดาษที่เหมาะสมสำหรับ Section"""
        section.page_width = Pt(width_pt)
        section.page_height = Pt(height_pt)

        # กำหนดทิศทางกระดาษ (Landscape vs Portrait)
        if width_pt > height_pt:
            section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
        else:
            section.orientation = docx.enum.section.WD_ORIENT.PORTRAIT

        # กำหนดขอบกระดาษอย่างชาญฉลาด (Smart Margins)
        # หากเป็นฉลาก/ใบเสร็จขนาดเล็ก (กว้าง < 350 pt หรือ ~4.8 นิ้ว) ใช้ขอบ 0.25 นิ้วเพื่อประหยัดพื้นที่
        if width_pt < 350:
            section.top_margin = Inches(0.25)
            section.bottom_margin = Inches(0.25)
            section.left_margin = Inches(0.25)
            section.right_margin = Inches(0.25)
        else:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

    def _get_current_printable_width_in(self) -> float:
        """คำนวณพื้นที่พิมพ์จริงแนวนอน (ความกว้างหน้ากระดาษ - ขอบซ้ายขวา) ของ Section ปัจจุบัน"""
        try:
            sec = self.doc.sections[-1]
            w = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches
            if w > 1.0:
                return round(w, 2)
        except Exception:
            pass
        return 7.27

        # สไตล์ Normal
        style_normal = self.doc.styles["Normal"]
        style_normal.font.name = self.font_name
        style_normal.font.size = Pt(self.font_sizes["body"])
        style_normal.paragraph_format.line_spacing = 1.10
        style_normal.paragraph_format.space_after = Pt(1.5)

    def parse_line_formatting(self, raw_line: str) -> Tuple[str, Optional[WD_ALIGN_PARAGRAPH], bool]:
        """
        ตรวจจับแท็กการจัดตำแหน่ง [CENTER], [RIGHT], [JUSTIFY] และแท็กขนาด [SMALL]
        คืนค่า: (cleaned_line, alignment, is_small)
        """
        alignment = None
        is_small = False
        cleaned = raw_line.strip()

        # ถอดรหัส HTML entities และลบแท็ก <hr>
        cleaned = cleaned.replace("&nbsp;", " ")
        cleaned = re.sub(r"</?hr\s*/?>", "", cleaned, flags=re.IGNORECASE).strip()

        # ตรวจสอบแท็กขนาดเล็ก
        if "[SMALL]" in cleaned.upper() or "[/SMALL]" in cleaned.upper():
            is_small = True
            cleaned = re.sub(r"\[/?SMALL\]", "", cleaned, flags=re.IGNORECASE).strip()

        # ลบแท็ก [BADGE] หากมีในข้อความทั่วไป
        if "[BADGE]" in cleaned.upper() or "[/BADGE]" in cleaned.upper():
            cleaned = re.sub(r"\[/?BADGE\]", "", cleaned, flags=re.IGNORECASE).strip()

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
        page_width_pt: Optional[float] = None,
        page_height_pt: Optional[float] = None,
    ):
        """
        แปลงเนื้อหาของหนึ่งหน้าลงในเอกสาร Word (รองรับทั้ง Structural Layout JSON และ Markdown)
        พร้อมแทรกรูปภาพที่สกัดได้จากหน้านั้นๆ ปรับขนาดหน้ากระดาษตามต้นฉบับ และแยก Header/Footer อัตโนมัติ
        """
        if is_first_page:
            if page_width_pt and page_height_pt:
                self._apply_page_dimensions(self.doc.sections[0], page_width_pt, page_height_pt)
        else:
            curr_sec = self.doc.sections[-1]
            needs_new_section = False
            if page_width_pt and page_height_pt:
                prev_w = curr_sec.page_width.pt
                prev_h = curr_sec.page_height.pt
                # หากขนาดต่างกันมากกว่า 2 pt หรือทิศทางเปลี่ยน ให้เปิด Section ใหม่
                if abs(prev_w - page_width_pt) > 2.0 or abs(prev_h - page_height_pt) > 2.0:
                    needs_new_section = True

            if needs_new_section:
                new_sec = self.doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
                self._apply_page_dimensions(new_sec, page_width_pt, page_height_pt)
                new_sec.header.is_linked_to_previous = False
                new_sec.footer.is_linked_to_previous = False
            else:
                self.doc.add_page_break()

        images_queue = list(images) if images else []

        # 1. ตรวจสอบและแปลง Structural Layout JSON (รองรับทั้ง Single Object, Array [...], และ Nested)
        blocks, page_type = self._parse_json_blocks(markdown_text)
        if blocks:
            self._render_structural_blocks(blocks, images_queue, page_type=page_type)
        else:
            # 2. ป้องกัน JSON รั่วไหล (Anti-Leak Guard):
            # หากข้อความเป็นโค้ด JSON ที่ชำรุด ห้ามส่งให้ _render_markdown_lines บรรทัดต่อบรรทัดเด็ดขาด
            if self._is_raw_json_code(markdown_text):
                logger.warning(f"[Page {page_num}] ตรวจพบโค้ด JSON ไม่สมบูรณ์ กำลังสกัดเฉพาะข้อความเพื่อป้องกันโค้ดหลุดลง Word")
                safe_lines = self._extract_text_lines_from_broken_json(markdown_text)
                self._render_markdown_lines(safe_lines, images_queue)
            else:
                self._render_markdown_lines(markdown_text.splitlines(), images_queue)

    def _parse_json_blocks(self, text: str) -> Tuple[Optional[List[dict]], str]:
        """
        พยายามถอดรหัสและแปลงข้อความเป็น Structural Blocks:
        - รองรับ Single Object: {"page_type": "...", "blocks": [...]}
        - รองรับ Array of Object: [{"page_type": "...", "blocks": [...]}]
        - รองรับ Array of Blocks: [{"type": "table", ...}, {"type": "paragraph", ...}]
        - รองรับ Markdown code block ```json ... ```
        - ซ่อมแซม trailing commas และข้อความนอก JSON อัตโนมัติ
        """
        if not text:
            return None, "document"
        clean_str = text.strip()
        if clean_str.startswith("```"):
            clean_str = re.sub(r"^```(?:json)?\s*", "", clean_str)
            clean_str = re.sub(r"\s*```$", "", clean_str).strip()

        data = None
        try:
            data = json.loads(clean_str)
        except Exception:
            for open_ch, close_ch in [("{", "}"), ("[", "]")]:
                first_i = clean_str.find(open_ch)
                last_i = clean_str.rfind(close_ch)
                if first_i != -1 and last_i != -1 and last_i > first_i:
                    sub = clean_str[first_i : last_i + 1]
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
            return None, "document"

        if isinstance(data, list):
            if not data:
                return None, "document"
            if isinstance(data[0], dict) and "blocks" in data[0] and isinstance(data[0]["blocks"], list):
                return data[0]["blocks"], data[0].get("page_type", "document")
            if any(isinstance(x, dict) and "type" in x for x in data):
                return [x for x in data if isinstance(x, dict)], "document"
        elif isinstance(data, dict):
            if "blocks" in data and isinstance(data["blocks"], list):
                return data["blocks"], data.get("page_type", "document")
            for val in data.values():
                if isinstance(val, dict) and "blocks" in val and isinstance(val["blocks"], list):
                    return val["blocks"], val.get("page_type", "document")

        return None, "document"

    def _is_raw_json_code(self, text: str) -> bool:
        """ตรวจสอบว่าข้อความดูเหมือนเป็นโค้ด JSON ที่ไม่สมบูรณ์หรือไม่ เพื่อป้องกันไม่ให้นำไปเรนเดอร์เป็นย่อหน้าใน Word"""
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

    def _extract_text_lines_from_broken_json(self, text: str) -> List[str]:
        """สกัดเฉพาะเนื้อหาข้อความจาก JSON ที่ชำรุด แทนที่จะปล่อยให้โค้ด JSON หลุดลงหน้าเอกสาร"""
        from .gemini_extractor import clean_thai_ocr_text
        lines = []
        matches = re.findall(r'"text"\s*:\s*"((?:\\.|[^"\\])*)"', text)
        for m in matches:
            val = m.replace('\\"', '"').replace('\\n', '\n').replace('\\t', ' ').strip()
            val = clean_thai_ocr_text(val)
            if val and not self._is_raw_json_code(val):
                lines.append(val)
        return lines

    def _render_structural_blocks(self, blocks: list, images_queue: List[ExtractedImage], page_type: str = "document"):
        """เรนเดอร์เนื้อหาจาก Structural Layout Blocks ที่มีสัดส่วนความกว้าง-สูง และการรวมเซลล์แบบแม่นยำ"""
        for block in blocks:
            b_type = str(block.get("type", "")).lower()
            if b_type == "table":
                self._create_structural_table(block, images_queue=images_queue)
            elif b_type == "heading":
                level = int(block.get("level", 1))
                text = str(block.get("text", "")).strip()
                align = str(block.get("align", "left")).lower()
                p = self.doc.add_paragraph()
                p.paragraph_format.keep_with_next = True
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(1.5)
                if "center" in align:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif "right" in align:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                h_size = self.font_sizes.get(f"h{level}", self.font_sizes["h1"])
                self._add_formatted_text_to_paragraph(p, text, size_pt=h_size, bold=True)
            elif b_type == "paragraph":
                text = str(block.get("text", "")).strip()
                align = str(block.get("align", "left")).lower()
                is_small = bool(block.get("is_small", False))
                p = self.doc.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(1.5)
                p.paragraph_format.line_spacing = 1.10
                if "center" in align:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif "right" in align:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif "justify" in align:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.paragraph_format.first_line_indent = Inches(0.4)
                sz = self.font_sizes["small"] if is_small else self.font_sizes["body"]
                self._add_formatted_text_to_paragraph(p, text, size_pt=sz)
            elif b_type == "list":
                ordered = bool(block.get("ordered", False))
                style = "List Number" if ordered else "List Bullet"
                for item in block.get("items", []):
                    p = self.doc.add_paragraph(style=style)
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(1)
                    p.paragraph_format.line_spacing = 1.10
                    self._add_formatted_text_to_paragraph(p, str(item).strip(), size_pt=self.font_sizes["body"])
            elif b_type == "image":
                if images_queue:
                    img = images_queue.pop(0)
                    self._insert_image(img)
            elif b_type == "header":
                self._render_header_block(block)
            elif b_type == "footer":
                self._render_footer_block(block)

    def _create_structural_table(
        self,
        table_block: dict,
        images_queue: Optional[List[ExtractedImage]] = None,
    ):
        """สร้างตารางตามพิกัดและสัดส่วนจริงของ Structural Table (Exact Widths, Heights, Merging)"""
        rows_data = table_block.get("rows", [])
        if not rows_data:
            return

        num_rows = len(rows_data)
        num_cols = 0
        for r_data in rows_data:
            cells = r_data.get("cells", [])
            cols_in_row = sum(max(1, int(c.get("colspan", 1))) for c in cells)
            if cols_in_row > num_cols:
                num_cols = cols_in_row

        if num_cols == 0 or num_rows == 0:
            return

        total_page_width_in = self._get_current_printable_width_in()
        raw_widths = table_block.get("col_widths_pct") or []
        if raw_widths and len(raw_widths) == num_cols:
            total_w = sum(raw_widths)
            if total_w > 0:
                col_widths = [(w / total_w) * total_page_width_in for w in raw_widths]
            else:
                col_widths = [total_page_width_in / num_cols] * num_cols
        else:
            col_widths = [total_page_width_in / num_cols] * num_cols

        table = self.doc.add_table(rows=num_rows, cols=num_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        border_style = str(table_block.get("border_style", "grid")).lower()
        if "borderless" in border_style or "no_border" in border_style:
            set_table_borderless(table)
        elif "horizontal" in border_style:
            set_table_horizontal_borders(table)
        else:
            table.style = "Table Grid"

        # ขอบเซลล์กะทัดรัด (Padding: 20 dxa บน-ล่าง, 60 dxa ซ้าย-ขวา)
        set_table_margins(table, top=20, bottom=20, left=60, right=60)

        for c_idx, col in enumerate(table.columns):
            if c_idx < len(col_widths):
                col.width = Inches(col_widths[c_idx])

        occupied = [[False for _ in range(num_cols)] for _ in range(num_rows)]
        cells_to_merge = []

        for r_idx, r_data in enumerate(rows_data):
            row_elem = table.rows[r_idx]
            trPr = row_elem._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            height_pt = r_data.get("height_pt")
            if height_pt and height_pt > 0:
                h_twips = int(height_pt * 20)
                trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{h_twips}" w:hRule="atLeast"/>'))

            c_cursor = 0
            for cell_data in r_data.get("cells", []):
                while c_cursor < num_cols and occupied[r_idx][c_cursor]:
                    c_cursor += 1
                if c_cursor >= num_cols:
                    break

                colspan = max(1, int(cell_data.get("colspan", 1)))
                rowspan = max(1, int(cell_data.get("rowspan", 1)))

                for dr in range(rowspan):
                    for dc in range(colspan):
                        if r_idx + dr < num_rows and c_cursor + dc < num_cols:
                            occupied[r_idx + dr][c_cursor + dc] = True

                if colspan > 1 or rowspan > 1:
                    end_r = min(num_rows - 1, r_idx + rowspan - 1)
                    end_c = min(num_cols - 1, c_cursor + colspan - 1)
                    cells_to_merge.append((r_idx, c_cursor, end_r, end_c))

                target_cell = row_elem.cells[c_cursor]
                merged_width = sum(col_widths[c_cursor + dc] for dc in range(colspan) if c_cursor + dc < len(col_widths))
                target_cell.width = Inches(merged_width)

                self._populate_structural_cell(target_cell, cell_data, images_queue, merged_width)
                c_cursor += colspan

        for (sr, sc, er, ec) in cells_to_merge:
            try:
                table.cell(sr, sc).merge(table.cell(er, ec))
            except Exception as e:
                logger.warning(f"Cell merge failed ({sr}, {sc}) -> ({er}, {ec}): {e}")

        # ย่อหน้าคั่นตาราง 1 pt ป้องกันตารางรวมกัน
        sep_p = self.doc.add_paragraph()
        sep_p.paragraph_format.space_before = Pt(0)
        sep_p.paragraph_format.space_after = Pt(0)
        sep_p.paragraph_format.line_spacing = Pt(1)
        run = sep_p.add_run()
        run.font.size = Pt(1)

    def _populate_structural_cell(
        self,
        cell,
        cell_data: dict,
        images_queue: Optional[List[ExtractedImage]],
        cell_width_in: float,
    ):
        """หยอดข้อมูล ข้อความ ป้าย Badge และรูปภาพลงในเซลล์ตารางพร้อมจัดรูปแบบ"""
        valign_str = str(cell_data.get("valign", "center")).lower()
        if "top" in valign_str:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        elif "bottom" in valign_str:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
        else:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        bg_color = cell_data.get("bg_color")
        is_badge = bool(cell_data.get("is_badge", False))
        text_val = str(cell_data.get("text", "")).strip()
        if not is_badge and text_val.upper() in ["W", "RR", "PICK UP", "COD"]:
            is_badge = True

        if is_badge:
            bg_color = bg_color or "595959"
            set_cell_background(cell, bg_color)
        elif bg_color:
            set_cell_background(cell, str(bg_color).lstrip("#"))

        cell.text = ""
        has_image = bool(cell_data.get("has_image", False)) or ("[IMAGE]" in text_val.upper())

        from .gemini_extractor import clean_thai_ocr_text
        text_val = clean_thai_ocr_text(text_val)
        text_val = text_val.replace("&nbsp;", " ")
        text_val = re.sub(r"</?hr\s*/?>", "<br>", text_val, flags=re.IGNORECASE)

        lines = re.split(r"<br\s*/?>|\n", text_val, flags=re.IGNORECASE)
        while lines and not lines[0].strip():
            lines.pop(0)
        while lines and not lines[-1].strip():
            lines.pop()
        if not lines:
            lines = [""]

        align_str = str(cell_data.get("align", "left")).lower()
        if "center" in align_str or is_badge:
            p_align = WD_ALIGN_PARAGRAPH.CENTER
        elif "right" in align_str:
            p_align = WD_ALIGN_PARAGRAPH.RIGHT
        elif "justify" in align_str:
            p_align = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            p_align = WD_ALIGN_PARAGRAPH.LEFT

        bold = bool(cell_data.get("bold", False)) or is_badge
        font_size = cell_data.get("font_size_pt")
        if not font_size:
            font_size = 11.5 if is_badge else max(9.0, self.font_sizes["body"] - 4.5)

        color_rgb = RGBColor(255, 255, 255) if (is_badge or bg_color in ["595959", "#595959"]) else None

        for line_idx, line in enumerate(lines):
            line_str = line.strip()
            p = cell.paragraphs[0] if (line_idx == 0 and cell.paragraphs) else cell.add_paragraph()

            p.paragraph_format.space_before = Pt(0.5)
            p.paragraph_format.space_after = Pt(0.5)
            p.paragraph_format.line_spacing = 1.0
            p.alignment = p_align

            if "[IMAGE]" in line_str.upper() or (has_image and line_idx == 0):
                remaining = re.sub(r"\[IMAGE\]", "", line_str, flags=re.IGNORECASE).strip()
                if images_queue:
                    img = images_queue.pop(0)
                    self._insert_image_to_paragraph(p, img, max_width_inches=min(cell_width_in, 3.0))
                if remaining:
                    p2 = cell.add_paragraph()
                    p2.paragraph_format.space_before = Pt(0.5)
                    p2.paragraph_format.space_after = Pt(0.5)
                    p2.paragraph_format.line_spacing = 1.0
                    p2.alignment = p_align
                    self._add_formatted_text_to_paragraph(p2, remaining, size_pt=font_size, bold=bold, color_rgb=color_rgb)
                continue

            if not line_str:
                r_empty = p.add_run("")
                set_run_font(r_empty, self.font_name, size_pt=font_size)
                continue

            self._add_formatted_text_to_paragraph(p, line_str, size_pt=font_size, bold=bold, color_rgb=color_rgb)

    def _render_header_block(self, block: dict):
        """บรรจุเนื้อหาหัวกระดาษลงใน Section Header ของ Word โดยตรง"""
        text = str(block.get("text", "")).strip()
        if not text:
            return
        sec = self.doc.sections[-1]
        existing_text = "".join(p.text for p in sec.header.paragraphs).strip()
        if existing_text == text:
            return

        header = sec.header
        header.is_linked_to_previous = False
        p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        p.text = ""
        align = str(block.get("align", "left")).lower()
        if "center" in align:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif "right" in align:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        font_sz = max(8.5, min(10.5, self.font_sizes["body"] - 4.0))
        self._add_formatted_text_to_paragraph(p, text, size_pt=font_sz, bold=bool(block.get("bold", False)))

    def _render_footer_block(self, block: dict):
        """บรรจุเนื้อหาท้ายกระดาษ/เลขหน้าลงใน Section Footer ของ Word โดยตรง"""
        text = str(block.get("text", "")).strip()
        if not text:
            return
        sec = self.doc.sections[-1]
        existing_text = "".join(p.text for p in sec.footer.paragraphs).strip()
        if existing_text == text:
            return

        footer = sec.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.text = ""
        align = str(block.get("align", "center")).lower()
        if "right" in align:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif "left" in align:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(0)
        font_sz = max(8.5, min(10.0, self.font_sizes["body"] - 4.5))

        is_page_number = bool(block.get("is_page_number", False))
        self._render_footer_text(p, text, font_sz, is_page_number)

    def _render_footer_text(self, p, text: str, font_sz: float, is_page_number: bool):
        """เรนเดอร์ข้อความท้ายกระดาษ พร้อมรองรับหมายเลขหน้าแบบ Dynamic Word Field (<w:fldSimple w:instr="PAGE"/>)"""
        page_pattern = r"(หน้า(?:ที่)?\s*)\d+(\s*/\s*\d+)?"
        page_en_pattern = r"(Page\s*)\d+(\s*(?:of|/)\s*\d+)?"

        m_th = re.search(page_pattern, text, flags=re.IGNORECASE)
        m_en = re.search(page_en_pattern, text, flags=re.IGNORECASE)

        if m_th or m_en or is_page_number:
            match = m_th or m_en
            if match:
                prefix = text[:match.start()]
                p_label = match.group(1)
                suffix_total = match.group(2) or ""
                remainder = text[match.end():]

                if prefix:
                    self._add_formatted_text_to_paragraph(p, prefix, size_pt=font_sz)

                r1 = p.add_run(p_label)
                set_run_font(r1, self.font_name, font_sz)

                r_num = p.add_run()
                set_run_font(r_num, self.font_name, font_sz)
                fld = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>')
                r_num._r.append(fld)

                if suffix_total:
                    r_tot = p.add_run(suffix_total)
                    set_run_font(r_tot, self.font_name, font_sz)

                if remainder:
                    self._add_formatted_text_to_paragraph(p, remainder, size_pt=font_sz)
                return

        self._add_formatted_text_to_paragraph(p, text, size_pt=font_sz)

    def _render_markdown_lines(self, lines: List[str], images_queue: List[ExtractedImage]):
        """เรนเดอร์เอกสารตามบรรทัด Markdown แบบเดิม (Legacy Fallback)"""
        i = 0
        while i < len(lines):
            raw_line = lines[i].strip()

            # 1. บรรทัดว่าง
            if not raw_line:
                i += 1
                continue

            # 1.1 ตรวจสอบแท็ก [HEADER] หรือ [FOOTER]
            if "[HEADER]" in raw_line.upper():
                hdr_text = re.sub(r"\[/?HEADER\]", "", raw_line, flags=re.IGNORECASE).strip()
                self._render_header_block({"text": hdr_text, "align": "left"})
                i += 1
                continue
            if "[FOOTER]" in raw_line.upper():
                ftr_text = re.sub(r"\[/?FOOTER\]", "", raw_line, flags=re.IGNORECASE).strip()
                self._render_footer_block({"text": ftr_text, "align": "center"})
                i += 1
                continue

            # 2. ตรวจสอบตาราง (Markdown Table) และแท็กกำหนดเส้นขอบ
            is_borderless_table = False
            is_horizontal_only_table = False
            if "[BORDERLESS]" in raw_line.upper() or "[NO_BORDER]" in raw_line.upper():
                is_borderless_table = True
                i += 1
                if i < len(lines):
                    raw_line = lines[i].strip()
            elif "[HORIZONTAL_ONLY]" in raw_line.upper() or "[HORIZONTAL_BORDERS]" in raw_line.upper():
                is_horizontal_only_table = True
                i += 1
                if i < len(lines):
                    raw_line = lines[i].strip()

            if raw_line.startswith("|") and ("|" in raw_line[1:]):
                table_lines = []
                while i < len(lines) and lines[i].strip().startswith("|") and ("|" in lines[i].strip()[1:]):
                    table_lines.append(lines[i].strip())
                    i += 1
                self._create_word_table(
                    table_lines,
                    images_queue=images_queue,
                    is_borderless=is_borderless_table,
                    is_horizontal_only=is_horizontal_only_table,
                )
                continue

            # 3. ตรวจสอบแท็กรูปภาพเดี่ยว [IMAGE] (ที่ไม่ได้อยู่ในตาราง)
            if "[IMAGE]" in raw_line.upper():
                if images_queue:
                    img = images_queue.pop(0)
                    self._insert_image(img)
                i += 1
                continue

            # ตรวจสอบและแยกแท็กจัดตำแหน่งและขนาดก่อน
            line, alignment, is_small = self.parse_line_formatting(raw_line)
            if not line:
                i += 1
                continue

            # 4. ตรวจสอบหัวข้อ (Headings)
            if line.startswith("### "):
                p = self.doc.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
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
                p.paragraph_format.space_before = Pt(2.5)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.keep_with_next = True
                if alignment:
                    p.alignment = alignment
                font_sz = self.font_sizes["h2"] if not is_small else self.font_sizes["small"]
                self._add_formatted_text_to_paragraph(p, line[3:].strip(), size_pt=font_sz, bold=True)
                i += 1
                continue
            elif line.startswith("# "):
                p = self.doc.add_paragraph()
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(1.5)
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
                p.paragraph_format.line_spacing = 1.10
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
                p.paragraph_format.line_spacing = 1.10
                if alignment:
                    p.alignment = alignment
                font_sz = self.font_sizes["body"] if not is_small else self.font_sizes["small"]
                self._add_formatted_text_to_paragraph(p, num_match.group(2).strip(), size_pt=font_sz)
                i += 1
                continue

            # 7. ย่อหน้าปกติ (Paragraph)
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.line_spacing = 1.10
            if alignment:
                p.alignment = alignment
                if alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                    p.paragraph_format.first_line_indent = Inches(0.4)
            elif (
                len(line) > 80
                and not any(pfx in line for pfx in [":", "：", "____", "  ", "\t", "|"])
                and not any(line.startswith(pfx) for pfx in ["ที่ ", "เรื่อง ", "เรียน ", "ถึง ", "จาก ", "วันที่ ", "เอกสาร "])
                and len(line.split()) >= 10
            ):
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
        color_rgb: Optional[RGBColor] = None,
    ):
        """แยกแท็กตัวหนา (**ข้อความ**) และตัวเอียง (*ข้อความ*) ออกมาใส่ใน Run พร้อมกำหนดขนาดและสี"""
        from .gemini_extractor import clean_thai_ocr_text
        text = clean_thai_ocr_text(text)
        text = re.sub(r"\[/?BADGE\]", "", text, flags=re.IGNORECASE)
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
                set_run_font(run, self.font_name, actual_size, bold=True, italic=italic, color_rgb=color_rgb)
            elif token.startswith("*") and token.endswith("*") and len(token) >= 2:
                run = paragraph.add_run(token[1:-1])
                set_run_font(run, self.font_name, actual_size, bold=bold, italic=True, color_rgb=color_rgb)
            else:
                run = paragraph.add_run(token)
                set_run_font(run, self.font_name, actual_size, bold=bold, italic=italic, color_rgb=color_rgb)

    def _create_word_table(
        self,
        table_lines: List[str],
        images_queue: Optional[List[ExtractedImage]] = None,
        is_borderless: bool = False,
        is_horizontal_only: bool = False,
    ):
        """แปลงตาราง Markdown เป็น Table Object ของ Word พร้อมจัดขอบตาราง ถอดรหัสแท็ก และแทรกรูปในเซลล์"""
        if not table_lines:
            return

        parsed_rows = []
        for line in table_lines:
            # ตรวจสอบแท็ก [BORDERLESS] หรือ [HORIZONTAL_ONLY] ในบรรทัดตาราง
            if "[BORDERLESS]" in line.upper() or "[NO_BORDER]" in line.upper():
                is_borderless = True
                continue
            if "[HORIZONTAL_ONLY]" in line.upper() or "[HORIZONTAL_BORDERS]" in line.upper():
                is_horizontal_only = True
                continue

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

        if num_cols == 0 or num_rows == 0:
            return

        # เติมเต็มเซลล์ว่างให้ทุกแถวมีจำนวนคอลัมน์เท่ากับ num_cols พอดี ป้องกัน IndexError: list index out of range
        for r in parsed_rows:
            while len(r) < num_cols:
                r.append("")

        # ตรวจสอบและตรวจจับอัตโนมัติ (Smart Auto-detection สำหรับรายงานผลแล็บและบล็อกลายเซ็น)
        all_text = " ".join(" ".join(r) for r in parsed_rows).lower()
        if not is_borderless and not is_horizontal_only:
            # 1. ตรวจจับบล็อกลงลายมือชื่อและข้อมูลผู้ป่วย (Signature & Patient blocks) -> ไร้ขอบ 100%
            sig_keywords = ["reported by", "approved by", "medical technologist", "ผู้รายงาน", "ผู้รับรอง", "ผู้มีอำนาจลงนาม", "ลายมือชื่อ"]
            patient_keywords = ["patient's name", "hospital no", "opd / ward", "lab barcode"]
            if any(k in all_text for k in sig_keywords) or any(k in all_text for k in patient_keywords):
                is_borderless = True
            # 2. ตรวจจับตารางผลตรวจแล็บ / ทางการแพทย์ (มี Test Name / Result / Reference Range) -> เฉพาะเส้นแนวนอน
            elif any(k in all_text for k in ["test name", "parameters", "reference range"]) and not any(k in all_text for k in ["shopee", "barcode", "tracking"]):
                is_horizontal_only = True
            elif len(parsed_rows) == 1 and any("PICK UP" in c.upper() or "SPX" in c.upper() for c in parsed_rows[0]):
                is_borderless = True

        table = self.doc.add_table(rows=num_rows, cols=num_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        if is_borderless:
            set_table_borderless(table)
        elif is_horizontal_only:
            set_table_horizontal_borders(table)
        else:
            table.style = "Table Grid"
        # ขอบเซลล์แบบกะทัดรัด (Compact cell padding)
        set_table_margins(table, top=40, bottom=40, left=80, right=80)

        # คำนวณความกว้างคอลัมน์แบบสัดส่วนตามเนื้อหาจริง (Smart Proportional Column Widths)
        total_page_width_in = self._get_current_printable_width_in()
        col_weights = []
        for col_idx in range(num_cols):
            col_cells = [r[col_idx] for r in parsed_rows if col_idx < len(r)]
            max_len = max((len(c) for c in col_cells), default=1)
            has_img = any("[IMAGE]" in c.upper() for c in col_cells)
            if has_img:
                weight = max(30.0, min(float(max_len), 70.0))
            else:
                weight = max(12.0, min(float(max_len), 150.0))
            col_weights.append(weight)

        total_weight = sum(col_weights) if sum(col_weights) > 0 else 1.0
        col_widths = [(w / total_weight) * total_page_width_in for w in col_weights]
        col_widths = [max(0.8, w) for w in col_widths]
        total_w = sum(col_widths)
        col_widths = [(w / total_w) * total_page_width_in for w in col_widths]

        table.autofit = False
        for c_idx, col in enumerate(table.columns):
            if c_idx < len(col_widths):
                col.width = Inches(col_widths[c_idx])

        # ตรวจจับคอลัมน์ที่มีแถวว่างต่อเนื่อง (Auto-merge empty vertical cells)
        # เช่น ฝั่งผู้รับ (TO) หรือบาร์โค้ดที่มีแถวย่อยหลายแถวในฝั่งขวา
        merged_cells_coords = set()
        for col in range(num_cols):
            r = 0
            while r < num_rows:
                if col < len(parsed_rows[r]) and parsed_rows[r][col].strip():
                    start_r = r
                    r += 1
                    while (
                        r < num_rows
                        and col < len(parsed_rows[r])
                        and not parsed_rows[r][col].strip()
                        and any((c < len(parsed_rows[r]) and bool(parsed_rows[r][c].strip())) for c in range(num_cols))
                    ):
                        r += 1
                    if r - 1 > start_r:
                        table.rows[start_r].cells[col].merge(table.rows[r - 1].cells[col])
                        for mr in range(start_r + 1, r):
                            merged_cells_coords.add((mr, col))
                else:
                    r += 1

        # คำนวณขนาดตัวอักษรสำหรับตารางให้กะทัดรัด พอดีกับช่องเอกสาร
        table_font_size = max(9.0, min(11.5, self.font_sizes["body"] - 4.5))
        header_font_size = max(10.0, min(12.5, self.font_sizes["body"] - 3.5))

        for row_idx, row_data in enumerate(parsed_rows):
            # ใช้ tblHeader เฉพาะตารางข้อมูลยาว (>= 6 แถว) เพื่อไม่ให้วนซ้ำในตารางขั้นตอนสั้นๆ
            is_header = (row_idx == 0 and len(parsed_rows) >= 6)
            row = table.rows[row_idx]

            # กำหนดคุณสมบัติไม่ให้แถวแตกข้ามหน้าถ้าไม่จำเป็น
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            # ถ้าเป็นหัวตารางของตารางยาว ให้วนซ้ำหัวตารางเมื่อขึ้นหน้าใหม่ (Repeat Header Row)
            if is_header:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

            for col_idx in range(num_cols):
                if (row_idx, col_idx) in merged_cells_coords:
                    continue

                cell = row.cells[col_idx]
                if col_idx < len(col_widths):
                    cell.width = Inches(col_widths[col_idx])
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                raw_cell_content = row_data[col_idx] if col_idx < len(row_data) else ""
                cell.text = ""  # เคลียร์พารากราฟเริ่มต้น

                # แทนที่ &nbsp; และแปลง <hr> เป็น <br>
                raw_cell_content = raw_cell_content.replace("&nbsp;", " ")
                raw_cell_content = re.sub(r"</?hr\s*/?>", "<br>", raw_cell_content, flags=re.IGNORECASE)

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
                        p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                    else:
                        p = cell.add_paragraph()

                    p.paragraph_format.space_before = Pt(0.5)
                    p.paragraph_format.space_after = Pt(0.5)
                    p.paragraph_format.line_spacing = 1.0

                    # ตรวจสอบแท็กรูปภาพ [IMAGE] ในเซลล์ตาราง (เช่น บาร์โค้ดหรือโลโก้)
                    if "[IMAGE]" in line_str.upper():
                        remaining_text = re.sub(r"\[IMAGE\]", "", line_str, flags=re.IGNORECASE).strip()
                        if images_queue:
                            img = images_queue.pop(0)
                            self._insert_image_to_paragraph(p, img, max_width_inches=3.0)
                        if remaining_text:
                            # มีข้อความอื่นอยู่ในบรรทัดเดียวกับ [IMAGE] เช่น [BADGE]PICK UP[/BADGE]
                            line_str = remaining_text
                            p = cell.add_paragraph()
                            p.paragraph_format.space_before = Pt(0.5)
                            p.paragraph_format.space_after = Pt(0.5)
                            p.paragraph_format.line_spacing = 1.0
                        else:
                            continue

                    # ตรวจสอบป้ายกำกับ [BADGE]
                    is_badge = False
                    if "[BADGE]" in line_str.upper():
                        is_badge = True
                        line_str = re.sub(r"\[/?BADGE\]", "", line_str, flags=re.IGNORECASE).strip()

                    # แยกแท็กจัดตำแหน่ง [CENTER], [RIGHT], [SMALL]
                    clean_text, align, is_small = self.parse_line_formatting(line_str)

                    # ตรวจจับป้ายกำกับแบบฟอร์มที่เป็นสัญลักษณ์ทึบอัตโนมัติ (เช่น W, RR, PICK UP, COD)
                    if clean_text.strip().upper() in ["W", "RR", "PICK UP", "COD"]:
                        is_badge = True

                    # ตรวจสอบหัวข้อในเซลล์ (เช่น # W_0_A2_227_HBKAE-B หรือ -# I17)
                    is_cell_heading = False
                    cell_heading_match = re.match(r"^[-*]?\s*(#{1,3})\s+(.*)$", clean_text)
                    if cell_heading_match:
                        h_level = len(cell_heading_match.group(1))
                        clean_text = cell_heading_match.group(2).strip()
                        is_cell_heading = True
                        if h_level == 1:
                            curr_size = 20.0
                        elif h_level == 2:
                            curr_size = 15.0
                        else:
                            curr_size = 13.0
                    elif is_small:
                        curr_size = self.font_sizes["small"]
                    elif is_header:
                        curr_size = header_font_size
                    else:
                        curr_size = table_font_size

                    # กำหนดการจัดตำแหน่งในเซลล์
                    if align:
                        p.alignment = align
                    elif is_header or is_badge or is_cell_heading:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif re.match(r"^[\$฿€¥]?\s*[\d,]+(\.\d+)?%?$", clean_text):
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    if is_badge:
                        has_cell_img = any("[IMAGE]" in cl.upper() for cl in cell_lines) or "[IMAGE]" in raw_cell_content.upper()
                        if not has_cell_img:
                            set_cell_background(cell, "595959")
                            self._add_formatted_text_to_paragraph(p, clean_text, size_pt=max(curr_size, 11.5), bold=True, color_rgb=RGBColor(255, 255, 255))
                        else:
                            self._add_formatted_text_to_paragraph(p, clean_text, size_pt=max(curr_size, 11.5), bold=True)
                    elif is_header:
                        set_cell_background(cell, "F2F2F2")
                        self._add_formatted_text_to_paragraph(p, clean_text, size_pt=curr_size, bold=True)
                    elif is_cell_heading:
                        self._add_formatted_text_to_paragraph(p, clean_text, size_pt=curr_size, bold=True)
                    else:
                        self._add_formatted_text_to_paragraph(p, clean_text, size_pt=curr_size)

        # เพิ่ม Paragraph คั่นตารางขนาด 1 pt ป้องกันไม่ให้ Microsoft Word หลอมรวมตารางที่อยู่ติดกันเป็นตารางเดียว
        sep_p = self.doc.add_paragraph()
        sep_p.paragraph_format.space_before = Pt(0)
        sep_p.paragraph_format.space_after = Pt(0)
        sep_p.paragraph_format.line_spacing = Pt(1)
        run = sep_p.add_run()
        run.font.size = Pt(1)

    def _insert_image_to_paragraph(self, p, img: ExtractedImage, max_width_inches: float = 6.0):
        """แทรกรูปภาพลงใน Paragraph ที่ระบุ พร้อมจำกัดความกว้างและความสูงไม่ให้ล้นตาราง"""
        try:
            image_stream = io.BytesIO(img.image_bytes)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            if img.bbox and len(img.bbox) == 4:
                bbox_w_pt = abs(img.bbox[2] - img.bbox[0])
                bbox_h_pt = abs(img.bbox[3] - img.bbox[1])
                width_in_inches = min(bbox_w_pt / 72.0, max_width_inches)
                # ควบคุมความสูงสูงสุดในช่องตารางไม่ให้เกิน 1.05 นิ้ว เพื่อประหยัดพื้นที่แนวตั้ง
                if bbox_w_pt > 0:
                    est_height = (bbox_h_pt / bbox_w_pt) * width_in_inches
                    if est_height > 1.05:
                        width_in_inches = (1.05 / est_height) * width_in_inches
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

    def to_bytes(self) -> bytes:
        """ส่งคืนเนื้อหาเอกสาร Word ในรูปแบบ bytes จากหน่วยความจำ"""
        import io
        bio = io.BytesIO()
        self.doc.save(bio)
        return bio.getvalue()

    def save(self, output_path: str | Path):
        """บันทึกเอกสาร Word ลงในไฟล์เป้าหมาย"""
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(path))
