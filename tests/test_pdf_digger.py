import io
import os
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch
import docx
import pymupdf as fitz
from PIL import Image

from src.config import DEFAULT_FONT, FALLBACK_MODELS
from src.docx_builder import DocxBuilder
from src.gemini_extractor import GeminiExtractor
from src.pdf_processor import ExtractedImage, PDFProcessor, enhance_document_image


class TestPDFDigger(unittest.TestCase):
    def setUp(self):
        self.test_dir = Path("./test_output")
        self.test_dir.mkdir(exist_ok=True)

    def tearDown(self):
        # ลบไฟล์ทดสอบ
        for f in self.test_dir.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
        try:
            self.test_dir.rmdir()
        except Exception:
            pass

    def test_docx_builder_thai_font_and_tables(self):
        """ทดสอบการสร้างไฟล์ Word: ตรวจสอบฟอนต์ Cordia New, โครงสร้างตาราง และการแทรกรูป"""
        builder = DocxBuilder(font_name="Cordia New")

        markdown_content = (
            "# รายงานสรุปผลการดำเนินงานประจำปี\n"
            "## ข้อมูลทั่วไปและการทดสอบภาษาไทย\n"
            "นี่คือข้อความภาษาไทยที่มีสระบน สระล่าง เช่น ที่นี่ ปี่ น้้ำ และ **ข้อความตัวหนา** กับ *ตัวเอียง*\n\n"
            "| ลำดับ | รายการ | จำนวน | ราคาต่อหน่วย | รวมเป็นเงิน |\n"
            "|---|---|---|---|---|\n"
            "| 1 | ค่าบริการจัดทำเอกสาร | 2 | 1,500.00 | 3,000.00 |\n"
            "| 2 | ค่าแปลภาษาไทย-อังกฤษ | 5 | 800.00 | 4,000.00 |\n\n"
            "[IMAGE]\n\n"
            "- รายการหัวข้อย่อยที่ 1\n"
            "- รายการหัวข้อย่อยที่ 2\n"
        )

        # สร้างภาพจำลอง
        img_buffer = io.BytesIO()
        pil_img = Image.new("RGB", (200, 100), color="blue")
        pil_img.save(img_buffer, format="PNG")
        dummy_image = ExtractedImage(
            page_num=0,
            image_index=1,
            image_bytes=img_buffer.getvalue(),
            ext="png",
            width=200,
            height=100,
        )

        builder.add_page_content(
            markdown_text=markdown_content,
            page_num=1,
            images=[dummy_image],
            is_first_page=True,
        )

        out_path = self.test_dir / "test_doc.docx"
        builder.save(out_path)

        self.assertTrue(out_path.exists())

        # ตรวจสอบโครงสร้างไฟล์ docx ที่สร้างขึ้น
        saved_doc = docx.Document(str(out_path))

        # ตรวจสอบว่ามีตารางอย่างน้อย 1 ตาราง
        self.assertEqual(len(saved_doc.tables), 1)
        tbl = saved_doc.tables[0]
        self.assertEqual(len(tbl.rows), 3)  # หัวตาราง 1 แถว + ข้อมูล 2 แถว
        self.assertEqual(len(tbl.columns), 5)

        # ตรวจสอบหัวตาราง
        self.assertEqual(tbl.rows[0].cells[1].text.strip(), "รายการ")
        self.assertEqual(tbl.rows[1].cells[1].text.strip(), "ค่าบริการจัดทำเอกสาร")

        # ตรวจสอบว่ามีการตั้งค่า Cordia New ใน XML Complex Script (cs)
        xml_content = saved_doc._body._element.xml
        self.assertIn('w:cs="Cordia New"', xml_content)
        self.assertIn('w:ascii="Cordia New"', xml_content)
        print("\n✅ ทดสอบ DocxBuilder (ฟอนต์ Cordia New, ตาราง, รูปภาพ) สำเร็จ")

    def test_pdf_processor_rendering(self):
        """ทดสอบการเปิดและเรนเดอร์หน้า PDF เป็นภาพ 200 DPI"""
        test_pdf = self.test_dir / "sample.pdf"

        # สร้าง PDF ตัวอย่างด้วย PyMuPDF
        doc = fitz.open()
        page = doc.new_page()
        page.insert_text((50, 72), "Test Document / ทดสอบเอกสารภาษาไทย", fontsize=16)

        # แทรกรูปภาพทดสอบลงใน PDF
        img_buffer = io.BytesIO()
        pil_img = Image.new("RGB", (100, 100), color="red")
        pil_img.save(img_buffer, format="PNG")
        page.insert_image(fitz.Rect(50, 100, 150, 200), stream=img_buffer.getvalue())

        doc.save(str(test_pdf))
        doc.close()

        processor = PDFProcessor(test_pdf, dpi=200)
        self.assertEqual(processor.get_page_count(), 1)

        page_data = processor.process_page(0)
        self.assertGreater(len(page_data.rendered_image_bytes), 0)
        self.assertGreaterEqual(len(page_data.embedded_images), 1)
        self.assertEqual(page_data.embedded_images[0].ext, "png")
        print("✅ ทดสอบ PDFProcessor (เรนเดอร์ 200 DPI และสกัดภาพ) สำเร็จ")

    def test_gemini_fallback_mechanism(self):
        """ทดสอบระบบ Seamless Fallback เมื่อโมเดลแรกติด 429 Quota Exceeded"""
        fallback_log = []

        def fallback_cb(prev, nxt, reason):
            fallback_log.append((prev, nxt, reason))

        with patch("src.gemini_extractor.genai.Client") as mock_client_cls:
            mock_client = MagicMock()
            mock_client_cls.return_value = mock_client

            call_count = 0

            def mock_generate_content(model, contents, config):
                nonlocal call_count
                call_count += 1
                if model == "gemini-3.6-flash":
                    # จำลองข้อผิดพลาด HTTP 429 Quota Exceeded
                    raise Exception("429 ResourceExhausted: Quota exceeded for model gemini-3.6-flash")
                else:
                    # โมเดลสำรองทำงานสำเร็จ
                    mock_resp = MagicMock()
                    mock_resp.text = "# หัวข้อที่แกะได้จากโมเดลสำรอง\nเนื้อหาทดสอบ"
                    return mock_resp

            mock_client.models.generate_content.side_effect = mock_generate_content

            extractor = GeminiExtractor(
                api_key="test-api-key",
                models=["gemini-3.6-flash", "gemini-3.7-flash", "gemini-flash-latest"],
                on_fallback=fallback_cb,
            )

            # ทดสอบเรียกประมวลผล
            dummy_bytes = b"fake-image-png-bytes"
            result_text, used_model, _ = extractor.extract_page_markdown(dummy_bytes, 0)

            # ตรวจสอบว่าสลับไปใช้ gemini-3.7-flash สำเร็จ
            self.assertEqual(used_model, "gemini-3.7-flash")
            self.assertIn("หัวข้อที่แกะได้จากโมเดลสำรอง", result_text)

            # ตรวจสอบว่ามีการบันทึกเหตุการณ์ fallback
            self.assertEqual(len(fallback_log), 1)
            self.assertEqual(fallback_log[0][0], "gemini-3.6-flash")
            self.assertEqual(fallback_log[0][1], "gemini-3.7-flash")
            self.assertIn("429 Quota Exceeded", fallback_log[0][2])

            print("✅ ทดสอบ Seamless Model Fallback (สลับโมเดลอัตโนมัติเมื่อโควตาเต็ม) สำเร็จ")

    def test_font_detector_normalization(self):
        """ทดสอบการทำความสะอาดและตรวจจับชื่อฟอนต์มาตรฐาน"""
        from src.font_detector import clean_font_name, detect_dominant_font

        # 1. ทดสอบการตัด prefix และแปลงชื่อฟอนต์
        self.assertEqual(clean_font_name("ABCDEF+THSarabunPSK-Bold"), "TH Sarabun New")
        self.assertEqual(clean_font_name("CordiaNew-Regular"), "Cordia New")
        self.assertEqual(clean_font_name("AngsanaUPC,Bold"), "Angsana New")
        self.assertEqual(clean_font_name("BrowalliaNew"), "Browallia New")
        self.assertEqual(clean_font_name("Tahoma-Bold"), "Tahoma")
        self.assertEqual(clean_font_name("Calibri"), "Calibri")

        # 2. ทดสอบตรวจจับฟอนต์จาก PDF ดิจิทัล
        test_pdf = self.test_dir / "font_sample.pdf"
        doc = fitz.open()
        page = doc.new_page()
        page.insert_text((50, 72), "This is a document with Helvetica or Times font", fontname="helv", fontsize=14)
        doc.save(str(test_pdf))
        doc.close()

        detected = detect_dominant_font(test_pdf)
        self.assertIsNotNone(detected)
        self.assertEqual(detected, "Arial")  # helv mapped to Arial
        print("✅ ทดสอบ Font Detector (ตรวจจับและแปลงชื่อฟอนต์ต้นฉบับ) สำเร็จ")

    def test_image_support_and_enhancement(self):
        """ทดสอบการรองรับไฟล์รูปภาพโดยตรง (.png) และระบบ Image Enhancement"""
        test_img_path = self.test_dir / "sample_scan.png"
        img = Image.new("RGB", (800, 600), color=(240, 235, 220))  # กระดาษอมเหลือง
        img.save(test_img_path)

        # 1. ตรวจสอบว่า PDFProcessor จัดการรูปภาพได้ถูกต้อง
        processor = PDFProcessor(test_img_path)
        self.assertTrue(processor.is_image)
        self.assertEqual(processor.get_page_count(), 1)
        self.assertIsNone(processor.detect_font())

        # 2. ตรวจสอบการแปลงและการเปิดโหมด enhance
        p_data = processor.process_page(0, enhance=True)
        self.assertGreater(len(p_data.rendered_image_bytes), 0)

        # 3. ตรวจสอบฟังก์ชัน enhance_document_image โดยตรง
        enhanced_bytes = enhance_document_image(p_data.rendered_image_bytes)
        self.assertGreater(len(enhanced_bytes), 0)
        print("✅ ทดสอบ Direct Image Input & Image Enhancement สำเร็จ")

    def test_translation_prompt_generation(self):
        """ทดสอบการสร้างคำสั่งแปลภาษา (Target Language) ใน GeminiExtractor"""
        with patch("src.gemini_extractor.genai.Client") as mock_client_cls:
            mock_client = MagicMock()
            mock_client_cls.return_value = mock_client

            captured_prompt = None

            def mock_generate_content(model, contents, config):
                nonlocal captured_prompt
                captured_prompt = contents[1]  # prompt string
                mock_resp = MagicMock()
                mock_resp.text = "# Translated Content in Thai\nเนื้อหาที่แปลเป็นไทย"
                return mock_resp

            mock_client.models.generate_content.side_effect = mock_generate_content

            extractor = GeminiExtractor(api_key="test-api-key")
            dummy_bytes = b"fake-image-bytes"

            # ทดสอบแปลเป็นไทย
            result, _, _ = extractor.extract_page_markdown(dummy_bytes, 0, target_language="th")
            self.assertIn("แปลเนื้อหาทั้งหมดเป็นภาษาไทย", captured_prompt)
            self.assertIn("Translated Content", result)

            # ทดสอบแปลเป็นอังกฤษ
            result_en, _, _ = extractor.extract_page_markdown(dummy_bytes, 0, target_language="en")
            self.assertIn("แปลเนื้อหาทั้งหมดเป็นภาษาอังกฤษ", captured_prompt)

            print("✅ ทดสอบ Translation Prompt Generation (แปลไทย/อังกฤษ) สำเร็จ")


if __name__ == "__main__":
    unittest.main()
