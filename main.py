#!/usr/bin/env python3
"""
PDF Digger CLI - เครื่องมือแปลง PDF และรูปภาพ ภาษาไทย/อังกฤษ เป็น Word (.docx) ด้วย Gemini AI
"""

import argparse
import logging
import sys
from pathlib import Path
from tqdm import tqdm

from src.config import DEFAULT_FONT, FALLBACK_MODELS
from src.pdf_processor import SUPPORTED_IMAGE_EXTENSIONS
from src.pipeline import PDFToWordPipeline

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("PDFDigger")


def main():
    parser = argparse.ArgumentParser(
        description="PDF Digger: แปลง PDF และรูปภาพ (ดิจิทัล/ภาพสแกน) เป็น Word (.docx) พร้อมแปลภาษาด้วย Gemini AI"
    )
    parser.add_argument("input", help="ที่อยู่ไฟล์ PDF/รูปภาพ ต้นทาง หรือโฟลเดอร์ที่มีไฟล์เอกสาร")
    parser.add_argument("-o", "--output", help="ที่อยู่ไฟล์ Word (.docx) ปลายทาง (ค่าเริ่มต้น: ชื่อเดียวกับต้นฉบับ)")
    parser.add_argument("--key", help="Gemini API Key (หากไม่ระบุจะดึงจากตัวแปรใน .env)")
    parser.add_argument(
        "--font",
        default=None,
        help=f"แบบอักษรสำหรับไฟล์ Word (ค่าเริ่มต้น: ตรวจจับจากต้นฉบับอัตโนมัติ หรือใช้ {DEFAULT_FONT})",
    )
    parser.add_argument("--model", help=f"โมเดลเริ่มต้นที่ต้องการใช้ (ค่าเริ่มต้น: {FALLBACK_MODELS[0]})")
    parser.add_argument(
        "-t",
        "--translate",
        choices=["original", "th", "en", "zh", "ja"],
        default="original",
        help="แปลภาษาเนื้อหาเอกสาร (original: ไม่แปล, th: แปลเป็นไทย, en: แปลเป็นอังกฤษ, zh: แปลเป็นจีน, ja: แปลเป็นญี่ปุ่น)",
    )
    parser.add_argument(
        "--enhance",
        action="store_true",
        help="เปิดโหมดเพิ่มความคมชัดพิเศษสำหรับภาพสแกนหรือภาพถ่ายมือถือ (Auto-contrast & Sharpening)",
    )
    parser.add_argument("--batch", action="store_true", help="แปลงไฟล์เอกสารและรูปภาพทุกไฟล์ที่อยู่ในโฟลเดอร์ที่ระบุ")

    args = parser.parse_args()

    # จัดการลำดับโมเดล
    models_to_use = FALLBACK_MODELS.copy()
    if args.model:
        if args.model in models_to_use:
            models_to_use.remove(args.model)
        models_to_use.insert(0, args.model)

    pipeline = PDFToWordPipeline(
        api_key=args.key,
        font_name=args.font,
        auto_detect_font=(args.font is None),
        models=models_to_use,
        target_language=args.translate,
        enhance_image=args.enhance,
    )

    input_path = Path(args.input)
    if not input_path.exists():
        logger.error(f"ไม่พบไฟล์หรือโฟลเดอร์: {input_path}")
        sys.exit(1)

    target_files = []
    valid_exts = {".pdf"}.union(SUPPORTED_IMAGE_EXTENSIONS)

    if input_path.is_dir() or args.batch:
        for f in sorted(input_path.iterdir()):
            if f.is_file() and f.suffix.lower() in valid_exts:
                target_files.append(f)

        if not target_files:
            logger.warning(f"ไม่พบไฟล์เอกสารหรือรูปภาพที่รองรับในโฟลเดอร์ {input_path}")
            sys.exit(0)
        logger.info(f"พบไฟล์ทั้งหมด {len(target_files)} ไฟล์ กำลังเริ่มประมวลผล...")
    else:
        target_files = [input_path]

    for file_path in target_files:
        print("\n" + "=" * 60)
        logger.info(f"กำลังแปลงไฟล์: {file_path.name}")

        output_file = None
        if args.output and not (input_path.is_dir() or args.batch):
            output_file = args.output

        pbar = None

        def update_progress(current: int, total: int, status_msg: str):
            nonlocal pbar
            if pbar is None:
                pbar = tqdm(total=total, desc="หน้า", unit="หน้า")
            pbar.n = current
            pbar.set_postfix_str(status_msg[:40], refresh=True)

        try:
            report = pipeline.convert(
                pdf_path=file_path,
                output_docx_path=output_file,
                progress_callback=update_progress,
                target_language=args.translate,
                enhance_image=args.enhance,
            )
            if pbar:
                pbar.close()

            print("\n" + "-" * 40)
            print(" สรุปผลการทำงาน:")
            print(f"  • ไฟล์ต้นฉบับ: {report.pdf_path}")
            print(f"  • ไฟล์ Word ปลายทาง: {report.output_docx_path}")
            font_info = f"{report.applied_font}"
            if report.detected_font:
                font_info += f" (ตรวจพบจากต้นฉบับ: {report.detected_font})"
            else:
                font_info += " (ค่าเริ่มต้น)"
            print(f"  • แบบอักษร (Font): {font_info}")
            print(f"  • จำนวนหน้าที่สำเร็จ: {report.successful_pages}/{report.total_pages} หน้า")
            print(f"  • รูปภาพที่สกัดได้: {report.total_images_extracted} รูป")
            if report.target_language != "original":
                print(f"  • การแปลภาษา: แปลเป็น {report.target_language}")
            if report.enhanced:
                print("  • โหมดปรับความคมชัดภาพ: เปิดใช้งาน (Enhanced)")
            print(f"  • โมเดลที่ใช้: {', '.join(report.models_used)}")
            if report.fallback_events:
                print("  • การสลับโมเดลสำรอง (Fallback):")
                for fb in report.fallback_events:
                    print(f"    - {fb}")
            print("-" * 40)

        except Exception as e:
            if pbar:
                pbar.close()
            logger.error(f"เกิดข้อผิดพลาดขณะแปลงไฟล์ {file_path.name}: {e}")

    print("\nแปลงไฟล์ทั้งหมดเรียบร้อยแล้ว!")


if __name__ == "__main__":
    main()
