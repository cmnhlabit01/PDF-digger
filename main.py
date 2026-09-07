#!/usr/bin/env python3
"""
PDF Digger CLI - เครื่องมือแปลง PDF ภาษาไทย/อังกฤษ เป็น Word (.docx) ด้วย Gemini AI
"""

import argparse
import logging
import sys
from pathlib import Path
from tqdm import tqdm

from src.config import DEFAULT_FONT, FALLBACK_MODELS
from src.pipeline import PDFToWordPipeline

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("PDFDigger")


def main():
    parser = argparse.ArgumentParser(
        description="PDF Digger: แปลง PDF (ดิจิทัล/ภาพสแกน) เป็น Word (.docx) ภาษาไทย-อังกฤษ ด้วย Gemini AI"
    )
    parser.add_argument("input", help="ที่อยู่ไฟล์ PDF ต้นทาง หรือโฟลเดอร์ที่มีไฟล์ PDF")
    parser.add_argument("-o", "--output", help="ที่อยู่ไฟล์ Word (.docx) ปลายทาง (ค่าเริ่มต้น: ชื่อเดียวกับต้นฉบับ)")
    parser.add_argument("--key", help="Gemini API Key (หากไม่ระบุจะดึงจากตัวแปรใน .env)")
    parser.add_argument(
        "--font",
        default=None,
        help=f"แบบอักษรสำหรับไฟล์ Word (ค่าเริ่มต้น: ตรวจจับจากต้นฉบับอัตโนมัติ หรือใช้ {DEFAULT_FONT})",
    )
    parser.add_argument("--model", help=f"โมเดลเริ่มต้นที่ต้องการใช้ (ค่าเริ่มต้น: {FALLBACK_MODELS[0]})")
    parser.add_argument("--batch", action="store_true", help="แปลงไฟล์ PDF ทุกไฟล์ที่อยู่ในโฟลเดอร์ที่ระบุ")

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
    )

    input_path = Path(args.input)
    if not input_path.exists():
        logger.error(f"ไม่พบไฟล์หรือโฟลเดอร์: {input_path}")
        sys.exit(1)

    pdf_files = []
    if input_path.is_dir() or args.batch:
        pdf_files = sorted(list(input_path.glob("*.pdf")) + list(input_path.glob("*.PDF")))
        if not pdf_files:
            logger.warning(f"ไม่พบไฟล์ .pdf ในโฟลเดอร์ {input_path}")
            sys.exit(0)
        logger.info(f"พบไฟล์ PDF ทั้งหมด {len(pdf_files)} ไฟล์ กำลังเริ่มประมวลผล...")
    else:
        pdf_files = [input_path]

    for pdf_file in pdf_files:
        print("\n" + "=" * 60)
        logger.info(f"กำลังแปลงไฟล์: {pdf_file.name}")

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
                pdf_path=pdf_file,
                output_docx_path=output_file,
                progress_callback=update_progress,
            )
            if pbar:
                pbar.close()

            print("\n" + "-" * 40)
            print(" สรุปผลการทำงาน:")
            print(f"  • ไฟล์ PDF ต้นฉบับ: {report.pdf_path}")
            print(f"  • ไฟล์ Word ปลายทาง: {report.output_docx_path}")
            font_info = f"{report.applied_font}"
            if report.detected_font:
                font_info += f" (ตรวจพบจากต้นฉบับ: {report.detected_font})"
            else:
                font_info += " (ค่าเริ่มต้น)"
            print(f"  • แบบอักษร (Font): {font_info}")
            print(f"  • จำนวนหน้าที่สำเร็จ: {report.successful_pages}/{report.total_pages} หน้า")
            print(f"  • รูปภาพที่สกัดได้: {report.total_images_extracted} รูป")
            print(f"  • โมเดลที่ใช้: {', '.join(report.models_used)}")
            if report.fallback_events:
                print("  • การสลับโมเดลสำรอง (Fallback):")
                for fb in report.fallback_events:
                    print(f"    - {fb}")
            print("-" * 40)

        except Exception as e:
            if pbar:
                pbar.close()
            logger.error(f"เกิดข้อผิดพลาดขณะแปลงไฟล์ {pdf_file.name}: {e}")

    print("\nแปลงไฟล์ทั้งหมดเรียบร้อยแล้ว!")


if __name__ == "__main__":
    main()
