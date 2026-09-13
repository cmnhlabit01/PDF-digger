import io
import os
import tempfile
import threading
import time
from pathlib import Path
from PIL import Image
import streamlit as st
import pymupdf as fitz

from src.config import DEFAULT_FONT, FALLBACK_MODELS
from src.pdf_processor import PDFProcessor, SUPPORTED_IMAGE_EXTENSIONS
from src.pipeline import PDFToWordPipeline

st.set_page_config(
    page_title="PDF Digger - แกะ PDF / รูปภาพ เป็น Word",
    page_icon="📄",
    layout="wide",
)

# Custom CSS สำหรับปรับแต่งหน้าตาให้สวยงามและอ่านภาษาไทยสบายตา
st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.5rem;
    }
    .sub-title {
        font-size: 1.1rem;
        color: #4B5563;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #F3F4F6;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #3B82F6;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def main():
    st.markdown('<div class="main-title">📄 PDF Digger</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">ระบบแกะข้อความภาษาไทย/อังกฤษ ตาราง และรูปภาพจาก PDF และไฟล์รูปภาพ '
        '(ทั้งดิจิทัลและภาพสแกน) เป็น Word (.docx) พร้อมแปลภาษาและตรวจจับฟอนต์อัตโนมัติ</div>',
        unsafe_allow_html=True,
    )

    # --- โหลด API Key จากระบบหลังบ้าน (st.secrets หรือ .env) โดยไม่แสดงใน UI ---
    api_key = ""
    try:
        if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets:
            api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass

    if not api_key:
        api_key = os.getenv("GEMINI_API_KEY", "")

    # --- Sidebar: การตั้งค่า ---
    with st.sidebar:
        st.header("⚙️ การตั้งค่าระบบ")

        # 1. แปลภาษา
        st.subheader("🌐 การแปลภาษา (Translation)")
        translate_dict = {
            "original": "📄 คงภาษาตามต้นฉบับ (ไม่แปล)",
            "th": "🇹🇭 แปลเป็นภาษาไทย (Translate to Thai)",
            "en": "🇬🇧 แปลเป็นภาษาอังกฤษ (Translate to English)",
            "zh": "🇨🇳 แปลเป็นภาษาจีน (Translate to Chinese)",
            "ja": "🇯🇵 แปลเป็นภาษาญี่ปุ่น (Translate to Japanese)",
        }
        target_language = st.selectbox(
            "เลือกภาษาเป้าหมาย",
            options=list(translate_dict.keys()),
            format_func=lambda x: translate_dict[x],
            index=0,
            help="ระบบจะแปลเนื้อหาทั้งหมดรวมถึงตารางเป็นภาษาที่เลือกในรอบเดียวโดยยังคงโครงสร้างตารางและรูปภาพครบถ้วน",
        )

        # 2. โหมดปรับความคมชัดภาพ
        st.subheader("✨ ปรับแต่งภาพสแกน (Image Enhancement)")
        enhance_image = st.checkbox(
            "เปิดโหมดเพิ่มความคมชัดพิเศษ",
            value=False,
            help="เหมาะสำหรับภาพสแกนกระดาษเหลือง หมึกจาง หรือภาพถ่ายจากมือถือ ระบบจะปรับ Auto-contrast และ Sharpness ให้ตัวหนังสือเด่นชัดขึ้น",
        )

        # 3. แบบอักษร Word
        st.subheader("🔤 แบบอักษร Word")
        font_options = [
            "🔍 ตรวจจับจากต้นฉบับอัตโนมัติ (Auto-detect)",
            "Cordia New",
            "TH Sarabun New",
            "Angsana New",
            "Tahoma",
            "Calibri",
            "Arial",
        ]
        font_choice = st.selectbox(
            "ฟอนต์เริ่มต้นของไฟล์ .docx",
            options=font_options,
            index=0,
            help="หากเลือก Auto-detect ระบบจะดึงฟอนต์ที่ใช้จริงจาก PDF ต้นฉบับมาใช้ใน Word ทันที",
        )

    # --- Main Area ---
    if not api_key or api_key == "your_gemini_api_key_here":
        st.error("⚠️ ไม่พบการกำหนดค่า API ในระบบหลังบ้าน กรุณาตรวจสอบการตั้งค่า GEMINI_API_KEY ใน Environment หรือ Server Secrets")
        return

    allowed_types = ["pdf", "png", "jpg", "jpeg", "webp", "bmp", "tiff", "tif"]
    uploaded_file = st.file_uploader(
        "เลือกหรือลากไฟล์ PDF หรือไฟล์รูปภาพ (.png, .jpg, .webp) มาวางที่นี่",
        type=allowed_types,
    )

    if uploaded_file is not None:
        file_ext = Path(uploaded_file.name).suffix.lower()
        is_direct_image = file_ext in SUPPORTED_IMAGE_EXTENSIONS
        file_bytes = uploaded_file.getvalue()

        # Reset active job if uploaded file changes
        if st.session_state.get("active_file_name") != uploaded_file.name:
            st.session_state.active_file_name = uploaded_file.name
            st.session_state.conversion_job = None

        with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_file:
            tmp_file.write(file_bytes)
            tmp_file_path = Path(tmp_file.name)

        try:
            if is_direct_image:
                total_pages = 1
                file_type_label = "รูปภาพเดี่ยว"
            else:
                with fitz.open(tmp_file_path) as doc:
                    total_pages = len(doc)
                file_type_label = "เอกสาร PDF"

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("ชื่อไฟล์", uploaded_file.name)
            with col2:
                st.metric("ประเภท / จำนวนหน้า", f"{file_type_label} ({total_pages} หน้า)")
            with col3:
                file_size_mb = len(file_bytes) / (1024 * 1024)
                st.metric("ขนาดไฟล์", f"{file_size_mb:.2f} MB")

            # ตรวจสอบฟอนต์ต้นฉบับเบื้องต้น
            temp_processor = PDFProcessor(tmp_file_path)
            pre_detected_font = temp_processor.detect_font()
            if pre_detected_font:
                st.info(f"🔤 **ตรวจพบฟอนต์ต้นฉบับใน PDF:** `{pre_detected_font}` (ระบบจะนำฟอนต์นี้ไปจัดหน้าใน Word ให้อัตโนมัติ)")
            else:
                st.caption("ℹ️ ภาพสแกน/รูปภาพ ระบบจะวิเคราะห์ฟอนต์จากลักษณะตัวอักษร หรือใช้ Cordia New เป็นค่าเริ่มต้นที่ปลอดภัย")

            # แสดงตัวอย่างหน้าแรก
            with st.expander("🔍 ดูตัวอย่างหน้าเอกสาร (หน้า 1)", expanded=False):
                if is_direct_image:
                    st.image(file_bytes, caption=uploaded_file.name, use_container_width=True)
                else:
                    with fitz.open(tmp_file_path) as doc:
                        if len(doc) > 0:
                            page_0 = doc[0]
                            pix = page_0.get_pixmap(dpi=150)
                            img_bytes = pix.tobytes(output="png")
                            st.image(img_bytes, caption="ตัวอย่างหน้า 1", use_container_width=True)
        finally:
            if tmp_file_path.exists():
                tmp_file_path.unlink()

        st.divider()

        job = st.session_state.get("conversion_job")

        # ถ้ายังไม่มี job ที่กำลังทำงาน ให้แสดงปุ่มเริ่มแปลง
        if job is None:
            btn_label = "🚀 เริ่มแปลงเอกสารเป็น Word (.docx)"
            if target_language != "original":
                btn_label += f" [แปลเป็น: {translate_dict[target_language]}]"
            start_btn = st.button(btn_label, type="primary", use_container_width=True)

            if start_btn:
                models_to_use = FALLBACK_MODELS.copy()
                is_auto_font = font_choice.startswith("🔍")
                target_font = None if is_auto_font else font_choice

                job_input = Path(tempfile.NamedTemporaryFile(delete=False, suffix=file_ext).name)
                job_input.write_bytes(file_bytes)
                job_output = job_input.with_suffix(".docx")

                pipeline = PDFToWordPipeline(
                    api_key=api_key,
                    font_name=target_font,
                    auto_detect_font=is_auto_font,
                    models=models_to_use,
                    target_language=target_language,
                    enhance_image=enhance_image,
                )

                job_data = {
                    "pipeline": pipeline,
                    "thread": None,
                    "progress": 0,
                    "message": "กำลังเตรียมความพร้อม...",
                    "status": "running",
                    "report": None,
                    "error_msg": None,
                    "docx_bytes": None,
                    "output_filename": f"{Path(uploaded_file.name).stem}_converted.docx",
                    "job_input": job_input,
                    "job_output": job_output,
                    "balloons_shown": False,
                }

                def run_worker(j):
                    def progress_cb(cur, tot, msg):
                        pct = int((cur / tot) * 100) if tot > 0 else 0
                        j["progress"] = pct
                        j["message"] = f"[{cur}/{tot}] {msg}"

                    try:
                        rep = j["pipeline"].convert(
                            pdf_path=j["job_input"],
                            output_docx_path=j["job_output"],
                            progress_callback=progress_cb,
                            target_language=target_language,
                            enhance_image=enhance_image,
                        )
                        if j["pipeline"].is_cancelled:
                            j["status"] = "cancelled"
                            j["message"] = "🛑 ยกเลิกการแปลงเอกสารเรียบร้อยแล้ว"
                        else:
                            j["status"] = "completed"
                            j["report"] = rep
                            if j["job_output"].exists():
                                with open(j["job_output"], "rb") as f_docx:
                                    j["docx_bytes"] = f_docx.read()
                    except Exception as exc:
                        if j["pipeline"].is_cancelled:
                            j["status"] = "cancelled"
                            j["message"] = "🛑 ยกเลิกการแปลงเอกสารเรียบร้อยแล้ว"
                        else:
                            j["status"] = "error"
                            j["error_msg"] = str(exc)
                    finally:
                        try:
                            if j["job_input"].exists():
                                j["job_input"].unlink()
                            if j["job_output"].exists():
                                j["job_output"].unlink()
                        except Exception:
                            pass

                t = threading.Thread(target=run_worker, args=(job_data,), daemon=True)
                job_data["thread"] = t
                t.start()
                st.session_state.conversion_job = job_data
                st.rerun()

        else:
            # มี Job ที่กำลังทำงาน หรือเสร็จสิ้นแล้ว
            status = job["status"]

            if status in ("running", "paused"):
                st.progress(job["progress"])

                if job["pipeline"].is_paused:
                    st.warning(f"⏸️ **สถานะ: หยุดชั่วคราว** — {job['message']}")
                else:
                    st.info(f"⏳ **สถานะ: กำลังประมวลผล** — {job['message']}")

                c1, c2 = st.columns(2)
                with c1:
                    if job["pipeline"].is_paused:
                        if st.button("▶️ เริ่มทำงานต่อ (Resume)", type="primary", use_container_width=True):
                            job["pipeline"].resume()
                            job["status"] = "running"
                            st.rerun()
                    else:
                        if st.button("⏸️ หยุดชั่วคราว (Pause)", use_container_width=True):
                            job["pipeline"].pause()
                            job["status"] = "paused"
                            st.rerun()
                with c2:
                    if st.button("🛑 ยกเลิกการแปลง (Cancel)", use_container_width=True):
                        job["pipeline"].cancel()
                        job["status"] = "cancelled"
                        st.rerun()

                # Polling update
                if job["thread"] and job["thread"].is_alive():
                    time.sleep(0.5)
                    st.rerun()
                else:
                    st.rerun()

            elif status == "completed":
                report = job["report"]
                st.progress(100)
                st.success("🎉 แปลงเอกสารเป็น Word สำเร็จเรียบร้อยแล้ว!")
                if not job.get("balloons_shown", False):
                    st.balloons()
                    job["balloons_shown"] = True

                # สรุปผล
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("หน้าที่สำเร็จ", f"{report.successful_pages}/{report.total_pages} หน้า")
                with c2:
                    st.metric("รูปภาพที่แทรก", f"{report.total_images_extracted} รูป")
                with c3:
                    font_lbl = report.applied_font
                    if report.detected_font:
                        font_lbl += " (ตรวจพบ)"
                    st.metric("ฟอนต์ใน Word", font_lbl)

                if report.target_language != "original":
                    st.info(f"🌐 **เอกสารถูกแปลเป็น:** {translate_dict.get(report.target_language, report.target_language)}")

                if report.enhanced:
                    st.caption("✨ ประมวลผลด้วยโหมดปรับความคมชัดพิเศษ (Enhanced Mode)")

                # ปุ่มดาวน์โหลดไฟล์ Word
                st.download_button(
                    label=f"📥 ดาวน์โหลดไฟล์ Word: {job['output_filename']}",
                    data=job["docx_bytes"],
                    file_name=job["output_filename"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                )

                # แสดงเนื้อหาที่แกะได้แต่ละหน้า
                with st.expander("📝 ดูข้อความและโครงสร้างที่แกะได้ (Markdown Preview)"):
                    for p_res in report.page_results:
                        st.markdown(f"#### หน้า {p_res.page_num}")
                        st.markdown(p_res.markdown_text)
                        st.divider()

                if st.button("🔄 แปลงไฟล์อื่น / เริ่มต้นใหม่", use_container_width=True):
                    st.session_state.conversion_job = None
                    st.rerun()

            elif status == "cancelled":
                st.warning("🛑 ยกเลิกการแปลงเอกสารเรียบร้อยแล้ว")
                if st.button("🔄 เริ่มแปลงใหม่อีกครั้ง", type="primary", use_container_width=True):
                    st.session_state.conversion_job = None
                    st.rerun()

            elif status == "error":
                st.error(f"❌ เกิดข้อผิดพลาดระหว่างประมวลผล: {job['error_msg']}")
                if st.button("🔄 ลองใหม่อีกครั้ง", type="primary", use_container_width=True):
                    st.session_state.conversion_job = None
                    st.rerun()


if __name__ == "__main__":
    main()
