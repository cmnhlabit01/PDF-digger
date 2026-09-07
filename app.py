import io
import os
import tempfile
from pathlib import Path
from PIL import Image
import streamlit as st
import pymupdf as fitz

from src.config import DEFAULT_FONT, FALLBACK_MODELS
from src.pdf_processor import PDFProcessor
from src.pipeline import PDFToWordPipeline

st.set_page_config(
    page_title="PDF Digger - แกะ PDF ภาษาไทย/อังกฤษ เป็น Word",
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
        '<div class="sub-title">ระบบแกะข้อความภาษาไทย/อังกฤษ ตาราง และรูปภาพจาก PDF '
        '(ทั้งดิจิทัลและภาพสแกน) เป็น Word (.docx) พร้อมตรวจจับฟอนต์ต้นฉบับอัตโนมัติ</div>',
        unsafe_allow_html=True,
    )

    # --- Sidebar: การตั้งค่า ---
    with st.sidebar:
        st.header("⚙️ การตั้งค่าระบบ")

        # 1. API Key (รองรับ st.secrets บน Streamlit Cloud หรือไฟล์ .env)
        embedded_key = ""
        try:
            if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets:
                embedded_key = st.secrets["GEMINI_API_KEY"]
        except Exception:
            pass

        if not embedded_key:
            embedded_key = os.getenv("GEMINI_API_KEY", "")

        has_valid_key = bool(embedded_key and embedded_key != "your_gemini_api_key_here")

        if has_valid_key:
            st.success("🔒 ระบบเชื่อมต่อ Gemini AI เรียบร้อยแล้ว (พร้อมใช้งาน)")
            with st.expander("⚙️ ต้องการเปลี่ยน API Key อื่น? (ค่าเริ่มต้นถูกฝังไว้แล้ว)"):
                api_key_input = st.text_input(
                    "ระบุ API Key ใหม่",
                    value="",
                    type="password",
                    help="หากเว้นว่างไว้ ระบบจะใช้คีย์หลักที่ฝังไว้ในระบบ",
                )
                api_key = api_key_input.strip() if api_key_input.strip() else embedded_key
        else:
            api_key_input = st.text_input(
                "กรอก Gemini API Key",
                value="",
                type="password",
                help="รับฟรีได้จาก https://aistudio.google.com/",
            )
            api_key = api_key_input.strip()

        # 2. เลือกลำดับโมเดล
        st.subheader("🤖 โมเดล AI")
        primary_model = st.selectbox(
            "โมเดลเริ่มต้นที่ต้องการใช้",
            options=FALLBACK_MODELS,
            index=0,
            help="ระบบจะใช้โมเดลนี้ก่อน หากติด 429 Quota Exceeded จะสลับไปตัวถัดไปให้อัตโนมัติ",
        )

        st.info(
            "🔄 **Seamless Fallback:**\n"
            f"หากโมเดลหลักโควตาเต็ม ระบบจะสลับไป `{FALLBACK_MODELS[1]}` ➔ `{FALLBACK_MODELS[2]}` ให้อัตโนมัติทันที"
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
        st.warning("⚠️ กรุณาระบุ **Gemini API Key** ในแถบด้านซ้ายก่อนเริ่มใช้งาน")
        st.info("💡 คุณสามารถขอรับ Gemini API Key ฟรีได้จาก [Google AI Studio](https://aistudio.google.com/)")
        return

    uploaded_file = st.file_uploader("เลือกหรือลากไฟล์ PDF มาวางที่นี่", type=["pdf"])

    if uploaded_file is not None:
        # บันทึกไฟล์ชั่วคราวเพื่ออ่านด้วย fitz
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.read())
            tmp_pdf_path = Path(tmp_pdf.name)

        doc = fitz.open(tmp_pdf_path)
        total_pages = len(doc)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("ชื่อไฟล์", uploaded_file.name)
        with col2:
            st.metric("จำนวนหน้าทั้งหมด", f"{total_pages} หน้า")
        with col3:
            file_size_mb = uploaded_file.size / (1024 * 1024)
            st.metric("ขนาดไฟล์", f"{file_size_mb:.2f} MB")

        # ตรวจสอบฟอนต์ต้นฉบับเบื้องต้น
        temp_processor = PDFProcessor(tmp_pdf_path)
        pre_detected_font = temp_processor.detect_font()
        if pre_detected_font:
            st.info(f"🔤 **ตรวจพบฟอนต์ต้นฉบับใน PDF:** `{pre_detected_font}` (ระบบจะนำฟอนต์นี้ไปจัดหน้าใน Word ให้อัตโนมัติ)")
        else:
            st.caption("ℹ️ เอกสารอาจเป็นภาพสแกน ระบบจะวิเคราะห์ฟอนต์จากลักษณะตัวอักษร หรือใช้ Cordia New เป็นค่าเริ่มต้น")

        # แสดงตัวอย่างหน้าแรก
        with st.expander("🔍 ดูตัวอย่างหน้าเอกสาร PDF (หน้า 1)", expanded=False):
            page_0 = doc[0]
            pix = page_0.get_pixmap(dpi=150)
            img_bytes = pix.tobytes(output="png")
            st.image(img_bytes, caption="ตัวอย่างหน้า 1", use_container_width=True)

        doc.close()

        st.divider()

        # ปุ่มเริ่มการแปลงไฟล์
        start_btn = st.button("🚀 เริ่มแปลงเอกสารเป็น Word (.docx)", type="primary", use_container_width=True)

        if start_btn:
            # จัดเตรียมลำดับโมเดลตามที่เลือก
            models_to_use = FALLBACK_MODELS.copy()
            if primary_model in models_to_use:
                models_to_use.remove(primary_model)
            models_to_use.insert(0, primary_model)

            progress_bar = st.progress(0)
            status_text = st.empty()
            fallback_alerts = st.empty()

            is_auto_font = font_choice.startswith("🔍")
            target_font = None if is_auto_font else font_choice

            output_docx_path = tmp_pdf_path.with_suffix(".docx")

            def update_progress_ui(current: int, total: int, msg: str):
                progress_bar.progress(int((current / total) * 100))
                status_text.info(f"⏳ **[{current}/{total}]** {msg}")

            try:
                pipeline = PDFToWordPipeline(
                    api_key=api_key,
                    font_name=target_font,
                    auto_detect_font=is_auto_font,
                    models=models_to_use,
                )

                report = pipeline.convert(
                    pdf_path=tmp_pdf_path,
                    output_docx_path=output_docx_path,
                    progress_callback=update_progress_ui,
                )

                progress_bar.progress(100)
                status_text.success("🎉 แปลงเอกสารเป็น Word สำเร็จเรียบร้อยแล้ว!")
                st.balloons()

                # สรุปผล
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("หน้าที่สำเร็จ", f"{report.successful_pages}/{report.total_pages} หน้า")
                with c2:
                    st.metric("รูปภาพที่แทรก", f"{report.total_images_extracted} รูป")
                with c3:
                    st.metric("โมเดลที่ใช้", ", ".join(report.models_used))
                with c4:
                    font_lbl = report.applied_font
                    if report.detected_font:
                        font_lbl += " (ตรวจพบ)"
                    st.metric("ฟอนต์ใน Word", font_lbl)

                if report.fallback_events:
                    st.warning("⚠️ **บันทึกการสลับโมเดลสำรอง (Seamless Fallback):**")
                    for rec in report.fallback_events:
                        st.write(f"- {rec}")

                # อ่านไฟล์ Word ที่สร้างขึ้น
                with open(output_docx_path, "rb") as f_docx:
                    docx_bytes = f_docx.read()

                # ปุ่มดาวน์โหลดไฟล์ Word
                output_filename = f"{Path(uploaded_file.name).stem}_converted.docx"
                st.download_button(
                    label=f"📥 ดาวน์โหลดไฟล์ Word: {output_filename}",
                    data=docx_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                )

                # แสดงเนื้อหาที่แกะได้แต่ละหน้า
                with st.expander("📝 ดูข้อความและโครงสร้างที่แกะได้ (Markdown Preview)"):
                    for p_res in report.page_results:
                        st.markdown(f"#### หน้า {p_res.page_num} (แกะด้วย {p_res.model_used})")
                        st.markdown(p_res.markdown_text)
                        st.divider()

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดระหว่างประมวลผล: {e}")
            finally:
                if tmp_pdf_path.exists():
                    tmp_pdf_path.unlink()
                if output_docx_path.exists():
                    output_docx_path.unlink()


if __name__ == "__main__":
    main()
