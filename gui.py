import os
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox
import customtkinter as ctk

from src.config import DEFAULT_FONT, FALLBACK_MODELS
from src.pdf_processor import PDFProcessor, SUPPORTED_IMAGE_EXTENSIONS
from src.pipeline import PDFToWordPipeline

# ตั้งค่าธีม CustomTkinter ให้เข้ากับ macOS
ctk.set_appearance_mode("System")  # System, Dark, Light
ctk.set_default_color_theme("blue")


class PDFDiggerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF Digger - ระบบแปลง PDF / รูปภาพ เป็น Word ด้วย Gemini AI")
        self.geometry("800x640")
        self.minsize(720, 560)

        self.selected_pdf: Path | None = None
        self.output_docx: Path | None = None
        self.is_converting = False

        self._setup_ui()

    def _setup_ui(self):
        # Grid layout configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # 1. Header Frame
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=24, pady=(16, 8), sticky="ew")

        title_label = ctk.CTkLabel(
            header_frame,
            text="📄 PDF Digger",
            font=ctk.CTkFont(family="Cordia New", size=28, weight="bold"),
        )
        title_label.pack(anchor="w")

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="แกะข้อความไทย-อังกฤษ ตาราง รูปภาพ แปลภาษา พร้อมตรวจจับฟอนต์ต้นฉบับอัตโนมัติ",
            font=ctk.CTkFont(family="Cordia New", size=16),
            text_color="gray",
        )
        subtitle_label.pack(anchor="w")

        # 2. File Selection Card
        file_card = ctk.CTkFrame(self, corner_radius=10)
        file_card.grid(row=1, column=0, padx=24, pady=8, sticky="ew")
        file_card.grid_columnconfigure(0, weight=1)

        file_title = ctk.CTkLabel(
            file_card,
            text="1. เลือกไฟล์เอกสาร (PDF หรือ รูปภาพ)",
            font=ctk.CTkFont(family="Cordia New", size=18, weight="bold"),
        )
        file_title.grid(row=0, column=0, columnspan=2, padx=16, pady=(10, 4), sticky="w")

        file_input_frame = ctk.CTkFrame(file_card, fg_color="transparent")
        file_input_frame.grid(row=1, column=0, columnspan=2, padx=16, pady=(0, 6), sticky="ew")
        file_input_frame.grid_columnconfigure(0, weight=1)

        self.file_path_entry = ctk.CTkEntry(
            file_input_frame,
            placeholder_text="ยังไม่ได้เลือกไฟล์... คลิกปุ่มเลือกไฟล์ด้านขวา (PDF, PNG, JPG, WEBP)",
            height=36,
            font=ctk.CTkFont(size=13),
        )
        self.file_path_entry.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        browse_btn = ctk.CTkButton(
            file_input_frame,
            text="📁 เลือกไฟล์",
            command=self._browse_pdf,
            width=130,
            height=36,
            font=ctk.CTkFont(family="Cordia New", size=16, weight="bold"),
        )
        browse_btn.grid(row=0, column=1)

        # Label รายละเอียดไฟล์ PDF ที่ตรวจพบ
        self.pdf_info_label = ctk.CTkLabel(
            file_card,
            text="💡 รองรับทั้ง PDF ดิจิทัล, เอกสารสแกน, และไฟล์รูปภาพเดี่ยว (.png, .jpg, .webp)",
            font=ctk.CTkFont(family="Cordia New", size=14),
            text_color="gray",
        )
        self.pdf_info_label.grid(row=2, column=0, columnspan=2, padx=16, pady=(0, 10), sticky="w")

        # 3. Settings Card
        settings_card = ctk.CTkFrame(self, corner_radius=10)
        settings_card.grid(row=2, column=0, padx=24, pady=8, sticky="ew")
        settings_card.grid_columnconfigure(1, weight=1)

        settings_title = ctk.CTkLabel(
            settings_card,
            text="2. การตั้งค่าระบบ (แปลภาษา, ฟอนต์)",
            font=ctk.CTkFont(family="Cordia New", size=18, weight="bold"),
        )
        settings_title.grid(row=0, column=0, columnspan=3, padx=16, pady=(10, 6), sticky="w")

        # Translation
        translate_label = ctk.CTkLabel(
            settings_card,
            text="แปลภาษา (Translation):",
            font=ctk.CTkFont(family="Cordia New", size=15),
        )
        translate_label.grid(row=1, column=0, padx=(16, 10), pady=4, sticky="w")

        self.translate_map = {
            "📄 คงภาษาตามต้นฉบับ (ไม่แปล)": "original",
            "🇹🇭 แปลเป็นภาษาไทย (Translate to Thai)": "th",
            "🇬🇧 แปลเป็นภาษาอังกฤษ (Translate to English)": "en",
            "🇨🇳 แปลเป็นภาษาจีน (Translate to Chinese)": "zh",
            "🇯🇵 แปลเป็นภาษาญี่ปุ่น (Translate to Japanese)": "ja",
        }
        self.translate_combo = ctk.CTkComboBox(
            settings_card,
            values=list(self.translate_map.keys()),
            height=30,
            font=ctk.CTkFont(size=13),
        )
        self.translate_combo.set("📄 คงภาษาตามต้นฉบับ (ไม่แปล)")
        self.translate_combo.grid(row=1, column=1, columnspan=2, padx=(0, 16), pady=4, sticky="ew")

        # Font Selection
        font_label = ctk.CTkLabel(
            settings_card,
            text="แบบอักษร Word:",
            font=ctk.CTkFont(family="Cordia New", size=15),
        )
        font_label.grid(row=2, column=0, padx=(16, 10), pady=4, sticky="w")

        font_options = [
            "🔍 ตรวจจับจากต้นฉบับอัตโนมัติ (Auto-detect)",
            "Cordia New",
            "TH Sarabun New",
            "Angsana New",
            "Tahoma",
            "Calibri",
            "Arial",
        ]
        self.font_combo = ctk.CTkComboBox(
            settings_card,
            values=font_options,
            height=30,
            font=ctk.CTkFont(size=13),
        )
        self.font_combo.set(font_options[0])
        self.font_combo.grid(row=2, column=1, columnspan=2, padx=(0, 16), pady=4, sticky="ew")

        # Image Enhancement Checkbox
        self.enhance_var = ctk.BooleanVar(value=False)
        self.enhance_chk = ctk.CTkCheckBox(
            settings_card,
            text="✨ เปิดโหมดเพิ่มความคมชัดพิเศษสำหรับภาพสแกน/ภาพถ่าย (Auto-contrast & Sharpening)",
            variable=self.enhance_var,
            font=ctk.CTkFont(family="Cordia New", size=14),
        )
        self.enhance_chk.grid(row=3, column=0, columnspan=3, padx=(16, 16), pady=(6, 10), sticky="w")

        # 4. Action & Progress Area
        action_card = ctk.CTkFrame(self, corner_radius=10)
        action_card.grid(row=3, column=0, padx=24, pady=8, sticky="nsew")
        action_card.grid_columnconfigure(0, weight=1)

        self.start_btn = ctk.CTkButton(
            action_card,
            text="🚀 เริ่มแปลงเอกสารเป็น Word (.docx)",
            command=self._start_conversion_thread,
            height=44,
            font=ctk.CTkFont(family="Cordia New", size=18, weight="bold"),
        )
        self.start_btn.pack(fill="x", padx=16, pady=(14, 8))

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(action_card, height=12)
        self.progress_bar.pack(fill="x", padx=16, pady=(4, 6))
        self.progress_bar.set(0)

        # Status text
        self.status_label = ctk.CTkLabel(
            action_card,
            text="พร้อมเริ่มการทำงาน",
            font=ctk.CTkFont(family="Cordia New", size=15),
            text_color="gray",
        )
        self.status_label.pack(anchor="w", padx=16, pady=(2, 4))

        # Action buttons frame (appear after success)
        self.result_frame = ctk.CTkFrame(action_card, fg_color="transparent")
        self.result_frame.pack(fill="x", padx=16, pady=(2, 12))
        self.result_frame.grid_columnconfigure(0, weight=1)
        self.result_frame.grid_columnconfigure(1, weight=1)

        self.open_doc_btn = ctk.CTkButton(
            self.result_frame,
            text="📝 เปิดไฟล์ Word ทันที",
            command=self._open_output_docx,
            height=36,
            font=ctk.CTkFont(family="Cordia New", size=15, weight="bold"),
            fg_color="#10B981",
            hover_color="#059669",
        )

        self.open_folder_btn = ctk.CTkButton(
            self.result_frame,
            text="📂 เปิดโฟลเดอร์ไฟล์",
            command=self._open_output_folder,
            height=36,
            font=ctk.CTkFont(family="Cordia New", size=15),
            fg_color="gray40",
            hover_color="gray50",
        )

    def _browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ PDF หรือไฟล์รูปภาพ",
            filetypes=[
                ("ไฟล์เอกสารและรูปภาพ", "*.pdf *.png *.jpg *.jpeg *.webp *.bmp *.tiff"),
                ("PDF Files", "*.pdf"),
                ("Image Files", "*.png *.jpg *.jpeg *.webp *.bmp *.tiff"),
                ("All Files", "*.*"),
            ],
        )
        if not file_path:
            return

        self.selected_pdf = Path(file_path)
        self.file_path_entry.delete(0, "end")
        self.file_path_entry.insert(0, str(self.selected_pdf))

        # ตรวจสอบจำนวนหน้าและฟอนต์เบื้องต้น
        try:
            processor = PDFProcessor(self.selected_pdf)
            total_pages = processor.get_page_count()
            detected_font = processor.detect_font()

            type_desc = "รูปภาพเดี่ยว" if processor.is_image else f"{total_pages} หน้า"
            info_text = f"📄 ประเภท: {type_desc} | "
            if detected_font:
                info_text += f"🔤 ตรวจพบฟอนต์ต้นฉบับ: {detected_font}"
            else:
                info_text += "ℹ️ ตรวจไม่พบฟอนต์ดิจิทัล (จะใช้ฟอนต์ที่กำหนดหรือ Cordia New)"

            self.pdf_info_label.configure(text=info_text, text_color="#2563EB")
        except Exception as e:
            self.pdf_info_label.configure(text=f"ข้อผิดพลาดในการอ่านไฟล์: {e}", text_color="red")

    def _start_conversion_thread(self):
        if self.is_converting:
            return

        if not self.selected_pdf or not self.selected_pdf.exists():
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกไฟล์ PDF หรือไฟล์รูปภาพก่อน")
            return

        api_key = os.getenv("GEMINI_API_KEY", "").strip()
        if not api_key or api_key == "your_gemini_api_key_here":
            messagebox.showerror(
                "แจ้งเตือน",
                "ไม่พบการตั้งค่า API Key ในระบบหลังบ้าน\nกรุณาตั้งค่า GEMINI_API_KEY ในไฟล์ .env หรือ System Environment",
            )
            return

        self.is_converting = True
        self.start_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="กำลังเริ่มต้นกระบวนการ...", text_color="#3B82F6")
        self.open_doc_btn.grid_forget()
        self.open_folder_btn.grid_forget()

        # เริ่มทำงานใน Background Thread
        thread = threading.Thread(target=self._run_conversion, daemon=True)
        thread.start()

    def _run_conversion(self):
        api_key = os.getenv("GEMINI_API_KEY", "").strip()
        models_to_use = FALLBACK_MODELS.copy()
        font_choice = self.font_combo.get()
        target_lang = self.translate_map.get(self.translate_combo.get(), "original")
        enhance = self.enhance_var.get()

        is_auto_font = font_choice.startswith("🔍")
        target_font = None if is_auto_font else font_choice

        self.output_docx = self.selected_pdf.with_suffix(".docx")

        def progress_callback(current: int, total: int, msg: str):
            progress_ratio = current / max(total, 1)
            self.after(0, lambda: self._update_progress_ui(progress_ratio, current, total, msg))

        try:
            pipeline = PDFToWordPipeline(
                api_key=api_key,
                font_name=target_font,
                auto_detect_font=is_auto_font,
                models=models_to_use,
                target_language=target_lang,
                enhance_image=enhance,
            )

            report = pipeline.convert(
                pdf_path=self.selected_pdf,
                output_docx_path=self.output_docx,
                progress_callback=progress_callback,
                target_language=target_lang,
                enhance_image=enhance,
            )

            self.after(0, lambda: self._on_success(report))

        except Exception as e:
            self.after(0, lambda: self._on_error(str(e)))

    def _update_progress_ui(self, ratio: float, current: int, total: int, msg: str):
        self.progress_bar.set(ratio)
        self.status_label.configure(text=f"⏳ [{current}/{total}] {msg}", text_color="#2563EB")

    def _on_success(self, report):
        self.is_converting = False
        self.start_btn.configure(state="normal")
        self.progress_bar.set(1.0)

        font_info = report.applied_font
        if report.detected_font:
            font_info += f" (ตรวจพบจากต้นฉบับ)"

        lang_str = f" | แปลภาษา: {report.target_language}" if report.target_language != "original" else ""
        enhance_str = " | โหมดภาพคมชัด" if report.enhanced else ""

        success_text = (
            f"🎉 แปลงสำเร็จ! ({report.successful_pages}/{report.total_pages} หน้า) | "
            f"ฟอนต์: {font_info} | รูปภาพ: {report.total_images_extracted} รูป{lang_str}{enhance_str}"
        )
        self.status_label.configure(text=success_text, text_color="#059669")

        # แสดงปุ่มเปิดไฟล์
        self.open_doc_btn.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        self.open_folder_btn.grid(row=0, column=1, padx=(8, 0), sticky="ew")

    def _on_error(self, err_msg: str):
        self.is_converting = False
        self.start_btn.configure(state="normal")
        self.status_label.configure(text=f"❌ เกิดข้อผิดพลาด: {err_msg}", text_color="red")
        messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถแปลงไฟล์ได้:\n{err_msg}")

    def _open_output_docx(self):
        if self.output_docx and self.output_docx.exists():
            subprocess.run(["open", str(self.output_docx)])

    def _open_output_folder(self):
        if self.output_docx and self.output_docx.exists():
            subprocess.run(["open", "-R", str(self.output_docx)])


if __name__ == "__main__":
    app = PDFDiggerApp()
    app.mainloop()
