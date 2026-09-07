import os
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox
import customtkinter as ctk

from src.config import DEFAULT_FONT, FALLBACK_MODELS
from src.pdf_processor import PDFProcessor
from src.pipeline import PDFToWordPipeline

# ตั้งค่าธีม CustomTkinter ให้เข้ากับ macOS
ctk.set_appearance_mode("System")  # System, Dark, Light
ctk.set_default_color_theme("blue")


class PDFDiggerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF Digger - ระบบแปลง PDF เป็น Word ด้วย Gemini AI")
        self.geometry("780x680")
        self.minsize(700, 600)

        self.selected_pdf: Path | None = None
        self.output_docx: Path | None = None
        self.is_converting = False

        self._setup_ui()
        self._load_saved_api_key()

    def _setup_ui(self):
        # Grid layout configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # 1. Header Frame
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=24, pady=(20, 10), sticky="ew")

        title_label = ctk.CTkLabel(
            header_frame,
            text="📄 PDF Digger",
            font=ctk.CTkFont(family="Cordia New", size=28, weight="bold"),
        )
        title_label.pack(anchor="w")

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="แกะข้อความไทย-อังกฤษ ตาราง รูปภาพ พร้อมตรวจจับฟอนต์ต้นฉบับอัตโนมัติ",
            font=ctk.CTkFont(family="Cordia New", size=16),
            text_color="gray",
        )
        subtitle_label.pack(anchor="w")

        # 2. File Selection Card
        file_card = ctk.CTkFrame(self, corner_radius=10)
        file_card.grid(row=1, column=0, padx=24, pady=10, sticky="ew")
        file_card.grid_columnconfigure(0, weight=1)

        file_title = ctk.CTkLabel(
            file_card,
            text="1. เลือกไฟล์เอกสาร PDF",
            font=ctk.CTkFont(family="Cordia New", size=18, weight="bold"),
        )
        file_title.grid(row=0, column=0, columnspan=2, padx=16, pady=(12, 6), sticky="w")

        file_input_frame = ctk.CTkFrame(file_card, fg_color="transparent")
        file_input_frame.grid(row=1, column=0, columnspan=2, padx=16, pady=(0, 8), sticky="ew")
        file_input_frame.grid_columnconfigure(0, weight=1)

        self.file_path_entry = ctk.CTkEntry(
            file_input_frame,
            placeholder_text="ยังไม่ได้เลือกไฟล์... คลิกปุ่มเลือกไฟล์ด้านขวา",
            height=36,
            font=ctk.CTkFont(size=13),
        )
        self.file_path_entry.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        browse_btn = ctk.CTkButton(
            file_input_frame,
            text="📁 เลือกไฟล์ PDF",
            command=self._browse_pdf,
            width=130,
            height=36,
            font=ctk.CTkFont(family="Cordia New", size=16, weight="bold"),
        )
        browse_btn.grid(row=0, column=1)

        # Label รายละเอียดไฟล์ PDF ที่ตรวจพบ
        self.pdf_info_label = ctk.CTkLabel(
            file_card,
            text="💡 รองรับทั้ง PDF ที่แปลงจาก Word และภาพสแกน/รูปถ่าย",
            font=ctk.CTkFont(family="Cordia New", size=14),
            text_color="gray",
        )
        self.pdf_info_label.grid(row=2, column=0, columnspan=2, padx=16, pady=(0, 12), sticky="w")

        # 3. Settings Card
        settings_card = ctk.CTkFrame(self, corner_radius=10)
        settings_card.grid(row=2, column=0, padx=24, pady=10, sticky="ew")
        settings_card.grid_columnconfigure(1, weight=1)

        settings_title = ctk.CTkLabel(
            settings_card,
            text="2. การตั้งค่าระบบ (API Key & ฟอนต์)",
            font=ctk.CTkFont(family="Cordia New", size=18, weight="bold"),
        )
        settings_title.grid(row=0, column=0, columnspan=3, padx=16, pady=(12, 8), sticky="w")

        # API Key
        api_label = ctk.CTkLabel(
            settings_card,
            text="Gemini API Key:",
            font=ctk.CTkFont(family="Cordia New", size=15),
        )
        api_label.grid(row=1, column=0, padx=(16, 10), pady=6, sticky="w")

        self.api_key_entry = ctk.CTkEntry(
            settings_card,
            placeholder_text="AIzaSy... (รับฟรีได้จาก https://aistudio.google.com/)",
            show="•",
            height=32,
            font=ctk.CTkFont(size=12),
        )
        self.api_key_entry.grid(row=1, column=1, padx=(0, 10), pady=6, sticky="ew")

        save_key_btn = ctk.CTkButton(
            settings_card,
            text="บันทึกคีย์",
            command=self._save_api_key,
            width=80,
            height=32,
            font=ctk.CTkFont(family="Cordia New", size=14),
            fg_color="gray30",
            hover_color="gray40",
        )
        save_key_btn.grid(row=1, column=2, padx=(0, 16), pady=6)

        # Model Selection
        model_label = ctk.CTkLabel(
            settings_card,
            text="โมเดลหลัก (Fallback):",
            font=ctk.CTkFont(family="Cordia New", size=15),
        )
        model_label.grid(row=2, column=0, padx=(16, 10), pady=6, sticky="w")

        self.model_combo = ctk.CTkComboBox(
            settings_card,
            values=FALLBACK_MODELS,
            height=32,
            font=ctk.CTkFont(size=13),
        )
        self.model_combo.set(FALLBACK_MODELS[0])
        self.model_combo.grid(row=2, column=1, columnspan=2, padx=(0, 16), pady=6, sticky="ew")

        # Font Selection
        font_label = ctk.CTkLabel(
            settings_card,
            text="แบบอักษร Word:",
            font=ctk.CTkFont(family="Cordia New", size=15),
        )
        font_label.grid(row=3, column=0, padx=(16, 10), pady=(6, 14), sticky="w")

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
            height=32,
            font=ctk.CTkFont(size=13),
        )
        self.font_combo.set(font_options[0])
        self.font_combo.grid(row=3, column=1, columnspan=2, padx=(0, 16), pady=(6, 14), sticky="ew")

        # 4. Action & Progress Area
        action_card = ctk.CTkFrame(self, corner_radius=10)
        action_card.grid(row=3, column=0, padx=24, pady=10, sticky="nsew")
        action_card.grid_columnconfigure(0, weight=1)

        self.start_btn = ctk.CTkButton(
            action_card,
            text="🚀 เริ่มแปลงเอกสารเป็น Word (.docx)",
            command=self._start_conversion_thread,
            height=46,
            font=ctk.CTkFont(family="Cordia New", size=18, weight="bold"),
        )
        self.start_btn.pack(fill="x", padx=16, pady=(16, 10))

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(action_card, height=12)
        self.progress_bar.pack(fill="x", padx=16, pady=(6, 6))
        self.progress_bar.set(0)

        # Status text
        self.status_label = ctk.CTkLabel(
            action_card,
            text="พร้อมเริ่มการทำงาน",
            font=ctk.CTkFont(family="Cordia New", size=15),
            text_color="gray",
        )
        self.status_label.pack(anchor="w", padx=16, pady=(2, 6))

        # Fallback notification text
        self.fallback_label = ctk.CTkLabel(
            action_card,
            text="",
            font=ctk.CTkFont(family="Cordia New", size=14),
            text_color="#F59E0B",
        )
        self.fallback_label.pack(anchor="w", padx=16, pady=(0, 10))

        # Action buttons frame (appear after success)
        self.result_frame = ctk.CTkFrame(action_card, fg_color="transparent")
        self.result_frame.pack(fill="x", padx=16, pady=(4, 16))
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

    def _load_saved_api_key(self):
        """โหลด API Key จากไฟล์ .env หากมี"""
        env_key = os.getenv("GEMINI_API_KEY", "")
        if env_key and env_key != "your_gemini_api_key_here":
            self.api_key_entry.delete(0, "end")
            self.api_key_entry.insert(0, env_key)

    def _save_api_key(self):
        """บันทึก API Key ลงใน .env"""
        key = self.api_key_entry.get().strip()
        if not key:
            messagebox.showwarning("คำเตือน", "กรุณากรอก API Key ก่อนบันทึก")
            return

        env_path = Path(".env")
        lines = []
        if env_path.exists():
            for line in env_path.read_text(encoding="utf-8").splitlines():
                if line.startswith("GEMINI_API_KEY="):
                    continue
                lines.append(line)
        lines.append(f"GEMINI_API_KEY={key}")
        env_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        os.environ["GEMINI_API_KEY"] = key
        messagebox.showinfo("สำเร็จ", "บันทึก Gemini API Key เรียบร้อยแล้ว")

    def _browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
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

            info_text = f"📄 จำนวน: {total_pages} หน้า | "
            if detected_font:
                info_text += f"🔤 ตรวจพบฟอนต์ต้นฉบับ: {detected_font}"
            else:
                info_text += "ℹ️ ตรวจไม่พบฟอนต์ดิจิทัล (อาจเป็นภาพสแกน)"

            self.pdf_info_label.configure(text=info_text, text_color="#2563EB")
        except Exception as e:
            self.pdf_info_label.configure(text=f"ข้อผิดพลาดในการอ่านไฟล์: {e}", text_color="red")

    def _start_conversion_thread(self):
        if self.is_converting:
            return

        if not self.selected_pdf or not self.selected_pdf.exists():
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกไฟล์ PDF ที่ต้องการแปลงก่อน")
            return

        api_key = self.api_key_entry.get().strip()
        if not api_key or api_key == "your_gemini_api_key_here":
            messagebox.showwarning(
                "แจ้งเตือน",
                "กรุณาระบุ Gemini API Key ก่อนเริ่มใช้งาน\n(สามารถขอรับคีย์ฟรีได้ที่ https://aistudio.google.com/)",
            )
            return

        self.is_converting = True
        self.start_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="กำลังเริ่มต้นกระบวนการ...", text_color="#3B82F6")
        self.fallback_label.configure(text="")
        self.open_doc_btn.grid_forget()
        self.open_folder_btn.grid_forget()

        # เริ่มทำงานใน Background Thread
        thread = threading.Thread(target=self._run_conversion, daemon=True)
        thread.start()

    def _run_conversion(self):
        api_key = self.api_key_entry.get().strip()
        primary_model = self.model_combo.get()
        font_choice = self.font_combo.get()

        models_to_use = FALLBACK_MODELS.copy()
        if primary_model in models_to_use:
            models_to_use.remove(primary_model)
        models_to_use.insert(0, primary_model)

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
            )

            report = pipeline.convert(
                pdf_path=self.selected_pdf,
                output_docx_path=self.output_docx,
                progress_callback=progress_callback,
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

        success_text = (
            f"🎉 แปลงสำเร็จ! ({report.successful_pages}/{report.total_pages} หน้า) | "
            f"ฟอนต์: {font_info} | รูปภาพ: {report.total_images_extracted} รูป"
        )
        self.status_label.configure(text=success_text, text_color="#059669")

        if report.fallback_events:
            fb_text = "🔄 มีการสลับโมเดลอัตโนมัติ: " + " | ".join(report.fallback_events)
            self.fallback_label.configure(text=fb_text)

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
