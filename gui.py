# gui.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from pathlib import Path
from pdf_tools import merge_pdfs, split_pdf, compress_pdf, pdf_to_docx, pdf_to_excel, pdf_to_pptx, pdf_to_images, images_to_pdf, rotate_pdf, protect_pdf, unlock_pdf, add_text_watermark, ocr_pdf
from office_tools import convert_with_soffice

class FileToolboxGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FileToolbox")
        self.geometry("400x250")
        self.resizable(False, False)
        self.create_widgets()

    def create_widgets(self):
        # ---------- SCROLLABLE AREA ----------
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
        vscroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)

        vscroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Create an inner frame *inside* the canvas
        self.inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=self.inner, anchor="nw")

        # Resize the inner frame whenever canvas size changes
        def _on_frame_config(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Stretch inner frame’s width to canvas width
            canvas.itemconfigure(inner_id, width=canvas.winfo_width())

        self.inner.bind("<Configure>", _on_frame_config)

        # ---------- WIDGETS GO INSIDE self.inner ----------
        style = ttk.Style(self)
        style.configure("TButton", padding=6, font=("Segoe UI", 10))

        ttk.Label(self.inner, text="PDF Tools", font=("Segoe UI", 12, "bold")).pack(pady=10)
        ttk.Button(self.inner, text="Merge PDFs",  width=30, command=self.merge_pdfs_dialog).pack(pady=3)
        ttk.Button(self.inner, text="Split PDF",   width=30, command=self.split_pdf_dialog).pack(pady=3)
        ttk.Button(self.inner, text="Compress PDF",width=30, command=self.compress_pdf_dialog).pack(pady=3)

        ttk.Separator(self.inner, orient="horizontal").pack(fill="x", pady=8)

        ttk.Label(self.inner, text="Office ↔ PDF", font=("Segoe UI", 11, "bold")).pack(pady=6)
        ttk.Button(self.inner, text="Word → PDF",        width=30,
                command=lambda: self.office_convert_dialog("pdf",  [("Word files", "*.docx;*.doc")])).pack(pady=3)
        ttk.Button(self.inner, text="Excel → PDF",       width=30,
                command=lambda: self.office_convert_dialog("pdf",  [("Excel files", "*.xlsx;*.xls")])).pack(pady=3)
        ttk.Button(self.inner, text="PowerPoint → PDF",  width=30,
                command=lambda: self.office_convert_dialog("pdf",  [("PowerPoint files", "*.pptx;*.ppt")])).pack(pady=3)
        # ttk.Button(self.inner, text="PDF → Word",        width=30,
        #         command=lambda: self.office_convert_dialog("docx", [("PDF files", "*.pdf")])).pack(pady=3)

        ttk.Separator(self.inner, orient="horizontal").pack(fill="x", pady=8)

        ttk.Label(self.inner, text="PDF ↔ Office", font=("Segoe UI", 11, "bold")).pack(pady=6)
        ttk.Button(self.inner, text="PDF → Word", width=30,
           command=self.pdf_to_word_dialog).pack(pady=3)
        ttk.Button(self.inner, text="PDF → Excel", width=30,
           command=self.pdf_to_excel_dialog).pack(pady=3)
        ttk.Button(self.inner, text="PDF → PowerPoint", width=30,
           command=self.pdf_to_pptx_dialog).pack(pady=3)
        
        ttk.Separator(self.inner, orient="horizontal").pack(fill="x", pady=8)

        ttk.Label(self.inner, text="PDF ↔ image", font=("Segoe UI", 11, "bold")).pack(pady=6)
        ttk.Button(self.inner, text="PDF → JPG", width=30,
           command=self.pdf_to_jpg_dialog).pack(pady=3)
        ttk.Button(self.inner, text="JPG → PDF", width=30,
            command=self.jpg_to_pdf_dialog).pack(pady=3)

        ttk.Separator(self.inner, orient="horizontal").pack(fill="x", pady=8)

        ttk.Label(self.inner, text="PDF Utilities", font=("Segoe UI", 11, "bold")).pack(pady=10)
        ttk.Button(self.inner, text="Rotate PDF", width=30,
                command=self.rotate_pdf_dialog).pack(pady=3)
        ttk.Button(self.inner, text="Protect PDF (Encrypt)", width=30,
                command=self.protect_pdf_dialog).pack(pady=3)
        ttk.Button(self.inner, text="Unlock PDF (Decrypt)", width=30,
                command=self.unlock_pdf_dialog).pack(pady=3)
        ttk.Button(self.inner, text="Watermark PDF", width=30,
           command=self.watermark_pdf_dialog).pack(pady=3)
        ttk.Button(self.inner, text="OCR Scanned PDF", width=30,
           command=self.ocr_pdf_dialog).pack(pady=3)


        # -- add more buttons below as you implement functions --

        # Give mouse‑wheel scrolling
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * int(e.delta / 120), "units"))

    # ----- Actions -----
    def merge_pdfs_dialog(self):
        files = filedialog.askopenfilenames(
            title="Select PDF files to merge",
            filetypes=[("PDF files", "*.pdf")])
        if not files:
            return
        output = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save merged PDF as")
        if not output:
            return
        try:
            merge_pdfs([Path(f) for f in files], output)
            messagebox.showinfo("Success", f"Merged PDF saved to:\n{output}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def split_pdf_dialog(self):
        src = filedialog.askopenfilename(
            title="Select a PDF to split",
            filetypes=[("PDF files", "*.pdf")])
        if not src:
            return
        outdir = filedialog.askdirectory(title="Choose output folder for pages")
        if not outdir:
            return
        try:
            pages = split_pdf(Path(src), outdir)
            messagebox.showinfo("Success", f"Created {len(pages)} separate PDFs in:\n{outdir}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def compress_pdf_dialog(self):
        input_file = filedialog.askopenfilename(
            title="Select a PDF to compress",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not input_file:
            return

        # Ask for target size
        target_str = simpledialog.askstring(
            "Target size",
            "Desired maximum size in MB (leave blank for default compression):",
            parent=self
        )
        if target_str is None:  # user cancelled
            return
        try:
            target_mb = float(target_str) if target_str.strip() else None
        except ValueError:
            messagebox.showerror("Invalid input", "Please enter a numeric value.")
            return

        output_file = filedialog.asksaveasfilename(
            title="Save compressed PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_file:
            return

        try:
            compress_pdf(input_file, output_file, target_mb)
            messagebox.showinfo(
                "Success",
                f"Compressed PDF saved to:\n{output_file}"
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def office_convert_dialog(self, target_ext: str, filetypes):
        src = filedialog.askopenfilename(title="Choose file", filetypes=filetypes)
        if not src:
            return
        try:
            result = convert_with_soffice(src, target_ext)
            messagebox.showinfo("Success", f"Saved as:\n{result}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def pdf_to_word_dialog(self):
        input_file = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not input_file:
            return

        output_file = filedialog.asksaveasfilename(
            title="Save Word file as",
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")]
        )
        if not output_file:
            return

        try:
            pdf_to_docx(input_file, output_file)
            messagebox.showinfo("Success", f"Saved Word document:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    def pdf_to_excel_dialog(self):
        input_file = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not input_file:
            return

        output_file = filedialog.asksaveasfilename(
            title="Save Excel file as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not output_file:
            return

        try:
            pdf_to_excel(input_file, output_file)
            messagebox.showinfo("Success", f"Saved Excel file:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    def pdf_to_pptx_dialog(self):
        input_file = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not input_file:
            return

        output_file = filedialog.asksaveasfilename(
            title="Save PowerPoint file as",
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx")]
        )
        if not output_file:
            return

        try:
            pdf_to_pptx(input_file, output_file)
            messagebox.showinfo("Success", f"Saved PowerPoint:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def pdf_to_jpg_dialog(self):
        input_file = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not input_file:
            return

        output_folder = filedialog.askdirectory(
            title="Select folder to save images"
        )
        if not output_folder:
            return

        try:
            image_paths = pdf_to_images(input_file, output_folder)
            messagebox.showinfo("Success", f"Saved {len(image_paths)} images to:\n{output_folder}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


    def jpg_to_pdf_dialog(self):
        image_files = filedialog.askopenfilenames(
            title="Select JPG images",
            filetypes=[("JPEG images", "*.jpg *.jpeg")],
            multiple=True
        )
        if not image_files:
            return

        output_pdf = filedialog.asksaveasfilename(
            title="Save as PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_pdf:
            return

        try:
            images_to_pdf(list(image_files), output_pdf)
            messagebox.showinfo("Success", f"PDF created:\n{output_pdf}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def rotate_pdf_dialog(self):
        input_file = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf")])
        if not input_file:
            return
        output_file = filedialog.asksaveasfilename(title="Save rotated PDF", defaultextension=".pdf")
        if not output_file:
            return
        angle = simpledialog.askinteger("Rotate", "Enter rotation angle (90, 180, 270):", minvalue=0, maxvalue=360)
        if not angle:
            return
        try:
            rotate_pdf(input_file, output_file, angle)
            messagebox.showinfo("Success", f"Rotated PDF saved:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


    def protect_pdf_dialog(self):
        input_file = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf")])
        if not input_file:
            return
        output_file = filedialog.asksaveasfilename(title="Save encrypted PDF", defaultextension=".pdf")
        if not output_file:
            return
        password = simpledialog.askstring("Set Password", "Enter password to encrypt PDF:", show="*")
        if not password:
            return
        try:
            protect_pdf(input_file, output_file, password)
            messagebox.showinfo("Success", f"Protected PDF saved:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


    def unlock_pdf_dialog(self):
        input_file = filedialog.askopenfilename(title="Select encrypted PDF", filetypes=[("PDF files", "*.pdf")])
        if not input_file:
            return
        output_file = filedialog.asksaveasfilename(title="Save unlocked PDF", defaultextension=".pdf")
        if not output_file:
            return
        password = simpledialog.askstring("Unlock PDF", "Enter password to decrypt PDF:", show="*")
        if not password:
            return
        try:
            unlock_pdf(input_file, output_file, password)
            messagebox.showinfo("Success", f"Unlocked PDF saved:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    def watermark_pdf_dialog(self):
        input_file = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf")])
        if not input_file:
            return
        output_file = filedialog.asksaveasfilename(title="Save watermarked PDF", defaultextension=".pdf")
        if not output_file:
            return
        watermark_text = simpledialog.askstring("Watermark Text", "Enter watermark text (e.g., CONFIDENTIAL):")
        if not watermark_text:
            return
        try:
            add_text_watermark(input_file, output_file, watermark_text)
            messagebox.showinfo("Success", f"Watermarked PDF saved:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def ocr_pdf_dialog(self):
        input_file = filedialog.askopenfilename(title="Select scanned PDF", filetypes=[("PDF files", "*.pdf")])
        if not input_file:
            return
        output_file = filedialog.asksaveasfilename(title="Save searchable PDF", defaultextension=".pdf")
        if not output_file:
            return
        try:
            ocr_pdf(input_file, output_file)
            messagebox.showinfo("Success", f"OCR complete. Saved PDF:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
if __name__ == "__main__":
    FileToolboxGUI().mainloop()
