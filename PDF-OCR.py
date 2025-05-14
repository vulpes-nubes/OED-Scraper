import os
import pytesseract
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2image import convert_from_path
from PIL import Image
import fitz  # PyMuPDF
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor

# Ubuntu default Tesseract path
pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"

class PDFOCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF OCR Tool")

        self.pdf_paths = []
        self.lock = threading.Lock()
        self.completed = 0

        self.create_widgets()

    def create_widgets(self):
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(fill="both", expand=True)

        tk.Button(frame, text="Select PDF File(s)", command=self.select_files).grid(row=0, column=0, sticky="ew")
        tk.Button(frame, text="Select Folder", command=self.select_folder).grid(row=0, column=1, sticky="ew")

        tk.Label(frame, text="Output Folder:").grid(row=1, column=0, sticky="w")
        self.output_dir = tk.StringVar()
        tk.Entry(frame, textvariable=self.output_dir, width=50).grid(row=1, column=1, sticky="ew")
        tk.Button(frame, text="Browse", command=self.select_output_dir).grid(row=1, column=2)

        tk.Label(frame, text="OCR Language:").grid(row=2, column=0, sticky="w")
        self.lang = tk.StringVar(value='eng')
        lang_menu = ttk.Combobox(frame, textvariable=self.lang,
                                 values=['eng', 'deu', 'fra', 'spa', 'ita'], state="readonly")
        lang_menu.grid(row=2, column=1, sticky="w")

        self.progress = ttk.Progressbar(frame, length=300)
        self.progress.grid(row=3, column=0, columnspan=3, pady=10)

        tk.Button(frame, text="Start OCR", command=self.start_ocr_thread).grid(row=4, column=0, columnspan=3, sticky="ew")

    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        self.pdf_paths = list(paths)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.pdf_paths = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.pdf')]

    def select_output_dir(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir.set(folder)

    def start_ocr_thread(self):
        if not self.pdf_paths:
            messagebox.showerror("No Files", "Please select PDF files or a folder.")
            return
        if not self.output_dir.get():
            messagebox.showerror("No Output Directory", "Please select an output directory.")
            return

        self.progress["maximum"] = len(self.pdf_paths)
        self.progress["value"] = 0
        self.completed = 0

        thread = threading.Thread(target=self.run_multithreaded_ocr)
        thread.start()

    def run_multithreaded_ocr(self):
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = []
            for pdf_path in self.pdf_paths:
                futures.append(executor.submit(self.convert_pdf_to_searchable, pdf_path))

            for future in futures:
                future.result()  # Wait for each thread

        self.root.after(0, lambda: messagebox.showinfo("Done", "OCR process completed!"))

    def update_progress(self):
        with self.lock:
            self.completed += 1
            self.root.after(0, lambda: self.progress.configure(value=self.completed))

    def convert_pdf_to_searchable(self, input_pdf_path):
        try:
            print(f"[INFO] Processing: {os.path.basename(input_pdf_path)}")
            images = convert_from_path(input_pdf_path, dpi=300)
            output_pdf_path = os.path.join(self.output_dir.get(), os.path.basename(input_pdf_path))
            doc = fitz.open()

            for i, image in enumerate(images):
                print(f"  [PAGE {i+1}] OCR...")
                ocr_pdf_bytes = pytesseract.image_to_pdf_or_hocr(image, extension='pdf', lang=self.lang.get())
                page_doc = fitz.open("pdf", ocr_pdf_bytes)
                doc.insert_pdf(page_doc)

            doc.save(output_pdf_path)
            print(f"  ✅ Saved to: {output_pdf_path}")
        except Exception as e:
            print(f"  ❌ Error: {input_pdf_path} – {e}")
        finally:
            self.update_progress()

# Main
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFOCRApp(root)
    root.mainloop()
