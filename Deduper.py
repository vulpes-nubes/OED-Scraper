import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import xml.etree.ElementTree as ET

class DeduplicatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Deduplicator")
        self.root.geometry("400x250")
        self.selected_path = ""
        self.mode = "file"  # "file" or "folder"
        self.selected_format = tk.StringVar(value="TXT")

        # GUI Layout
        self.path_label = tk.Label(root, text="No file or folder selected.", wraplength=380)
        self.path_label.pack(pady=5)

        # File and folder selection buttons
        tk.Button(root, text="Select Single File", command=self.select_file).pack(pady=5)
        tk.Button(root, text="Select Folder (Recursive)", command=self.select_folder).pack(pady=5)

        # Format selection
        format_frame = tk.Frame(root)
        format_frame.pack(pady=5)
        tk.Label(format_frame, text="Select File Format: ").pack(side=tk.LEFT)
        self.format_dropdown = ttk.Combobox(format_frame, textvariable=self.selected_format,
                                            values=["TXT", "CSV", "Excel", "XML"], state="readonly", width=10)
        self.format_dropdown.pack(side=tk.LEFT)

        # Deduplicate button
        tk.Button(root, text="Deduplicate", command=self.run_deduplication).pack(pady=15)

    def select_file(self):
        path = filedialog.askopenfilename()
        if path:
            self.selected_path = path
            self.mode = "file"
            self.path_label.config(text=f"Selected file: {path}")

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.selected_path = path
            self.mode = "folder"
            self.path_label.config(text=f"Selected folder: {path}")

    def run_deduplication(self):
        if not self.selected_path:
            messagebox.showwarning("No Path", "Please select a file or folder first.")
            return

        file_format = self.selected_format.get()
        files = []

        if self.mode == "file":
            if self._match_format(self.selected_path, file_format):
                files = [self.selected_path]
        elif self.mode == "folder":
            for root, _, filenames in os.walk(self.selected_path):
                for file in filenames:
                    if self._match_format(file, file_format):
                        files.append(os.path.join(root, file))

        if not files:
            messagebox.showinfo("No Files", f"No {file_format} files found.")
            return

        success_count = 0
        for file in files:
            try:
                self.deduplicate_file(file, file_format)
                success_count += 1
            except Exception as e:
                messagebox.showerror("Error", f"Failed on {file}:\n{e}")
                continue

        messagebox.showinfo("Done", f"Deduplication completed on {success_count} file(s).")

    def _match_format(self, filename, file_format):
        ext_map = {
            "TXT": [".txt"],
            "CSV": [".csv"],
            "Excel": [".xlsx", ".xls"],
            "XML": [".xml"]
        }
        return any(filename.lower().endswith(ext) for ext in ext_map[file_format])

    def deduplicate_file(self, filepath, file_format):
        if file_format == "TXT":
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            unique_lines = list(dict.fromkeys(line.strip() for line in lines))
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("\n".join(unique_lines))

        elif file_format == "CSV":
            df = pd.read_csv(filepath)
            df.drop_duplicates(inplace=True)
            df.to_csv(filepath, index=False)

        elif file_format == "Excel":
            df = pd.read_excel(filepath)
            df.drop_duplicates(inplace=True)
            df.to_excel(filepath, index=False)

        elif file_format == "XML":
            tree = ET.parse(filepath)
            root = tree.getroot()
            seen = set()
            unique_children = []

            for child in root:
                rep = ET.tostring(child, encoding="unicode")
                if rep not in seen:
                    seen.add(rep)
                    unique_children.append(child)

            root.clear()
            root.extend(unique_children)
            tree.write(filepath, encoding="utf-8", xml_declaration=True)

# Run the GUI
if __name__ == "__main__":
    root = tk.Tk()
    app = DeduplicatorApp(root)
    root.mainloop()
