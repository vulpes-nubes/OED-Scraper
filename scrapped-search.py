import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import xml.etree.ElementTree as ET
import PyPDF2
import tempfile

# --- Helper functions ---
def log(message, console, verbose=True):
    if verbose:
        print(message)
        if console:
            console.insert(tk.END, message + "\n")
            console.see(tk.END)

def detect_file_type(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext in ['.xlsx', '.xls']: return 'excel'
    elif ext == '.csv': return 'csv'
    elif ext == '.txt': return 'txt'
    elif ext == '.pdf': return 'pdf'
    elif ext == '.xml': return 'xml'
    elif ext == '.tei': return 'tei-xml'
    else: return None

def extract_text_from_pdf(filepath):
    text = ""
    with open(filepath, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text() + "\n"
    return text

def deduplicate_results(results):
    return list(set(results))

def get_text_lines(text):
    return text.splitlines()

def search_dataframe_with_context_and_headword(df, pattern, case_sensitive=False):
    results = []
    flags = 0 if case_sensitive else re.IGNORECASE
    columns = df.columns.tolist()
    for idx, row in df.iterrows():
        for i, col in enumerate(columns):
            cell = str(row[col])
            if re.search(pattern, cell, flags):
                headword = str(row['Headword']) if 'Headword' in row else ''
                before = str(row[columns[i - 1]]) if i > 0 else ''
                match = re.sub(f"({pattern})", r"**\\1**", cell, flags=flags)
                after = str(row[columns[i + 1]]) if i < len(columns) - 1 else ''
                results.append((headword, before, match, after))
                break
    return results

def search_xml_with_context_and_headword(filepath, pattern, case_sensitive=False):
    tree = ET.parse(filepath)
    root = tree.getroot()
    flags = 0 if case_sensitive else re.IGNORECASE
    results = []

    for parent in root.iter():
        children = list(parent)
        for i, child in enumerate(children):
            if child.text and re.search(pattern, child.text.strip(), flags):
                headword = ''
                for tag in ['headword', 'hw']:
                    hw_elem = parent.find(tag)
                    if hw_elem is not None:
                        headword = hw_elem.text.strip()
                        break
                before = children[i - 1].text.strip() if i > 0 and children[i - 1].text else ''
                match = re.sub(f"({pattern})", r"**\\1**", child.text.strip(), flags=flags)
                after = children[i + 1].text.strip() if i < len(children) - 1 and children[i + 1].text else ''
                results.append((headword, before, match, after))
    return results

def prompt_column_selection(headers):
    selection_window = tk.Toplevel()
    selection_window.title("Select Headers")
    selected = {}
    for header in headers:
        var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(selection_window, text=header, variable=var)
        chk.pack(anchor='w')
        selected[header] = var

    def on_ok():
        selection_window.selected_headers = [h for h, v in selected.items() if v.get()]
        selection_window.destroy()

    tk.Button(selection_window, text="OK", command=on_ok).pack()
    selection_window.wait_window()
    return selection_window.selected_headers

def search_file(filepath, pattern, export_format, dedup, console, case_sensitive):
    filetype = detect_file_type(filepath)
    results = []

    log(f"Reading {filepath} as {filetype}...", console)

    if filetype in ['excel', 'csv']:
        df = pd.read_excel(filepath) if filetype == 'excel' else pd.read_csv(filepath)
        results = search_dataframe_with_context_and_headword(df, pattern, case_sensitive)

    elif filetype in ['xml', 'tei-xml']:
        results = search_xml_with_context_and_headword(filepath, pattern, case_sensitive)

    else:
        log("Unsupported file type for structured context output.", console)
        return

    if dedup:
        results = deduplicate_results(results)
        log(f"Deduplicated results. Total unique matches: {len(results)}", console)

    if not results:
        messagebox.showinfo("Search Complete", "No matches found.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx" if export_format == "Excel" else ".xml",
                                             filetypes=[("Excel", "*.xlsx"), ("XML", "*.xml"), ("TEI-XML", "*.xml")])
    if not save_path:
        return

    headers = ["Headword", "Before", "Match", "After"]

    if export_format == "Excel":
        df = pd.DataFrame(results, columns=headers)
        df.to_excel(save_path, index=False, engine="openpyxl")

    elif export_format in ["XML", "TEI-XML"]:
        root = ET.Element("Results")
        for row in results:
            entry = ET.SubElement(root, "Entry")
            for tag, val in zip(headers, row):
                ET.SubElement(entry, tag).text = str(val)
        tree = ET.ElementTree(root)
        tree.write(save_path, encoding="utf-8", xml_declaration=True)

    log(f"Results saved to {save_path}", console)
    messagebox.showinfo("Search Complete", f"Results exported to {save_path}")

# --- GUI Setup ---
root = tk.Tk()
root.title("Regex Search with Context and Headword")

filepath_var = tk.StringVar()
pattern_var = tk.StringVar()
dedup_var = tk.BooleanVar(value=True)
case_sensitive_var = tk.BooleanVar(value=False)
export_format_var = tk.StringVar(value="Excel")

frame = ttk.Frame(root)
frame.pack(padx=10, pady=10, fill='both', expand=True)

ttk.Button(frame, text="Choose File", command=lambda: filepath_var.set(filedialog.askopenfilename())).grid(row=0, column=0)
ttk.Entry(frame, textvariable=filepath_var, width=50).grid(row=0, column=1)

ttk.Label(frame, text="Regex Pattern:").grid(row=1, column=0)
ttk.Entry(frame, textvariable=pattern_var).grid(row=1, column=1)

ttk.Checkbutton(frame, text="Deduplicate Results", variable=dedup_var).grid(row=2, column=0, sticky='w')
ttk.Checkbutton(frame, text="Case Sensitive", variable=case_sensitive_var).grid(row=2, column=1, sticky='w')

export_label = ttk.Label(frame, text="Export format:")
export_label.grid(row=3, column=0, sticky='w')
ttk.Radiobutton(frame, text="Excel", variable=export_format_var, value="Excel").grid(row=3, column=1, sticky='w')
ttk.Radiobutton(frame, text="XML", variable=export_format_var, value="XML").grid(row=3, column=2, sticky='w')
ttk.Radiobutton(frame, text="TEI-XML", variable=export_format_var, value="TEI-XML").grid(row=3, column=3, sticky='w')

console = tk.Text(root, height=10)
console.pack(fill='both', expand=True)

def on_search():
    filepath = filepath_var.get()
    pattern = pattern_var.get()
    dedup = dedup_var.get()
    case_sensitive = case_sensitive_var.get()
    export_format = export_format_var.get()

    if not os.path.exists(filepath):
        messagebox.showerror("Error", "File path is invalid.")
        return

    search_file(filepath, pattern, export_format, dedup, console, case_sensitive)

ttk.Button(frame, text="Start Search", command=on_search).grid(row=4, column=0, columnspan=3)

root.mainloop()
