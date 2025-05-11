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

def find_matches_in_lines(lines, pattern, context=0, case_sensitive=False):
    matches = []
    flags = 0 if case_sensitive else re.IGNORECASE
    for i, line in enumerate(lines):
        if re.search(pattern, line, flags):
            nearby = lines[max(i-context, 0):i+context+1]
            highlighted = [re.sub(f"({pattern})", r"**\\1**", l, flags=flags) for l in nearby]
            matches.append("\n".join(highlighted))
    return matches

def search_dataframe(df, pattern, context=0, headers_only=False, retained_headers=None, case_sensitive=False):
    results = []
    flags = 0 if case_sensitive else re.IGNORECASE
    for index, row in df.iterrows():
        columns_to_search = retained_headers if headers_only and retained_headers else df.columns
        for i, col in enumerate(columns_to_search):
            if col in row and pd.notna(row[col]) and re.search(pattern, str(row[col]), flags):
                nearby_data = []
                col_index = df.columns.get_loc(col)
                start = max(0, col_index - context)
                end = min(len(df.columns), col_index + context + 1)
                for nearby_col in df.columns[start:end]:
                    value = str(row[nearby_col])
                    highlighted = re.sub(f"({pattern})", r"**\\1**", value, flags=flags)
                    nearby_data.append(highlighted)
                results.append(tuple(nearby_data))
                break
    return results

def parse_xml_file(filepath, pattern, context=0, case_sensitive=False):
    tree = ET.parse(filepath)
    root = tree.getroot()
    matches = []
    flags = 0 if case_sensitive else re.IGNORECASE
    for elem in root.iter():
        if elem.text and re.search(pattern, elem.text, flags):
            text = elem.text.strip()
            matches.append(text)
    return find_matches_in_lines(matches, pattern, context, case_sensitive)

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

def search_file(filepath, pattern, export_format, context, dedup, full_doc_search, headers_only, retained_headers, console, case_sensitive):
    filetype = detect_file_type(filepath)
    results = []

    log(f"Reading {filepath} as {filetype}...", console)

    if filetype in ['excel', 'csv']:
        df = pd.read_excel(filepath) if filetype == 'excel' else pd.read_csv(filepath)
        results = search_dataframe(df, pattern, context, headers_only, retained_headers, case_sensitive)
        headers = df.columns.tolist()

    elif filetype == 'txt':
        with open(filepath, 'r', encoding='utf-8') as f:
            text = f.read()
        lines = get_text_lines(text)
        results = find_matches_in_lines(lines, pattern, context, case_sensitive)
        headers = ["Matched Text"]

    elif filetype == 'pdf':
        text = extract_text_from_pdf(filepath)
        lines = get_text_lines(text)
        results = find_matches_in_lines(lines, pattern, context, case_sensitive)
        headers = ["Matched Text"]

    elif filetype in ['xml', 'tei-xml']:
        results = parse_xml_file(filepath, pattern, context, case_sensitive)
        headers = ["Matched Text"]

    else:
        log("Unsupported file type.", console)
        return

    if dedup:
        results = deduplicate_results(results)
        log(f"Deduplicated results. Total unique matches: {len(results)}", console)

    if not results:
        messagebox.showinfo("Search Complete", "No matches found.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx" if export_format=="Excel" else ".xml",
                                             filetypes=[("Excel", "*.xlsx"), ("XML", "*.xml"), ("TEI-XML", "*.xml")])
    if not save_path:
        return

    if export_format == "Excel":
        if isinstance(results[0], (list, tuple)):
            df = pd.DataFrame(results, columns=headers[:len(results[0])])
        else:
            df = pd.DataFrame(results, columns=["Matched Text"])
        df.to_excel(save_path, index=False, engine="openpyxl")

    elif export_format in ["XML", "TEI-XML"]:
        root = ET.Element("Results")
        for item in results:
            entry = ET.SubElement(root, "Entry")
            if isinstance(item, (tuple, list)):
                for i, val in enumerate(item):
                    ET.SubElement(entry, f"Field{i+1}").text = str(val)
            else:
                ET.SubElement(entry, "Match").text = str(item)
        tree = ET.ElementTree(root)
        tree.write(save_path, encoding="utf-8", xml_declaration=True)

    log(f"Results saved to {save_path}", console)
    messagebox.showinfo("Search Complete", f"Results exported to {save_path}")

# --- GUI Setup ---
root = tk.Tk()
root.title("Single File Regex Search")

filepath_var = tk.StringVar()
pattern_var = tk.StringVar()
context_var = tk.IntVar(value=0)
dedup_var = tk.BooleanVar(value=True)
headers_only_var = tk.BooleanVar(value=False)
full_doc_search_var = tk.BooleanVar(value=True)
case_sensitive_var = tk.BooleanVar(value=False)

export_format_var = tk.StringVar(value="Excel")
retained_headers = []

frame = ttk.Frame(root)
frame.pack(padx=10, pady=10, fill='both', expand=True)

ttk.Button(frame, text="Choose File", command=lambda: filepath_var.set(filedialog.askopenfilename())).grid(row=0, column=0)
ttk.Entry(frame, textvariable=filepath_var, width=50).grid(row=0, column=1)

ttk.Label(frame, text="Regex Pattern:").grid(row=1, column=0)
ttk.Entry(frame, textvariable=pattern_var).grid(row=1, column=1)

ttk.Checkbutton(frame, text="Deduplicate Results", variable=dedup_var).grid(row=2, column=0, sticky='w')
ttk.Label(frame, text="Nearby Context:").grid(row=2, column=1, sticky='e')
ttk.Spinbox(frame, from_=0, to=5, textvariable=context_var, width=5).grid(row=2, column=2)

ttk.Checkbutton(frame, text="Case Sensitive", variable=case_sensitive_var).grid(row=3, column=0, sticky='w')

search_mode = tk.StringVar(value="full")
ttk.Radiobutton(frame, text="Search entire document", variable=search_mode, value="full").grid(row=4, column=0, sticky='w')
ttk.Radiobutton(frame, text="Search only in selected headers", variable=search_mode, value="headers").grid(row=4, column=1, sticky='w')

export_label = ttk.Label(frame, text="Export format:")
export_label.grid(row=5, column=0, sticky='w')
ttk.Radiobutton(frame, text="Excel", variable=export_format_var, value="Excel").grid(row=5, column=1, sticky='w')
ttk.Radiobutton(frame, text="XML", variable=export_format_var, value="XML").grid(row=5, column=2, sticky='w')
ttk.Radiobutton(frame, text="TEI-XML", variable=export_format_var, value="TEI-XML").grid(row=5, column=3, sticky='w')

console = tk.Text(root, height=10)
console.pack(fill='both', expand=True)

def on_search():
    filepath = filepath_var.get()
    pattern = pattern_var.get()
    context = context_var.get()
    dedup = dedup_var.get()
    case_sensitive = case_sensitive_var.get()
    headers_only = search_mode.get() == "headers"
    full_doc_search = search_mode.get() == "full"
    export_format = export_format_var.get()

    if not os.path.exists(filepath):
        messagebox.showerror("Error", "File path is invalid.")
        return

    filetype = detect_file_type(filepath)
    if filetype in ['excel', 'csv']:
        df = pd.read_excel(filepath) if filetype == 'excel' else pd.read_csv(filepath)
        global retained_headers
        retained_headers = prompt_column_selection(df.columns.tolist())

    search_file(filepath, pattern, export_format, context, dedup, full_doc_search, headers_only, retained_headers, console, case_sensitive)

ttk.Button(frame, text="Start Search", command=on_search).grid(row=6, column=0, columnspan=3)

root.mainloop()
