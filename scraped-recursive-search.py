import os
import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET

def get_search_parameters():
    user_input = {
        "folder": None,
        "query": None,
        "column": None,
        "search_all": False,
        "only_excel": False,
        "only_csv": False,
        "only_xml": False,
        "only_tei": False,
        "case_insensitive": False,
        "export_format": "CSV"  # Default export format
    }

    def select_folder():
        path = filedialog.askdirectory(title="Select Folder with Files")
        if path:
            entry_folder.delete(0, tk.END)
            entry_folder.insert(0, path)

    def submit():
        folder = entry_folder.get()
        query = entry_query.get()
        column = entry_column.get()

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Please select a valid folder.")
            return
        if not query:
            messagebox.showerror("Error", "Please enter a search query.")
            return

        user_input.update({
            "folder": folder,
            "query": query,
            "column": column,
            "search_all": bool(var_search_all.get()),
            "only_excel": bool(var_excel_only.get()),
            "only_csv": bool(var_csv_only.get()),
            "only_xml": bool(var_xml_only.get()),
            "only_tei": bool(var_tei_only.get()),
            "case_insensitive": bool(var_case_insensitive.get()),
            "export_format": export_format.get()  # Export format selection
        })

        root.quit()
        root.destroy()

    root = tk.Tk()
    root.title("Search Configuration")

    tk.Label(root, text="Folder Containing Files:").grid(row=0, column=0, sticky="e")
    entry_folder = tk.Entry(root, width=50)
    entry_folder.grid(row=0, column=1)
    tk.Button(root, text="Browse", command=select_folder).grid(row=0, column=2)

    tk.Label(root, text="Search Query (Regex or Word):").grid(row=1, column=0, sticky="e")
    entry_query = tk.Entry(root, width=50)
    entry_query.grid(row=1, column=1, columnspan=2)

    tk.Label(root, text="Column Name / Tag (optional):").grid(row=2, column=0, sticky="e")
    entry_column = tk.Entry(root, width=50)
    entry_column.grid(row=2, column=1, columnspan=2)

    var_search_all = tk.IntVar()
    tk.Checkbutton(root, text="Search Entire Document", variable=var_search_all).grid(row=3, columnspan=3, sticky="w", padx=10)

    var_case_insensitive = tk.IntVar()
    tk.Checkbutton(root, text="Case-Insensitive Search", variable=var_case_insensitive).grid(row=4, columnspan=3, sticky="w", padx=10)

    # Multi-select filetype checkboxes
    var_excel_only = tk.IntVar()
    var_csv_only = tk.IntVar()
    var_xml_only = tk.IntVar()
    var_tei_only = tk.IntVar()

    tk.Label(root, text="Include File Types:").grid(row=5, column=0, sticky="w", padx=10, pady=(10, 0))
    tk.Checkbutton(root, text="Excel (.xlsx, .xls)", variable=var_excel_only).grid(row=6, column=0, sticky="w", padx=20)
    tk.Checkbutton(root, text="CSV (.csv)", variable=var_csv_only).grid(row=6, column=1, sticky="w", padx=20)
    tk.Checkbutton(root, text="XML (.xml)", variable=var_xml_only).grid(row=7, column=0, sticky="w", padx=20)
    tk.Checkbutton(root, text="TEI-XML (.tei, .tei.xml)", variable=var_tei_only).grid(row=7, column=1, sticky="w", padx=20)

    # Export format selection (CSV, Excel, XML)
    export_format = tk.StringVar(value="CSV")
    tk.Label(root, text="Export Format:").grid(row=8, column=0, sticky="w", padx=10, pady=(10, 0))
    tk.Radiobutton(root, text="CSV", variable=export_format, value="CSV").grid(row=9, column=0, sticky="w", padx=20)
    tk.Radiobutton(root, text="Excel", variable=export_format, value="Excel").grid(row=9, column=1, sticky="w", padx=20)
    tk.Radiobutton(root, text="XML", variable=export_format, value="XML").grid(row=9, column=2, sticky="w", padx=20)

    tk.Button(root, text="Start Search", command=submit).grid(row=10, columnspan=3, pady=15)

    root.mainloop()

    return user_input if user_input["folder"] else None

def search_files(config):
    filetypes = []
    if config["only_excel"]:
        filetypes += [".xlsx", ".xls"]
    if config["only_csv"]:
        filetypes += [".csv"]
    if config["only_xml"]:
        filetypes += [".xml"]
    if config["only_tei"]:
        filetypes += [".tei", ".tei.xml"]

    # Default to all file types if none selected
    if not filetypes:
        filetypes = [".xlsx", ".xls", ".csv", ".xml", ".tei", ".tei.xml"]

    all_files = []
    for root_dir, _, files in os.walk(config["folder"]):
        for file in files:
            if any(file.endswith(ext) for ext in filetypes):
                all_files.append(os.path.join(root_dir, file))

    if not all_files:
        messagebox.showerror("Error", "No files matching selected filters were found.")
        return

    output = []
    for file in all_files:
        print(f"Processing: {file}")  # Verbose output to the console for user tracking

        if file.endswith((".xlsx", ".xls")):  # Process Excel files
            df = pd.read_excel(file)
            process_file(df, file, config, output)
        elif file.endswith(".csv"):  # Process CSV files
            df = pd.read_csv(file)
            process_file(df, file, config, output)
        elif file.endswith(".xml"):  # Process XML files
            with open(file, 'r') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                process_xml_file(root, file, config, output)
        elif file.endswith((".tei", ".tei.xml")):  # Process TEI-XML files
            with open(file, 'r') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                process_xml_file(root, file, config, output)

    # Write results to selected export format
    write_to_export_format(output, config["export_format"])

def process_file(df, file, config, output):
    print(f"Processing file: {file}")
    total_cells = 0
    matched_cells = 0
    occurrences = 0

    for col in df.columns:
        if config["search_all"] or (config["column"] and col == config["column"]):
            total_cells += df[col].notna().sum()  # Count non-empty cells
            for i, cell in enumerate(df[col]):
                if isinstance(cell, str):
                    cell_text = cell.lower() if config["case_insensitive"] else cell
                    if re.search(config["query"].lower(), cell_text if config["case_insensitive"] else cell_text):
                        matched_cells += 1
                        occurrences += len(re.findall(config["query"], cell_text))
                        output.append([file, cell, occurrences, (matched_cells / total_cells) * 100 if total_cells else 0])

def process_xml_file(root, file, config, output):
    total_cells = 0
    matched_cells = 0
    occurrences = 0

    for elem in root.iter():
        if config["search_all"] or (config["column"] and elem.tag == config["column"]):
            total_cells += 1
            if isinstance(elem.text, str):
                cell_text = elem.text.lower() if config["case_insensitive"] else elem.text
                if re.search(config["query"].lower(), cell_text if config["case_insensitive"] else cell_text):
                    matched_cells += 1
                    occurrences += len(re.findall(config["query"], cell_text))
                    output.append([file, elem.text, occurrences, (matched_cells / total_cells) * 100 if total_cells else 0])

def write_to_export_format(output, export_format):
    if output:
        if export_format == "CSV":
            output_file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if output_file:
                with open(output_file, mode='w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(["File Name", "Matched Content", "Occurrences", "Percentage of Cells"])
                    writer.writerows(output)
                messagebox.showinfo("Success", "Search results have been saved to CSV.")
        
        elif export_format == "Excel":
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if output_file:
                df = pd.DataFrame(output, columns=["File Name", "Matched Content", "Occurrences", "Percentage of Cells"])
                df.to_excel(output_file, index=False)
                messagebox.showinfo("Success", "Search results have been saved to Excel.")
        
        elif export_format == "XML":
            output_file = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")])
            if output_file:
                root = ET.Element("SearchResults")
                for row in output:
                    result = ET.SubElement(root, "Result")
                    ET.SubElement(result, "FileName").text = row[0]
                    ET.SubElement(result, "MatchedContent").text = row[1]
                    ET.SubElement(result, "Occurrences").text = str(row[2])
                    ET.SubElement(result, "Percentage").text = str(row[3])

                tree = ET.ElementTree(root)
                tree.write(output_file)
                messagebox.showinfo("Success", "Search results have been saved to XML.")
    else:
        messagebox.showinfo("No Matches", "No matches found for the provided query.")

# Main execution
if __name__ == "__main__":
    config = get_search_parameters()
    if config:
        search_files(config)
