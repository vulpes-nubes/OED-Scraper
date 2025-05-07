import os
import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def get_search_parameters():
    user_input = {
        "folder": None,
        "query": None,
        "column": None,
        "search_all": False,
        "only_excel": False,
        "only_csv": False,
        "case_insensitive": False
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
        search_all = var_search_all.get()
        only_excel = var_excel_only.get()
        only_csv = var_csv_only.get()
        case_insensitive = var_case_insensitive.get()

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Please select a valid folder.")
            return
        if not query:
            messagebox.showerror("Error", "Please enter a search query.")
            return
        if only_excel and only_csv:
            messagebox.showerror("Error", "Cannot select both 'Only Excel' and 'Only CSV'.")
            return

        user_input.update({
            "folder": folder,
            "query": query,
            "column": column,
            "search_all": bool(search_all),
            "only_excel": bool(only_excel),
            "only_csv": bool(only_csv),
            "case_insensitive": bool(case_insensitive)
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

    tk.Label(root, text="Column Name (optional):").grid(row=2, column=0, sticky="e")
    entry_column = tk.Entry(root, width=50)
    entry_column.grid(row=2, column=1, columnspan=2)

    var_search_all = tk.IntVar()
    tk.Checkbutton(root, text="Search Entire Document", variable=var_search_all).grid(row=3, columnspan=3, sticky="w", padx=10)

    var_case_insensitive = tk.IntVar()
    tk.Checkbutton(root, text="Case-Insensitive Search", variable=var_case_insensitive).grid(row=4, columnspan=3, sticky="w", padx=10)

    var_excel_only = tk.IntVar()
    tk.Checkbutton(root, text="Only Excel Files", variable=var_excel_only).grid(row=5, columnspan=3, sticky="w", padx=10)

    var_csv_only = tk.IntVar()
    tk.Checkbutton(root, text="Only CSV Files", variable=var_csv_only).grid(row=6, columnspan=3, sticky="w", padx=10)

    tk.Button(root, text="Start Search", command=submit).grid(row=7, columnspan=3, pady=10)

    root.mainloop()

    # Ensure all inputs were properly collected
    if user_input["folder"] is None:
        return None

    print("\n[INFO] Configuration Summary:")
    for k, v in user_input.items():
        print(f"  - {k.replace('_', ' ').capitalize()}: {v}")

    return user_input

def search_files(config):
    results = []
    folder_path = config["folder"]
    query = config["query"]
    column_name = config["column"]
    search_all = config["search_all"]
    case_insensitive = config["case_insensitive"]

    filetypes = []
    if config["only_excel"]:
        filetypes = [".xlsx", ".xls"]
    elif config["only_csv"]:
        filetypes = [".csv"]
    else:
        filetypes = [".xlsx", ".xls", ".csv"]

    # Prepare regex
    regex_flags = re.IGNORECASE if case_insensitive else 0
    pattern = re.compile(query, flags=regex_flags)

    print("\n[INFO] Beginning search...\n")

    for filename in os.listdir(folder_path):
        if not any(filename.endswith(ext) for ext in filetypes):
            continue

        filepath = os.path.join(folder_path, filename)
        print(f"[PROCESSING] File: {filename}")

        try:
            if filename.endswith(".csv"):
                df = pd.read_csv(filepath, dtype=str, encoding='utf-8', engine='python')
            else:
                df = pd.read_excel(filepath, dtype=str)
        except Exception as e:
            print(f"[ERROR] Could not read {filename}: {e}")
            continue

        df.fillna("", inplace=True)
        matches = []
        total_non_empty_cells = df.astype(bool).sum().sum()
        total_matches = 0

        if search_all:
            for i, row in df.iterrows():
                for j, cell in enumerate(row):
                    if pattern.search(str(cell)):
                        matches.append(str(cell))
                        total_matches += 1
                        print(f"  [MATCH] Row {i + 1}, Column {df.columns[j]}: {cell}")
        else:
            if column_name not in df.columns:
                print(f"[WARNING] Column '{column_name}' not found in {filename}")
                continue
            for i, cell in enumerate(df[column_name]):
                if pattern.search(str(cell)):
                    matches.append(str(cell))
                    total_matches += 1
                    print(f"  [MATCH] Row {i + 1}: {cell}")

        if total_matches > 0:
            percentage = (total_matches / total_non_empty_cells) * 100 if total_non_empty_cells > 0 else 0
            print(f"[RESULT] {total_matches} matches, {percentage:.2f}% of non-empty cells in {filename}")
            results.append({
                "File Name": filename,
                "Matched Cells": " | ".join(matches),
                "Occurrences": total_matches,
                "Percentage": f"{percentage:.2f}%"
            })
        else:
            print(f"[INFO] No matches found in {filename}.")

    return results

def save_results(results, output_path="search_results.csv"):
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["File Name", "Matched Cells", "Occurrences", "Percentage"])
        writer.writeheader()
        for row in results:
            writer.writerow(row)
    print(f"\n[INFO] Results saved to: {output_path}")

def main():
    config = get_search_parameters()
    if not config:
        messagebox.showinfo("Cancelled", "Operation cancelled.")
        return

    results = search_files(config)

    if results:
        save_results(results)
        messagebox.showinfo("Done", f"Search complete. Found matches in {len(results)} file(s).")
    else:
        print("[INFO] No matches found.")
        messagebox.showinfo("No Matches", "No matches found in the selected files.")

if __name__ == "__main__":
    main()
