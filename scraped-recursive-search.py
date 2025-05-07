import os
import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import pandas as pd

def select_folder():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Select Folder with Excel Files")

def get_search_parameters():
    def submit():
        query = entry_query.get()
        column = entry_column.get()
        search_all = var_search_all.get()
        if not query:
            messagebox.showerror("Missing Query", "Please enter a search query.")
            return
        root.destroy()
        root.query = query
        root.column = column
        root.search_all = bool(search_all)

    root = tk.Tk()
    root.title("Search Parameters")

    tk.Label(root, text="Search Query (Regex or Word):").grid(row=0, column=0, sticky="e")
    entry_query = tk.Entry(root, width=40)
    entry_query.grid(row=0, column=1)

    tk.Label(root, text="Column Name (leave blank if searching all columns):").grid(row=1, column=0, sticky="e")
    entry_column = tk.Entry(root, width=40)
    entry_column.grid(row=1, column=1)

    var_search_all = tk.IntVar()
    tk.Checkbutton(root, text="Search Entire Document", variable=var_search_all).grid(row=2, columnspan=2)

    tk.Button(root, text="OK", command=submit).grid(row=3, columnspan=2, pady=10)

    root.mainloop()
    return root.query, root.column, root.search_all

def search_excel_files(folder_path, query, column_name, search_all):
    results = []

    for filename in os.listdir(folder_path):
        if not filename.endswith((".xlsx", ".xls")):
            continue

        filepath = os.path.join(folder_path, filename)
        try:
            df = pd.read_excel(filepath, dtype=str)
        except Exception as e:
            print(f"Error reading {filename}: {e}")
            continue

        df.fillna("", inplace=True)
        matches = []
        total_non_empty_cells = df.astype(bool).sum().sum()
        total_matches = 0

        if search_all:
            for row in df.itertuples(index=False):
                for cell in row:
                    if re.search(query, str(cell)):
                        matches.append(str(cell))
                        total_matches += 1
        else:
            if column_name not in df.columns:
                print(f"Column '{column_name}' not found in {filename}")
                continue
            for cell in df[column_name]:
                if re.search(query, str(cell)):
                    matches.append(str(cell))
                    total_matches += 1

        if total_matches > 0:
            percentage = (total_matches / total_non_empty_cells) * 100 if total_non_empty_cells > 0 else 0
            results.append({
                "File Name": filename,
                "Matched Cells": " | ".join(matches),
                "Occurrences": total_matches,
                "Percentage": f"{percentage:.2f}%"
            })

    return results

def save_results(results, output_path="search_results.csv"):
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["File Name", "Matched Cells", "Occurrences", "Percentage"])
        writer.writeheader()
        for row in results:
            writer.writerow(row)
    print(f"Results saved to {output_path}")

def main():
    folder = select_folder()
    if not folder:
        messagebox.showinfo("Cancelled", "No folder selected.")
        return

    query, column_name, search_all = get_search_parameters()
    results = search_excel_files(folder, query, column_name, search_all)

    if results:
        save_results(results)
        messagebox.showinfo("Done", f"Search complete. Found matches in {len(results)} file(s).")
    else:
        messagebox.showinfo("No Matches", "No matches found in the selected files.")

if __name__ == "__main__":
    main()
