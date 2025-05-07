import os
import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def get_search_parameters():
    user_input = {"folder": None, "query": None, "column": None, "search_all": False}

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

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Please select a valid folder.")
            return
        if not query:
            messagebox.showerror("Error", "Please enter a search query.")
            return

        user_input["folder"] = folder
        user_input["query"] = query
        user_input["column"] = column
        user_input["search_all"] = bool(search_all)
        root.destroy()

    root = tk.Tk()
    root.title("Search Configuration")

    tk.Label(root, text="Folder Containing Excel/CSV Files:").grid(row=0, column=0, sticky="e")
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
    tk.Checkbutton(root, text="Search Entire Document", variable=var_search_all).grid(row=3, columnspan=3)

    tk.Button(root, text="Start Search", command=submit).grid(row=4, columnspan=3, pady=10)

    root.mainloop()
    print(f"[INFO] Selected folder: {user_input['folder']}")
    print(f"[INFO] Search query: {user_input['query']}")
    print(f"[INFO] Column name: {user_input['column'] if user_input['column'] else '(none specified)'}")
    print(f"[INFO] Search all columns: {'Yes' if user_input['search_all'] else 'No'}")

    return user_input["folder"], user_input["query"], user_input["column"], user_input["search_all"]

def search_files(folder_path, query, column_name, search_all):
    results = []
    print("[INFO] Beginning search across Excel and CSV files...")

    for filename in os.listdir(folder_path):
        if not filename.endswith((".xlsx", ".xls", ".csv")):
            continue

        filepath = os.path.join(folder_path, filename)
        print(f"\n[PROCESSING] File: {filename}")

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
                    if re.search(query, str(cell)):
                        matches.append(str(cell))
                        total_matches += 1
                        print(f"  [MATCH] Row {i + 1}, Column {df.columns[j]}: {cell}")
        else:
            if column_name not in df.columns:
                print(f"[WARNING] Column '{column_name}' not found in {filename}")
                continue
            for i, cell in enumerate(df[column_name]):
                if re.search(query, str(cell)):
                    matches.append(str(cell))
                    total_matches += 1
                    print(f"  [MATCH] Row {i + 1}: {cell}")

        if total_matches > 0:
            percentage = (total_matches / total_non_empty_cells) * 100 if total_non_empty_cells > 0 else 0
            print(f"[RESULT] {total_matches} matches in {filename}, {percentage:.2f}% of non-empty cells")
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
    print(f"\n[INFO] Results written to {output_path}")

def main():
    folder, query, column_name, search_all = get_search_parameters()
    results = search_files(folder, query, column_name, search_all)

    if results:
        save_results(results)
        messagebox.showinfo("Done", f"Search complete. Found matches in {len(results)} file(s).")
    else:
        print("[INFO] No matches found across all files.")
        messagebox.showinfo("No Matches", "No matches found in the selected files.")

if __name__ == "__main__":
    main()
