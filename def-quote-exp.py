import os
import pandas as pd

# Directory where the Excel files are stored
excel_directory = '/path here'

# Output directory for the text files
output_directory = '/path here'
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Loop through all Excel files in the directory
for file in os.listdir(excel_directory):
    if file.endswith(".xlsx"):  # Ensure you're only working with Excel files
        file_path = os.path.join(excel_directory, file)
        
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Extract 'Meaning' column, remove duplicates, and write to file
        if 'Meaning' in df.columns:
            meaning_file = os.path.join(output_directory, f"{os.path.splitext(file)[0]}_Meaning.txt")
            unique_meanings = df['Meaning'].dropna().drop_duplicates()
            unique_meanings.to_csv(meaning_file, index=False, header=False, sep='\n')
        
        # Extract 'Quotation Text' column, remove duplicates, and write to file
        if 'Quotation Text' in df.columns:
            quotation_file = os.path.join(output_directory, f"{os.path.splitext(file)[0]}_Quotation.txt")
            unique_quotations = df['Quotation Text'].dropna().drop_duplicates()
            unique_quotations.to_csv(quotation_file, index=False, header=False, sep='\n')
