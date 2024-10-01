import os
import pandas as pd

# Directory where the Excel files are stored
excel_directory = 'path_to_your_excel_files'

# Output directory for the text files
output_directory = 'path_to_your_output_files'
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Loop through all Excel files in the directory
for file in os.listdir(excel_directory):
    if file.endswith(".xlsx"):  # Ensure you're only working with Excel files
        file_path = os.path.join(excel_directory, file)
        
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Extract 'Meaning' and 'Quotation Text' columns, if they exist
        if 'Meaning' in df.columns:
            meaning_file = os.path.join(output_directory, f"{os.path.splitext(file)[0]}_Meaning.txt")
            df['Meaning'].dropna().to_csv(meaning_file, index=False, header=False, sep='\n')

        if 'Quotation Text' in df.columns:
            quotation_file = os.path.join(output_directory, f"{os.path.splitext(file)[0]}_Quotation.txt")
            df['Quotation Text'].dropna().to_csv(quotation_file, index=False, header=False, sep='\n')
