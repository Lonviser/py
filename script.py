import os
import pandas as pd
import re

def sanitize_filename(filename):
    # Replace characters that are not allowed in Windows filenames with underscores
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

def rename_files_based_on_excel_data(folder_path, excel_path):
    # Load the Excel data into a DataFrame
    df = pd.read_excel(excel_path)

    # Iterate through files in the specified folder
    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)

        # Check if the file is a photo and the filename contains digits
        if filename.lower().endswith(('.jpg', '.jpeg', '.png')) and any(char.isdigit() for char in filename):
            # Extract digits from the filename
            digits = ''.join(filter(str.isdigit, filename))

            # Check if the digits are present in column A of the Excel data
            match = df[df['A'].astype(str).str.contains(digits, na=False)]

            # If there's a match and the filename is less than 10 characters
            if not match.empty and len(filename) < 10:
                # Construct the new filename based on values from columns B and C
                new_filename = f"{filename.split('.')[0]}_{sanitize_filename(match['B'].values[0])}_{sanitize_filename(match['C'].values[0])}.{filename.split('.')[1]}"
                new_filepath = os.path.join(folder_path, new_filename)

                # Rename the file
                os.rename(filepath, new_filepath)
                print(f"Renamed '{filename}' to '{new_filename}'.")

# Specify the folder containing the files and the Excel file path
folder_path = r'C:\poznay'
excel_path = os.path.join(folder_path, 'book.xlsx')

# Rename files based on the provided algorithm
rename_files_based_on_excel_data(folder_path, excel_path)
