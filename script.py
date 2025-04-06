import pandas as pd
import os

# Function to remove the file extension
def remove_extension(filename):
    return os.path.splitext(filename)[0]

# Load the Excel file
file_path = r'C:\Users\JJ\Documents\mr_moses\New folder (3)\##.xlsx'  # Corrected path
df = pd.read_excel(file_path)

# Get the values in the first column of the Excel file and remove the file extensions
excel_files = [remove_extension(file) for file in df.iloc[:, 0].tolist()]

# Define the folder path
folder_path = r'C:\Users\JJ\Documents\mr_moses\New folder (3)\New folder\marked'

# Get the list of files in the folder and remove the file extensions
folder_files = [remove_extension(file) for file in os.listdir(folder_path)]

# Check which files are missing in the folder compared to the Excel file
missing_in_folder = [file for file in excel_files if file not in folder_files]

# Check which files are missing in the Excel file compared to the folder
missing_in_excel = [file for file in folder_files if file not in excel_files]

# Print the results
if missing_in_folder:
    print("\nFiles listed in Excel but missing in the folder:")
    for file in missing_in_folder:
        print(file)
else:
    print("\nNo files are missing from the folder that are listed in the Excel file.")

if missing_in_excel:
    print("\nFiles in the folder but missing in the Excel file:")
    for file in missing_in_excel:
        print(file)
else:
    print("\nNo files are missing from the Excel file that are present in the folder.")
