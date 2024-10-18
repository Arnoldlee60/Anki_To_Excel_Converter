import sqlite3
import zipfile
import os
import pandas as pd
import re
from html import unescape
import random
from datetime import datetime

# Directory containing the .apkg files
apkg_directory = './apkg/'

# Ensure the output directory exists
output_dir = "Created_Files"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Get all .apkg files in the folder
apkg_files = [f for f in os.listdir(apkg_directory) if f.endswith('.apkg')]

# Function to clean up illegal characters and HTML entities
def clean_text(text):
    text = unescape(text)
    text = re.sub(r'<.*?>', '', text)
    text = text.replace("\xa0", " ") 
    text = re.sub(r'[^\x20-\x7E]+', '', text)
    return text

def clean_text_with_increment(text):
    counter = [1]
    def replacement(match):
        result = f"(..{counter[0]}..)"
        counter[0] += 1
        return result
    text = re.sub(r'{{c\d+::(.*?)}}', replacement, text)
    return text

def format_sfld(text):
    counter = [1]
    def replacement(match):
        result = f"({counter[0]}: {match.group(1)})"
        counter[0] += 1
        return result
    formatted_text = re.sub(r'{{c\d+::(.*?)}}', replacement, text)
    return formatted_text

# Get the current date in YYYY_MM_DD format
current_date = datetime.now().strftime("%Y_%m_%d")

# Loop through each .apkg file
for apkg_file in apkg_files:
    # Unzip the .apkg file
    apkg_file_path = os.path.join(apkg_directory, apkg_file)
    with zipfile.ZipFile(apkg_file_path, 'r') as zip_ref:
        extract_folder = f'extracted_apkg_{random.randint(1, 100000)}'
        zip_ref.extractall(extract_folder)

    # Path to the extracted .sqlite file
    extracted_db = os.path.join(extract_folder, 'collection.anki2')

    # Check if the extracted database file exists
    if not os.path.exists(extracted_db):
        print(f"Extracted database not found for {apkg_file}")
        continue

    # Connect to the SQLite database
    conn = sqlite3.connect(extracted_db)
    cursor = conn.cursor()

    # Query the notes table to extract card data
    cursor.execute("SELECT tags, flds, sfld FROM notes")
    notes = cursor.fetchall()

    data = []

    # Loop through the notes to extract the required data
    for tags, flds, sfld in notes:
        cloze_deletion = re.findall(r'{{.*?}}', flds)
        flds = clean_text_with_increment(flds)
        flds = clean_text(flds)
        sfld = clean_text(format_sfld(sfld))
        cloze_deletion = clean_text(", ".join(cloze_deletion))
        cloze_deletion = re.sub(r'{{c\d+::(.*?)}}', r'\1', cloze_deletion)
        cloze_deletion = cloze_deletion.split(",")
        for i in range(len(cloze_deletion)):
            cloze_deletion[i] = ("(" + str(i + 1) + ": " + cloze_deletion[i] + ")")
        cloze_deletion = ", ".join(cloze_deletion)

        # Append the extracted data to the list
        data.append({
            "tags": tags.split("::")[-1],
            "front of card": flds,
            "back of card": sfld,
            "cloze deletion": cloze_deletion
        })

    # Convert the list of dictionaries into a pandas DataFrame
    df = pd.DataFrame(data)

    # Get the base name of the .apkg file without extension
    base_name = os.path.splitext(apkg_file)[0]

    # Create the Excel file name with the same name as the .apkg file and date
    excel_file_name = f"{base_name}_{current_date}.xlsx"

    # Write the DataFrame to an Excel file
    df.to_excel(os.path.join(output_dir, excel_file_name), index=False)

    print(f"Excel file created for {apkg_file} as {excel_file_name}!")

    # Close the connection
    conn.close()

    # Cleanup extracted files
    if os.path.exists(extract_folder):
        for root, dirs, files in os.walk(extract_folder, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(extract_folder)

print("All .apkg files processed successfully!")
