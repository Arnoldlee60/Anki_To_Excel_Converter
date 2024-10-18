import sqlite3
import zipfile
import os
import pandas as pd
import re
from html import unescape
import random

# Path to the .apkg file
apkg_file = './apkg/AnatomyHY_blank.apkg'

# Unzip the .apkg file
with zipfile.ZipFile(apkg_file, 'r') as zip_ref:
    zip_ref.extractall('extracted_apkg')

# Path to the extracted .sqlite file (Anki's database file is usually called 'collection.anki2')
extracted_db = 'extracted_apkg/collection.anki2'

# Check if the extracted database file exists
if os.path.exists(extracted_db):
    print("Extracted database found")
else:
    print("Extracted database not found")

# Connect to the SQLite database in the extracted .anki2 file
conn = sqlite3.connect(extracted_db)
cursor = conn.cursor()

# Query the notes table to extract card data
cursor.execute("SELECT tags, flds, sfld FROM notes")
notes = cursor.fetchall()

data = []

# Function to clean up illegal characters and HTML entities
def clean_text(text):
    # Unescape HTML entities like &nbsp;, &lt;, etc.
    text = unescape(text)
    # Replace html entities like <div> and such idk might be broken
    text = re.sub(r'<.*?>', '', text)
    # Replace non-breaking space and any illegal characters
    text = text.replace("\xa0", " ")  # Replace non-breaking space
    # Remove non-ASCII characters and symbols like '▼' that cause issues
    text = re.sub(r'[^\x20-\x7E]+', '', text)  # Remove non-ASCII printable characters
    return text

def clean_text_with_increment(text):
    # Initialize a counter
    counter = [1]
    
    # Function to increment and replace
    def replacement(match):
        # Get the current counter value and increment it
        result = f"(..{counter[0]}..)"
        counter[0] += 1
        return result

    # Use re.sub with a callable replacement to increment numbers
    text = re.sub(r'{{c\d+::(.*?)}}', replacement, text)
    return text

def format_sfld(text):
    # Initialize a counter for cloze deletions
    counter = [1]
    
    # Function to replace each cX with (X) (Content)
    def replacement(match):
        # Get the current counter value and increment it
        result = f"({counter[0]}: {match.group(1)})"
        counter[0] += 1
        return result

    # Regular expression to find patterns like {{c1::String}}, {{c2::String}}, etc.
    formatted_text = re.sub(r'{{c\d+::(.*?)}}', replacement, text)
    
    return formatted_text

# Loop through the notes to extract the required data
for tags, flds, sfld in notes:
    # Extract cloze deletions
    cloze_deletion = re.findall(r'{{.*?}}', flds)
    
    # Replace cloze deletions
    flds = clean_text_with_increment(flds)
    
    # Clean up the front and back of card text
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


counter = 0

# for i in data:
#     print(counter, i)
#     counter += 1
    
# Convert the list of dictionaries into a pandas DataFrame
df = pd.DataFrame(data)

random_number = random.randint(1, 100000)

# Ensure the output directory exists
output_dir = "Created_Files"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Write the DataFrame to an Excel file
df.to_excel(os.path.join(output_dir, f"anki_cards_{random_number}.xlsx"), index=False)

print("Excel file created successfully!")

#Close the connection
conn.close()
