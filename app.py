import os
import sqlite3
import zipfile
import pandas as pd
import re
import random
from flask import Flask, request, render_template, send_file
from html import unescape
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'Created_Files'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

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

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    if file:
        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)
        excel_file = process_apkg(file_path)
        return send_file(excel_file, as_attachment=True)

def process_apkg(apkg_file_path):
    with zipfile.ZipFile(apkg_file_path, 'r') as zip_ref:
        extract_folder = f'extracted_apkg_{random.randint(1, 100000)}'
        zip_ref.extractall(extract_folder)

    extracted_db = os.path.join(extract_folder, 'collection.anki2')

    if not os.path.exists(extracted_db):
        return "Extracted database not found"

    conn = sqlite3.connect(extracted_db)
    cursor = conn.cursor()

    cursor.execute("SELECT tags, flds, sfld FROM notes")
    notes = cursor.fetchall()

    data = []
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

        data.append({
            "tags": tags.split("::")[-1],
            "front of card": flds,
            "back of card": sfld,
            "cloze deletion": cloze_deletion
        })

    df = pd.DataFrame(data)
    base_name = os.path.splitext(os.path.basename(apkg_file_path))[0]
    current_date = datetime.now().strftime("%Y_%m_%d")
    excel_file_name = f"{base_name}_{current_date}.xlsx"
    excel_file_path = os.path.join(OUTPUT_FOLDER, excel_file_name)
    df.to_excel(excel_file_path, index=False)

    conn.close()

    if os.path.exists(extract_folder):
        for root, dirs, files in os.walk(extract_folder, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(extract_folder)

    return excel_file_path

if __name__ == '__main__':
    app.run(debug=True)
