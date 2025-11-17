import os
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, session
import pandas as pd
import uuid 
import base64
import fitz  
import json
import requests 
import re
from pdf2image import convert_from_bytes
import io
import time


app = Flask(__name__)
app.secret_key = '7c5bf2a2e99d01369d411c7f11e6a5f8'

DOWNLOAD_FOLDER = 'temp_downloads'

if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

@app.route('/', methods=['GET'])
def index():
    session.pop('download_filename', None) 
    return render_template('index.html', data=None)

@app.route('/process_sheet', methods=['POST'])
def process_sheet():

    if 'file' not in request.files:
        flash('No file part in the request.')
        return redirect(url_for('index'))
    file = request.files['file']
    sheet_choice = request.form.get('sheet_choice')
    if file.filename == '':
        flash('No file selected.')
        return redirect(url_for('index'))
    if file and file.filename.endswith('.xlsx'):
        try:
            df_cleaned = None
            headers = []
            if sheet_choice == 'Sheet1':
                df = pd.read_excel(file, sheet_name='Sheet1', header=None, dtype=str)
                headers = ['Name', 'Email', 'ID', 'Assigned To', 'Status', 'Date', 'Source', 'Type']
                all_records = []
                for index, row in df.iterrows():
                    values = row.dropna().tolist()
                    if values and values[0] != 'Preview':
                        record_values = values[:len(headers)]
                        while len(record_values) < len(headers):
                            record_values.append('')
                        record = dict(zip(headers, record_values))
                        all_records.append(record)
                df_cleaned = pd.DataFrame(all_records)
                if not df_cleaned.empty:
                    df_cleaned['ID'] = df_cleaned['ID'].astype(str).str.replace(r'[^0-9]', '', regex=True)
                    df_cleaned['Email'] = df_cleaned['Email'].astype(str).str.replace(r'[_\s]', '', regex=True)
            elif sheet_choice == 'Sheet2':
                df = pd.read_excel(file, sheet_name='Sheet2', header=None, dtype=str)
                headers = ['Client Name', 'Email', 'Phone No', 'Owned By', 'Status', 'Created At', 'Source', 'Contact Type']
                final_data_list = []
                block_data_list = df.iloc[0:1535, 0].dropna().tolist()
                for i in range(0, len(block_data_list), 9):
                    if i + 8 < len(block_data_list):
                        chunk = block_data_list[i : i + 9]
                        record = {'Client Name': chunk[0], 'Email': chunk[2], 'Phone No': chunk[3], 'Owned By': chunk[4], 'Status': chunk[5], 'Created At': chunk[6], 'Source': chunk[7], 'Contact Type': chunk[8]}
                        final_data_list.append(record)
                structured_df = df.iloc[1535:1668].copy()
                if structured_df.shape[1] >= 10:
                    structured_df_data = structured_df.iloc[:, 2:10]
                    structured_df_data.columns = headers
                    structured_df_cleaned = structured_df_data.dropna(subset=['Client Name'])
                    final_data_list.extend(structured_df_cleaned.to_dict(orient='records'))
                df_cleaned = pd.DataFrame(final_data_list)
                if not df_cleaned.empty:
                    df_cleaned['Phone No'] = df_cleaned['Phone No'].astype(str).str.replace(r'[^0-9]', '', regex=True)
            if df_cleaned is not None and not df_cleaned.empty:
                unique_filename = f"{uuid.uuid4()}.xlsx"
                filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], unique_filename)
                df_cleaned.to_excel(filepath, index=False)
                session['download_filename'] = unique_filename
                data_for_html = df_cleaned.to_dict(orient='records')
                return render_template('index.html', data=data_for_html, headers=headers)
            else:
                flash('No data was processed from the sheet.')
                return redirect(url_for('index'))
        except Exception as e:
            flash(f"An error occurred while processing the sheet: {e}")
            return redirect(url_for('index'))
    else:
        flash('Please upload a valid .xlsx file.')
        return redirect(url_for('index'))

# IMAGE PROCESSING ROUTE 
@app.route('/process_image', methods=['POST'])
def process_image():
    
    uploaded_files = request.files.getlist('image_file')

    if not uploaded_files or uploaded_files[0].filename == '':
        flash('No images selected.')
        return redirect(url_for('index'))

    all_extracted_data = []
    
    for file in uploaded_files:
        if file:
            try:
                
                image_bytes = file.read()
                base64_image = base64.b64encode(image_bytes).decode('utf-8')
                mime_type = file.mimetype

                # API KEY 
                api_key = ""
                
                # API URL 
                api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
                
               
                prompt = """
                Analyze the following image of a spreadsheet. Extract all data from all visible rows and columns precisely as they appear.
                Return the result as a valid JSON array of objects. Each object represents a row.
                Use the actual column headers visible in the image (e.g., "AIHN", "Customer", "Product") as the keys for the JSON objects.
                Include rows with partial data, filling missing values with empty strings.
                Do not skip any rows or columns.
                """

                payload = {
                    "contents": [{
                        "parts": [
                            {"text": prompt},
                            {"inline_data": {"mime_type": mime_type, "data": base64_image}}
                        ]
                    }]
                }

                
                response = requests.post(api_url, json=payload)
                response.raise_for_status()
                result = response.json()

                
                json_text = result['candidates'][0]['content']['parts'][0]['text']
                if json_text.strip().startswith("```json"):
                    json_text = json_text.strip()[7:-3]
                
                
                all_extracted_data.extend(json.loads(json_text))

            except Exception as e:
                flash(f"An error occurred while processing one of the images ({file.filename}): {e}")
                
                continue
    
    
    if not all_extracted_data:
        flash("The AI model could not extract any data from the uploaded images.")
        return redirect(url_for('index'))

    
    df_cleaned = pd.DataFrame(all_extracted_data)
    headers = list(df_cleaned.columns)

   
    for col_name in ['Mobile', 'Phone No', 'ID', 'AI']:
        if col_name in df_cleaned.columns:
            df_cleaned[col_name] = "" + df_cleaned[col_name].astype(str)
    
    unique_filename = f"{uuid.uuid4()}.xlsx"
    filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], unique_filename)
    df_cleaned.to_excel(filepath, index=False)
    session['download_filename'] = unique_filename
    data_for_html = df_cleaned.to_dict(orient='records')
    
    return render_template('index.html', data=data_for_html, headers=headers)


@app.route('/process_whatsapp', methods=['POST'])
def process_whatsapp():
    uploaded_files = request.files.getlist('whatsapp_file')

    if not uploaded_files or uploaded_files[0].filename == '':
        flash('No WhatsApp images selected.', 'error')
        return redirect(url_for('index'))

    all_phone_numbers = []
    
    for file in uploaded_files:
        if file:
            try:
                
                image_bytes = file.read()
                base64_image = base64.b64encode(image_bytes).decode('utf-8')
                mime_type = file.mimetype

                
                api_key = "" 
                api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
                
                
                prompt = "Extract all text from the following image."

                payload = {
                    "contents": [{
                        "parts": [
                            {"text": prompt},
                            {"inline_data": {"mime_type": mime_type, "data": base64_image}}
                        ]
                    }]
                }

                
                response = requests.post(api_url, json=payload)
                response.raise_for_status()
                result = response.json()

                
                raw_text = result['candidates'][0]['content']['parts'][0]['text']
                
               
                phone_pattern = r'\+?\d[\d\s\(\)-]{8,}\d'
                found_numbers = re.findall(phone_pattern, raw_text)
                
                
                for num in found_numbers:
                    clean_num = re.sub(r'\D', '', num)
                    if len(clean_num) > 7: 
                        all_phone_numbers.append(clean_num)

            except Exception as e:
                flash(f"An error occurred while processing one of the images ({file.filename}): {e}", 'error')
                continue
    
    if not all_phone_numbers:
        flash("Could not extract any phone numbers from the uploaded images.", 'error')
        return redirect(url_for('index'))
    df_cleaned = pd.DataFrame(all_phone_numbers, columns=['Extracted Phone Numbers'])
    
    
    unique_filename = f"{uuid.uuid4()}.xlsx"
    filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], unique_filename)
    df_cleaned.to_excel(filepath, index=False)
    session['download_filename'] = unique_filename
    
    
    headers = list(df_cleaned.columns)
    data_for_html = df_cleaned.to_dict(orient='records')

    flash(f'Successfully extracted {len(df_cleaned)} phone numbers! Your file is ready for download.', 'success')

    
    return render_template('index.html', data=data_for_html, headers=headers)


@app.route('/process_pdf', methods=['POST'])
def process_pdf():
    if 'pdf_file' not in request.files:
        flash('No file part in the request.', 'error')
        return redirect(url_for('index'))

    file = request.files['pdf_file']
    if file.filename == '':
        flash('No PDF file selected.', 'error')
        return redirect(url_for('index'))

    if file and file.filename.endswith('.pdf'):
        try:
            
            api_key = ""
            api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
            
            all_records = []
            
            
            prompt = """
            Analyze the image of this document page. Extract all tabular data you can find.
            Return the data as a valid JSON array of objects, where each object is a row from the table.
            Use the table's headers as the keys for the JSON objects.
            If there is no table on this page, return an empty array [].
            """

            
            pdf_bytes = file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")

            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(dpi=200) 
                img_bytes = pix.tobytes("png")

                if not img_bytes:
                    continue

                base64_image = base64.b64encode(img_bytes).decode('utf-8')

                payload = {
                    "contents": [{
                        "parts": [
                            {"text": prompt},
                            {"inline_data": {"mime_type": "image/png", "data": base64_image}}
                        ]
                    }]
                }
                
                params = {'key': api_key}
                response = requests.post(api_url, params=params, json=payload)
                response.raise_for_status()

                result = response.json()

                if 'candidates' in result and result['candidates'][0]['content']['parts']:
                    json_text = result['candidates'][0]['content']['parts'][0]['text']
                    
                    if json_text.strip().startswith("```json"):
                        json_text = json_text.strip()[7:-3]
                    
                    try:
                        page_data = json.loads(json_text) 
                        all_records.extend(page_data)
                    except json.JSONDecodeError:
                        flash(f"Warning: AI returned invalid JSON for page {page_num + 1}. Skipping.", 'warning')
                        continue
                else:
                    flash(f"Warning: AI returned no data for page {page_num + 1}. It might be blank.", 'warning')
                
                time.sleep(1) 

            if not all_records:
                flash('The AI could not extract any data from the PDF.', 'error')
                return redirect(url_for('index'))

            df_cleaned = pd.DataFrame(all_records)
            headers = list(df_cleaned.columns) 

            unique_filename = f"{uuid.uuid4()}.xlsx"
            filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], unique_filename)
            df_cleaned.to_excel(filepath, index=False)
            
            session['download_filename'] = unique_filename
            data_for_html = df_cleaned.to_dict(orient='records')
            flash(f'Successfully extracted {len(df_cleaned)} records from the PDF.', 'success')
            
            return render_template('index.html', data=data_for_html, headers=headers)

        except Exception as e:
            flash(f"An error occurred while processing the PDF: {e}", "error")
            return redirect(url_for('index'))
    else:
        flash('Please upload a valid .pdf file.', 'error')
        return redirect(url_for('index'))


@app.route('/download')
def download_file():
    filename = session.get('download_filename', None)
    if filename is None:
        flash("No file available to download. Please process a file first.")
        return redirect(url_for('index'))
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

