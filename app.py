import os
from flask import Flask, request, redirect, url_for, render_template, flash, session, send_from_directory
from werkzeug.utils import secure_filename
import pytesseract
import fitz  # PyMuPDF
from PIL import Image, ImageFilter, ImageEnhance
import re
from openpyxl import Workbook

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'pdf'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'supersecretkey'

# Dummy user credentials for authentication
users = {'admin': 'password'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username in users and users[username] == password:
            session['username'] = username
            flash('Login successful')
            return redirect(url_for('upload_file'))
        else:
            flash('Invalid credentials')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('username', None)
    flash('Logged out successfully')
    return redirect(url_for('home'))

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if 'username' not in session:
        flash('Please login first')
        return redirect(url_for('login'))

    if request.method == 'POST':
        if 'files[]' not in request.files:
            flash('No file part')
            return redirect(request.url)
        files = request.files.getlist('files[]')
        if not files or files[0].filename == '':
            flash('No selected files')
            return redirect(request.url)

        filenames = []
        all_text = ""
        extracted_data = []
        action = request.form.get('action')

        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                filenames.append(filename)

                if filename.rsplit('.', 1)[1].lower() == 'pdf':
                    text = extract_text_from_pdf(file_path)
                else:
                    processed_image_path = process_image(file_path, action)
                    text = ocr_from_image(processed_image_path)
                
                all_text += f"\nText from {filename}:\n{text}\n"
                extracted_data.append(extract_gate_data(text+"\n"+str(filename)))

        # Create a spreadsheet with extracted data
        spreadsheet_path = create_spreadsheet(extracted_data)

        return render_template('upload.html', filename=filenames, text=all_text, data=extracted_data, spreadsheet_path=spreadsheet_path)

    return render_template('upload.html', filename=None, text=None, data=None, spreadsheet_path=None)

def process_image(image_path, action):
    with Image.open(image_path) as img:
        if action == 'gray':
            img = img.convert('L')
        elif action == 'blur':
            img = img.filter(ImageFilter.BLUR)
        elif action == 'sharpen':
            enhancer = ImageEnhance.Sharpness(img)
            img = enhancer.enhance(2.0)
        elif action == 'resize':
            img = img.resize((img.width // 2, img.height // 2))
        else:
            flash('Invalid action selected')
            return image_path
        
        processed_image_path = os.path.join(app.config['UPLOAD_FOLDER'], 'processed_' + os.path.basename(image_path))
        img.save(processed_image_path)
        return processed_image_path

def ocr_from_image(image_path):
    with Image.open(image_path) as img:
        text = pytesseract.image_to_string(img)
    return text

def extract_text_from_pdf(pdf_path):
    text = ""
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            text += page.get_text()
    return text

import re

def extract_gate_data(text):
    """Extracts GATE scorecard data from the provided text based on specific places."""
    lines = text.splitlines()

    data = {
        'roll_no': '',
        'score': '',
        'marks_out_of_100': '',
        'name': '',
        'test_paper': '',
        'all_india_rank': '',
        'additional_info': '',
        'date': '',
        'roll_no':''
    }
    
    if len(lines) >= 12:
        data['register_no'] = lines[0].strip()
        data['test_paper'] = lines[1].strip()
        data['name'] = lines[2].strip()
        if((lines[3].strip().isdigit())):
            data['all_india_rank'] = int(lines[3].strip())
        else:
            data['all_india_rank'] = ''
        data['additional_info'] = lines[4].strip()
        if((lines[9].strip().isdigit())):
            data['marks_out_of_100'] = float(lines[9].strip())
        else:
            data['marks_out_of_100'] = ''
        if((lines[8].strip().isdigit())):
            data['score'] = int(lines[8].strip())
        else:
            data['score'] = ''
        data['date'] = lines[12].strip() if len(lines) > 7 else ''
        # pattern = r'_([A-Z0-9]+)'
        # match = re.search(pattern, lines[-1].strip())
        # if match:
        #     data['roll_no'] = match.group(1)
        # else:
        #     data['roll_no'] = ''
    res=""
    f=0
    for i in lines[-1].strip():
        if((i=='-') or (f==1 and len(res) < 11)):
            f=1
            if(i!='-'):
                if(i.isalpha()):
                    res=res+i.upper()
                else:
                    res=res+i
        elif(len(res) >= 11):
            break
    data['roll_no'] = str(res[:10])
            

    return data

def create_spreadsheet(data):
    """Creates an Excel spreadsheet with the extracted data."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'GATE Scorecard Data'
    
    # Define headers
    headers = ['Name of Candidates', 'Register Number', 'GATE Score', 'ALL INDIA Rank', 'Test Paper', 'Date', 'Marks out of 100','Roll No']
    ws.append(headers)
    
    for entry in data:
        row = [
            entry.get('name', ''),
            entry.get('register_no', ''),
            entry.get('score', ''),
            entry.get('all_india_rank', ''),
            entry.get('test_paper', ''),
            entry.get('date', ''),
            entry.get('marks_out_of_100', ''),
            entry.get('roll_no', '')
        ]
        ws.append(row)
    
    spreadsheet_path = os.path.join(UPLOAD_FOLDER, 'gate_scorecard_data.xlsx')
    wb.save(spreadsheet_path)
    
    return 'gate_scorecard_data.xlsx'

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
