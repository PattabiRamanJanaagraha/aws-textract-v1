from flask import Flask, request, render_template, send_file, jsonify
import boto3, os, time, threading
from openpyxl import Workbook
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'

# AWS setup
s3_bucket = 'replace with bucket name'
textract = boto3.client('textract')
s3 = boto3.client('s3')

os.makedirs('uploads', exist_ok=True)
os.makedirs('output', exist_ok=True)

progress = {}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    uploaded_files = request.files.getlist("files[]")
    file_names = []

    for file in uploaded_files:
        filename = secure_filename(file.filename)
        local_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(local_path)

        s3.upload_file(local_path, s3_bucket, filename)
        file_names.append(filename)

        progress[filename] = {"status": "Queued", "percent": 0}
        thread = threading.Thread(target=process_file, args=(filename,))
        thread.start()

    return jsonify({"files": file_names})

@app.route('/progress/<filename>')
def get_progress(filename):
    return jsonify(progress.get(filename, {}))

@app.route('/download/<filename>')
def download(filename):
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    return send_file(file_path, as_attachment=True)

def process_file(filename):
    progress[filename]['status'] = 'Uploading to Textract...'
    start_time = time.time()

    response = textract.start_document_analysis(
        DocumentLocation={'S3Object': {'Bucket': s3_bucket, 'Name': filename}},
        FeatureTypes=["TABLES"]
    )
    job_id = response['JobId']
    progress[filename]['status'] = 'Processing with Textract...'
    percent = 10

    while True:
        result = textract.get_document_analysis(JobId=job_id)
        status = result['JobStatus']
        if status in ['SUCCEEDED', 'FAILED']:
            break
        progress[filename]['percent'] = min(percent, 90)
        percent += 5
        time.sleep(4)

    if status != 'SUCCEEDED':
        progress[filename]['status'] = 'Failed'
        progress[filename]['percent'] = 100
        return

    # Gather all blocks
    blocks = []
    next_token = None
    while True:
        if next_token:
            result = textract.get_document_analysis(JobId=job_id, NextToken=next_token)
        blocks.extend(result['Blocks'])
        next_token = result.get('NextToken')
        if not next_token:
            break

    block_map = {b['Id']: b for b in blocks}
    table_blocks = [b for b in blocks if b['BlockType'] == 'TABLE']

    # Single Excel sheet setup
    wb = Workbook()
    ws = wb.active
    ws.title = 'All_Pages'
    current_row = 1

    for table in table_blocks:
        cell_blocks = []
        if 'Relationships' in table:
            for rel in table['Relationships']:
                if rel['Type'] == 'CHILD':
                    for child_id in rel['Ids']:
                        child = block_map[child_id]
                        if child['BlockType'] == 'CELL':
                            cell_blocks.append(child)

        for cell in cell_blocks:
            row = cell['RowIndex'] + current_row
            col = cell['ColumnIndex']
            text = ''
            if 'Relationships' in cell:
                for rel in cell['Relationships']:
                    if rel['Type'] == 'CHILD':
                        text = ' '.join([
                            block_map[child_id]['Text']
                            for child_id in rel['Ids']
                            if block_map[child_id]['BlockType'] == 'WORD'
                        ])
            ws.cell(row=row, column=col, value=text)

        current_row += max([c['RowIndex'] for c in cell_blocks], default=0) + 3

    output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename.replace('.pdf', '.xlsx'))
    wb.save(output_path)

    end_time = time.time()
    progress[filename]['status'] = f"Done in {round(end_time - start_time)}s"
    progress[filename]['percent'] = 100

if __name__ == '__main__':
    app.run(debug=True)