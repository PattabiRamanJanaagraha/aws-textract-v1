from flask import Flask, request, render_template, send_file, jsonify
import boto3, math, os, time, threading, traceback
from botocore.exceptions import ClientError
from collections import defaultdict
from openpyxl import Workbook
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'

# AWS setup
s3_bucket = 'cf-digitization'
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
    start_time = time.time()

    try:
        progress[filename]['status'] = 'Uploading to Textract...'

        response = textract.start_document_analysis(
            DocumentLocation={'S3Object': {'Bucket': s3_bucket, 'Name': filename}},
            FeatureTypes=["TABLES"]
        )
        job_id = response['JobId']
        progress[filename]['status'] = 'Processing with Textract...'
        percent = 10
        terminal_statuses = ['SUCCEEDED', 'FAILED', 'PARTIAL_SUCCESS']

        while True:
            result = textract.get_document_analysis(JobId=job_id, MaxResults=1000)
            status = result['JobStatus']
            if status in terminal_statuses:
                break
            progress[filename]['percent'] = min(percent, 90)
            percent += 5
            time.sleep(4)

        if status == 'FAILED':
            status_message = result.get('StatusMessage', 'Textract job failed')
            progress[filename]['status'] = f'Failed: {status_message}'
            progress[filename]['percent'] = 100
            return

        progress[filename]['status'] = 'Downloading Textract results...'
        progress[filename]['percent'] = 92

        blocks, document_metadata, warnings = collect_textract_blocks(job_id, result)
        progress[filename]['status'] = 'Writing Excel...'
        progress[filename]['percent'] = 96

        orientation_pages = write_blocks_to_excel(filename, blocks, document_metadata)

        end_time = time.time()
        status_parts = [f"Done in {round(end_time - start_time)}s"]
        if status == 'PARTIAL_SUCCESS':
            status_parts.append('partial Textract success')
        if warnings:
            status_parts.append('Textract warnings')
        if orientation_pages:
            status_parts.append(f"orientation warnings on pages {format_page_list(orientation_pages)}")

        progress[filename]['status'] = '; '.join(status_parts)
        progress[filename]['percent'] = 100

    except ClientError as exc:
        traceback.print_exc()
        progress.setdefault(filename, {})
        progress[filename]['status'] = format_aws_client_error(exc)
        progress[filename]['percent'] = 100

    except Exception as exc:
        traceback.print_exc()
        progress.setdefault(filename, {})
        progress[filename]['status'] = f'Failed: {exc}'
        progress[filename]['percent'] = 100


def format_aws_client_error(exc):
    error = exc.response.get('Error', {})
    code = error.get('Code', 'AWS error')
    message = error.get('Message', str(exc))

    if 'AccessDenied' in code or 'not authorized' in message:
        if 'textract:StartDocumentAnalysis' in message:
            return (
                'Failed: AWS IAM permission missing for textract:StartDocumentAnalysis. '
                'Attach an identity-based policy that allows Textract async analysis for this IAM user.'
            )
        return f'Failed: AWS IAM permission missing ({code}). {message}'

    return f'Failed: AWS {code}. {message}'


def collect_textract_blocks(job_id, first_result):
    blocks = []
    warnings = []
    document_metadata = first_result.get('DocumentMetadata', {})
    result = first_result

    while True:
        blocks.extend(result.get('Blocks', []))
        warnings.extend(result.get('Warnings', []))
        document_metadata = result.get('DocumentMetadata', document_metadata)

        next_token = result.get('NextToken')
        if not next_token:
            break

        result = textract.get_document_analysis(
            JobId=job_id,
            MaxResults=1000,
            NextToken=next_token
        )

    return blocks, document_metadata, warnings


def write_blocks_to_excel(filename, blocks, document_metadata):
    block_map = {b['Id']: b for b in blocks if 'Id' in b}
    lines_by_page = defaultdict(list)
    tables_by_page = defaultdict(list)
    words_by_page = defaultdict(list)
    pages = set()

    for block in blocks:
        page = block.get('Page')
        if page:
            pages.add(page)

        block_type = block.get('BlockType')
        if block_type == 'LINE' and page:
            lines_by_page[page].append(block)
        elif block_type == 'TABLE' and page:
            tables_by_page[page].append(block)
        elif block_type == 'WORD' and page:
            words_by_page[page].append(block)

    page_count = document_metadata.get('Pages')
    if page_count:
        page_numbers = range(1, page_count + 1)
    else:
        page_numbers = sorted(pages)

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title='All_Pages')
    orientation_pages = []

    for page in page_numbers:
        orientation_warning = get_orientation_warning(words_by_page.get(page, []))
        page_title = f'Page {page}'
        if orientation_warning:
            page_title = f'{page_title} - orientation warning: {orientation_warning}'
            orientation_pages.append(page)

        ws.append([page_title])

        page_tables = tables_by_page.get(page, [])
        table_boxes = [get_bounding_box(table) for table in page_tables]
        outside_lines = [
            line for line in lines_by_page.get(page, [])
            if not is_line_inside_any_table(line, table_boxes)
        ]

        page_items = [('line', line) for line in outside_lines]
        page_items.extend(('table', table) for table in page_tables)
        page_items.sort(key=lambda item: get_block_position(item[1]))

        for item_type, item in page_items:
            if item_type == 'line':
                ws.append([item.get('Text', '')])
            else:
                append_table_rows(ws, item, block_map)
                ws.append([None])

        ws.append([None])

    output_name = os.path.splitext(filename)[0] + '.xlsx'
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_name)
    wb.save(output_path)

    return orientation_pages


def append_table_rows(ws, table, block_map):
    cell_blocks = get_child_blocks(table, block_map, 'CELL')
    if not cell_blocks:
        return

    max_row = max(cell.get('RowIndex', 0) for cell in cell_blocks)
    max_col = max(cell.get('ColumnIndex', 0) for cell in cell_blocks)
    table_rows = [[None for _ in range(max_col)] for _ in range(max_row)]

    for cell in cell_blocks:
        row_index = cell.get('RowIndex', 1) - 1
        col_index = cell.get('ColumnIndex', 1) - 1
        if row_index >= 0 and col_index >= 0:
            table_rows[row_index][col_index] = get_cell_text(cell, block_map)

    for row in table_rows:
        ws.append(row)


def get_child_blocks(block, block_map, block_type=None):
    child_blocks = []
    for relationship in block.get('Relationships', []):
        if relationship.get('Type') != 'CHILD':
            continue
        for child_id in relationship.get('Ids', []):
            child = block_map.get(child_id)
            if child and (block_type is None or child.get('BlockType') == block_type):
                child_blocks.append(child)
    return child_blocks


def get_cell_text(cell, block_map):
    words = [
        child.get('Text', '')
        for child in get_child_blocks(cell, block_map, 'WORD')
        if child.get('Text')
    ]
    return ' '.join(words)


def get_bounding_box(block):
    return block.get('Geometry', {}).get('BoundingBox')


def get_block_position(block):
    box = get_bounding_box(block) or {}
    return (box.get('Top', 1), box.get('Left', 1))


def is_line_inside_any_table(line, table_boxes):
    line_box = get_bounding_box(line)
    if not line_box:
        return False

    center_x = line_box.get('Left', 0) + (line_box.get('Width', 0) / 2)
    center_y = line_box.get('Top', 0) + (line_box.get('Height', 0) / 2)

    for table_box in table_boxes:
        if not table_box:
            continue
        left = table_box.get('Left', 0)
        right = left + table_box.get('Width', 0)
        top = table_box.get('Top', 0)
        bottom = top + table_box.get('Height', 0)
        if left <= center_x <= right and top <= center_y <= bottom:
            return True

    return False


def get_orientation_warning(words):
    angles = []
    for word in words:
        angle = get_word_angle(word)
        if angle is not None:
            angles.append(abs(normalize_angle(angle)))

    if not angles:
        return None

    rotated_or_tilted = [angle for angle in angles if angle > 5]
    if len(rotated_or_tilted) < 3 or len(rotated_or_tilted) / len(angles) < 0.2:
        return None

    if max(rotated_or_tilted) >= 45:
        return 'rotated text detected'

    return 'tilted text detected'


def get_word_angle(word):
    geometry = word.get('Geometry', {})
    rotation_angle = geometry.get('RotationAngle')
    if rotation_angle is not None:
        return rotation_angle

    polygon = geometry.get('Polygon', [])
    if len(polygon) < 2:
        return None

    first = polygon[0]
    second = polygon[1]
    delta_x = second.get('X', 0) - first.get('X', 0)
    delta_y = second.get('Y', 0) - first.get('Y', 0)
    if delta_x == 0 and delta_y == 0:
        return None

    return math.degrees(math.atan2(delta_y, delta_x))


def normalize_angle(angle):
    return (angle + 180) % 360 - 180


def format_page_list(pages, limit=10):
    pages = sorted(pages)
    visible_pages = ', '.join(str(page) for page in pages[:limit])
    if len(pages) > limit:
        return f'{visible_pages}, +{len(pages) - limit} more'
    return visible_pages

if __name__ == '__main__':
    app.run(debug=True)
