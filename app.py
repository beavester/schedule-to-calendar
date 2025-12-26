import sys
import os
from flask import Flask, render_template, request, jsonify, send_file, session
import tempfile
import logging
from shift_converter import process_excel_file, generate_ics_file, SHIFT_MAP

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__,
            template_folder=resource_path('templates'),
            static_folder=resource_path('static'))
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': 'Please upload an Excel file (.xlsx)'}), 400

        # Save to temp location
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'schedule.xlsx')
        file.save(temp_path)
        session['excel_file_path'] = temp_path

        # Process file
        employees, start_date = process_excel_file(temp_path)

        return jsonify({
            'success': True,
            'employees': employees,
            'start_date': start_date.strftime('%Y-%m-%d')
        })

    except Exception as e:
        logger.error(f"Upload error: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/generate-calendar', methods=['POST'])
def generate_calendar():
    try:
        data = request.json
        employee = data.get('employee')

        if not employee:
            return jsonify({'error': 'Please select an employee'}), 400

        file_path = session.get('excel_file_path')
        if not file_path or not os.path.exists(file_path):
            return jsonify({'error': 'Please upload the file again'}), 400

        # Generate ICS
        ics_path = generate_ics_file(file_path, employee, SHIFT_MAP)

        return send_file(
            ics_path,
            mimetype='text/calendar',
            as_attachment=True,
            download_name=f'{employee}_schedule.ics'
        )

    except Exception as e:
        logger.error(f"Calendar error: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
