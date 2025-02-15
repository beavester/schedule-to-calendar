import os
from flask import Flask, render_template, request, jsonify, send_file
import tempfile
import logging
from shift_converter import process_excel_file, generate_ics_file

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")

# Configure max file size (16MB)
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
            
        if not file.filename.endswith('.xlsx'):
            return jsonify({'error': 'Invalid file format. Please upload an Excel file (.xlsx)'}), 400

        # Save uploaded file to temporary location
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'schedule.xlsx')
        file.save(temp_path)

        # Process the Excel file
        employees, start_date = process_excel_file(temp_path)
        
        return jsonify({
            'success': True,
            'employees': employees,
            'start_date': start_date.strftime('%Y-%m-%d')
        })

    except Exception as e:
        logger.error(f"Error processing upload: {str(e)}", exc_info=True)
        return jsonify({'error': str(e)}), 500

@app.route('/generate-calendar', methods=['POST'])
def generate_calendar():
    try:
        data = request.json
        employee = data.get('employee')
        file_path = data.get('file_path')

        if not employee or not file_path:
            return jsonify({'error': 'Missing required parameters'}), 400

        # Generate ICS file
        ics_path = generate_ics_file(file_path, employee)
        
        return send_file(
            ics_path,
            mimetype='text/calendar',
            as_attachment=True,
            download_name=f'{employee}_schedule.ics'
        )

    except Exception as e:
        logger.error(f"Error generating calendar: {str(e)}", exc_info=True)
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
