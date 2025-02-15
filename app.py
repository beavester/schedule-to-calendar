import os
from flask import Flask, render_template, request, jsonify, send_file, session
import tempfile
import logging
from shift_converter import process_excel_file, generate_ics_file, SHIFT_MAP

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

@app.route('/shifts', methods=['GET'])
def get_shifts():
    # Get shift mappings from session or use default
    shift_mappings = session.get('shift_mappings', SHIFT_MAP)
    return jsonify(shift_mappings)

@app.route('/shifts', methods=['POST'])
def update_shifts():
    try:
        new_mappings = request.json
        if not isinstance(new_mappings, dict):
            return jsonify({'error': 'Invalid format. Expected dictionary of shift mappings'}), 400

        # Validate the format of each mapping (HHMM-HHMM or OFF)
        for shift_code, time_range in new_mappings.items():
            if time_range != "OFF":
                try:
                    start, end = time_range.split('-')
                    # Validate time format
                    if not (len(start) == 4 and len(end) == 4 and
                           start.isdigit() and end.isdigit() and
                           0 <= int(start[:2]) <= 23 and 0 <= int(end[:2]) <= 23 and
                           0 <= int(start[2:]) <= 59 and 0 <= int(end[2:]) <= 59):
                        raise ValueError
                except:
                    return jsonify({'error': f'Invalid time format for shift code {shift_code}. Use HHMM-HHMM or OFF'}), 400

        # Store in session
        session['shift_mappings'] = new_mappings
        return jsonify({'message': 'Shift mappings updated successfully'})
    except Exception as e:
        logger.error(f"Error updating shift mappings: {str(e)}", exc_info=True)
        return jsonify({'error': str(e)}), 500

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

        # Store the file path in session
        session['excel_file_path'] = temp_path

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

        if not employee:
            return jsonify({'error': 'Missing employee name'}), 400

        file_path = session.get('excel_file_path')
        if not file_path:
            return jsonify({'error': 'No Excel file found. Please upload the file again.'}), 400

        if not os.path.exists(file_path):
            return jsonify({'error': 'Excel file not found. Please upload the file again.'}), 400

        # Use custom shift mappings if available
        shift_mappings = session.get('shift_mappings', SHIFT_MAP)

        # Generate ICS file
        logger.info(f"Generating calendar for employee: {employee} using file: {file_path}")
        ics_path = generate_ics_file(file_path, employee, shift_mappings)

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