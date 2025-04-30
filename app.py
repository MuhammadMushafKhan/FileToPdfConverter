import os
import logging
import time
from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import uuid
import converter
import tempfile
import shutil

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev_secret_key")

# File upload configuration
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {
    'doc', 'docx',  # Word documents
    'xls', 'xlsx',  # Excel spreadsheets
    'ppt', 'pptx',  # PowerPoint presentations
    'jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff',  # Images
    'html', 'htm',  # HTML files
    'txt', 'rtf',   # Text files
    'csv'           # CSV files
}
MAX_CONTENT_LENGTH = 10 * 1024 * 1024  # 10MB limit

# Ensure upload folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
    
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """Render the main page"""
    # Clear any temp files from previous sessions
    if 'temp_files' in session:
        for file_path in session['temp_files']:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as e:
                logger.error(f"Error removing temp file {file_path}: {e}")
        session.pop('temp_files', None)
        
    session['temp_files'] = []
    return render_template('index.html', allowed_extensions=list(ALLOWED_EXTENSIONS))

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and conversion"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': f'File type not supported. Allowed types: {", ".join(ALLOWED_EXTENSIONS)}'}), 400
    
    try:
        filename = secure_filename(file.filename)
        unique_id = str(uuid.uuid4())
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{filename}")
        file.save(input_path)
        
        if 'temp_files' not in session:
            session['temp_files'] = []
        session['temp_files'].append(input_path)
        
        # Get file extension and convert to PDF
        file_ext = filename.rsplit('.', 1)[1].lower()
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_output.pdf")
        
        success, error_msg = converter.convert_to_pdf(input_path, output_path, file_ext)
        
        if success:
            session['temp_files'].append(output_path)
            output_filename = f"{os.path.splitext(filename)[0]}.pdf"
            return jsonify({
                'success': True,
                'message': 'File converted successfully',
                'output_id': unique_id,
                'output_filename': output_filename
            })
        else:
            return jsonify({'error': error_msg}), 500
            
    except Exception as e:
        logger.exception("Error during file upload and conversion")
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

@app.route('/download/<string:file_id>', methods=['GET'])
def download_file(file_id):
    """Download the converted PDF file"""
    try:
        # Look for the output PDF file directly in the uploads folder
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{file_id}_output.pdf")
        
        if not os.path.exists(output_path):
            logger.error(f"Output PDF file not found: {output_path}")
            
            # Try to find any file with the file_id in the uploads folder as a fallback
            for filename in os.listdir(app.config['UPLOAD_FOLDER']):
                if file_id in filename and filename.endswith('_output.pdf'):
                    output_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    logger.info(f"Found alternative file: {output_path}")
                    break
        
        if not os.path.exists(output_path):
            return jsonify({'error': 'PDF file not found. The file may have been deleted or the conversion failed.'}), 404
            
        # Get original filename
        original_filename = request.args.get('filename', 'converted_file.pdf')
        
        # Log information about the download
        logger.info(f"Downloading file: {output_path} as {original_filename}")
            
        return send_file(
            output_path,
            as_attachment=True,
            download_name=original_filename,
            mimetype='application/pdf'
        )
    except Exception as e:
        logger.exception("Error during file download")
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

@app.errorhandler(413)
def request_entity_too_large(error):
    """Handle file size exceeding the limit"""
    return jsonify({'error': 'File too large. Maximum size is 10MB'}), 413

# Schedule automatic cleanup of old files (runs daily)
def cleanup_old_files():
    """Remove files older than 24 hours from the uploads folder"""
    try:
        now = time.time()
        for filename in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            # If the file is older than 24 hours (86400 seconds)
            if os.path.isfile(file_path) and os.stat(file_path).st_mtime < now - 86400:
                try:
                    os.remove(file_path)
                    logger.info(f"Removed old file: {file_path}")
                except Exception as e:
                    logger.error(f"Error removing old file {file_path}: {e}")
    except Exception as e:
        logger.error(f"Error in scheduled cleanup: {e}")

# Clean up temporary files for current request
@app.after_request
def cleanup_request_files(response):
    """Clean up files after each request completes"""
    try:
        # Only try to access session if in request context
        if 'temp_files' in session:
            for file_path in session.get('temp_files', []):
                # Don't delete output PDFs right away
                if '_output.pdf' not in file_path:
                    try:
                        if os.path.exists(file_path):
                            os.remove(file_path)
                            logger.debug(f"Removed temp file: {file_path}")
                    except Exception as e:
                        logger.error(f"Error removing temp file {file_path}: {e}")
            # We keep the output PDFs in session for download
    except Exception as e:
        logger.error(f"Error in request cleanup: {e}")
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
