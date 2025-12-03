import os
import secrets
from flask import Flask, render_template, request, send_file, jsonify, session
from engine import Config, PDFExtractor, AIAgent, CommentOptimizer, ExcelGenerator

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# Configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        api_key = request.form.get('api_key')
        
        if not file or file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not api_key:
            return jsonify({'error': 'API Key is required'}), 400
            
        # Save uploaded file
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        # Initialize components
        config = Config(api_key)
        pdf_extractor = PDFExtractor()
        ai_agent = AIAgent(config.model)
        optimizer = CommentOptimizer()
        excel_generator = ExcelGenerator()
        
        # Process
        print(f"DEBUG APP: Extracting text from {file_path}")
        raw_text = pdf_extractor.extract_text(file_path)
        print(f"DEBUG APP: Extracted text length: {len(raw_text)}")
        
        print("DEBUG APP: Calling AI agent...")
        structured_data = ai_agent.extract_key_values(raw_text)
        print(f"DEBUG APP: AI returned {len(structured_data)} items")
        
        # Add row numbers before optimization
        for idx, item in enumerate(structured_data, 1):
            item['#'] = idx
        
        print("DEBUG APP: Optimizing comments...")
        # Optimize comments (deduplicate)
        optimized_data = optimizer.optimize_comments(structured_data)
        print(f"DEBUG APP: After optimization: {len(optimized_data)} items")
        
        # Generate Excel
        output_filename = f"Output_{os.path.splitext(file.filename)[0]}.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        print(f"DEBUG APP: Creating Excel at {output_path}")
        excel_generator.create_excel(optimized_data, output_path)
        print(f"DEBUG APP: Excel created successfully")
        
        # Store output path in session for download
        session['last_output'] = output_path
        session['last_filename'] = output_filename
        
        return jsonify({
            'success': True,
            'data': optimized_data,
            'download_url': f'/download/{output_filename}'
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
