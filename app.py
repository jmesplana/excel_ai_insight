import os
from flask import Flask, request, render_template, jsonify, send_file
import pandas as pd
# from openai import OpenAI  # Import the OpenAI client
import openai

from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configure upload folder and allowed extensions
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Read Excel file
        xls = pd.ExcelFile(filepath)
        sheets = {}
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(filepath, sheet_name=sheet_name, nrows=5)  # Preview first 5 rows
            sheets[sheet_name] = {
                "columns": df.columns.tolist(),
                "data": df.to_dict('records')
            }
        
        return jsonify({"filename": filename, "sheets": sheets})
    return jsonify({"error": "Invalid file type"}), 400

@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        app.logger.info("Analyze route called")
        data = request.json
        app.logger.info(f"Received data: {data}")
        
        filename = data['filename']
        sheet_name = data['sheetName']
        api_key = data['apiKey']
        general_instructions = data['generalInstructions']
        column_configs = data['columnConfigs']
        
        # Initialize the OpenAI client with the provided API key
        # client = OpenAI(api_key=api_key)
        openai.api_key = api_key

        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        
        for config in column_configs:
            column = config['column']
            prompt = config['prompt']
            full_prompt = f"{general_instructions}\n\nColumn-specific instructions: {prompt}"
            df[f'{column}_analysis'] = df[column].apply(lambda x: analyze_text(client, str(x), full_prompt))
        
        # Save the updated DataFrame
        output_filename = f"analyzed_{filename}"
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        with pd.ExcelWriter(output_filepath) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        app.logger.info("Analysis complete")
        return jsonify({"message": "Analysis complete", "filename": output_filename})
    except Exception as e:
        app.logger.error(f"Error in analyze: {str(e)}")
        return jsonify({"error": str(e)}), 500

def analyze_text(client, text, prompt):
    """Use OpenAI's API to analyze the text based on the given prompt."""
    try:
        # Use the client instance to create a chat completion
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that analyzes text."},
                {"role": "user", "content": f"{prompt}\n\nText to analyze: {text}"}
            ],
            max_tokens=100  # Adjust as needed
        )
        
        # Correctly accessing the content using dot notation
        message_content = response.choices[0].message.content.strip()
        return message_content

    except Exception as e:
        app.logger.error(f"Error in analyze_text: {str(e)}")
        return f"Error in analysis: {str(e)}"



@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
