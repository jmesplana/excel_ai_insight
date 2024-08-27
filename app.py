import os
from flask import Flask, request, render_template, jsonify, send_file
import pandas as pd
from openai import OpenAI
from openai.types.chat import ChatCompletion
from openai import AuthenticationError, APIError, APIConnectionError
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configure upload folder and allowed extensions
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
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
        
        try:
            # 1. Read the file based on its extension
            if filename.rsplit('.', 1)[1].lower() == 'csv':
                app.logger.info("CSV file detected and being processed.")
                df = pd.read_csv(filepath, nrows=5)  # Preview first 5 rows
            else:
                app.logger.info("Excel file detected and being processed.")
                xls = pd.ExcelFile(filepath)
                df = pd.read_excel(filepath, sheet_name=xls.sheet_names[0], nrows=5)  # Preview first 5 rows
            
            # 2. Check for NaN values and log a warning if found
            if df.isnull().values.any():
                app.logger.warning("Data contains NaN values, converting NaNs to None for JSON serialization.")

            # 3. Replace Inf, -Inf values with None, and handle NaNs correctly
            df.replace([float('inf'), float('-inf')], pd.NA, inplace=True)
            
            # Convert DataFrame to use pandas NA type before filling
            df = df.convert_dtypes()  
            
            # Now fill NaNs
            df.fillna(value=pd.NA, inplace=True)
            
            # 4. Convert the cleaned DataFrame to JSON-compatible format
            sheets = {
                "Sheet1": {
                    "columns": df.columns.tolist(),
                    "data": df.to_dict('records')
                }
            }
            
            # 5. Return the JSON response
            return jsonify({"filename": filename, "sheets": sheets})
        
        except Exception as e:
            app.logger.error(f"Error processing file: {str(e)}")
            return jsonify({"error": f"Error processing file: {str(e)}"}), 500
    
    return jsonify({"error": "Invalid file type"}), 400


@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        app.logger.info("Analyze route called")
        data = request.json
        app.logger.info(f"Received data: {data}")

        # Validate incoming data
        if not data:
            raise ValueError("No data received in request.")

        # Extract and validate data from the request
        filename = data.get('filename')
        sheet_name = data.get('sheetName')
        api_key = data.get('apiKey')
        general_instructions = data.get('generalInstructions')
        column_configs = data.get('columnConfigs')

        # Check if all necessary keys are present and not empty
        if not filename:
            return jsonify({"error": "Missing 'filename' in the request data."}), 400
        if not sheet_name:
            return jsonify({"error": "Missing 'sheetName' in the request data."}), 400
        if not api_key:
            return jsonify({"error": "Missing 'apiKey' in the request data."}), 400
        if not general_instructions:
            return jsonify({"error": "Missing 'generalInstructions' in the request data."}), 400
        if not column_configs or len(column_configs) == 0:
            return jsonify({"error": "Missing 'columnConfigs' in the request data."}), 400

        # Set the OpenAI API key
        client = OpenAI(api_key=api_key)

        # Construct the full file path
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        # Check if the file exists
        if not os.path.exists(filepath):
            return jsonify({"error": f"The file {filename} does not exist in the upload folder."}), 404

        # Read the file based on its extension
        try:
            if filename.endswith(('xlsx', 'xls')):
                df = pd.read_excel(filepath, sheet_name=sheet_name)
            else:
                df = pd.read_csv(filepath)
        except Exception as e:
            return jsonify({"error": f"Error reading file {filename}: {e}"}), 500

        # Process each column configuration
        for config in column_configs:
            column = config.get('column')
            prompt = config.get('prompt')
            
            # Check if the column exists in the DataFrame
            if column not in df.columns:
                return jsonify({"error": f"Column '{column}' not found in the file."}), 400
            
            full_prompt = f"{general_instructions}\n\nColumn-specific instructions: {prompt}"

            # Catch NaN and other potential data issues
            try:
                df[f'{column}_analysis'] = df[column].apply(
            lambda x: analyze_text(client, str(x), full_prompt) if pd.notna(x) else "NaN detected"
        )
            except Exception as e:
                app.logger.error(f"Error analyzing column '{column}': {str(e)}")
                return jsonify({"error": f"Error analyzing column '{column}': {str(e)}"}), 500

        # Save the updated DataFrame
        output_filename = f"analyzed_{filename}"
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        try:
            if filename.endswith(('xlsx', 'xls')):
                with pd.ExcelWriter(output_filepath) as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                df.to_csv(output_filepath, index=False)
        except Exception as e:
            app.logger.error(f"Error saving the analyzed file: {str(e)}")
            return jsonify({"error": f"Error saving the analyzed file: {str(e)}"}), 500

        app.logger.info("Analysis complete")
        return jsonify({"message": "Analysis complete", "filename": output_filename})

    except Exception as e:
        app.logger.error(f"Error in analyze: {str(e)}")
        return jsonify({"error": str(e)}), 500

def analyze_text(client, text, prompt):
    """Use OpenAI's API to analyze the text based on the given prompt."""
    try:
        # Ensure client is an instance of OpenAI
        if not isinstance(client, OpenAI):
            client = OpenAI(api_key=client)  # Assuming client might be the API key string

        # Use the client to create a chat completion
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that analyzes text."},
                {"role": "user", "content": f"{prompt}\n\nText to analyze: {text}"}
            ],
            max_tokens=100  # Adjust as needed
        )

        # Access the content
        message_content = response.choices[0].message.content.strip()
        return message_content

    except AuthenticationError as e:
        app.logger.error(f"Authentication error: {e}")
        raise ValueError("Invalid API key provided. Please check your API key and try again.")

    except (APIError, APIConnectionError) as e:
        app.logger.error(f"OpenAI API error: {str(e)}")
        raise

    except Exception as e:
        app.logger.error(f"Unexpected error in analyze_text: {str(e)}")
        raise


@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
