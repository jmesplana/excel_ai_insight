import os
from flask import Flask, request, render_template, jsonify, send_file
import pandas as pd
from openai import OpenAI
from openai.types.chat import ChatCompletion
from openai import AuthenticationError, APIError, APIConnectionError
from werkzeug.utils import secure_filename
import time
import threading

app = Flask(__name__)

# Global variables for progress tracking
analysis_in_progress = False
analysis_progress = {
    "total": 0,
    "completed": 0,
    "current_column": None
}

# Lock for thread safety when updating progress
progress_lock = threading.Lock()

# Configure upload folder and allowed extensions
# For Vercel compatibility, check if we're in production and use /tmp if so
if os.environ.get('VERCEL') == '1' or os.environ.get('VERCEL_ENV') == 'production':
    UPLOAD_FOLDER = '/tmp'
else:
    UPLOAD_FOLDER = 'uploads'

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_dataframe(df, sheet_name):
    """Process a DataFrame to handle NaN values and convert to JSON-compatible format."""
    # Check for NaN values and log a warning if found
    if df.isnull().values.any():
        app.logger.warning(f"Data in sheet '{sheet_name}' contains NaN values, handling them for proper display.")

    # Replace Inf, -Inf values with None, and handle NaNs correctly
    df.replace([float('inf'), float('-inf')], pd.NA, inplace=True)

    # Convert DataFrame to use pandas NA type before filling
    df = df.convert_dtypes()

    # Now fill NaNs
    df.fillna(value=pd.NA, inplace=True)

    # Convert the cleaned DataFrame to JSON-compatible format
    return {
        sheet_name: {
            "columns": df.columns.tolist(),
            "data": df.to_dict('records')
        }
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/how-it-works')
def how_it_works():
    return render_template('how-it-works.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if file and allowed_file(file.filename):
        try:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

            # Debug info about file paths
            app.logger.info(f"Upload folder: {UPLOAD_FOLDER}, File path: {filepath}")

            # Ensure directory exists
            os.makedirs(os.path.dirname(filepath), exist_ok=True)

            # Save the file
            file.save(filepath)

            # Verify file was saved
            if not os.path.exists(filepath):
                app.logger.error(f"File was not saved to {filepath}")
                return jsonify({"error": "Failed to save file to server"}), 500

            app.logger.info(f"File saved successfully to {filepath}")
        except Exception as save_error:
            app.logger.error(f"Error saving file: {str(save_error)}")
            return jsonify({"error": f"Error saving file: {str(save_error)}"}), 500

        try:
            # 1. Read the file based on its extension
            if filename.rsplit('.', 1)[1].lower() == 'csv':
                app.logger.info("CSV file detected and being processed.")
                df = pd.read_csv(filepath, nrows=10)  # Preview first 10 rows

                # Process CSV as a single sheet
                sheets = process_dataframe(df, "Sheet1")
            else:
                app.logger.info("Excel file detected and being processed.")

                # Improved error handling for Excel files
                try:
                    xls = pd.ExcelFile(filepath)

                    # Process each sheet in the Excel file
                    sheets = {}
                    for sheet_name in xls.sheet_names:
                        try:
                            df = pd.read_excel(filepath, sheet_name=sheet_name, nrows=10)  # Preview first 10 rows
                            sheet_data = process_dataframe(df, sheet_name)
                            sheets.update(sheet_data)
                        except Exception as sheet_error:
                            app.logger.error(f"Error processing sheet '{sheet_name}': {str(sheet_error)}")
                            # Continue with other sheets if one fails
                except Exception as excel_error:
                    app.logger.error(f"Error opening Excel file: {str(excel_error)}")
                    return jsonify({"error": f"Cannot open Excel file: {str(excel_error)}"}), 500

                # Check if we successfully processed at least one sheet
                if not sheets:
                    return jsonify({"error": "Could not process any sheets in the Excel file"}), 500

            # Return the JSON response
            return jsonify({"filename": filename, "sheets": sheets})

        except Exception as e:
            app.logger.error(f"Error processing file: {str(e)}")
            return jsonify({"error": f"Error processing file: {str(e)}"}), 500

    return jsonify({"error": "Invalid file type"}), 400


@app.route('/detect_patterns', methods=['POST'])
def detect_patterns():
    try:
        app.logger.info("Detect patterns route called")
        data = request.json
        app.logger.info(f"Received pattern detection data: {data}")

        if not data:
            raise ValueError("No data received in request.")

        # Extract and validate data from the request
        filename = data.get('filename')
        sheet_name = data.get('sheetName')
        api_key = data.get('apiKey')
        column = data.get('column')
        pattern_prompt = data.get('patternPrompt')
        num_categories = data.get('numCategories', 5)

        # Check if all necessary keys are present and not empty
        if not all([filename, sheet_name, api_key, column, pattern_prompt]):
            missing = [k for k in ['filename', 'sheetName', 'apiKey', 'column', 'patternPrompt']
                      if not data.get(k)]
            return jsonify({"error": f"Missing required fields: {', '.join(missing)}"}), 400

        # Set the OpenAI API key
        client = OpenAI(api_key=api_key)

        # Construct the full file path
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        # Check if the file exists
        if not os.path.exists(filepath):
            return jsonify({"error": f"The file {filename} does not exist in the upload folder."}), 404

        # Read the file based on its extension
        try:
            if filename.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(filepath, sheet_name=sheet_name)
            else:
                df = pd.read_csv(filepath)
        except Exception as e:
            return jsonify({"error": f"Error reading file {filename}: {e}"}), 500

        # Check if the column exists in the DataFrame
        if column not in df.columns:
            return jsonify({"error": f"Column '{column}' not found in the file."}), 400

        # Get a sample of the data (up to 100 values) for pattern detection
        # Fix: Get non-null values first, then sample from those
        non_null_values = df[column].dropna()
        if len(non_null_values) == 0:
            return jsonify({"error": f"Column '{column}' contains no valid data."}), 400

        sample_size = min(100, len(non_null_values))
        sample_values = non_null_values.sample(sample_size).tolist()

        # Format the sample values for the AI
        sample_text = "\n".join([f"- {str(val)}" for val in sample_values])

        # Create the full prompt for pattern detection
        full_prompt = f"""
        {pattern_prompt}

        Please analyze the following sample data from the column '{column}' and suggest {num_categories} distinct categories:

        {sample_text}

        Based on the above sample, provide exactly {num_categories} categories in the following JSON format:
        {{
          "categories": [
            "Category1",
            "Category2",
            ...
          ],
          "explanation": "Explanation of your categorization logic and how to apply these categories to new data."
        }}
        """

        try:
            # Use the client to create a chat completion
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are an AI assistant that helps analyze data patterns and create categorization systems."},
                    {"role": "user", "content": full_prompt}
                ],
                response_format={"type": "json_object"}
            )

            # Get the response content
            categories_json = response.choices[0].message.content
            return jsonify({"result": categories_json})

        except Exception as e:
            app.logger.error(f"Error in pattern detection: {str(e)}")
            return jsonify({"error": f"Error in pattern detection: {str(e)}"}), 500

    except Exception as e:
        app.logger.error(f"Error in detect_patterns: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/analysis_progress')
def get_analysis_progress():
    """Return the current progress of analysis for client-side updates."""
    with progress_lock:
        return jsonify({
            "in_progress": analysis_in_progress,
            "progress": analysis_progress
        })

@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        app.logger.info("Analyze route called")
        data = request.json
        app.logger.info(f"Received data: {data}")

        # Initialize progress tracking
        global analysis_in_progress, analysis_progress
        with progress_lock:
            analysis_in_progress = True
            analysis_progress = {
                "total": 0,
                "completed": 0,
                "current_column": None
            }

        # Validate incoming data
        if not data:
            raise ValueError("No data received in request.")

        # Extract and validate data from the request
        filename = data.get('filename')
        sheet_name = data.get('sheetName')
        api_key = data.get('apiKey')
        general_instructions = data.get('generalInstructions')
        column_configs = data.get('columnConfigs')
        is_test_run = data.get('isTestRun', False)

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

        # Validate OpenAI API key before proceeding
        try:
            client = OpenAI(api_key=api_key)
            # Make a small test request to verify the API key
            client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": "Test"}],
                max_tokens=5
            )
        except AuthenticationError:
            return jsonify({"error": "Invalid OpenAI API key. Please check your API key and try again."}), 401
        except Exception as e:
            return jsonify({"error": f"Error connecting to OpenAI: {str(e)}"}), 500

        # Construct the full file path
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        # Check if the file exists
        if not os.path.exists(filepath):
            return jsonify({"error": f"The file {filename} does not exist in the upload folder."}), 404

        # Read the file based on its extension
        try:
            if filename.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(filepath, sheet_name=sheet_name)
            else:
                df = pd.read_csv(filepath)
        except Exception as e:
            return jsonify({"error": f"Error reading file {filename}: {e}"}), 500

        # Get row limit based on test mode
        row_limit = 5 if is_test_run else len(df)

        # Track progress for logging and UI updates
        total_operations = len(column_configs) * min(row_limit, len(df))
        operations_completed = 0
        error_count = 0

        # Update progress tracking
        with progress_lock:
            analysis_progress["total"] = total_operations
            analysis_progress["completed"] = 0

        # Validate columns before processing
        for config in column_configs:
            column = config.get('column')
            columns = config.get('columns', [column])

            for col in columns:
                if col not in df.columns:
                    return jsonify({"error": f"Column '{col}' not found in the file."}), 400

        # Process each column configuration
        for i, config in enumerate(column_configs):
            column = config.get('column')
            prompt = config.get('prompt')
            config_id = config.get('id', str(i))  # Get unique ID for each config or use index
            output_column_name = config.get('outputColumnName')  # Get user-specified column name

            # Update current column in progress tracking
            with progress_lock:
                analysis_progress["current_column"] = column

            full_prompt = f"{general_instructions}\n\nColumn-specific instructions: {prompt}"

            # Catch NaN and other potential data issues
            try:
                # Use custom name if provided, otherwise use auto-generated name
                if output_column_name:
                    analysis_column_name = output_column_name
                    # Check if column name already exists and make it unique if needed
                    counter = 1
                    original_name = analysis_column_name
                    while analysis_column_name in df.columns:
                        analysis_column_name = f"{original_name}_{counter}"
                        counter += 1
                else:
                    analysis_column_name = f'{column}_analysis_{config_id}'

                # Check for multiple column analysis
                columns = config.get('columns', [column])
                multiple_columns = len(columns) > 1

                # Function to process a single row with multiple columns
                def process_row(row_idx):
                    nonlocal operations_completed, error_count

                    try:
                        # Handle multiple columns if present
                        if multiple_columns:
                            # Combine the values from multiple columns
                            combined_values = []
                            column_headers = []

                            for col in columns:
                                if col in df.columns and pd.notna(df.at[row_idx, col]):
                                    combined_values.append(str(df.at[row_idx, col]))
                                    column_headers.append(col)

                            if not combined_values:
                                operations_completed += 1
                                # Update progress tracking
                                with progress_lock:
                                    analysis_progress["completed"] = operations_completed
                                return "No valid data in selected columns"

                            # Create a formatted input with column headers
                            formatted_input = "\n".join([f"{col}: {val}" for col, val in zip(column_headers, combined_values)])
                            result = analyze_text(client, formatted_input, full_prompt)
                            operations_completed += 1
                            # Update progress tracking
                            with progress_lock:
                                analysis_progress["completed"] = operations_completed
                            return result
                        else:
                            # Original single column processing
                            value = df.at[row_idx, column]
                            if pd.notna(value):
                                result = analyze_text(client, str(value), full_prompt)
                                operations_completed += 1
                                # Update progress tracking
                                with progress_lock:
                                    analysis_progress["completed"] = operations_completed
                                return result
                            else:
                                operations_completed += 1
                                # Update progress tracking
                                with progress_lock:
                                    analysis_progress["completed"] = operations_completed
                                return "No data (empty cell)"
                    except Exception as e:
                        error_count += 1
                        operations_completed += 1
                        # Update progress tracking
                        with progress_lock:
                            analysis_progress["completed"] = operations_completed
                        app.logger.error(f"Error processing row {row_idx} for column '{column}': {str(e)}")
                        return f"Error: {str(e)[:50]}..."

                # Ensure the analysis column exists
                if analysis_column_name not in df.columns:
                    df[analysis_column_name] = pd.NA

                # Process rows based on test mode
                for idx in range(min(row_limit, len(df))):
                    # Log progress for every 10% completion
                    progress_percentage = int((operations_completed / total_operations) * 100) if total_operations > 0 else 0
                    if progress_percentage % 10 == 0 and operations_completed > 0:
                        app.logger.info(f"Analysis progress: {progress_percentage}%, errors: {error_count}")

                    try:
                        df.at[idx, analysis_column_name] = process_row(idx)
                    except Exception as row_error:
                        app.logger.error(f"Error processing row {idx}: {str(row_error)}")
                        df.at[idx, analysis_column_name] = f"Error: {str(row_error)[:50]}..."
                        error_count += 1

            except Exception as e:
                app.logger.error(f"Error analyzing column '{column}': {str(e)}")
                return jsonify({"error": f"Error analyzing column '{column}': {str(e)}"}), 500

        # Save the updated DataFrame
        output_filename = f"analyzed_{filename}"
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        try:
            if filename.endswith(('.xlsx', '.xls')):
                with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                df.to_csv(output_filepath, index=False)
        except Exception as e:
            app.logger.error(f"Error saving the analyzed file: {str(e)}")
            return jsonify({"error": f"Error saving the analyzed file: {str(e)}"}), 500

        app.logger.info(f"Analysis complete. Processed {operations_completed} cells with {error_count} errors.")

        # Reset progress tracking
        with progress_lock:
            analysis_in_progress = False
            analysis_progress["completed"] = analysis_progress["total"]  # Ensure 100%

        return jsonify({
            "message": "Analysis complete!" + (" (Test run on 5 rows)" if is_test_run else ""),
            "filename": output_filename,
            "stats": {
                "processed": operations_completed,
                "errors": error_count
            }
        })

    except Exception as e:
        app.logger.error(f"Error in analyze: {str(e)}")
        # Reset progress tracking on error
        with progress_lock:
            analysis_in_progress = False
        return jsonify({"error": str(e)}), 500

def analyze_text(client, text, prompt):
    """Use OpenAI's API to analyze the text based on the given prompt."""
    try:
        # Ensure client is an instance of OpenAI
        if not isinstance(client, OpenAI):
            client = OpenAI(api_key=client)  # Assuming client might be the API key string

        # Truncate text if it's too long (to avoid token limits)
        MAX_TEXT_LENGTH = 8000  # Conservative limit for gpt-4o-mini
        if len(text) > MAX_TEXT_LENGTH:
            text = text[:MAX_TEXT_LENGTH] + "... [text truncated due to length]"

        # Define a system message that sets expectations for the response
        system_message = """You are an expert data analyst helping to analyze text data.
        Provide concise, insightful analysis based on the user's instructions.
        Keep your response focused and under 100 words.
        If the text is unclear or lacks sufficient information, say so briefly."""

        # Add retry logic for rate limits
        max_retries = 3
        retry_delay = 2  # seconds

        for attempt in range(max_retries):
            try:
                # Use the client to create a chat completion
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": system_message},
                        {"role": "user", "content": f"{prompt}\n\nText to analyze: {text}"}
                    ],
                    max_tokens=150,  # Increased slightly to allow more detailed analysis
                    temperature=0.3   # Lower temperature for more consistent, focused responses
                )

                # Access the content
                message_content = response.choices[0].message.content.strip()
                return message_content

            except APIConnectionError as e:
                if "rate limit" in str(e).lower() and attempt < max_retries - 1:
                    # Exponential backoff
                    sleep_time = retry_delay * (2 ** attempt)
                    app.logger.warning(f"Rate limit hit, retrying in {sleep_time} seconds...")
                    time.sleep(sleep_time)
                    continue
                else:
                    raise

    except AuthenticationError as e:
        app.logger.error(f"Authentication error: {e}")
        raise ValueError("Invalid API key provided. Please check your API key and try again.")

    except (APIError, APIConnectionError) as e:
        app.logger.error(f"OpenAI API error: {str(e)}")
        if "rate limit" in str(e).lower():
            raise ValueError("OpenAI rate limit reached. Please wait a moment before trying again.")
        else:
            raise

    except Exception as e:
        app.logger.error(f"Unexpected error in analyze_text: {str(e)}")
        raise


@app.route('/get_analyzed_data/<filename>')
def get_analyzed_data(filename):
    """Fetch analyzed file data as JSON for display in the results viewer."""
    try:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        # Check if the file exists
        if not os.path.exists(filepath):
            return jsonify({"error": f"File {filename} not found"}), 404

        # Read the file based on its extension
        try:
            if filename.endswith(('.xlsx', '.xls')):
                # Read all sheets from Excel file
                xls = pd.ExcelFile(filepath)
                sheets_data = {}

                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    # Process the dataframe to handle NaN values
                    sheets_data.update(process_dataframe(df, sheet_name))

                return jsonify({"sheets": sheets_data})
            else:
                # Read CSV file
                df = pd.read_csv(filepath)
                sheets_data = process_dataframe(df, "Sheet1")
                return jsonify({"sheets": sheets_data})

        except Exception as e:
            app.logger.error(f"Error reading file {filename}: {str(e)}")
            return jsonify({"error": f"Error reading file: {str(e)}"}), 500

    except Exception as e:
        app.logger.error(f"Error in get_analyzed_data: {str(e)}")
        return jsonify({"error": str(e)}), 500

def analyze_data_with_code(df, question):
    """
    Perform data analysis operations based on the question.
    Returns the result of the analysis as a string.
    """
    import re

    # Helper function to safely get column values
    def get_column_info(col_name):
        if col_name not in df.columns:
            return f"Column '{col_name}' not found."

        col = df[col_name]
        unique_vals = col.dropna().unique()

        info = {
            'name': col_name,
            'dtype': str(col.dtype),
            'non_null_count': col.notna().sum(),
            'null_count': col.isna().sum(),
            'unique_count': len(unique_vals),
            'unique_values': unique_vals.tolist() if len(unique_vals) <= 100 else unique_vals[:100].tolist()
        }

        # Add numeric stats if applicable
        if pd.api.types.is_numeric_dtype(col):
            info['mean'] = col.mean()
            info['median'] = col.median()
            info['std'] = col.std()
            info['min'] = col.min()
            info['max'] = col.max()

        return info

    # Detect if question asks about unique values
    if any(keyword in question.lower() for keyword in ['unique', 'distinct', 'different', 'all values', 'list of']):
        # Try to identify column name in question
        for col in df.columns:
            if col.lower() in question.lower():
                return get_column_info(col)

    # Detect if question asks about column information
    if any(keyword in question.lower() for keyword in ['columns', 'fields', 'what data', 'what columns']):
        return {
            'columns': df.columns.tolist(),
            'total_rows': len(df),
            'total_columns': len(df.columns)
        }

    # Detect if question asks for summary statistics
    if any(keyword in question.lower() for keyword in ['summary', 'statistics', 'stats', 'describe']):
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        if numeric_cols:
            return df[numeric_cols].describe().to_dict()

    # Detect if question asks about correlations
    if 'correlation' in question.lower() or 'correlate' in question.lower():
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        if len(numeric_cols) >= 2:
            return df[numeric_cols].corr().to_dict()

    # Detect value counts requests (including "which has the most")
    if any(keyword in question.lower() for keyword in ['count', 'frequency', 'distribution', 'how many', 'which', 'most', 'least', 'lot of']):
        for col in df.columns:
            if col.lower() in question.lower():
                value_counts = df[col].value_counts().head(50).to_dict()
                return {
                    'column': col,
                    'value_counts': value_counts,
                    'total_unique': df[col].nunique()
                }

    return None

@app.route('/chat_with_data', methods=['POST'])
def chat_with_data():
    """Handle AI-powered questions about the analyzed data with streaming support."""
    try:
        from flask import Response, stream_with_context
        import json

        data = request.json
        app.logger.info(f"Chat request received: {data}")

        # Extract request parameters
        filename = data.get('filename')
        question = data.get('question')
        api_key = data.get('apiKey')

        # Validate required fields
        if not filename:
            return jsonify({"error": "Missing 'filename' in the request data."}), 400
        if not question:
            return jsonify({"error": "Missing 'question' in the request data."}), 400
        if not api_key:
            return jsonify({"error": "Missing 'apiKey' in the request data."}), 400

        # Validate API key
        try:
            client = OpenAI(api_key=api_key)
        except Exception as e:
            return jsonify({"error": f"Invalid API key: {str(e)}"}), 401

        # Load the analyzed file
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if not os.path.exists(filepath):
            return jsonify({"error": f"File {filename} not found"}), 404

        try:
            # Read the file and prepare data context
            if filename.endswith(('.xlsx', '.xls')):
                xls = pd.ExcelFile(filepath)
                # For simplicity, we'll focus on the first sheet for chat
                # Future enhancement: allow users to specify which sheet
                first_sheet = xls.sheet_names[0]
                df = pd.read_excel(filepath, sheet_name=first_sheet)
            else:
                df = pd.read_csv(filepath)

            # Try to perform direct data analysis first
            analysis_result = analyze_data_with_code(df, question)

            # Create a comprehensive data summary for the AI
            data_summary = f"""
Dataset Information:
- Total rows: {len(df)}
- Total columns: {len(df.columns)}
- Columns: {', '.join(df.columns.tolist())}

Column Details:
"""
            for col in df.columns:
                dtype = df[col].dtype
                non_null = df[col].notna().sum()
                unique_count = df[col].nunique()
                data_summary += f"  - {col} ({dtype}): {non_null} non-null values, {unique_count} unique values\n"

                # Add some sample values for context (increased from 3 to 5)
                sample_values = df[col].dropna().head(5).tolist()
                if sample_values:
                    data_summary += f"    Sample values: {sample_values}\n"

            # Add basic statistics for numeric columns
            numeric_cols = df.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 0:
                data_summary += "\nNumeric Column Statistics:\n"
                stats = df[numeric_cols].describe().to_string()
                data_summary += stats

            # Add the analysis result if available
            if analysis_result:
                data_summary += f"\n\nDirect Analysis Result:\n{json.dumps(analysis_result, indent=2, default=str)}\n"

            # Limit data summary size to avoid token limits
            MAX_SUMMARY_LENGTH = 8000  # Increased from 4000
            if len(data_summary) > MAX_SUMMARY_LENGTH:
                data_summary = data_summary[:MAX_SUMMARY_LENGTH] + "\n... [truncated for length]"

            # Create the AI prompt
            system_message = """You are an expert data analyst helping users understand their data.

When answering questions:
1. UNDERSTAND THE QUESTION FIRST - don't just list data
   - "List countries" → Show the complete list
   - "Which country has the most?" → Analyze and answer which one
   - "How many X?" → Count and answer

2. For LIST questions (list, show all, what are):
   - If "unique_values" exist in analysis result, show EVERY value
   - Never use "[Other values...]" or "[Additional...]" placeholders

3. For ANALYTICAL questions (which, most, least, best):
   - Actually ANALYZE the data using value_counts or statistics
   - Give a direct answer with supporting numbers
   - Don't just repeat the list

4. Be conversational and helpful, not robotic

Format responses in clear markdown."""

            user_message = f"""Dataset Information:

{data_summary}

User Question: {question}

Analyze the question and provide an appropriate answer. If it's asking for analysis (which/most/least), actually analyze the data. If it's asking for a list, provide the complete list."""

            # Stream the response using SSE (Server-Sent Events)
            def generate():
                try:
                    response = client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[
                            {"role": "system", "content": system_message},
                            {"role": "user", "content": user_message}
                        ],
                        max_tokens=2000,  # Increased from 500 to allow longer responses
                        temperature=0.7,
                        stream=True
                    )

                    for chunk in response:
                        if chunk.choices[0].delta.content:
                            content = chunk.choices[0].delta.content
                            # Send as SSE format with proper JSON encoding
                            json_data = json.dumps({'content': content}, ensure_ascii=False)
                            yield f"data: {json_data}\n\n"

                    # Send completion signal
                    yield f"data: {json.dumps({'done': True})}\n\n"

                except Exception as stream_error:
                    app.logger.error(f"Streaming error: {str(stream_error)}")
                    error_data = json.dumps({'error': str(stream_error)}, ensure_ascii=False)
                    yield f"data: {error_data}\n\n"

            return Response(
                stream_with_context(generate()),
                mimetype='text/event-stream',
                headers={
                    'Cache-Control': 'no-cache',
                    'X-Accel-Buffering': 'no',
                    'Connection': 'keep-alive'
                }
            )

        except Exception as file_error:
            app.logger.error(f"Error reading file for chat: {str(file_error)}")
            return jsonify({"error": f"Error reading file: {str(file_error)}"}), 500

    except Exception as e:
        app.logger.error(f"Error in chat_with_data: {str(e)}")
        return jsonify({"error": str(e)}), 500


@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
