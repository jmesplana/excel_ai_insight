"""Chat routes and domain operations."""
import logging
from flask import Blueprint, request, jsonify
from llm_provider import (DEFAULT_OPENAI_MODEL, build_client as build_client_for,
    LLMConfigError, get_client_and_model, provider_label, resolve_config)
from openai import OpenAI, AuthenticationError, APIError, APIConnectionError
from .errors import public_error
import pandas as pd

logger = logging.getLogger(__name__)

bp = Blueprint("chat", __name__)

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

@bp.route('/chat_with_data', methods=['POST'])
def chat_with_data():
    """Handle AI-powered questions about the analyzed data with streaming support."""
    try:
        from flask import Response, stream_with_context
        import json

        data = request.json

        # Extract request parameters. The browser holds the (analyzed) data and
        # sends the rows directly, so chat is stateless on the server.
        question = data.get('question')
        rows = data.get('rows')
        columns = data.get('columns')

        # Validate required fields
        if not rows:
            return jsonify({"error": "Missing 'rows' in the request data."}), 400
        if not question:
            return jsonify({"error": "Missing 'question' in the request data."}), 400

        # Build the provider client (OpenAI or Azure) from the request/env.
        try:
            client, model = get_client_and_model(data)
        except LLMConfigError as e:
            return jsonify({"error": public_error(e)}), 400
        except Exception as e:
            return jsonify({"error": f"Invalid provider configuration: {public_error(e)}"}), 401

        try:
            # Build a DataFrame from the posted rows, preserving column order.
            # JSON preserves numeric types from the browser, so pandas infers
            # numeric dtypes automatically (statistics work as before).
            df = pd.DataFrame(rows)
            if columns:
                ordered = [c for c in columns if c in df.columns]
                if ordered:
                    df = df[ordered]

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
                        model=model,
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
                    logger.error("Operation failed (%s)", type(stream_error).__name__)
                    error_data = json.dumps({'error': public_error(stream_error)}, ensure_ascii=False)
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

        except Exception as data_error:
            logger.error("Operation failed (%s)", type(data_error).__name__)
            return jsonify({"error": f"Error processing data: {public_error(data_error)}"}), 500

    except Exception as e:
        logger.error("Operation failed (%s)", type(e).__name__)
        return jsonify({"error": public_error(e)}), 500


