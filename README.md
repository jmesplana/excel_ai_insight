
# Excel AI Insight
<p align="center">
  <img src="https://github.com/jmesplana/excel_ai_insight/blob/main/excel_ai_insight_logo.webp" alt="excel AI Insight Logo" width=25%/>
</p>

Excel AI Insight is a web application that allows you to upload Excel files and generate insightful analyses using advanced AI models. This tool helps uncover patterns, trends, and actionable insights quickly and efficiently.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [Configuration](#configuration)
  - [Uploading Files](#uploading-files)
  - [Analyzing Data](#analyzing-data)
  - [Downloading Results](#downloading-results)
  - [Sample Prompt for Data Categorization](#instruction-for-categorizing-text-data)
- [Running Tests](#running-tests)
- [Contributing](#contributing)
- [License](#license)

## Features

- **Excel File Upload**: Upload `.xlsx` or `.xls` files.
- **Preview Data**: Preview the first 5 rows of any sheet within the uploaded Excel file.
- **AI-Powered Analysis**: Use OpenAI's GPT model to analyze text data in the Excel sheets.
- **Download Results**: Download the analyzed data as an Excel file.

## Prerequisites

- Python 3.7 or higher
- [OpenAI API key](https://beta.openai.com/signup/)

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/jmesplana/excel-ai-insight.git
   cd excel-ai-insight
   ```

2. **Create a virtual environment:**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Set up the OpenAI API key:** (Optional)
   Create a `.env` file in the project root and add your OpenAI API key:
   ```env
   OPENAI_API_KEY=your-openai-api-key
   ```

## Usage

1. **Run the Flask application:**
   ```bash
   python app.py
   ```

2. **Access the web app:**
   Open your web browser and go to `http://127.0.0.1:5000/`.

### Configuration

- Enter your OpenAI API key and any general instructions for the analysis.

### Uploading Files

- Click the "Upload Excel File" button and select your `.xlsx`, `.xls` or `.csv` file.

### Analyzing Data

- Choose the sheet you want to analyze.
- Configure which columns you want to analyze and provide specific prompts.
- Click the "Analyze Columns" button to start the analysis.

### Downloading Results

- After the analysis is complete, a download link will appear. Click it to download the analyzed Excel file.

### Instruction for Categorizing Text Data
  ```bash
Instruction for Categorizing Text Data
To categorize text data effectively, follow these instructions:

**Objective**: For each entry in the text data, identify the most appropriate category from a predefined list of categories.

**Categories**: Below is the default list of categories. You can modify, add, or remove categories as needed for your specific use case:
- Category 1
- Category 2
- Category 3
- Category 4
- Category 5
- (Add or modify categories as needed)
**Configuration**: To configure the categories:

- Edit the list of categories to match the needs of your analysis.
- Ensure each category is distinct and clearly defined to avoid overlap.

**Output Requirement**: Provide only the category name as the output for each entry. Do not include explanations, the original text, or any additional information.

By following these steps, you will ensure that each text entry is accurately categorized according to the most relevant category.
   ```

### Instructions for Sentiment Analysis
**General Instructions**
1. Upload the Dataset: Upload your file (e.g., Excel, CSV) containing feedback data.
2. Select the Feedback Column: Specify the column name that contains the feedback entries (e.g., "Customer Feedback," "Patient Feedback").
3. Run the Analysis: Use the tool to perform sentiment analysis, which will classify feedback as Positive, Negative, or Mixed.
**Prompt Example**
To perform sentiment analysis, use the following prompt format:

```bash
Analyze the '[COLUMN_NAME]' column and return only the sentiment result as 'Positive,' 'Negative,' or 'Mixed' without additional text.
```
This structure ensures clarity for both the general process and specific usage.

## Running Tests

You can run tests to ensure everything is working correctly. Use `unittest` or `pytest`:

```bash
pytest
```

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Make your changes.
4. Commit your changes (`git commit -am 'Add new feature'`).
5. Push to the branch (`git push origin feature-branch`).
6. Open a pull request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
