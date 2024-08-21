
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

- Click the "Upload Excel File" button and select your `.xlsx` or `.xls` file.

### Analyzing Data

- Choose the sheet you want to analyze.
- Configure which columns you want to analyze and provide specific prompts.
- Click the "Analyze Columns" button to start the analysis.

### Downloading Results

- After the analysis is complete, a download link will appear. Click it to download the analyzed Excel file.

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
