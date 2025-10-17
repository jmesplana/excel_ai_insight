
# Aidstack Insights
<p align="center">
  <img src="https://github.com/jmesplana/excel_ai_insight/blob/main/excel_ai_insight_logo.webp" alt="Aidstack Insights Logo" width=25%/>
</p>

**Turn spreadsheets into insights in minutes.** Aidstack Insights transforms your Excel data with AI-powered analysis, sentiment detection, translations, and pattern recognition—automatically.

Perfect for anyone drowning in Excel data: e-commerce managers analyzing customer reviews, HR professionals categorizing survey responses, market researchers translating feedback, and data analysts extracting insights.

## 🚀 Key Features

- **AI-Powered Analysis**: Automatically analyze thousands of rows with GPT-4o-mini
- **Sentiment Analysis**: Extract sentiment (Positive/Negative/Neutral) from customer feedback
- **Translation Service**: Translate text to any language (English, Spanish, French, Japanese, etc.)
- **Pattern Detection**: Discover categories and themes in your data automatically
- **Multi-Column Analysis**: Analyze multiple columns together for deeper insights
- **Test Mode**: Test your prompts on 5 rows before running full analysis
- **Progress Tracking**: Real-time progress updates during analysis
- **Dark Mode**: Easy on the eyes for long analysis sessions
- **Privacy First**: API keys stored locally in your browser, never on our servers

## 📋 Table of Contents

- [Quick Start](#quick-start)
- [Use Cases](#use-cases)
- [Features in Detail](#features-in-detail)
- [Installation](#installation)
- [Usage Guide](#usage-guide)
- [Prompt Templates](#prompt-templates)
- [Deployment](#deployment)
- [Contributing](#contributing)
- [License](#license)

## ⚡ Quick Start

1. **Clone the repository:**
   ```bash
   git clone https://github.com/jmesplana/excel_ai_insight.git
   cd excel_ai_insight
   ```

2. **Create a virtual environment:**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application:**
   ```bash
   python app.py
   ```

5. **Open your browser:**
   Navigate to `http://127.0.0.1:5000/`

6. **Get your OpenAI API key:**
   - Sign up at [OpenAI Platform](https://platform.openai.com/)
   - Create an API key
   - Enter it in the Settings tab

## 💼 Use Cases

### E-commerce Manager: Customer Review Analysis
**Challenge**: 500+ customer reviews weekly—too many to read manually
**Solution**: AI analyzes all reviews for sentiment and key issues in 5 minutes

### HR Professional: Employee Survey Categorization
**Challenge**: 1,000 employee survey responses need categorization
**Solution**: AI categorizes by topic (benefits, culture, workload) and flags urgent concerns

### Market Researcher: Multi-Language Translation
**Challenge**: Customer feedback in 5 different languages needs translation
**Solution**: AI translates all 800 responses to English in minutes

### Data Analyst: Quality Assurance
**Challenge**: Need to identify data inconsistencies and errors across thousands of entries
**Solution**: AI checks for formatting errors, outliers, and missing information

## 🎯 Features in Detail

### Sentiment Analysis (Concise)
Get one-word sentiment without explanations:
```
Analyze the sentiment of this text. Return ONLY one word: Positive, Negative, or Neutral.
Do not include any explanation, reasoning, or additional text.
```

### Translation Service
Translate to any language:
```
Translate the following text to [TARGET_LANGUAGE]. Return ONLY the translation without any
additional commentary, notes, or explanations. If the text is already in [TARGET_LANGUAGE],
return it unchanged.
```

### Pattern Detection
Discover categories automatically:
- AI analyzes a sample of your data (up to 100 values)
- Suggests distinct categories based on patterns
- Provides explanation of categorization logic

### Multi-Column Analysis
Analyze relationships across columns:
- Select multiple columns for combined analysis
- AI considers all selected data together
- Perfect for context-dependent insights

## 🛠️ Installation

### Prerequisites

- Python 3.7 or higher
- [OpenAI API key](https://platform.openai.com/api-keys)

### Local Development

```bash
# Clone repository
git clone https://github.com/jmesplana/excel_ai_insight.git
cd excel_ai_insight

# Create virtual environment
python3 -m venv venv

# Activate virtual environment
# On macOS/Linux:
source venv/bin/activate
# On Windows:
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

Access the app at `http://127.0.0.1:5000/`

## 📖 Usage Guide

### Step 1: Configuration
1. Click the **Settings** button
2. Enter your **OpenAI API key**
3. Add **General Instructions** (optional, applies to all analyses)
4. Choose from suggested prompt templates or write your own

### Step 2: Upload Your File
1. Click **Get Started** or **Try It Free**
2. Upload your Excel (`.xlsx`, `.xls`) or CSV file
3. Preview your data to verify upload

### Step 3: Select Columns & Configure Analysis
1. Select the sheet you want to analyze
2. Choose columns for analysis
3. Add column-specific prompts
4. Optionally enable **Multi-Column Analysis** for related columns

### Step 4: Run Analysis
1. Click **Run Test (5 rows)** to test your prompts first (recommended)
2. Review test results
3. Click **Analyze All Rows** to process entire dataset
4. Monitor real-time progress

### Step 5: Download Results
1. Once complete, click **Download Analyzed File**
2. Open in Excel to see AI-generated insights in new columns

## 📝 Prompt Templates

### Business Analysis
```
You are a business analyst helping to extract insights from company data. Be concise, factual,
and focus on actionable business insights. Avoid fluff and marketing language. When possible,
quantify your observations.
```

### Data Quality Check
```
You are a data quality specialist. Examine each entry for inconsistencies, formatting errors,
outliers, or missing information. When you find issues, provide specific suggestions for
improvement. Be thorough but concise.
```

### Customer Feedback Analysis (Detailed)
```
You are analyzing customer feedback to improve products and services. Identify the sentiment
(positive/negative/neutral), extract key issues or praise points, and suggest one concrete
action that could address any concerns mentioned.
```

### Translation Service
```
Translate the following text to [TARGET_LANGUAGE]. Return ONLY the translation without any
additional commentary, notes, or explanations. If the text is already in [TARGET_LANGUAGE],
return it unchanged.
```

**💡 Tip**: For concise outputs (e.g., sentiment analysis), explicitly instruct the AI to provide ONLY the required value without explanations. Example: "Return only the sentiment (Positive/Negative/Neutral) without any additional text or explanation."

## 🚀 Deployment

### Deploy to Vercel

1. Install Vercel CLI:
   ```bash
   npm install -g vercel
   ```

2. Deploy:
   ```bash
   vercel
   ```

3. Follow the prompts to complete deployment

The `vercel.json` configuration is already included in the repository.

### Environment Variables

For production deployment, set:
- `VERCEL=1` (automatically set by Vercel)
- No need to set OpenAI API key—users provide their own

## 🔒 Privacy & Security

- **API keys stored locally**: Your OpenAI API key is stored in your browser's localStorage, never transmitted to our servers
- **No data retention**: Uploaded files are processed temporarily and automatically deleted
- **Client-side API calls**: OpenAI API calls are made directly from your browser to OpenAI
- **Full control**: You maintain complete control over your data and API usage

## 🐛 Troubleshooting

### Common Issues

**Error: "Cannot take a larger sample than population when 'replace=False'"**
- **Cause**: Trying to sample more rows than exist in the column
- **Solution**: Update to latest version (fixed in v1.1.0)

**Error: "Invalid OpenAI API key"**
- **Cause**: API key is incorrect or expired
- **Solution**: Generate a new API key from OpenAI Platform

**Analysis is slow**
- **Cause**: Processing many rows or complex prompts
- **Solution**: Use Test Mode first, optimize prompts for conciseness

## 🤝 Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [Flask](https://flask.palletsprojects.com/)
- Powered by [OpenAI GPT-4o-mini](https://openai.com/)
- UI components from [Bootstrap 5](https://getbootstrap.com/)
- Design system follows [Aidstack Brand Guidelines](aidstack-brand-guide.md)

## 📧 Support

- **Issues**: [GitHub Issues](https://github.com/jmesplana/excel_ai_insight/issues)
- **Discussions**: [GitHub Discussions](https://github.com/jmesplana/excel_ai_insight/discussions)

---

**Part of the Aidstack.ai ecosystem** | Visit [aidstack.ai](https://aidstack.ai) to explore more tools
