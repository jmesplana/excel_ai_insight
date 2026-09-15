
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
- **Medical Coding & Translation (WHO ICD-11)**: Map free-text diagnoses (any language, even with typos) to official WHO ICD-11 codes and translated terms, plus an optional derived ICD-10 mapping with a relationship indicator (same-as / broader-than / narrower-than / no-map)
- **Pattern Detection**: Discover categories and themes in your data automatically
- **Multi-Column Analysis**: Analyze multiple columns together for deeper insights
- **Test Mode**: Test your prompts on 5 rows before running full analysis
- **Progress Tracking**: Real-time progress updates during analysis
- **Dark Mode**: Easy on the eyes for long analysis sessions
- **Privacy First**: Files parsed locally; selected rows and credentials pass through the server for processing

## 📋 Table of Contents

- [Quick Start](#quick-start)
- [Use Cases](#use-cases)
- [Features in Detail](#features-in-detail)
- [Installation](#installation)
- [AI Provider Configuration](#-ai-provider-configuration)
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

2. **Start the app:**
   ```bash
   npm run dev
   ```
   This creates the Python virtual environment, installs dependencies, and starts
   the server. No manual `venv` or `pip` steps are needed.

   <details>
   <summary>Prefer plain Python?</summary>

   ```bash
   python3 -m venv venv
   source venv/bin/activate    # On Windows: venv\Scripts\activate
   pip install -r requirements.txt
   python app.py
   ```
   </details>

3. **Open your browser:**
   Navigate to `http://127.0.0.1:8080/`

4. **Configure your AI provider** in the API Settings dialog — either
   [OpenAI](https://platform.openai.com/) or **Azure AI Foundry**
   (see [AI Provider Configuration](#-ai-provider-configuration) below).

5. **(Optional) Get WHO ICD API credentials** — only for the Medical Translation (ICD-11) feature:
   - Register for free at the [WHO ICD API portal](https://icd.who.int/icdapi)
   - Copy your **Client ID** and **Client Secret**
   - Enter both in the same API Settings dialog

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

### Health Information Officer: Diagnosis Standardization
**Challenge**: A register of free-text diagnoses in mixed languages and spellings needs official ICD codes
**Solution**: The Medical Translation (ICD-11) mode maps each entry to a WHO ICD-11 code and translated term, with an optional ICD-10 mapping and relationship indicator

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

### Medical Coding & Translation (WHO ICD-11)
Map a column of medical diagnoses to authoritative WHO ICD-11 codes and translated terms. This is a dedicated mode (select **Medical Translation (ICD-11)** on the Configuration step) and it uses a **hybrid** of OpenAI + the official [WHO ICD-11 API](https://icd.who.int/icdapi):

- **OpenAI understands & selects** — it normalizes messy free text (typos, abbreviations, lay wording, and American vs. British spelling such as *diarrhea → diarrhoea*) into official search queries and disambiguates the best match.
- **WHO verifies code existence** — official lookups supply codes and terms. Model-generated fallback codes and gap-filled terms are suggestions, identified by per-field source columns and review status. A valid code does not establish that it correctly matches the input.

Capabilities:
- **Two input modes**: free text in any source language, **or** an existing code column (ICD-11 MMS or ICD-10).
- **Any source/target language** from the WHO-supported set (English, Spanish, French, Arabic, Chinese, and more).
- **Typo- and spelling-tolerant matching** via WHO flexisearch + autocode, synonym/index-term search, and AI normalization.
- **Optional ICD-10 mapping** column with a *derived, advisory* relationship indicator: same-as / broader-than / narrower-than / no-map.

#### Output columns

The mapped file adds the following columns (ICD-11 / ICD-10 code columns appear only for the outputs you selected):

| Column | Meaning | Example values |
|---|---|---|
| **ICD-11 Code** | The WHO ICD-11 (MMS) code the entry was matched to. | `NE83`, `5A11`, or blank |
| **ICD-11 Term (XX)** | The official ICD-11 title for that code, translated into your target language (`XX` = language code). | Official term text |
| **ICD-10 Code** | The corresponding ICD-10 code, derived from WHO's official crosswalk. | `T63.0`, `E11`, or blank |
| **ICD-10 Match** | Relationship between the ICD-11 and ICD-10 code (they rarely line up 1:1). | `same-as`, `broader-than`, `narrower-than`, `no-map` |
| **Type** | What the entry is — flags rows that are not true diagnoses. | `Diagnosis`, `Procedure (not codeable in ICD — see ICHI)`, `Other` |
| **Source Term** | The official ICD title (source language) the input was matched to, for verification. | Official term text |
| **Confidence** | How sure the tool is about the match. | `high`, `medium`, `low`, `none` |
| **Source** | Where the answer came from — the WHO/AI audit trail. | `WHO ICD-11 API`, `WHO ICD-11 API (+LLM: term, ICD-10)`, `LLM (gpt-4o-mini)` |
| **Notes** | Plain-language reason a row did not map to a clean diagnosis code. | `Procedure, not a diagnosis…`, `No ICD match found`, `Provided by LLM (not found in WHO API)`, or blank |

> **Reading the results:** `Confidence` + `Source` together tell you how much to trust a row — a `high` / `WHO ICD-11 API` row identifies a WHO code but still needs appropriate matching review, while a `low` / `LLM (gpt-4o-mini)` row is the model's best guess and should be reviewed. `Notes` explains non-diagnosis entries, e.g. procedures, which ICD does not code (use the WHO ICHI classification instead).

> ICD-11 is licensed by WHO under CC BY-ND 3.0 IGO. The ICD-10 mapping and relationship indicators are derived and provided for reference; verify before clinical or statistical use.

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
- Node.js 18+ *(only for the `npm run dev` launcher)*
- An AI provider: an [OpenAI API key](https://platform.openai.com/api-keys) **or** an Azure AI Foundry deployment
- *(Optional)* [WHO ICD API credentials](https://icd.who.int/icdapi) (Client ID + Secret) — only for the Medical Translation (ICD-11) feature

### Local Development

```bash
git clone https://github.com/jmesplana/excel_ai_insight.git
cd excel_ai_insight
npm run dev
```

`npm run dev` creates `./venv`, installs `requirements.txt` (only when it
changes), and starts Flask with auto-reload. Access the app at
`http://127.0.0.1:8080/`.

| Command | What it does |
|---|---|
| `npm run dev` | Set up if needed, then run with auto-reload |
| `npm start` | Run without the debug reloader |
| `npm run setup` | Create the venv and install dependencies, then exit |
| `npm run clean` | Remove `venv/` and Python caches |

Override the port or host with environment variables:

```bash
PORT=3000 npm run dev
```

## 🤖 AI Provider Configuration

The app works with **OpenAI** or a model deployed in your organization's
**Azure AI Foundry** (Azure OpenAI) tenant. Configure it in either place:

### Option A: In the browser (per user)

Open **API Settings**, choose your provider, and fill in the fields. Use
**Test Connection** to verify before running a whole file. Credentials are kept
in your browser's `localStorage` when requested, and sent to the server for processing. WHO authentication is temporarily cached in server memory.

For **Azure AI Foundry** you need three values from the Azure portal:

| Field | Where to find it |
|---|---|
| **Azure Endpoint** | Your resource → *Keys and Endpoint* (e.g. `https://your-resource.openai.azure.com`) |
| **Azure API Key** | Either key on that same page |
| **Deployment Name** | *Model deployments* in Azure AI Foundry — the name **your organization chose**, which is often different from the model name |
| **API Version** | Optional; defaults to `2024-10-21` |

### Option B: Server-side (shared deployment)

Set environment variables so users never have to enter credentials. Copy
`.env.example` to `.env` and fill in:

```bash
LLM_PROVIDER=azure
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com
AZURE_OPENAI_API_KEY=your-azure-key
AZURE_OPENAI_DEPLOYMENT=your-deployment-name
AZURE_OPENAI_API_VERSION=2024-10-21   # optional
```

Or for OpenAI:

```bash
LLM_PROVIDER=openai
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4o-mini   # optional
```

> **Precedence:** values entered in the browser always win. Environment
> variables are the fallback, so an organization can ship a working default
> while still letting individuals use their own key.

`.env` is gitignored — never commit real credentials.

## 📖 Usage Guide

### Step 1: Configuration
1. Click the **Settings** button
2. Choose your **AI provider** (OpenAI or Azure AI Foundry) and enter its
   credentials, then click **Test Connection**. You can skip this if the server
   already has a provider configured.
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

### Alternative: Medical Translation (ICD-11)
On the Configuration step, choose **Medical Translation (ICD-11)** instead of AI Analysis, then:
1. Make sure your **WHO ICD Client ID + Secret** are set in API Settings (and your OpenAI key, for best matching)
2. Upload your file and pick the **source column**
3. Choose the **input type** — *Free text* (with a source language) or *Existing code* (ICD-11 MMS or ICD-10)
4. Pick the **target language** and which outputs you want (ICD-11 term and/or ICD-10 code + relationship)
5. Click **Run** and download the mapped file

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
- AI credentials are optional — users can provide their own in the browser. To
  ship a shared deployment instead, set `LLM_PROVIDER` plus the matching
  `AZURE_OPENAI_*` or `OPENAI_*` variables in your Vercel project settings.

## 🔒 Privacy & Security

- **API keys stored locally**: Your AI provider credentials (OpenAI or Azure) and WHO ICD credentials can be saved in browser storage. Credentials pass through the server; WHO tokens and credential cache keys are temporarily held in memory
- **Bring your own model**: Point the app at your organization's Azure AI Foundry deployment so prompts stay inside your own Azure tenant
- **No data retention**: Uploaded files are parsed in your browser and rows are sent to the server only transiently for processing—no files are stored
- **Medical provenance**: Results distinguish WHO lookups from model suggestions. Review suggested matches before use; confidence is not a verified measure of matching accuracy.
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

**Error: "Invalid ICD API credentials" (Medical Translation mode)**
- **Cause**: WHO ICD Client ID/Secret missing or incorrect
- **Solution**: Register at the [WHO ICD API portal](https://icd.who.int/icdapi) and enter both the Client ID and Secret in API Settings

**ICD rows come back blank**
- **Cause**: Free-text terms that don't match WHO's (British English) terminology, or no OpenAI key for normalization
- **Solution**: Add your OpenAI API key (it normalizes spelling/typos before lookup); for coded input, confirm the correct source classification (ICD-11 MMS vs ICD-10) is selected

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
- Medical coding via the [WHO ICD-11 API](https://icd.who.int/icdapi) — ICD-11 © World Health Organization (CC BY-ND 3.0 IGO)
- UI components from [Bootstrap 5](https://getbootstrap.com/)
- Design system follows [Aidstack Brand Guidelines](aidstack-brand-guide.md)

## 📧 Support

- **Issues**: [GitHub Issues](https://github.com/jmesplana/excel_ai_insight/issues)
- **Discussions**: [GitHub Discussions](https://github.com/jmesplana/excel_ai_insight/discussions)

---

**Part of the Aidstack.ai ecosystem** | Visit [aidstack.ai](https://aidstack.ai) to explore more tools


## Development structure

- `app.py`: application factory, page routes, and feature registration.
- `insights/analysis.py`: connection testing, pattern detection, and row analysis.
- `insights/chat.py`: dataset statistics and streaming chat.
- `insights/icd.py`: WHO lookups, mapping, and per-field provenance.
- `insights/errors.py`: safe public error messages.
- `static/js/app.js`: workflow orchestration and event wiring.
- `static/js/results.js`: results tables and chat controls, initialized with explicit dependencies.
- `static/js/provider-settings.js`: provider configuration and form visibility.
- `static/js/batch-runner.js`: resumable in-memory batch execution.
- `static/js/spreadsheet.js`: header validation, typed imports, and table export.
- `static/js/rendering.js`: text escaping and sanitized Markdown.
- `templates/partials/`: landing page, settings, and individual workflow steps.
- `static/css/app.css`: application styles.

Run regression tests after changing feature boundaries:

```bash
npm test
npm run test:backend
npm run test:browser
```

The browser suite uses Chrome, starts its own local server on port 8091, and
mocks provider responses. Run `npm install` once to install the test dependency.

Analysis runs can stop after the current batch, download completed rows, and
resume from the last acknowledged batch using the original settings. Checkpoints
stay in memory: keep the page open, or download partial results before leaving.
Cell-level errors are retained in the export; resume retries an interrupted batch,
not individual cells that already returned an error. A request whose response was
lost may have incurred provider usage even though its batch must be retried.

Exports are **results tables**, containing the selected sheet's values and new
outputs. They do not preserve the original workbook's other sheets, formulas, or
formatting. Duplicate or blank column headers must be corrected before import.

Output length can be selected per analysis configuration. Oversized inputs and
outputs that hit the response limit return explicit cell errors instead of silently
truncated results. Medical exports include review status, WHO references, and
separate sources for ICD-11 codes, translated terms, and ICD-10 mappings. The
review filter changes the on-screen view; exports include every result row.

DOMPurify is vendored at `static/vendor/purify.es.mjs` (version and license in its
header). All AI Markdown must go through `renderMarkdown`; spreadsheet text and
attribute values must use `escapeHtml` or DOM `textContent`.
