# Release Notes

## Version 1.1.0 - October 17, 2025

### 🎉 Major Updates

#### Enhanced Landing Page with Mock UI
- **New Interactive Demo**: Added live mock spreadsheet showing real-time AI analysis examples
- **Relatable Use Cases**: Three persona-driven stories (E-commerce Manager, HR Professional, Market Researcher) demonstrating real-world applications
- **Pain-Focused Messaging**: Updated hero section to address user frustrations and time-wasting concerns
- **Visual Improvements**: Added gradient backgrounds, hover effects, and persona badges for better engagement

#### Translation Service Template
- **Universal Translation**: New generic translation template supporting any language
- **User-Friendly Placeholder**: Clear instructions to replace `[TARGET_LANGUAGE]` with desired language
- **Consolidated Design**: Replaced three language-specific templates with one flexible template

#### Prompt Writing Guidance
- **Best Practices Tip**: Added yellow alert box with prompt optimization tips
- **Concise Output Instructions**: Clear guidance on how to get short, focused AI responses
- **Example Prompts**: Demonstrated proper formatting for sentiment analysis and other concise outputs

### 🐛 Bug Fixes

#### Critical: Pattern Detection Sampling Error
- **Issue**: `"Cannot take a larger sample than population when 'replace=False'"` error when detecting patterns
- **Root Cause**: Code was attempting to sample more rows than available non-null values
- **Fix**:
  - Now calculates sample size based on actual non-null values count (app.py:193-198)
  - Added validation to check for empty columns before sampling
  - Returns clear error message if column contains no valid data
- **Files Changed**: `app.py` lines 191-198
- **Impact**: Pattern detection now works reliably on all dataset sizes

#### File Extension Validation
- **Issue**: File type checking for `.xlsx` and `.xls` files was failing
- **Fix**: Corrected `filename.endswith()` to include dots (`.xlsx` instead of `xlsx`)
- **Files Changed**: `app.py` lines 180, 312, 444
- **Impact**: Excel file uploads now work correctly on all routes

### 📝 Documentation

#### Comprehensive README Overhaul
- **New Sections Added**:
  - Quick Start guide for faster onboarding
  - Use Cases with specific personas and outcomes
  - Features in Detail with code examples
  - Troubleshooting section with common issues
  - Privacy & Security explanation
- **Better Organization**: Emoji icons, clear headings, and logical flow
- **Updated Instructions**: More detailed usage guide with step-by-step instructions
- **Prompt Templates**: All template examples with explanations

### 🎨 UI/UX Improvements

#### New CSS Styles
- **Mock UI Container**: Gradient background with professional styling
- **Mock Spreadsheet**: Realistic table design with AI result highlighting
- **Use Case Cards**: Interactive cards with hover effects and persona badges
- **Color Coding**: Green highlights for AI-generated results
- **Dark Mode Support**: All new components support dark mode

#### Template Suggestions
- **New Templates**:
  - Sentiment Analysis (Concise) - One-word outputs
  - Translation Service - Universal language translation
  - Customer Feedback Analysis (Detailed) - Comprehensive analysis
- **Improved Existing**:
  - Business Analysis
  - Data Quality Check

### 🔧 Technical Details

#### Files Modified
```
templates/index.html
- Lines 403-493: New CSS for mock UI and use cases
- Lines 553-598: New hero section with mock spreadsheet
- Lines 600-686: Relatable use cases section
- Lines 680-712: Updated prompt templates

app.py
- Lines 191-198: Fixed pattern detection sampling
- Lines 180, 312, 444: Fixed file extension checks

README.md
- Complete rewrite with better structure
- Added use cases, troubleshooting, and detailed guides

RELEASE_NOTES.md
- New file tracking all changes
```

### 📊 Performance & Reliability

- **No Performance Regression**: All fixes are logic-only, no performance impact
- **Improved Error Handling**: Better error messages for users
- **Enhanced Stability**: Pattern detection now works on edge cases

### 🚀 Deployment

- **Vercel Compatible**: All changes work seamlessly on Vercel deployment
- **No Breaking Changes**: Existing functionality preserved
- **Backward Compatible**: Old uploaded files and configurations still work

### 🔜 Coming Soon

- Export to multiple formats (PDF, JSON)
- Batch processing for multiple files
- Custom AI model selection
- API rate limit visualization
- Collaboration features

---

## Version 1.0.0 - Initial Release

### Features
- Excel and CSV file upload
- AI-powered text analysis
- Multi-column analysis
- Real-time progress tracking
- Dark mode support
- Pattern detection
- OpenAI GPT-4o-mini integration
- Vercel deployment support

---

**Installation & Upgrade Instructions**

To upgrade to v1.1.0:
```bash
git pull origin main
pip install -r requirements.txt
python app.py
```

For fresh installation, see the [README.md](README.md)

---

**Questions or Issues?**

- Report bugs: [GitHub Issues](https://github.com/jmesplana/excel_ai_insight/issues)
- Discussions: [GitHub Discussions](https://github.com/jmesplana/excel_ai_insight/discussions)
