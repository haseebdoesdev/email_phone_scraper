# 🚀 Company Contact Scraper with Gemini AI

An intelligent web scraper that extracts email addresses and phone numbers from company websites using **Selenium 4+** and **Google's Gemini AI**.

## ✨ Features

- 🤖 **AI-Powered Extraction**: Uses Google Gemini AI for intelligent contact information extraction
- 🔄 **Modern Selenium 4+**: Updated implementation with latest best practices
- 📊 **Excel Integration**: Reads from and writes to Excel files (.xlsm, .xlsx)
- 🍪 **Cookie Handling**: Automatically handles cookie consent dialogs
- 🔍 **Smart Navigation**: Finds and navigates to contact pages automatically
- 💾 **Auto-Save**: Saves progress after each website
- 📝 **Detailed Logging**: Comprehensive logging to both console and file
- 🛡️ **Error Handling**: Robust error handling and recovery

## 🎯 Key Improvements

### 1. **Gemini AI Integration**
- Intelligent extraction using Google's latest AI model
- Context-aware parsing of website content
- Better accuracy than regex-only approaches
- Handles multiple languages and formats

### 2. **Modern Selenium 4+**
- Updated WebDriver initialization
- Chrome DevTools Protocol (CDP) commands for stealth
- Improved element location strategies
- Better timeout and wait handling
- Modern `By` locators

### 3. **Code Quality**
- Type hints for better code clarity
- Dataclasses for structured data
- Separation of concerns (GeminiExtractor, CompanyContactScraper)
- Comprehensive error handling
- Clean, maintainable code structure

## 📋 Prerequisites

1. **Python 3.9+**
2. **Google Chrome** browser installed
3. **Gemini API Key** (get from [Google AI Studio](https://aistudio.google.com/app/apikey))

## 🔧 Installation

### Step 1: Clone or Download

Download the files to your desired directory.

### Step 2: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 3: Set Up Gemini API Key

Get your free API key from: https://aistudio.google.com/app/apikey

**Option A: Environment Variable (Recommended)**
```bash
# Windows
set GEMINI_API_KEY=your-api-key-here

# Linux/Mac
export GEMINI_API_KEY=your-api-key-here
```

**Option B: Direct in Code**
Edit `scraper.py` and set the API key in the `main()` function:
```python
gemini_api_key = "your-api-key-here"
```

## 📁 File Structure

```
gytis_email_phone_scraper/
├── scraper.py           # Main scraper script
├── requirements.txt     # Python dependencies
├── README.md           # This file
├── test.xlsm           # Your Excel file (place here)
└── scraper.log         # Generated log file
```

## 🚀 Usage

### Step 1: Prepare Your Excel File

Place your Excel file (`test.xlsm`) in the same directory. The file should have:
- A column with website URLs (will be auto-detected)
- Optional: "Email" column
- Optional: "Phone number" column

The scraper will automatically:
- Detect which column contains URLs
- Create Email and Phone columns if they don't exist
- Clean the data (remove headers, invalid entries)

### Step 2: Run the Scraper

```bash
python scraper.py
```

### Step 3: Follow the Prompts

1. The scraper will show you the detected columns and data
2. Enter how many websites you want to process
3. Sit back and watch the magic happen! ✨

## 📊 Excel File Format

Your Excel file can have any structure, but should contain URLs. Example:

| Row Labels | Email | Phone number |
|-----------|-------|--------------|
| example.com | | |
| company.lt | | |
| website.eu | | |

The scraper will:
- Auto-detect the URL column
- Fill in the Email and Phone columns
- Save progress after each website

## 🎯 How It Works

1. **Load Excel**: Reads your Excel file and detects URL column
2. **Clean Data**: Removes headers and invalid entries
3. **Initialize Browser**: Starts Chrome with stealth settings
4. **For Each Website**:
   - Navigate to the URL
   - Handle cookie consent
   - Try to find contact page
   - Extract page content
   - Use Gemini AI to extract emails and phones
   - Save results to Excel
5. **Save & Report**: Final summary and logs

## 🤖 Gemini AI Extraction

The scraper uses Google's Gemini 2.0 Flash model to:
- Parse HTML content intelligently
- Understand context and structure
- Extract valid emails and phone numbers
- Handle multiple languages
- Filter out invalid data

**Fallback**: If Gemini is unavailable, the scraper falls back to regex patterns.

## 📝 Logging

All activities are logged to:
- **Console**: Real-time progress
- **scraper.log**: Detailed log file

Log levels:
- ✅ INFO: Successful operations
- ⚠️ WARNING: Non-critical issues
- ❌ ERROR: Errors and failures

## 🛡️ Error Handling

The scraper handles:
- Missing Excel files
- Invalid URLs
- Website timeouts
- Cookie dialogs
- Network errors
- Gemini API errors
- Excel save errors

Progress is saved after each website, so you won't lose data if interrupted.

## ⚙️ Configuration

You can customize the scraper by modifying these variables in `scraper.py`:

```python
# Excel file name
self.excel_file = "test.xlsm"

# Column names
self.email_column = "Email"
self.phone_column = "Phone number"

# Gemini model
self.model = "gemini-2.0-flash-exp"

# Timeouts
self.driver.set_page_load_timeout(30)
self.driver.implicitly_wait(5)
```

## 🔍 Troubleshooting

### "GEMINI_API_KEY not found"
- Set the environment variable or add it directly in the code
- Get your API key from https://aistudio.google.com/app/apikey

### "Failed to start browser"
- Make sure Google Chrome is installed
- Try updating Chrome to the latest version
- Check if ChromeDriver is compatible

### "Excel file not found"
- Make sure `test.xlsm` is in the same directory
- Check the file name matches exactly

### "No URL column found"
- Verify your Excel file has a column with website URLs
- URLs should contain `.com`, `.lt`, `.eu`, etc.

## 📈 Performance

- **Speed**: ~3-5 seconds per website (with delays to be respectful)
- **Accuracy**: High accuracy with Gemini AI
- **Rate Limiting**: Built-in delays to avoid overwhelming servers

## 🔒 Privacy & Ethics

This scraper:
- Only extracts publicly available contact information
- Respects robots.txt (when possible)
- Includes delays between requests
- Does not store sensitive data
- Is intended for legitimate business use only

**Please use responsibly and ethically!**

## 📄 License

This project is provided as-is for educational and business purposes.

## 🤝 Contributing

Feel free to improve the code! Some ideas:
- Add proxy support
- Implement rate limiting
- Add more extraction patterns
- Support more file formats
- Add GUI interface

## 📞 Support

If you encounter issues:
1. Check `scraper.log` for detailed error messages
2. Verify all prerequisites are installed
3. Make sure your Gemini API key is valid
4. Try with a smaller number of websites first

## 🎉 Credits

Built with:
- [Selenium](https://www.selenium.dev/) - Web automation
- [Google Gemini AI](https://ai.google.dev/) - Intelligent extraction
- [Pandas](https://pandas.pydata.org/) - Data manipulation
- [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) - HTML parsing

---

**Happy Scraping! 🚀**
