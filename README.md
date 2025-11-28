# Company Contact Scraper

A web scraper that extracts email addresses and phone numbers from company websites using Selenium and regex patterns.

## Features

- **Smart Contact Page Detection**: Scores all internal links to find the best contact page
- **Multi-Source Extraction**: Extracts from `mailto:` links, `tel:` links, and page text
- **Resume Capability**: Re-run on the same file and it continues from where it left off
- **Auto-Save**: Saves progress after each website to `_output.xlsx`
- **Cookie Handling**: Automatically handles cookie consent dialogs
- **Page Scrolling**: Scrolls pages to load lazy content

## Prerequisites

- Python 3.9+
- Google Chrome browser

## Installation

```bash
pip install pandas selenium webdriver-manager beautifulsoup4 openpyxl
```

## Usage

1. Place your Excel file (`.xlsx` or `.xlsm`) in the same directory
2. Update the file name in `scraper.py` if needed:
   ```python
   self.excel_file = "your_file.xlsx"
   ```
3. Run:
   ```bash
   python scraper.py
   ```
4. Enter how many websites to process

## Excel Format

Your Excel file should have a column with website URLs. The scraper will:
- Auto-detect the URL column
- Create "Email" and "Phone number" columns if they don't exist
- Save results to `yourfile_output.xlsx`

## How It Works

1. Loads Excel file (or `_output.xlsx` if resuming)
2. Finds rows with empty email/phone fields
3. For each website:
   - Loads homepage, scrolls, extracts contacts
   - Finds best contact page using keyword scoring
   - Loads contact page, scrolls, extracts contacts
   - Saves results immediately

## Contact Page Detection

The scraper scores all internal links based on keywords:
- **+10 points**: Link text contains keyword
- **+5 points**: URL contains keyword

Keywords: `contact`, `kontakt`, `contacto`, `contatti`, `impressum`, `about`, `uber-uns`, etc.

## Extraction Patterns

**Emails**: Any valid email format (`user@domain.tld`)

**Phones**: International numbers starting with `+` (e.g., `+49 123 456789`, `+351 225400326`)

Also extracts from:
- `<a href="mailto:email@example.com">`
- `<a href="tel:+123456789">`

## Resume Feature

If you re-run the scraper on a file that was already processed:
- It loads from `_output.xlsx` (previous results)
- Only processes rows where email/phone is empty
- Skips rows with "Not found", "Error", or "Duplicate"

## Output

Results are saved to `yourfile_output.xlsx` with columns:
- Original URL column
- Email
- Phone number

## Troubleshooting

**"Failed to start browser"**
- Make sure Chrome is installed
- Try updating Chrome

**"File not found"**
- Check the file name in `scraper.py`

**"No URL column found"**
- Ensure your Excel has URLs containing `.com`, `.lt`, `.eu`, etc.

**Not finding contacts**
- Some sites block scrapers
- Contact info may be loaded via JavaScript (currently disabled for speed)

## Configuration

Edit these in `scraper.py`:

```python
self.excel_file = "test.xlsx"      # Input file
self.email_column = "Email"         # Email column name
self.phone_column = "Phone number"  # Phone column name
```

## License

For educational and business use.
