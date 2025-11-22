import pandas as pd
import re
import time
import os
import logging
import sys
from typing import Tuple, List, Optional
from dataclasses import dataclass

# Selenium imports (Selenium 4+)
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# BeautifulSoup for HTML parsing
from bs4 import BeautifulSoup

# Google Gemini AI
from google import genai

import urllib3
from urllib.parse import urlparse

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('scraper.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


@dataclass
class ContactInfo:
    """Data class to hold extracted contact information"""
    emails: List[str]
    phones: List[str]
    raw_text: str = ""


class GeminiExtractor:
    """Handles AI-powered extraction of contact information using Gemini"""
    
    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize Gemini client
        
        Args:
            api_key: Optional API key. If not provided, will use GEMINI_API_KEY env variable
        """
        try:
            if api_key:
                os.environ['GEMINI_API_KEY'] = api_key
            
            # Initialize Gemini client
            self.client = genai.Client()
            self.model = "gemini-2.5-flash"
            logger.info("✅ Gemini AI initialized successfully")
        except Exception as e:
            logger.error(f"❌ Failed to initialize Gemini: {str(e)}")
            self.client = None
    
    def extract_contacts(self, html_content: str, url: str) -> ContactInfo:
        """
        Use Gemini AI to intelligently extract emails and phone numbers
        
        Args:
            html_content: Raw HTML content
            url: Website URL for context
            
        Returns:
            ContactInfo object with extracted emails and phones
        """
        if not self.client:
            logger.warning("Gemini client not available, falling back to regex")
            return ContactInfo(emails=[], phones=[])
        
        try:
            # Parse HTML to get clean text
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Remove script and style elements
            for script in soup(["script", "style", "noscript"]):
                script.decompose()
            
            # Get text content
            text = soup.get_text(separator=' ', strip=True)
            
            # Limit text length to avoid token limits (keep first 8000 chars)
            text = text[:8000]
            
            # Create prompt for Gemini
            prompt = f"""You are an expert at extracting contact information from website content.

Website URL: {url}

Website Content:
{text}

Task: Extract ALL email addresses and phone numbers from this content.

Requirements:
1. Extract ONLY valid email addresses (format: user@domain.com)
2. Extract ONLY valid phone numbers (international format preferred, e.g., +370 123 45678)
3. Return results in JSON format ONLY
4. Do not include any explanations, just the JSON

Expected JSON format:
{{
    "emails": ["email1@example.com", "email2@example.com"],
    "phones": ["+370 123 45678", "+44 20 1234 5678"]
}}

If no emails or phones found, return empty arrays.
"""
            
            # Generate content using Gemini
            response = self.client.models.generate_content(
                model=self.model,
                contents=prompt
            )
            
            # Parse response
            response_text = response.text.strip()
            
            # Extract JSON from response (handle markdown code blocks)
            if '```json' in response_text:
                response_text = response_text.split('```json')[1].split('```')[0].strip()
            elif '```' in response_text:
                response_text = response_text.split('```')[1].split('```')[0].strip()
            
            # Parse JSON
            import json
            data = json.loads(response_text)
            
            emails = data.get('emails', [])
            phones = data.get('phones', [])
            
            logger.info(f"🤖 Gemini extracted: {len(emails)} emails, {len(phones)} phones")
            
            return ContactInfo(
                emails=emails[:5],  # Limit to 5 emails
                phones=phones[:5],  # Limit to 5 phones
                raw_text=text
            )
            
        except Exception as e:
            logger.error(f"❌ Gemini extraction failed: {str(e)}")
            return ContactInfo(emails=[], phones=[])


class CompanyContactScraper:
    """Main scraper class with modern Selenium 4+ implementation"""
    
    def __init__(self, excel_file: str = "test.xlsm", gemini_api_key: Optional[str] = None):
        self.excel_file = excel_file
        self.df = None
        self.driver = None
        self.scraped_urls = set()
        self.processed_count = 0
        
        # Column names
        self.url_column = None
        self.email_column = "Email"
        self.phone_column = "Phone number"
        
        # Initialize Gemini extractor
        self.gemini = GeminiExtractor(api_key=gemini_api_key)
        
        # Fallback regex patterns
        self.email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
        self.phone_patterns = [
            r'\+(?:[1-9]\d{0,3})[\s\-]?(?:\d{1,5}[\s\-]?){3,8}\d{1,5}',
            r'(?:\+?\(?\d{1,4}\)?[\s\-]?)?\(?\d{2,5}\)?[\s\-.]?\d{1,5}[\s\-.]?\d{1,5}[\s\-.]?\d{1,5}',
        ]
    
    def setup_driver(self) -> bool:
        """Setup Chrome browser with Selenium 4+ best practices"""
        try:
            chrome_options = Options()
            
            # Modern Selenium 4 options
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--log-level=3")
            
            # User agent to avoid detection
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            
            # Exclude automation flags
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # Initialize service
            service = Service(ChromeDriverManager().install())
            
            # Create driver
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # Set timeouts (Selenium 4 way)
            self.driver.set_page_load_timeout(30)
            self.driver.implicitly_wait(5)
            
            # Execute CDP commands to avoid detection
            self.driver.execute_cdp_cmd('Network.setUserAgentOverride', {
                "userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            })
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            logger.info("✅ Browser started successfully")
            return True
            
        except Exception as e:
            logger.error(f"❌ Failed to start browser: {str(e)}")
            return False
    
    def detect_url_column(self) -> Optional[str]:
        """Automatically detect which column contains URLs"""
        logger.info("🔍 Detecting URL column...")
        
        for col in self.df.columns:
            non_empty_values = self.df[col].dropna()
            
            if len(non_empty_values) > 0:
                for value in non_empty_values.head(5):
                    value_str = str(value).strip().lower()
                    
                    if any(domain in value_str for domain in ['.com', '.lt', '.eu', '.org', '.net', 'http', 'www.', '//']):
                        logger.info(f"✅ Found URL column: '{col}'")
                        return col
        
        if "Row Labels" in self.df.columns:
            logger.info("✅ Using 'Row Labels' column as URL source")
            return "Row Labels"
        
        logger.warning("❌ No URL column found automatically")
        return None
    
    def clean_dataframe(self) -> bool:
        """Clean the dataframe by removing header rows and non-URL entries"""
        logger.info("🧹 Cleaning data...")
        
        if self.url_column is None:
            return False
        
        original_count = len(self.df)
        
        # Remove header rows
        header_texts = ['Row Labels', 'unique website link', 'Count of date when found']
        for header in header_texts:
            self.df = self.df[self.df[self.url_column] != header]
        
        # Keep only rows with valid URLs
        def is_website(url):
            if pd.isna(url):
                return False
            url_str = str(url).strip().lower()
            return any(domain in url_str for domain in ['.com', '.lt', '.eu', '.org', '.net', 'http', 'www.'])
        
        self.df = self.df[self.df[self.url_column].apply(is_website)]
        self.df.reset_index(drop=True, inplace=True)
        
        cleaned_count = len(self.df)
        logger.info(f"✅ Cleaned data: {original_count} → {cleaned_count} rows")
        
        return True
    
    def load_excel_file(self) -> bool:
        """Load and validate Excel file"""
        try:
            if not os.path.exists(self.excel_file):
                logger.error(f"❌ Excel file '{self.excel_file}' not found!")
                return False
            
            logger.info("📖 Reading Excel file...")
            
            try:
                self.df = pd.read_excel(self.excel_file, engine='openpyxl')
            except Exception as e:
                logger.warning(f"First read attempt failed: {e}")
                self.df = pd.read_excel(self.excel_file, engine='openpyxl', header=1)
            
            logger.info(f"📊 Found columns: {list(self.df.columns)}")
            logger.info(f"📝 Total rows before cleaning: {len(self.df)}")
            
            # Detect URL column
            self.url_column = self.detect_url_column()
            if not self.url_column:
                logger.error("❌ Could not find URL column")
                return False
            
            # Clean dataframe
            if not self.clean_dataframe():
                return False
            
            # Create email and phone columns if needed
            if self.email_column not in self.df.columns:
                self.df[self.email_column] = ""
            if self.phone_column not in self.df.columns:
                self.df[self.phone_column] = ""
            
            logger.info(f"✅ Excel file loaded successfully with {len(self.df)} rows")
            return True
            
        except Exception as e:
            logger.error(f"❌ Error loading Excel file: {str(e)}")
            return False
    
    def handle_cookie_consent(self):
        """Handle cookie consent dialogs"""
        try:
            wait = WebDriverWait(self.driver, 3)
            
            # Common cookie button selectors
            selectors = [
                (By.XPATH, "//button[contains(translate(., 'ACCEPT', 'accept'), 'accept')]"),
                (By.XPATH, "//button[contains(translate(., 'AGREE', 'agree'), 'agree')]"),
                (By.XPATH, "//a[contains(translate(., 'ACCEPT', 'accept'), 'accept')]"),
                (By.ID, "accept-cookies"),
                (By.CLASS_NAME, "cookie-accept"),
            ]
            
            for by, selector in selectors:
                try:
                    button = wait.until(EC.element_to_be_clickable((by, selector)))
                    button.click()
                    logger.info("✅ Accepted cookies")
                    time.sleep(1)
                    return True
                except:
                    continue
            
            return False
            
        except Exception as e:
            logger.debug(f"Cookie handling: {str(e)}")
            return False
    
    def navigate_to_contact_page(self) -> bool:
        """Try to navigate to contact page"""
        try:
            contact_keywords = ['contact', 'kontakt', 'contacto', 'contatti', 'impressum']
            
            for keyword in contact_keywords:
                try:
                    # Use modern Selenium 4 syntax
                    wait = WebDriverWait(self.driver, 5)
                    link = wait.until(EC.element_to_be_clickable((
                        By.XPATH, 
                        f"//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{keyword}')]"
                    )))
                    link.click()
                    logger.info(f"✅ Navigated to contact page ({keyword})")
                    time.sleep(2)
                    return True
                except:
                    continue
            
            return False
            
        except Exception as e:
            logger.debug(f"Contact page navigation: {str(e)}")
            return False
    
    def scrape_website(self, url: str) -> Tuple[str, str]:
        """Scrape a single website for contact information"""
        logger.info(f"\n🌐 Scraping: {url}")
        
        if url in self.scraped_urls:
            logger.info("⏭️  Duplicate URL, skipping...")
            return "Duplicate", "Duplicate"
        
        try:
            # Ensure URL has protocol
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
            
            # Navigate to website
            self.driver.get(url)
            time.sleep(2)
            
            # Handle cookies
            self.handle_cookie_consent()
            
            # Try to navigate to contact page
            self.navigate_to_contact_page()
            
            # Get page source
            page_source = self.driver.page_source
            
            # Use Gemini AI for extraction
            contact_info = self.gemini.extract_contacts(page_source, url)
            
            # Format results
            email_result = ', '.join(contact_info.emails) if contact_info.emails else "Not found"
            phone_result = ', '.join(contact_info.phones) if contact_info.phones else "Not found"
            
            if email_result != "Not found":
                logger.info(f"✅ Found emails: {email_result}")
            else:
                logger.info("❌ No emails found")
            
            if phone_result != "Not found":
                logger.info(f"✅ Found phones: {phone_result}")
            else:
                logger.info("❌ No phones found")
            
            self.scraped_urls.add(url)
            return email_result, phone_result
            
        except Exception as e:
            logger.error(f"❌ Error scraping {url}: {str(e)}")
            return "Error", "Error"
    
    def find_empty_rows(self, count: int) -> List[Tuple[int, str]]:
        """Find rows that need to be processed"""
        empty_rows = []
        
        for index, row in self.df.iterrows():
            if len(empty_rows) >= count:
                break
            
            email_empty = pd.isna(row[self.email_column]) or str(row[self.email_column]).strip() in ['', 'Not found', 'Error', 'Duplicate']
            phone_empty = pd.isna(row[self.phone_column]) or str(row[self.phone_column]).strip() in ['', 'Not found', 'Error', 'Duplicate']
            
            url = row[self.url_column] if pd.notna(row[self.url_column]) else ""
            if (email_empty or phone_empty) and url and str(url).strip():
                empty_rows.append((index, str(url).strip()))
        
        return empty_rows
    
    def update_excel_row(self, index: int, email: str, phone: str):
        """Update a specific row in the Excel data"""
        try:
            # Handle Excel's + issue with phone numbers
            if phone and phone not in ['Not found', 'Error', 'Duplicate'] and phone.startswith('+'):
                phone = "'" + phone
            
            self.df.at[index, self.email_column] = email
            self.df.at[index, self.phone_column] = phone
            
        except Exception as e:
            logger.error(f"❌ Error updating row {index}: {str(e)}")
    
    def save_excel(self) -> bool:
        """Save the Excel file"""
        try:
            # Create backup
            backup_file = self.excel_file.replace('.xlsm', '_backup.xlsm')
            if os.path.exists(self.excel_file):
                import shutil
                shutil.copy2(self.excel_file, backup_file)
                logger.info("💾 Created backup file")
            
            # Save
            self.df.to_excel(self.excel_file, index=False, engine='openpyxl')
            logger.info("✅ Excel file saved successfully")
            return True
            
        except Exception as e:
            logger.error(f"❌ Error saving Excel file: {str(e)}")
            return False
    
    def show_progress(self):
        """Show current progress statistics"""
        if self.df is None or self.url_column is None:
            return
        
        total_rows = len(self.df)
        emails_filled = self.df[self.email_column].notna() & (self.df[self.email_column] != '') & (self.df[self.email_column] != 'Not found')
        phones_filled = self.df[self.phone_column].notna() & (self.df[self.phone_column] != '') & (self.df[self.phone_column] != 'Not found')
        
        emails_count = emails_filled.sum()
        phones_count = phones_filled.sum()
        completion = ((emails_filled | phones_filled).sum() / total_rows) * 100 if total_rows > 0 else 0
        
        print(f"\n📊 Current Progress:")
        print(f"   • Total rows: {total_rows}")
        print(f"   • Emails found: {emails_count}")
        print(f"   • Phones found: {phones_count}")
        print(f"   • Completion: {completion:.1f}%")
    
    def run(self):
        """Main application function"""
        print("\n" + "="*60)
        print("     COMPANY CONTACT SCRAPER WITH GEMINI AI")
        print("="*60)
        
        # Load Excel file
        if not self.load_excel_file():
            input("\nPress Enter to exit...")
            return
        
        self.show_progress()
        
        # Get user input
        try:
            num_websites = int(input("\n👉 How many websites would you like to process? "))
            if num_websites <= 0:
                print("❌ Please enter a positive number.")
                return
        except ValueError:
            print("❌ Please enter a valid number.")
            return
        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            return
        
        # Setup browser
        if not self.setup_driver():
            print("❌ Failed to start browser. Please check if Chrome is installed.")
            input("\nPress Enter to exit...")
            return
        
        try:
            # Find rows to process
            rows_to_process = self.find_empty_rows(num_websites)
            
            if not rows_to_process:
                print("🎉 All rows already have contact information!")
                input("\nPress Enter to exit...")
                return
            
            actual_count = min(num_websites, len(rows_to_process))
            print(f"\n🔍 Found {actual_count} rows to process")
            
            # Process websites
            for i, (index, url) in enumerate(rows_to_process, 1):
                if self.processed_count >= num_websites:
                    break
                
                print(f"\n📈 Progress: {i}/{actual_count}")
                
                # Scrape the website
                email, phone = self.scrape_website(url)
                
                # Update the data
                self.update_excel_row(index, email, phone)
                self.processed_count += 1
                
                # Save after each website
                if self.save_excel():
                    logger.info("💾 Progress saved to Excel")
                
                # Delay between requests
                time.sleep(3)
            
            # Final summary
            print("\n" + "="*60)
            print("                     SUMMARY")
            print("="*60)
            print(f"✅ Successfully processed: {self.processed_count} websites")
            print(f"💾 Data saved to: {self.excel_file}")
            print(f"📋 Check scraper.log for detailed logs")
            print("="*60)
            
        except KeyboardInterrupt:
            print("\n⚠️  Process interrupted by user. Saving progress...")
            self.save_excel()
        except Exception as e:
            logger.error(f"❌ Unexpected error: {str(e)}")
            print(f"\n❌ Unexpected error: {str(e)}")
        finally:
            # Cleanup
            if self.driver:
                self.driver.quit()
                logger.info("🌐 Browser closed")
            
            input("\nPress Enter to close this app...")


def main():
    """Entry point"""
    # You can set your Gemini API key here or use environment variable
    # Get API key from: https://aistudio.google.com/app/apikey
    
    gemini_api_key = os.getenv('GEMINI_API_KEY')  # Or set directly: "your-api-key-here"
    
    if not gemini_api_key:
        print("⚠️  Warning: GEMINI_API_KEY not found in environment variables")
        print("Please set it using: set GEMINI_API_KEY=your-api-key")
        print("Get your API key from: https://aistudio.google.com/app/apikey")
        print("\nContinuing without Gemini AI (will use regex fallback)...\n")
    
    scraper = CompanyContactScraper(gemini_api_key=gemini_api_key)
    scraper.run()


if __name__ == "__main__":
    main()
