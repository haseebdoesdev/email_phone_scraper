import pandas as pd
import re
import time
import os
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from urllib.parse import urlparse


class CompanyContactScraper:
    def __init__(self, excel_file: str = r"test.xlsx"):
        self.excel_file = excel_file
        self.df = None
        self.driver = None
        self.scraped_urls = set()
        self.processed_count = 0
        self.url_column = None
        self.email_column = "Email"
        self.phone_column = "Phone number"
        
        # Regex patterns - broad email TLDs (all common ones)
        self.email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[a-z]{2,}\b', re.IGNORECASE)
        # Phone pattern - any international number starting with +
        self.phone_pattern = re.compile(r'\+\d{1,4}[\s\-./]?(?:\(?\d{1,5}\)?[\s\-./]?){1,8}\d{1,5}')
    
    def setup_driver(self):
        """Setup Chrome browser"""
        try:
            chrome_options = Options()
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--log-level=3")
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-plugins")
            chrome_options.add_argument("--disable-images")
            chrome_options.add_argument("--disable-javascript")
            # service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.set_page_load_timeout(30)
            self.driver.implicitly_wait(2)
            
            print("✅ Browser started successfully")
            return True
        except Exception as e:
            print(f"❌ Failed to start browser: {e}")
            return False
    
    def extract_contacts(self, html_content):
        """Extract emails and phones using regex"""
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            
            emails = set()
            phones = set()
            
            # Extract from mailto: links
            for link in soup.find_all('a', href=True):
                href = link['href']
                if 'mailto:' in href:
                    email = href.replace('mailto:', '').split('?')[0].strip()
                    if '@' in email and '.' in email:
                        emails.add(email)
                if 'tel:' in href:
                    phone = href.replace('tel:', '').strip()
                    if len(re.sub(r'\D', '', phone)) >= 7:
                        phones.add(phone)
            
            # Remove scripts/styles for text extraction
            for script in soup(["script", "style", "noscript"]):
                script.decompose()
            
            text = soup.get_text(separator=' ', strip=True)
            
            # Extract emails from text
            for email in self.email_pattern.findall(text):
                emails.add(email)
            
            # Extract phones from text (any + number)
            for phone in self.phone_pattern.findall(text):
                digits = re.sub(r'\D', '', phone)
                if 9 <= len(digits) <= 15:
                    phones.add(phone)
            
            return list(emails), list(phones)
        except Exception as e:
            print(f"❌ Extraction error: {e}")
            return [], []
    
    def detect_url_column(self):
        """Find URL column"""
        for col in self.df.columns:
            non_empty = self.df[col].dropna()
            if len(non_empty) > 0:
                for value in non_empty.head(5):
                    val = str(value).strip().lower()
                    if any(d in val for d in ['.com', '.lt', '.eu', '.org', '.net', 'http', 'www.']):
                        return col
        
        if "Row Labels" in self.df.columns:
            return "Row Labels"
        return None
    
    def clean_dataframe(self):
        """Remove header rows and invalid URLs"""
        if not self.url_column:
            return False
        
        headers = ['Row Labels', 'unique website link', 'Count of date when found']
        for h in headers:
            self.df = self.df[self.df[self.url_column] != h]
        
        def is_website(url):
            if pd.isna(url):
                return False
            u = str(url).strip().lower()
            return any(d in u for d in ['.com', '.lt', '.eu', '.org', '.net', 'http', 'www.'])
        
        self.df = self.df[self.df[self.url_column].apply(is_website)]
        self.df.reset_index(drop=True, inplace=True)
        return True
    
    def load_excel_file(self):
        """Load Excel file"""
        try:
            output_file = self.excel_file.replace('.xlsm', '_output.xlsx').replace('.xlsx', '_output.xlsx')
            file_to_load = output_file if os.path.exists(output_file) else self.excel_file
            
            if not os.path.exists(file_to_load):
                print(f"❌ File not found: {self.excel_file}")
                return False
            
            print(f"📂 Loading: {file_to_load}")
            
            try:
                self.df = pd.read_excel(file_to_load, engine='openpyxl')
            except:
                self.df = pd.read_excel(file_to_load, engine='openpyxl', header=1)
            
            print(f"Found {len(self.df)} rows")
            
            self.url_column = self.detect_url_column()
            if not self.url_column:
                print("❌ Could not find URL column")
                return False
            
            if not self.clean_dataframe():
                return False
            
            if self.email_column not in self.df.columns:
                self.df[self.email_column] = ""
            if self.phone_column not in self.df.columns:
                self.df[self.phone_column] = ""
            
            return True
        except Exception as e:
            print(f"❌ Error loading file: {e}")
            return False
    
    def scroll_page(self):
        """Scroll down to load lazy content"""
        try:
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(0.3)
            self.driver.execute_script("window.scrollTo(0, 0);")
        except:
            pass
    
    def handle_cookies(self):
        """Try to accept cookie dialogs"""
        try:
            wait = WebDriverWait(self.driver, 1)
            selectors = [
                (By.XPATH, "//button[contains(translate(., 'ACCEPT', 'accept'), 'accept')]"),
                (By.XPATH, "//button[contains(translate(., 'AGREE', 'agree'), 'agree')]"),
            ]
            for by, sel in selectors:
                try:
                    btn = wait.until(EC.element_to_be_clickable((by, sel)))
                    btn.click()
                    time.sleep(0.5)
                    return
                except:
                    continue
        except:
            pass
    
    def navigate_to_contact(self):
        """Find contact page by scoring all internal links"""
        contact_keywords = ['contact', 'kontakt', 'contacto', 'contatti', 'impressum', 'kontakte', 'about', 'uber-uns', 'about-us', 'reach', 'connect']
        
        try:
            current_domain = urlparse(self.driver.current_url).netloc.replace('www.', '')
            
            # Get all links
            links = self.driver.find_elements(By.TAG_NAME, 'a')
            
            best_link = None
            best_score = 0
            
            for link in links:
                try:
                    href = link.get_attribute('href') or ''
                    text = link.text.lower().strip()
                    
                    # Skip if no href or external link
                    if not href or href.startswith(('javascript:', 'mailto:', 'tel:', '#')):
                        continue
                    
                    # Check if internal link (same domain)
                    link_domain = urlparse(href).netloc.replace('www.', '')
                    if link_domain and link_domain != current_domain:
                        continue
                    
                    # Score based on keyword matches
                    href_lower = href.lower()
                    score = 0
                    for kw in contact_keywords:
                        if kw in text:
                            score += 10  # Text match is strong signal
                        if kw in href_lower:
                            score += 5   # URL match is good signal
                    
                    if score > best_score:
                        best_score = score
                        best_link = link
                        
                except:
                    continue
            
            if best_link and best_score > 0:
                best_link.click()
                time.sleep(0.5)
                return True
                
        except Exception as e:
            pass
        
        return False
    
    def scrape_website(self, url, row_num):
        """Scrape a single website"""
        normalized = url.strip().lower().replace('www.', '')
        if normalized in self.scraped_urls:
            return "Duplicate", "Duplicate"
        
        try:
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
            
            domain = urlparse(url).netloc.replace('www.', '')
            print(f"Processing row {row_num}: {domain}")
            
            all_emails = set()
            all_phones = set()
            
            # Scrape homepage first
            self.driver.get(url)
            time.sleep(1)
            self.handle_cookies()
            self.scroll_page()
            
            emails, phones = self.extract_contacts(self.driver.page_source)
            all_emails.update(emails)
            all_phones.update(phones)
            
            # Try contact page
            if self.navigate_to_contact():
                self.handle_cookies()
                self.scroll_page()
                emails, phones = self.extract_contacts(self.driver.page_source)
                all_emails.update(emails)
                all_phones.update(phones)
            
            email_result = ', '.join(all_emails) if all_emails else "Not found"
            phone_result = ', '.join(all_phones) if all_phones else "Not found"
            
            print(f"Email: {'✅ ' + email_result if all_emails else '❎ Not found'}")
            print(f"Phone: {'✅ ' + phone_result if all_phones else '❎ Not found'}")
            print()
            
            self.scraped_urls.add(normalized)
            return email_result, phone_result
            
        except Exception as e:
            print(f"❌ Error: {e}\n")
            return "Error", "Error"
    
    def find_empty_rows(self, count):
        """Find rows needing processing"""
        rows = []
        for idx, row in self.df.iterrows():
            if len(rows) >= count:
                break
            
            email_empty = pd.isna(row[self.email_column]) or str(row[self.email_column]).strip() == ''
            phone_empty = pd.isna(row[self.phone_column]) or str(row[self.phone_column]).strip() == ''
            url = row[self.url_column] if pd.notna(row[self.url_column]) else ""
            
            if (email_empty or phone_empty) and url:
                rows.append((idx, str(url).strip()))
        return rows
    
    def update_row(self, idx, email, phone):
        """Update row with results"""
        # Escape + for Excel
        if phone and phone not in ['Not found', 'Error', 'Duplicate']:
            if ',' in phone:
                parts = [("'" + p.strip()) if p.strip().startswith('+') else p.strip() for p in phone.split(',')]
                phone = ', '.join(parts)
            elif phone.startswith('+'):
                phone = "'" + phone
        
        self.df.at[idx, self.email_column] = email
        self.df.at[idx, self.phone_column] = phone
    
    def save_excel(self):
        """Save results"""
        try:
            output = self.excel_file.replace('.xlsm', '_output.xlsx').replace('.xlsx', '_output.xlsx')
            self.df.to_excel(output, index=False, engine='openpyxl')
            return True
        except Exception as e:
            print(f"❌ Save error: {e}")
            return False
    
    def run(self):
        """Main loop"""
        if not self.load_excel_file():
            input("\nPress Enter to exit...")
            return
        
        try:
            num = int(input("\n👉 How many websites to process? "))
            if num <= 0:
                print("❌ Enter a positive number")
                return
        except ValueError:
            print("❌ Enter a valid number")
            return
        except KeyboardInterrupt:
            print("\n👋 Cancelled")
            return
        
        if not self.setup_driver():
            input("\nPress Enter to exit...")
            return
        
        try:
            rows = self.find_empty_rows(num)
            if not rows:
                print("🎉 All rows already have contact info!")
                input("\nPress Enter to exit...")
                return
            
            for i, (idx, url) in enumerate(rows, 1):
                if self.processed_count >= num:
                    break
                
                email, phone = self.scrape_website(url, i)
                self.update_row(idx, email, phone)
                self.processed_count += 1
                self.save_excel()
                time.sleep(1)
            
            output = self.excel_file.replace('.xlsm', '_output.xlsx').replace('.xlsx', '_output.xlsx')
            print(f"\n{'='*50}")
            print(f"✅ Processed: {self.processed_count} websites")
            print(f"💾 Saved to: {output}")
            print('='*50)
            
        except KeyboardInterrupt:
            print("\n⚠️ Interrupted. Saving...")
            self.save_excel()
        finally:
            if self.driver:
                self.driver.quit()
            input("\nPress Enter to close...")


if __name__ == "__main__":
    scraper = CompanyContactScraper()
    scraper.run()
