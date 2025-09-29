import requests
import re
import json
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urljoin, urlparse, quote, urlencode
import logging
import os
from difflib import SequenceMatcher

class SwissCompanyExtractor:
    def __init__(self):
        self.results = []
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
        self.driver = None

        # Setup logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)

        # Swiss UID and CH-ID patterns
        self.uid_patterns = [
            r'CHE[\s\-]?([0-9]{3})[\s\-\.]?([0-9]{3})[\s\-\.]?([0-9]{3})',  # CHE-123.456.789
            r'UID[\s:]*CHE[\s\-]?([0-9]{3})[\s\-\.]?([0-9]{3})[\s\-\.]?([0-9]{3})',  # UID: CHE-123.456.789
        ]

        self.ch_id_patterns = [
            r'CH[\s\-]?([0-9]{3})[\s\-\.]?([0-9]{1})[\s\-\.]?([0-9]{3})[\s\-\.]?([0-9]{3})[\s\-]?([0-9]{1})',  # CH-660.7.436.025-2
            r'CH[\-\s]?ID[\s:]*CH[\s\-]?([0-9]{3})[\s\-\.]?([0-9]{1})[\s\-\.]?([0-9]{3})[\s\-\.]?([0-9]{3})[\s\-]?([0-9]{1})',  # CH-ID: CH-660.7.436.025-2
        ]

    def similarity(self, a, b):
        """Calculate similarity between two strings"""
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    def extract_uid_from_text(self, text):
        """Extract Swiss UID from text"""
        for pattern in self.uid_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                if len(match.groups()) >= 3:
                    # Reconstruct UID format: CHE-123.456.789
                    uid = f"CHE-{match.group(1)}.{match.group(2)}.{match.group(3)}"
                    return uid
        return None

    def extract_ch_id_from_text(self, text):
        """Extract Swiss CH-ID from text"""
        for pattern in self.ch_id_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                if len(match.groups()) >= 5:
                    # Reconstruct CH-ID format: CH-660.7.436.025-2
                    ch_id = f"CH-{match.group(1)}.{match.group(2)}.{match.group(3)}.{match.group(4)}-{match.group(5)}"
                    return ch_id
        return None

    def search_auditorstats_by_name(self, company_name):
        """Search auditorstats.ch for company information"""
        try:
            self.logger.info(f"Searching auditorstats.ch for: {company_name}")

            # Setup Selenium for JavaScript-heavy site
            if not self.driver:
                options = Options()
                options.add_argument('--headless')
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')
                options.add_argument('--disable-gpu')
                options.add_argument('--window-size=1920,1080')
                options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')

                try:
                    self.driver = webdriver.Chrome(options=options)
                except Exception as e:
                    self.logger.error(f"Could not initialize Chrome driver: {str(e)}")
                    return None

            # Navigate to auditorstats.ch search page
            search_url = "https://auditorstats.ch/index.php?pid=4"
            self.driver.get(search_url)

            # Wait for page to load
            time.sleep(3)

            # Find search input field (adjust selector based on actual page structure)
            search_selectors = [
                'input[name="firm"]',
                'input[name="search"]',
                'input[name="company"]',
                'input[type="text"]',
                'input[placeholder*="firm"]',
                'input[placeholder*="company"]',
                '#search',
                '.search-input'
            ]

            search_input = None
            for selector in search_selectors:
                try:
                    search_input = self.driver.find_element(By.CSS_SELECTOR, selector)
                    break
                except:
                    continue

            if not search_input:
                # Try to find any text input
                try:
                    search_input = self.driver.find_element(By.TAG_NAME, "input")
                except:
                    self.logger.error("Could not find search input field")
                    return None

            # Clear and enter company name
            search_input.clear()
            search_input.send_keys(company_name)

            # Find and click search button
            search_button_selectors = [
                'input[type="submit"]',
                'button[type="submit"]',
                'button:contains("Search")',
                '.search-button',
                '#search-button'
            ]

            search_button = None
            for selector in search_button_selectors:
                try:
                    search_button = self.driver.find_element(By.CSS_SELECTOR, selector)
                    break
                except:
                    continue

            if search_button:
                search_button.click()
            else:
                # Try pressing Enter
                from selenium.webdriver.common.keys import Keys
                search_input.send_keys(Keys.RETURN)

            # Wait for results
            time.sleep(5)

            # Get page source and parse results
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')

            # Parse company information from results
            company_info = self.parse_auditorstats_results(soup, company_name)

            return company_info

        except Exception as e:
            self.logger.error(f"Error searching auditorstats.ch for {company_name}: {str(e)}")
            return None

    def parse_auditorstats_results(self, soup, original_company_name):
        """Parse company information from auditorstats.ch results"""
        try:
            # Look for company information in various table/div structures
            companies_found = []

            # Method 1: Look for table rows with company data
            rows = soup.find_all('tr')
            for row in rows:
                cells = row.find_all(['td', 'th'])
                if len(cells) >= 3:
                    row_text = ' '.join([cell.get_text(strip=True) for cell in cells])

                    # Check if this row contains company information
                    if any(suffix in row_text.upper() for suffix in ['SA', 'AG', 'GMBH', 'SARL', 'CHE-']):
                        company_data = self.extract_company_data_from_row(row_text, cells)
                        if company_data:
                            companies_found.append(company_data)

            # Method 2: Look for div/span structures with company data
            divs = soup.find_all(['div', 'span', 'p'])
            for div in divs:
                div_text = div.get_text(strip=True)
                if 'CHE-' in div_text and any(suffix in div_text.upper() for suffix in ['SA', 'AG', 'GMBH', 'SARL']):
                    company_data = self.extract_company_data_from_text(div_text)
                    if company_data:
                        companies_found.append(company_data)

            # Method 3: Extract from entire page text
            page_text = soup.get_text()
            company_data = self.extract_company_data_from_text(page_text)
            if company_data:
                companies_found.append(company_data)

            # Find best match based on similarity
            if companies_found:
                best_match = None
                best_score = 0

                for company in companies_found:
                    if company.get('company_name'):
                        score = self.similarity(company['company_name'], original_company_name)
                        if score > best_score:
                            best_score = score
                            best_match = company

                # Use best match if similarity > 0.4, otherwise use first result
                if best_match and best_score > 0.4:
                    return best_match
                elif companies_found:
                    return companies_found[0]

            return None

        except Exception as e:
            self.logger.error(f"Error parsing auditorstats results: {str(e)}")
            return None

    def extract_company_data_from_row(self, row_text, cells):
        """Extract company data from table row"""
        try:
            company_data = {}

            # Try to identify company name (usually contains SA, AG, GmbH, etc.)
            for cell in cells:
                cell_text = cell.get_text(strip=True)
                if any(suffix in cell_text.upper() for suffix in ['SA', 'AG', 'GMBH', 'SARL', 'LTD', 'LIMITED']):
                    company_data['company_name'] = cell_text
                    break

            # Extract UID and CH-ID from row text
            uid = self.extract_uid_from_text(row_text)
            if uid:
                company_data['uid'] = uid

            ch_id = self.extract_ch_id_from_text(row_text)
            if ch_id:
                company_data['ch_id'] = ch_id

            return company_data if company_data else None

        except Exception as e:
            self.logger.error(f"Error extracting company data from row: {str(e)}")
            return None

    def extract_company_data_from_text(self, text):
        """Extract company data from text block"""
        try:
            company_data = {}

            # Extract UID
            uid = self.extract_uid_from_text(text)
            if uid:
                company_data['uid'] = uid

            # Extract CH-ID
            ch_id = self.extract_ch_id_from_text(text)
            if ch_id:
                company_data['ch_id'] = ch_id

            # Try to extract company name (look for text with Swiss suffixes)
            swiss_suffixes = [
                r'\b([A-Za-z\s&\-\.]+(?:SA|AG|GmbH|Sàrl|SARL|Ltd|Limited))\b',
                r'\b([A-Za-z\s&\-\.]+\s+(?:SA|AG|GmbH|Sàrl|SARL))\b'
            ]

            for pattern in swiss_suffixes:
                matches = re.finditer(pattern, text, re.IGNORECASE)
                for match in matches:
                    candidate_name = match.group(1).strip()
                    if len(candidate_name) > 5 and not company_data.get('company_name'):
                        company_data['company_name'] = candidate_name
                        break

            return company_data if company_data else None

        except Exception as e:
            self.logger.error(f"Error extracting company data from text: {str(e)}")
            return None

    def extract_legal_name_structured_approach(self, html_content, company_name):
        """Extract legal name using Swiss company suffixes"""
        soup = BeautifulSoup(html_content, 'html.parser')
        found_names = []

        # Swiss legal suffixes - comprehensive list
        swiss_suffixes = [
            r'SA', r'AG', r'GmbH', r'Sàrl', r'SARL',
            r'Limited', r'Ltd', r'Ltée',
            r'Aktiengesellschaft', r'Gesellschaft mit beschränkter Haftung',
            r'Société Anonyme', r'Société à responsabilité limitée',
            r'Società Anonima', r'Società a Garanzia Limitata',
            # Variations with punctuation
            r'S\.A\.', r'A\.G\.', r'G\.m\.b\.H\.',
        ]

        suffix_pattern = '|'.join(swiss_suffixes)

        # Extract from various HTML elements
        for tag in soup.find_all(['title', 'h1', 'h2', 'h3', 'div', 'span']):
            text = tag.get_text(strip=True)
            if re.search(rf'\b({suffix_pattern})\b', text, re.IGNORECASE):
                found_names.append(text.strip())

        # Rank and select best match
        if found_names:
            scored_names = []
            for name in found_names:
                score = self.similarity(name, company_name)
                if re.search(rf'\b({suffix_pattern})\b', name, re.IGNORECASE):
                    score += 0.2
                scored_names.append((score, name))

            scored_names.sort(reverse=True)
            return scored_names[0][1] if scored_names else company_name

        return company_name

    def process_company_with_auditorstats(self, company_name, company_url=None):
        """Process a single company using auditorstats.ch search"""
        result = {
            'portfolio_company': company_name,
            'company_url': company_url or '',
            'legal_name': '',
            'uid': '',
            'ch_id': '',
            'source': 'auditorstats.ch',
            'search_method': 'auditorstats_web',
            'status': '',
            'auditorstats_searched': 'Yes'
        }

        try:
            self.logger.info(f"Processing Swiss company: {company_name}")

            # Search auditorstats.ch for company information
            company_info = self.search_auditorstats_by_name(company_name)

            if company_info:
                self.logger.info(f"Found company info: {company_info}")

                result.update({
                    'legal_name': company_info.get('company_name', company_name),
                    'uid': company_info.get('uid', ''),
                    'ch_id': company_info.get('ch_id', ''),
                    'status': 'Found'
                })
            else:
                self.logger.warning(f"Could not find company info for {company_name}")
                result['legal_name'] = company_name
                result['status'] = 'Not Found'

            # If we have the company URL, try to extract additional info from their website
            if company_url:
                try:
                    response = self.session.get(company_url, timeout=10)
                    if response.status_code == 200:
                        # Try to extract UID and CH-ID from company website
                        website_uid = self.extract_uid_from_text(response.text)
                        if website_uid and not result['uid']:
                            result['uid'] = website_uid
                            result['search_method'] = 'website_extraction'

                        website_ch_id = self.extract_ch_id_from_text(response.text)
                        if website_ch_id and not result['ch_id']:
                            result['ch_id'] = website_ch_id

                        # Extract legal name from website using Swiss suffixes
                        legal_name = self.extract_legal_name_structured_approach(response.text, company_name)
                        if legal_name and legal_name != company_name and not result['legal_name']:
                            result['legal_name'] = legal_name

                except Exception as e:
                    self.logger.error(f"Error processing company website {company_url}: {str(e)}")

        except Exception as e:
            self.logger.error(f"Error processing company {company_name}: {str(e)}")
            result['legal_name'] = company_name
            result['status'] = 'Error'

        return result

    def process_portfolio_companies(self, companies_data):
        """Process a list of portfolio companies"""
        self.results = []

        for company_data in companies_data:
            if isinstance(company_data, dict):
                company_name = company_data.get('name', '')
                company_url = company_data.get('url', '')
            else:
                company_name = str(company_data)
                company_url = ''

            if company_name:
                result = self.process_company_with_auditorstats(company_name, company_url)
                self.results.append(result)

                # Add delay to be respectful to auditorstats.ch
                time.sleep(2)

        return self.results

    def save_results_to_csv(self, filename='swiss_company_extraction_results.csv'):
        """Save results to CSV file"""
        if self.results:
            df = pd.DataFrame(self.results)
            df.to_csv(filename, index=False)
            self.logger.info(f"Results saved to {filename}")
            return filename
        else:
            self.logger.warning("No results to save")
            return None

    def get_results_summary(self):
        """Get a summary of extraction results"""
        if not self.results:
            return "No results available"

        total = len(self.results)
        with_uid = len([r for r in self.results if r['uid']])
        with_ch_id = len([r for r in self.results if r['ch_id']])
        with_legal_name = len([r for r in self.results if r['legal_name'] and r['legal_name'] != r['portfolio_company']])
        found_status = len([r for r in self.results if r['status'] == 'Found'])

        return f"""
        Total companies processed: {total}
        Companies found in auditorstats.ch: {found_status} ({found_status/total*100:.1f}%)
        Companies with UID extracted: {with_uid} ({with_uid/total*100:.1f}%)
        Companies with CH-ID extracted: {with_ch_id} ({with_ch_id/total*100:.1f}%)
        Companies with legal name extracted: {with_legal_name} ({with_legal_name/total*100:.1f}%)
        """

    def cleanup(self):
        """Clean up resources"""
        if self.driver:
            self.driver.quit()
            self.driver = None

# Example usage
if __name__ == "__main__":
    # Example Swiss portfolio companies data
    example_companies = pd.read_excel('swiss_companies.xlsx').to_dict(orient='records')

    extractor = SwissCompanyExtractor()

    try:
        results = extractor.process_portfolio_companies(example_companies)

        # Print results
        for result in results:
            print(f"Company: {result['portfolio_company']}")
            print(f"Legal Name: {result['legal_name']}")
            print(f"UID: {result['uid']}")
            print(f"CH-ID: {result['ch_id']}")
            print(f"Status: {result['status']}")
            print("-" * 50)

        # Save to CSV
        extractor.save_results_to_csv()

        # Print summary
        print(extractor.get_results_summary())

    finally:
        # Always cleanup
        extractor.cleanup()
