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
from urllib.parse import urljoin, urlparse, quote
import logging
import os
from difflib import SequenceMatcher

class GermanyTaxExtractor:
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
        
        # Germany-specific tax number patterns
        self.tax_patterns = [
            r'Steuernummer[\s#:]*([0-9]{2,3}\/[0-9]{3,4}\/[0-9]{4,5})',  # Standard Steuernummer format
            r'Steuer-Nr\.?[\s#:]*([0-9]{2,3}\/[0-9]{3,4}\/[0-9]{4,5})',  # Steuer-Nr. format
            r'St\.?\s*Nr\.?[\s#:]*([0-9]{2,3}\/[0-9]{3,4}\/[0-9]{4,5})',  # St. Nr. format
            r'Tax\s*ID[\s#:]*([0-9]{2,3}\/[0-9]{3,4}\/[0-9]{4,5})',  # English Tax ID
            r'Umsatzsteuer-ID[\s#:]*([A-Z]{2}[0-9]{9})',  # VAT ID (DE + 9 digits)
            r'USt-IdNr\.?[\s#:]*([A-Z]{2}[0-9]{9})',  # USt-IdNr format
            r'Handelsregister[\s#:]*([A-Z]{2,3}\s*[0-9]+)',  # Commercial register (HRB, HRA, etc.)
            r'HRB[\s#:]*([0-9]+)',  # Handelsregister B
            r'HRA[\s#:]*([0-9]+)',  # Handelsregister A
        ]

    def search_unternehmensregister(self, company_name):
        self.logger.info("Unternehmensregister search not implemented due to technical constraints (CAPTCHA).")
        return None, None


    def similarity(self, a, b):
        """Calculate similarity between two strings"""
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    def extract_tax_number_from_html(self, html_content):
        """Extract German tax number from HTML content"""
        found_numbers = []
        
        for pattern in self.tax_patterns:
            matches = re.finditer(pattern, html_content, re.IGNORECASE)
            for match in matches:
                code = match.group(1) if len(match.groups()) > 0 else match.group(0)
                clean_code = code.strip()
                
                # Validate different types of numbers
                if re.match(r'^[0-9]{2,3}\/[0-9]{3,4}\/[0-9]{4,5}$', clean_code):
                    # Valid Steuernummer format
                    found_numbers.append(('Steuernummer', clean_code))
                elif re.match(r'^[A-Z]{2}[0-9]{9}$', clean_code):
                    # Valid VAT ID format
                    found_numbers.append(('USt-IdNr', clean_code))
                elif re.match(r'^[A-Z]{2,3}\s*[0-9]+$', clean_code):
                    # Valid Handelsregister format
                    found_numbers.append(('Handelsregister', clean_code))
                elif re.match(r'^[0-9]+$', clean_code):
                    # Pure number (HRB/HRA)
                    found_numbers.append(('Handelsregister', clean_code))
        
        return found_numbers

    def extract_legal_name_structured_approach(self, html_content, company_name):
        """Extract legal name using the specified structured approach for German companies"""
        soup = BeautifulSoup(html_content, 'html.parser')
        found_names = []
        
        # German legal suffixes
        german_suffixes = [
            r'GmbH', r'AG', r'KG', r'OHG', r'GbR', r'UG', r'SE', r'KGaA', r'eG',
            r'G\.m\.b\.H\.?', r'A\.G\.?', r'K\.G\.?', r'O\.H\.G\.?', r'U\.G\.?',
            r'Gesellschaft\s+mit\s+beschränkter\s+Haftung',
            r'Aktiengesellschaft', r'Kommanditgesellschaft', r'Offene\s+Handelsgesellschaft'
        ]
        suffix_pattern = '|'.join(german_suffixes)
        
        # 1. Extract from <title> tag
        title_tag = soup.find('title')
        if title_tag:
            title_text = title_tag.get_text().strip()
            # Look for legal suffixes in title
            title_matches = re.findall(rf'([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß\s&.\'-]+\s+(?:{suffix_pattern}))', title_text, re.I)
            for match in title_matches:
                clean_name = re.sub(r'\s+', ' ', match.strip())
                if 10 <= len(clean_name) <= 100 and self.similarity(clean_name, company_name) > 0.3:
                    found_names.append((clean_name, 'title'))
        
        # 2. Extract from <meta name="og:site_name">
        og_site_name = soup.find('meta', attrs={'name': 'og:site_name'}) or soup.find('meta', attrs={'property': 'og:site_name'})
        if og_site_name and og_site_name.get('content'):
            site_name = og_site_name.get('content').strip()
            # Check if it contains legal suffix
            if re.search(rf'(?:{suffix_pattern})', site_name, re.I):
                if 10 <= len(site_name) <= 100 and self.similarity(site_name, company_name) > 0.3:
                    found_names.append((site_name, 'og:site_name'))
        
        # 3. Parse JSON-LD structured data
        json_ld_scripts = soup.find_all('script', type='application/ld+json')
        for script in json_ld_scripts:
            try:
                data = json.loads(script.string)
                # Handle both single objects and arrays
                if isinstance(data, list):
                    objects = data
                else:
                    objects = [data]
                
                for obj in objects:
                    if isinstance(obj, dict):
                        obj_type = obj.get('@type', '').lower()
                        if obj_type in ['organization', 'corporation', 'localbusiness']:
                            name = obj.get('name', '').strip()
                            legal_name = obj.get('legalName', '').strip()
                            
                            # Prefer legalName over name
                            candidate_name = legal_name if legal_name else name
                            if candidate_name:
                                # Check if it contains legal suffix
                                if re.search(rf'(?:{suffix_pattern})', candidate_name, re.I):
                                    if 10 <= len(candidate_name) <= 100 and self.similarity(candidate_name, company_name) > 0.3:
                                        found_names.append((candidate_name, 'json-ld'))
            except json.JSONDecodeError:
                continue
        
        # 4. Footer and copyright extraction
        footers = soup.find_all(['footer', 'div'], class_=re.compile(r'footer|impressum', re.I))
        for footer in footers:
            footer_text = footer.get_text()
            # Look for copyright patterns
            copyright_patterns = [
                rf'(?:©|Copyright|Alle\s+Rechte\s+vorbehalten)\s*(?:20[0-9]{{2}})?\s*([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß\s&.\'-]+(?:{suffix_pattern}))',
                rf'©\s*([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß\s&.\'-]+(?:{suffix_pattern}))'
            ]
            
            for pattern in copyright_patterns:
                matches = re.finditer(pattern, footer_text, re.IGNORECASE)
                for match in matches:
                    name = match.group(1).strip()
                    clean_name = re.sub(r'\s+', ' ', name)
                    if 10 <= len(clean_name) <= 100 and self.similarity(clean_name, company_name) > 0.3:
                        found_names.append((clean_name, 'copyright'))
        
        # 5. General legal suffix search in HTML
        legal_name_matches = re.findall(rf'([A-ZÄÖÜa-zäöüß][A-Za-zÄÖÜäöüß\s&.\'-]+\s+(?:{suffix_pattern}))', html_content, re.I)
        for match in legal_name_matches:
            clean_name = re.sub(r'\s+', ' ', match.strip())
            clean_name = re.sub(r'^[^\w]+|[^\w]+$', '', clean_name)
            if (10 <= len(clean_name) <= 150 and 
                not re.search(r'cookie|privacy|terms|contact|home|menu|impressum', clean_name, re.I) and
                len(clean_name.split()) >= 2 and
                self.similarity(clean_name, company_name) > 0.3):
                found_names.append((clean_name, 'general'))
        
        # Sort by similarity to company name and return best matches
        unique_names = {}
        for name, source in found_names:
            if name not in unique_names:
                unique_names[name] = source
        
        # Sort by similarity score
        scored_names = []
        for name, source in unique_names.items():
            score = self.similarity(name, company_name)
            scored_names.append((name, source, score))
        
        scored_names.sort(key=lambda x: x[2], reverse=True)
        return [(name, source) for name, source, score in scored_names]
    

        self.logger.info(f"searching TradeRegistry.de for: {company_name}")
        try:
            # Handelsregister.de search URL
            search_url = "https://www.traderegistry.de/rp_web/search/company.do"
            
            params = {
                'action': 'search',
                'companyName': company_name,
                'searchType': 'company'
            }
            
            response = self.session.get(search_url, params=params, timeout=15)
            
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # Look for company results in the response
                company_links = soup.find_all('a', href=re.compile(r'company\.do\?action=show'))
                
                if company_links:
                    best_match = None
                    best_score = 0
                    
                    for link in company_links[:5]:  # Check first 5 results
                        company_text = link.get_text().strip()
                        score = self.similarity(company_text, company_name)
                        
                        self.logger.info(f"  Handelsregister candidate: {company_text} (score: {score:.2f})")
                        
                        if score > best_score and score > 0.6:
                            best_match = {
                                'name': company_text,
                                'url': link.get('href')
                            }
                            best_score = score
                    
                    if best_match:
                        # Get detailed company info
                        detail_url = urljoin(search_url, best_match['url'])
                        detail_response = self.session.get(detail_url, timeout=10)
                        
                        if detail_response.status_code == 200:
                            detail_soup = BeautifulSoup(detail_response.text, 'html.parser')
                            
                            # Extract company details
                            legal_name = best_match['name']
                            
                            # Look for registration numbers in the details
                            reg_numbers = self.extract_tax_number_from_html(detail_response.text)
                            
                            self.logger.info(f"Found match in Handelsregister: {legal_name}")
                            return reg_numbers, legal_name
            
            time.sleep(1)
            
        except Exception as e:
            self.logger.debug(f"Handelsregister.de search error: {e}")
        
        return None, No
    def search_traderegistry(self, company_name):
        self.logger.info(f"Searching TradeRegistry.de for {company_name}")
        try:
            search_url = "https://traderegistry.de/company-search/"
            params = {"search": company_name}
            response = self.session.get(search_url, params=params, timeout=15)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")
                company_links = soup.find_all("a", href=True)
                best_match = None
                best_score = 0
                for link in company_links[:5]:  # Prendi i primi 5 risultati
                    text = link.text.strip()
                    score = self.similarity(text, company_name)
                    if score > best_score and score > 0.6:
                        best_match = {"name": text, "url": link["href"]}
                        best_score = score
                if best_match:
                    detail_url = urljoin(search_url, best_match["url"])
                    detail_response = self.session.get(detail_url, timeout=10)
                    if detail_response.status_code == 200:
                        legal_name = best_match["name"]
                        tax_numbers = self.extracttaxnumberfromhtml(detail_response.text)
                        self.logger.info(f"Found match in TradeRegistry {legal_name}")
                        return tax_numbers, legal_name
            return None, None
        except Exception as e:
            self.logger.error(f"TradeRegistry search error: {e}")
            return None, None


    def search_handelsregister(self, company_name):
        """Search for company on Handelsregister.de"""
        self.logger.info(f"Searching Handelsregister.de for: {company_name}")
        
        try:
            # Handelsregister.de search URL
            search_url = "https://www.handelsregister.de/rp_web/search/company.do"
            
            params = {
                'action': 'search',
                'companyName': company_name,
                'searchType': 'company'
            }
            
            response = self.session.get(search_url, params=params, timeout=15)
            
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # Look for company results in the response
                company_links = soup.find_all('a', href=re.compile(r'company\.do\?action=show'))
                
                if company_links:
                    best_match = None
                    best_score = 0
                    
                    for link in company_links[:5]:  # Check first 5 results
                        company_text = link.get_text().strip()
                        score = self.similarity(company_text, company_name)
                        
                        self.logger.info(f"  Handelsregister candidate: {company_text} (score: {score:.2f})")
                        
                        if score > best_score and score > 0.6:
                            best_match = {
                                'name': company_text,
                                'url': link.get('href')
                            }
                            best_score = score
                    
                    if best_match:
                        # Get detailed company info
                        detail_url = urljoin(search_url, best_match['url'])
                        detail_response = self.session.get(detail_url, timeout=10)
                        
                        if detail_response.status_code == 200:
                            detail_soup = BeautifulSoup(detail_response.text, 'html.parser')
                            
                            # Extract company details
                            legal_name = best_match['name']
                            
                            # Look for registration numbers in the details
                            reg_numbers = self.extract_tax_number_from_html(detail_response.text)
                            
                            self.logger.info(f"Found match in Handelsregister: {legal_name}")
                            return reg_numbers, legal_name
            
            time.sleep(1)
            
        except Exception as e:
            self.logger.debug(f"Handelsregister.de search error: {e}")
        
        return None, None

    def search_bundesanzeiger(self, company_name):
        """Search for company on Bundesanzeiger.de"""
        self.logger.info(f"Searching Bundesanzeiger.de for: {company_name}")
        
        try:
            # Bundesanzeiger.de search URL
            search_url = "https://www.bundesanzeiger.de/pub/de/suchergebnis"
            
            params = {
                'search.searchtext': company_name,
                'search.searchtype': 'company'
            }
            
            response = self.session.get(search_url, params=params, timeout=15)
            
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # Look for company results
                result_items = soup.find_all('div', class_=['result-item', 'search-result'])
                
                if result_items:
                    best_match = None
                    best_score = 0
                    
                    for item in result_items[:5]:  # Check first 5 results
                        company_name_elem = item.find(['h2', 'h3', 'a'], class_=re.compile(r'company|title'))
                        
                        if company_name_elem:
                            company_text = company_name_elem.get_text().strip()
                            score = self.similarity(company_text, company_name)
                            
                            self.logger.info(f"  Bundesanzeiger candidate: {company_text} (score: {score:.2f})")
                            
                            if score > best_score and score > 0.6:
                                best_match = {
                                    'name': company_text,
                                    'element': item
                                }
                                best_score = score
                    
                    if best_match:
                        # Extract additional info from the result item
                        item_html = str(best_match['element'])
                        reg_numbers = self.extract_tax_number_from_html(item_html)
                        
                        legal_name = best_match['name']
                        
                        self.logger.info(f"Found match in Bundesanzeiger: {legal_name}")
                        return reg_numbers, legal_name
            
            time.sleep(1)
            
        except Exception as e:
            self.logger.debug(f"Bundesanzeiger.de search error: {e}")
        
        return None, None

    def setup_selenium(self):
        """Setup Selenium WebDriver"""
        if self.driver is None:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
            
            try:
                self.driver = webdriver.Chrome(options=chrome_options)
                self.driver.set_page_load_timeout(15)
            except Exception as e:
                self.logger.error(f"Failed to setup Chrome driver: {e}")
                self.driver = None

    def fetch_with_requests(self, url, timeout=10):
        """Fetch page content with requests"""
        try:
            response = self.session.get(url, timeout=timeout)
            response.raise_for_status()
            return response.text
        except Exception as e:
            self.logger.debug(f"Requests failed for {url}: {e}")
            return None

    def fetch_with_selenium(self, url):
        """Fallback to Selenium for dynamic content"""
        if self.driver is None:
            self.setup_selenium()
        
        if self.driver is None:
            return None

        try:
            self.driver.get(url)
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            time.sleep(2)
            return self.driver.page_source
        except Exception as e:
            self.logger.debug(f"Selenium failed for {url}: {e}")
            return None

    def get_german_paths(self):
        """Return common paths for German websites"""
        return [
            '',  # Main page (homepage)
            '/kontakt', '/contact', '/kontaktieren',
            '/ueber-uns', '/uber-uns', '/about', '/about-us',
            '/impressum', '/legal', '/rechtliches', '/datenschutz'
        ]

    def search_website_for_info(self, base_url, company_name):
        """Search website for German tax numbers and legal name"""
        found_legal_names = []
        found_tax_numbers = []
        
        german_paths = self.get_german_paths()

        for path in german_paths:
            url = base_url.rstrip('/') + path
            self.logger.info(f"Checking: {url}")

            try:
                # Try requests first
                html_content = self.fetch_with_requests(url)
                
                # Fallback to Selenium
                if not html_content:
                    html_content = self.fetch_with_selenium(url)

                if html_content:
                    # Look for German tax numbers
                    tax_numbers = self.extract_tax_number_from_html(html_content)
                    if tax_numbers:
                        found_tax_numbers.extend(tax_numbers)
                    
                    # Extract legal names using structured approach
                    legal_names_with_sources = self.extract_legal_name_structured_approach(html_content, company_name)
                    found_legal_names.extend(legal_names_with_sources)
                    
                    # If we found tax numbers and legal names, we can break
                    if found_tax_numbers and found_legal_names:
                        break

            except Exception as e:
                self.logger.debug(f"Error processing {url}: {e}")

            time.sleep(0.5)

        return found_legal_names, found_tax_numbers

    def process_single_german_company(self, company):
        """Process a single German company to extract tax number"""
        self.logger.info(f"\n--- Processing German company: {company['name']} ---")
        
        website = company.get('website', '').strip()
        
        if website and not website.startswith('http'):
            website = 'https://' + website

        result = {
            'original_company_name': company['name'],
            'pe_name': company.get('pe_name', ''),
            'website': website,
            'pe_website': company.get('pe_website', ''),
            'target_geography': company.get('target_geography', ''),
            'target_industry': company.get('target_industry', ''),
            'target_sub_industry': company.get('target_sub_industry', ''),
            'entry_year': company.get('entry_year', ''),
            'legal_name_website': '',
            'legal_name_source': '',
            'legal_name_register': '',
            'steuernummer': '',
            'ust_idnr': '',
            'handelsregister': '',
            'status': 'not_found',
            'search_location': '',
            'error': None
        }

        try:
            # Step 1: Search Unternehmensregister first (currently not implemented due to technical constraints)
            self.logger.info("Searching Unternehmensregister...")
            register_tax_info, register_legal_name = self.search_unternehmensregister(company['name'])
            
            if register_tax_info and register_legal_name:
                result['legal_name_register'] = register_legal_name
                result['status'] = 'found'
                result['search_location'] = 'unternehmensregister'
                self.logger.info(f"Found via Unternehmensregister: Legal name: {register_legal_name}")
                return result
            
            # Step 2: If not found in register or no website, try website scraping
            if not website:
                result['status'] = 'no_website'
                result['error'] = 'No website provided and not found in Unternehmensregister'
                return result
            
            self.logger.info("Not found in Unternehmensregister. Searching website...")
            legal_names_with_sources, website_tax_numbers = self.search_website_for_info(website, company['name'])
            
            if legal_names_with_sources:
                result['legal_name_website'] = legal_names_with_sources[0][0]
                result['legal_name_source'] = legal_names_with_sources[0][1]
                self.logger.info(f"Found legal name on website: {legal_names_with_sources[0][0]} (source: {legal_names_with_sources[0][1]})")
            
            if website_tax_numbers:
                # Organize tax numbers by type
                for tax_type, tax_number in website_tax_numbers:
                    if tax_type == 'Steuernummer':
                        result['steuernummer'] = tax_number
                    elif tax_type == 'USt-IdNr':
                        result['ust_idnr'] = tax_number
                    elif tax_type == 'Handelsregister':
                        result['handelsregister'] = tax_number
                
                result['status'] = 'found'
                result['search_location'] = 'website'
                self.logger.info(f"Found tax information on website: {website_tax_numbers}")
            else:
                # If we found legal name on website but no tax numbers, try searching register with the legal name
                if legal_names_with_sources:
                    self.logger.info(f"Trying Unternehmensregister search with website legal name: {legal_names_with_sources[0][0]}")
                    legal_name_tax_info, _ = self.search_unternehmensregister(legal_names_with_sources[0][0])
                    if legal_name_tax_info:
                        result['status'] = 'found'
                        result['search_location'] = 'register_via_website_name'
                        self.logger.info(f"Found tax info via Unternehmensregister using website legal name")

        except Exception as error:
            result['error'] = str(error)
            result['status'] = 'error'
            self.logger.error(f"Error: {error}")

        return result

    def process_german_companies(self, companies):
        """Process multiple German companies"""
        self.logger.info(f"Starting German tax number extraction for {len(companies)} companies...")
        
        for i, company in enumerate(companies, 1):
            self.logger.info(f"\n[{i}/{len(companies)}]")
            result = self.process_single_german_company(company)
            self.results.append(result)
            
            # Respectful delay between companies
            time.sleep(2)

        return self.results

    def load_companies_from_excel(self, file_path):
        """Load German companies from Excel file"""
        try:
            df = pd.read_excel(file_path)
            
            column_mappings = {
                'pe_name': ['PE NAME'],
                'pe_country': ['Country (HQ)'],
                'pe_website': ['Website'],
                'company_name': ['Portfolio Companies'],
                'target_website': ['Target Website'],
                'target_geography': ['Target Geography'],
                'target_industry': ['Target Industry'],
                'target_sub_industry': ['Target Sub-Industry'],
                'entry_year': ['Entry']
            }
            
            found_columns = {}
            for key, possible_names in column_mappings.items():
                for col in df.columns:
                    if col in possible_names:
                        found_columns[key] = col
                        break
            
            required_columns = ['company_name']
            missing_required = [col for col in required_columns if col not in found_columns]
            if missing_required:
                raise ValueError(f"Missing required columns: {[column_mappings[col][0] for col in missing_required]}")
            
            companies = []
            for _, row in df.iterrows():
                company_name = row[found_columns['company_name']]
                
                if pd.isna(company_name) or str(company_name).strip() == '':
                    continue
                
                target_website = ''
                if 'target_website' in found_columns and not pd.isna(row[found_columns['target_website']]):
                    target_website = str(row[found_columns['target_website']]).strip()
                
                pe_website = ''  
                if 'pe_website' in found_columns and not pd.isna(row[found_columns['pe_website']]):
                    pe_website = str(row[found_columns['pe_website']]).strip()
                    
                company_data = {
                    'name': str(company_name).strip(),
                    'website': target_website,
                    'pe_name': str(row[found_columns['pe_name']]).strip() if 'pe_name' in found_columns and not pd.isna(row[found_columns['pe_name']]) else '',
                    'pe_website': pe_website,
                    'target_geography': str(row[found_columns['target_geography']]).strip() if 'target_geography' in found_columns and not pd.isna(row[found_columns['target_geography']]) else '',
                    'target_industry': str(row[found_columns['target_industry']]).strip() if 'target_industry' in found_columns and not pd.isna(row[found_columns['target_industry']]) else '',
                    'target_sub_industry': str(row[found_columns['target_sub_industry']]).strip() if 'target_sub_industry' in found_columns and not pd.isna(row[found_columns['target_sub_industry']]) else '',
                    'entry_year': str(row[found_columns['entry_year']]).strip() if 'entry_year' in found_columns and not pd.isna(row[found_columns['entry_year']]) else ''
                }
                companies.append(company_data)
            
            self.logger.info(f"Loaded {len(companies)} German companies from {file_path}")
            return companies
            
        except Exception as e:
            self.logger.error(f"Error loading Excel file: {e}")
            raise

    def save_results_to_excel(self, filename='germany_tax_results.xlsx'):
        """Save results to Excel"""
        try:
            data = []
            for result in self.results:
                data.append({
                    'Original Company Name': result['original_company_name'],
                    'PE Name': result.get('pe_name', ''),
                    'Target Website': result['website'],
                    'PE Website': result.get('pe_website', ''),
                    'Target Geography': result.get('target_geography', ''),
                    'Target Industry': result.get('target_industry', ''),
                    'Target Sub-Industry': result.get('target_sub_industry', ''),
                    'Entry Year': result.get('entry_year', ''),
                    'Legal Name (Website)': result['legal_name_website'],
                    'Legal Name Source': result['legal_name_source'],
                    'Legal Name (Register)': result['legal_name_register'],
                    'Steuernummer': result['steuernummer'],
                    'USt-IdNr': result['ust_idnr'],
                    'Handelsregister': result['handelsregister'],
                    'Status': result['status'],
                    'Found At': result['search_location'],
                    'Error': result['error'] or ''
                })
            
            df = pd.DataFrame(data)
            df.to_excel(filename, index=False, engine='openpyxl')
            self.logger.info(f"\nResults saved to {filename}")
            
        except Exception as e:
            self.logger.error(f"Error saving to Excel: {e}")
            # Fallback to JSON
            self.save_results_to_json()

    def save_results_to_json(self, filename='germany_tax_results.json'):
        """Save results to JSON file"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)
        self.logger.info(f"Results saved to {filename}")

    def generate_summary_report(self):
        """Generate comprehensive summary report"""
        total = len(self.results)
        found = len([r for r in self.results if r['status'] == 'found'])
        partial = len([r for r in self.results if r['status'] == 'partial'])
        not_found = len([r for r in self.results if r['status'] == 'not_found'])
        errors = len([r for r in self.results if r['status'] == 'error'])
        no_website = len([r for r in self.results if r['status'] == 'no_website'])
        
        website_found = len([r for r in self.results if r['status'] == 'found' and r['search_location'] == 'website'])
        handelsregister_found = len([r for r in self.results if r['status'] == 'found' and 'handelsregister' in r['search_location']])
        bundesanzeiger_found = len([r for r in self.results if r['status'] == 'found' and 'bundesanzeiger' in r['search_location']])
        
        print('\n' + '='*60)
        print('GERMANY TAX NUMBER EXTRACTION SUMMARY')
        print('='*60)
        print(f'Total companies processed: {total}')
        print(f'Tax information successfully found: {found} ({(found/total*100):.1f}%)')
        print(f'  - Found on company websites: {website_found}')
        print(f'  - Found via Handelsregister.de: {handelsregister_found}')
        print(f'  - Found via Bundesanzeiger.de: {bundesanzeiger_found}')
        print(f'Partial information found: {partial}')
        print(f'Not found: {not_found}')
        print(f'Errors: {errors}')
        print(f'No website: {no_website}')

        # Additional statistics
        steuernummer_found = len([r for r in self.results if r['steuernummer']])
        ust_idnr_found = len([r for r in self.results if r['ust_idnr']])
        handelsregister_found_count = len([r for r in self.results if r['handelsregister']])
        
        print(f'\nTax number types found:')
        print(f'  - Steuernummer: {steuernummer_found}')
        print(f'  - USt-IdNr: {ust_idnr_found}')
        print(f'  - Handelsregister: {handelsregister_found_count}')

    def close(self):
        """Clean up resources"""
        if self.driver:
            self.driver.quit()
        self.session.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


def main():
    excel_file = 'smalldbmachinery_de.xlsx'
    
    import sys
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found!")
        print("\nRequired Excel columns:")
        print("- Portfolio Companies (company name)")
        print("- Target Website (company website) - optional")
        print(f"\nUsage: python {sys.argv[0]} your_file.xlsx")
        return

    with GermanyTaxExtractor() as extractor:
        try:
            print("Loading German companies from Excel...")
            companies = extractor.load_companies_from_excel(excel_file)
            
            if not companies:
                print("No valid companies found in the Excel file!")
                return
            
            print(f"\nProcessing {len(companies)} German companies...")
            print("This extractor will:")
            print("1. First search company websites using rigorous German-specific patterns:")
            print("   - German paths: /impressum, /kontakt, /ueber-uns")
            print("   - German legal suffixes: GmbH, AG, UG, KG, etc.")
            print("   - German tax identifiers: Steuernummer, USt-IdNr, Handelsregister")
            print("2. If website search is incomplete or low confidence, search external sources:")
            print("   - Handelsregister.de (primary external source)")
            print("   - Bundesanzeiger.de (fallback external source)")
            print("3. Confidence-based approach: only use external sources when needed")
            print("=" * 60)
            
            results = extractor.process_german_companies(companies)
            
            # Save results
            extractor.save_results_to_excel()
            extractor.save_results_to_json()
            
            # Generate summary
            extractor.generate_summary_report()
            
            # Show detailed results for successful extractions
            print('\n' + '='*60)
            print('DETAILED RESULTS - SUCCESSFUL EXTRACTIONS')
            print('='*60)
            
            successful_results = [r for r in results if r['status'] == 'found']
            for result in successful_results:
                print(f"\n{result['original_company_name']}:")
                if result['steuernummer']:
                    print(f"  Steuernummer: {result['steuernummer']}")
                if result['ust_idnr']:
                    print(f"  USt-IdNr: {result['ust_idnr']}")
                if result['handelsregister']:
                    print(f"  Handelsregister: {result['handelsregister']}")
                print(f"  Found at: {result['search_location']}")
                if result['legal_name_register']:
                    print(f"  Legal name (Register): {result['legal_name_register']}")
                if result['legal_name_website']:
                    print(f"  Legal name (Website): {result['legal_name_website']} (source: {result['legal_name_source']})")
                if result['website']:
                    print(f"  Website: {result['website']}")
            
            # Show failures for debugging
            failed_results = [r for r in results if r['status'] not in ['found']]
            if failed_results:
                print(f'\n{"="*60}')
                print(f'INCOMPLETE EXTRACTIONS ({len(failed_results)} companies)')
                print('='*60)
                for result in failed_results:
                    print(f"\n{result['original_company_name']}: {result['status']}")
                    if result['error']:
                        print(f"  Error: {result['error']}")
                    if result['legal_name_website']:
                        print(f"  Found legal name on website: {result['legal_name_website']}")
                        if result['status'] == 'partial':
                            print(f"  (but no complete tax information found)")
                    if result['legal_name_register']:
                        print(f"  Found legal name in register: {result['legal_name_register']}")
                    
        except Exception as error:
            print(f'Fatal error: {error}')
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()