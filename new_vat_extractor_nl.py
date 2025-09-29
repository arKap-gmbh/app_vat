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
from selenium.webdriver.common.keys import Keys
from urllib.parse import urljoin, urlparse, quote
import logging
import os
from difflib import SequenceMatcher
import unidecode

class DutchKvKExtractor:
    def __init__(self):
        self.results = []
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        })
        self.driver = None
        
        # Setup logging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        self.logger = logging.getLogger(__name__)
        
        # Dutch-specific patterns for KvK, RSIN, LEI, BTW
        self.kvk_patterns = [
            r'KvK[\s#:]*([0-9]{8})',
            r'(?:Kamer\s*van\s*Koophandel|K\.v\.K\.?)[\s#:]*([0-9]{8})',
            r'KvK\s*:?\s*([0-9]{2}[\s\-]?[0-9]{2}[\s\-]?[0-9]{2}[\s\-]?[0-9]{2})',
            r'(?:Chamber\s*of\s*Commerce|CoC)[\s#:]*([0-9]{8})',
        ]
        
        self.rsin_patterns = [
            r'RSIN[\s#:]*([0-9]{9})',
            r'(?:Rechtspersonen\s*Samenwerkingsverbanden\s*Informatie\s*Nummer)[\s#:]*([0-9]{9})',
        ]
        
        self.lei_patterns = [
            r'LEI[\s#:]*([A-Z0-9]{20})',
            r'(?:Legal\s*Entity\s*Identifier)[\s#:]*([A-Z0-9]{20})',
        ]
        
        self.btw_patterns = [
            r'BTW[\s#:]*(?:NL)?([0-9]{9})B[0-9]{2}',
            r'VAT[\s#:]*(?:NL)?([0-9]{9})B[0-9]{2}',
            r'(?:BTW[-\s]*(?:nummer|number)|VAT[-\s]*(?:nummer|number))[\s#:]*(?:NL)?([0-9]{9})B[0-9]{2}',
        ]

    def validate_kvk_number(self, kvk_number):
        """Validate Dutch KvK number (8 digits)"""
        if not kvk_number or len(kvk_number) != 8 or not kvk_number.isdigit():
            return False
        return not kvk_number.startswith('0')

    def validate_rsin_number(self, rsin_number):
        """Validate Dutch RSIN number (9 digits with checksum)"""
        if not rsin_number or len(rsin_number) != 9 or not rsin_number.isdigit():
            return False
        
        digits = [int(d) for d in rsin_number]
        checksum = sum(digits[i] * (9 - i) for i in range(8)) % 11
        
        if checksum < 2:
            return digits[8] == checksum
        else:
            return digits[8] == 11 - checksum

    def validate_btw_number(self, btw_number):
        """Validate Dutch BTW number"""
        if not btw_number or len(btw_number) != 9 or not btw_number.isdigit():
            return False
        return self.validate_rsin_number(btw_number)

    def validate_lei_code(self, lei_code):
        """Validate LEI code (20 alphanumeric characters)"""
        if not lei_code or len(lei_code) != 20:
            return False
        return re.match(r'^[A-Z0-9]{20}$', lei_code.upper()) is not None

    def similarity(self, a, b):
        """Calculate similarity between two strings with normalization"""
        if not a or not b:
            return 0
        
        # Normalize strings for better comparison
        def normalize(s):
            s = unidecode.unidecode(s.lower())
            s = re.sub(r'[^\w\s]', ' ', s)
            s = re.sub(r'\s+', ' ', s).strip()
            return s
        
        norm_a = normalize(a)
        norm_b = normalize(b)
        
        return SequenceMatcher(None, norm_a, norm_b).ratio()

    def extract_legal_name_from_website(self, html_content, company_name):
        """Enhanced legal name extraction from website"""
        soup = BeautifulSoup(html_content, 'html.parser')
        found_names = []
        
        # Dutch legal suffixes (more comprehensive)
        dutch_suffixes = [
            r'BV', r'NV', r'VOF', r'CV', r'Eenmanszaak', r'Maatschap', 
            r'Commanditaire\s+Vennootschap', r'Vennootschap\s+onder\s+Firma', 
            r'Besloten\s+Vennootschap', r'Naamloze\s+Vennootschap',
            r'B\.V\.?', r'N\.V\.?', r'V\.O\.F\.?', r'C\.V\.?',
            r'Stichting', r'Vereniging', r'Coöperatie', r'Mutual'
        ]
        suffix_pattern = '|'.join(dutch_suffixes)
        
        # 1. Title tag
        title_tag = soup.find('title')
        if title_tag:
            title_text = title_tag.get_text().strip()
            title_matches = re.findall(rf'([A-Za-z][A-Za-z\s&.\'-]+\s+(?:{suffix_pattern}))', title_text, re.I)
            for match in title_matches:
                clean_name = re.sub(r'\s+', ' ', match.strip())
                if 5 <= len(clean_name) <= 150 and self.similarity(clean_name, company_name) > 0.2:
                    found_names.append((clean_name, 'title', self.similarity(clean_name, company_name)))

        # 2. Meta tags
        meta_tags = ['og:site_name', 'og:title', 'twitter:title', 'application-name']
        for meta_name in meta_tags:
            meta_tag = soup.find('meta', attrs={'name': meta_name}) or soup.find('meta', attrs={'property': meta_name})
            if meta_tag and meta_tag.get('content'):
                content = meta_tag.get('content').strip()
                if re.search(rf'(?:{suffix_pattern})', content, re.I):
                    if 5 <= len(content) <= 150 and self.similarity(content, company_name) > 0.2:
                        found_names.append((content, f'meta:{meta_name}', self.similarity(content, company_name)))

        # 3. JSON-LD structured data
        json_ld_scripts = soup.find_all('script', type='application/ld+json')
        for script in json_ld_scripts:
            try:
                data = json.loads(script.string)
                if isinstance(data, list):
                    objects = data
                else:
                    objects = [data]
                
                for obj in objects:
                    if isinstance(obj, dict):
                        obj_type = obj.get('@type', '').lower()
                        if obj_type in ['organization', 'corporation', 'localbusiness', 'company']:
                            for name_field in ['legalName', 'name', 'alternateName']:
                                name = obj.get(name_field, '').strip()
                                if name and re.search(rf'(?:{suffix_pattern})', name, re.I):
                                    if 5 <= len(name) <= 150 and self.similarity(name, company_name) > 0.2:
                                        found_names.append((name, f'json-ld:{name_field}', self.similarity(name, company_name)))
            except json.JSONDecodeError:
                continue

        # 4. Header elements (h1, h2, h3)
        for header_tag in soup.find_all(['h1', 'h2', 'h3']):
            header_text = header_tag.get_text().strip()
            if re.search(rf'(?:{suffix_pattern})', header_text, re.I):
                if 5 <= len(header_text) <= 150 and self.similarity(header_text, company_name) > 0.2:
                    found_names.append((header_text, f'header:{header_tag.name}', self.similarity(header_text, company_name)))

        # 5. Footer and copyright
        footers = soup.find_all(['footer', 'div'], class_=re.compile(r'footer|voet|copyright', re.I))
        for footer in footers:
            footer_text = footer.get_text()
            copyright_patterns = [
                rf'(?:©|Copyright|Alle\s+rechten\s+voorbehouden)\s*(?:20[0-9]{{2}})?\s*([A-Za-z][A-Za-z\s&.\'-]+(?:{suffix_pattern}))',
                rf'©\s*([A-Za-z][A-Za-z\s&.\'-]+(?:{suffix_pattern}))'
            ]
            
            for pattern in copyright_patterns:
                matches = re.finditer(pattern, footer_text, re.IGNORECASE)
                for match in matches:
                    name = match.group(1).strip()
                    clean_name = re.sub(r'\s+', ' ', name)
                    if 5 <= len(clean_name) <= 150 and self.similarity(clean_name, company_name) > 0.2:
                        found_names.append((clean_name, 'copyright', self.similarity(clean_name, company_name)))

        # 6. General text search for legal names
        text_content = soup.get_text()
        legal_name_matches = re.findall(rf'([A-Z][A-Za-z\s&.\'-]+\s+(?:{suffix_pattern}))', text_content, re.I)
        for match in legal_name_matches:
            clean_name = re.sub(r'\s+', ' ', match.strip())
            clean_name = re.sub(r'^[^\w]+|[^\w]+$', '', clean_name)
            if (5 <= len(clean_name) <= 150 and 
                not re.search(r'cookie|privacy|voorwaarden|contact|home|menu|login|search', clean_name, re.I) and
                len(clean_name.split()) >= 2 and
                self.similarity(clean_name, company_name) > 0.2):
                found_names.append((clean_name, 'text_search', self.similarity(clean_name, company_name)))

        # Remove duplicates and sort by similarity
        unique_names = {}
        for name, source, similarity in found_names:
            if name not in unique_names or unique_names[name][1] < similarity:
                unique_names[name] = (source, similarity)

        # Sort by similarity score
        sorted_names = [(name, source, sim) for name, (source, sim) in unique_names.items()]
        sorted_names.sort(key=lambda x: x[2], reverse=True)
        
        return [(name, source) for name, source, sim in sorted_names[:5]]  # Return top 5

    def setup_selenium(self):
        """Setup Selenium WebDriver with better options"""
        if self.driver is None:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument('--disable-web-security')
            chrome_options.add_argument('--allow-running-insecure-content')
            
            try:
                self.driver = webdriver.Chrome(options=chrome_options)
                self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
                self.driver.set_page_load_timeout(20)
                self.logger.info("Selenium WebDriver initialized successfully")
            except Exception as e:
                self.logger.error(f"Failed to setup Chrome driver: {e}")
                self.driver = None

    def search_kvk_register_enhanced(self, company_name):
        """Enhanced KvK register search with detailed result extraction"""
        self.logger.info(f"Searching KvK register for: {company_name}")
        
        if self.driver is None:
            self.setup_selenium()
        
        if self.driver is None:
            return None, None, []

        try:
            # Navigate to KvK search page
            search_url = "https://www.kvk.nl/en/search/"
            self.driver.get(search_url)
            time.sleep(2)
            
            # Wait for search input and enter company name
            search_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search'], input[name='q'], input[placeholder*='search'], input[placeholder*='Search']"))
            )
            
            search_input.clear()
            search_input.send_keys(company_name)
            search_input.send_keys(Keys.ENTER)
            
            time.sleep(3)
            
            # Wait for results to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # Look for search results
            search_results = []
            
            # Try different selectors for search results
            result_selectors = [
                'div[class*="search-result"]',
                'div[class*="result"]',
                'li[class*="search"]',
                'div[class*="company"]',
                'article',
                'div[class*="item"]'
            ]
            
            for selector in result_selectors:
                results = soup.select(selector)
                if results:
                    self.logger.info(f"Found {len(results)} results using selector: {selector}")
                    
                    for result in results[:5]:  # Limit to top 5 results
                        result_text = result.get_text()
                        
                        # Look for company names and KvK numbers
                        kvk_matches = re.findall(r'KvK[\s#:]*([0-9]{8})', result_text, re.IGNORECASE)
                        
                        # Look for links to detailed pages
                        links = result.find_all('a', href=True)
                        for link in links:
                            link_text = link.get_text().strip()
                            if len(link_text) > 5 and self.similarity(link_text, company_name) > 0.3:
                                result_url = link.get('href')
                                if result_url.startswith('/'):
                                    result_url = 'https://www.kvk.nl' + result_url
                                
                                search_results.append({
                                    'name': link_text,
                                    'url': result_url,
                                    'similarity': self.similarity(link_text, company_name),
                                    'kvk_preview': kvk_matches[0] if kvk_matches else None
                                })
                    
                    if search_results:
                        break
            
            # Sort by similarity and try to get detailed info from best matches
            search_results.sort(key=lambda x: x['similarity'], reverse=True)
            
            for result in search_results[:3]:  # Try top 3 results
                self.logger.info(f"Checking detailed page for: {result['name']} (similarity: {result['similarity']:.2f})")
                
                kvk_number, legal_name = self.extract_kvk_from_detail_page(result['url'])
                if kvk_number:
                    return kvk_number, legal_name, search_results
            
            time.sleep(1)
            
        except Exception as e:
            self.logger.debug(f"KvK register search error: {e}")
        
        return None, None, []

    def extract_kvk_from_detail_page(self, detail_url):
        """Extract KvK number from detailed company page"""
        try:
            self.logger.info(f"Extracting from detail page: {detail_url}")
            self.driver.get(detail_url)
            time.sleep(3)
            
            page_source = self.driver.page_source
            
            # Look for the specific VAT number section you mentioned
            vat_pattern = r'<span[^>]*class="[^"]*font-size-base[^"]*"[^>]*>.*?VAT number.*?([0-9]{8}).*?</span>'
            vat_match = re.search(vat_pattern, page_source, re.IGNORECASE | re.DOTALL)
            
            if vat_match:
                kvk_number = vat_match.group(1)
                if self.validate_kvk_number(kvk_number):
                    # Try to extract company name from page
                    soup = BeautifulSoup(page_source, 'html.parser')
                    
                    # Look for company name in various places
                    name_selectors = ['h1', 'h2', '.company-name', '[class*="name"]', 'title']
                    legal_name = None
                    
                    for selector in name_selectors:
                        elements = soup.select(selector)
                        for element in elements:
                            text = element.get_text().strip()
                            if len(text) > 5 and len(text) < 150:
                                legal_name = text
                                break
                        if legal_name:
                            break
                    
                    return kvk_number, legal_name
            
            # Fallback: look for any KvK patterns
            kvk_patterns_extended = [
                r'KvK[\s#:]*([0-9]{8})',
                r'(?:Chamber\s*of\s*Commerce|Kamer\s*van\s*Koophandel)[\s#:]*([0-9]{8})',
                r'VAT\s*number[\s#:]*([0-9]{8})',
                r'BTW[\s#:]*(?:NL)?([0-9]{8})'
            ]
            
            for pattern in kvk_patterns_extended:
                matches = re.findall(pattern, page_source, re.IGNORECASE)
                for match in matches:
                    if self.validate_kvk_number(match):
                        soup = BeautifulSoup(page_source, 'html.parser')
                        title_element = soup.find('title')
                        legal_name = title_element.get_text().strip() if title_element else None
                        return match, legal_name
            
        except Exception as e:
            self.logger.debug(f"Error extracting from detail page {detail_url}: {e}")
        
        return None, None

    def search_lei_lookup(self, company_name):
        """Search for LEI code using lei-lookup.com"""
        self.logger.info(f"Searching LEI lookup for: {company_name}")
        
        if self.driver is None:
            self.setup_selenium()
        
        if self.driver is None:
            return None, None

        try:
            # Navigate to LEI lookup
            lei_url = "https://www.lei-lookup.com/"
            self.driver.get(lei_url)
            time.sleep(2)
            
            # Find search input
            search_selectors = [
                "input[type='search']",
                "input[name='search']",
                "input[placeholder*='search']",
                "input[placeholder*='Search']",
                "#search",
                ".search-input"
            ]
            
            search_input = None
            for selector in search_selectors:
                try:
                    search_input = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    break
                except:
                    continue
            
            if not search_input:
                self.logger.warning("Could not find search input on LEI lookup site")
                return None, None
            
            search_input.clear()
            search_input.send_keys(company_name)
            search_input.send_keys(Keys.ENTER)
            
            time.sleep(3)
            
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # Look for LEI codes in results
            lei_patterns = [
                r'LEI[\s#:]*([A-Z0-9]{20})',
                r'([A-Z0-9]{20})',  # Generic 20-character code
            ]
            
            search_results = []
            for pattern in lei_patterns:
                matches = re.findall(pattern, page_source)
                for match in matches:
                    if self.validate_lei_code(match):
                        # Try to find associated company name
                        result_context = self.find_text_context(page_source, match)
                        
                        # Look for company name in the context
                        context_soup = BeautifulSoup(result_context, 'html.parser')
                        context_text = context_soup.get_text()
                        
                        # Simple heuristic: take text before or after LEI that looks like company name
                        company_name_match = re.search(rf'([A-Z][A-Za-z\s&.\'-]+(?:BV|NV|B\.V\.?|N\.V\.?)).*?{re.escape(match)}|{re.escape(match)}.*?([A-Z][A-Za-z\s&.\'-]+(?:BV|NV|B\.V\.?|N\.V\.?))', context_text, re.I)
                        
                        associated_name = None
                        if company_name_match:
                            associated_name = (company_name_match.group(1) or company_name_match.group(2)).strip()
                        
                        search_results.append({
                            'lei': match,
                            'name': associated_name,
                            'similarity': self.similarity(associated_name, company_name) if associated_name else 0
                        })
            
            # Sort by similarity and return best match
            if search_results:
                search_results.sort(key=lambda x: x['similarity'], reverse=True)
                best_result = search_results[0]
                return best_result['lei'], best_result['name']
            
        except Exception as e:
            self.logger.debug(f"LEI lookup error: {e}")
        
        return None, None

    def find_text_context(self, html_content, search_term, context_size=500):
        """Find text context around a search term"""
        try:
            index = html_content.find(search_term)
            if index == -1:
                return ""
            
            start = max(0, index - context_size)
            end = min(len(html_content), index + len(search_term) + context_size)
            
            return html_content[start:end]
        except:
            return ""

    def extract_codes_from_html(self, html_content):
        """Extract all codes from HTML with enhanced validation"""
        codes = {'kvk': None, 'rsin': None, 'lei': None, 'btw': None}
        
        # Extract KvK
        for pattern in self.kvk_patterns:
            matches = re.finditer(pattern, html_content, re.IGNORECASE)
            for match in matches:
                code = match.group(1) if len(match.groups()) > 0 else match.group(0)
                clean_code = re.sub(r'[^0-9]', '', code)
                if self.validate_kvk_number(clean_code):
                    codes['kvk'] = clean_code
                    break
            if codes['kvk']:
                break
        
        # Extract RSIN
        for pattern in self.rsin_patterns:
            matches = re.finditer(pattern, html_content, re.IGNORECASE)
            for match in matches:
                code = match.group(1) if len(match.groups()) > 0 else match.group(0)
                clean_code = re.sub(r'[^0-9]', '', code)
                if self.validate_rsin_number(clean_code):
                    codes['rsin'] = clean_code
                    break
            if codes['rsin']:
                break
        
        # Extract LEI
        for pattern in self.lei_patterns:
            matches = re.finditer(pattern, html_content, re.IGNORECASE)
            for match in matches:
                code = match.group(1) if len(match.groups()) > 0 else match.group(0)
                clean_code = re.sub(r'[^A-Z0-9]', '', code.upper())
                if self.validate_lei_code(clean_code):
                    codes['lei'] = clean_code
                    break
            if codes['lei']:
                break
        
        # Extract BTW
        for pattern in self.btw_patterns:
            matches = re.finditer(pattern, html_content, re.IGNORECASE)
            for match in matches:
                code = match.group(1) if len(match.groups()) > 0 else match.group(0)
                clean_code = re.sub(r'[^0-9]', '', code)
                if self.validate_btw_number(clean_code):
                    codes['btw'] = clean_code
                    break
            if codes['btw']:
                break
        
        return codes

    def fetch_with_requests(self, url, timeout=10):
        """Enhanced fetch with better error handling"""
        try:
            response = self.session.get(url, timeout=timeout, allow_redirects=True)
            response.raise_for_status()
            
            # Check content type
            content_type = response.headers.get('content-type', '').lower()
            if 'html' in content_type or 'xml' in content_type:
                return response.text
            else:
                self.logger.debug(f"Non-HTML content type for {url}: {content_type}")
                return None
                
        except requests.exceptions.RequestException as e:
            self.logger.debug(f"Requests failed for {url}: {e}")
            return None

    def fetch_with_selenium(self, url):
        """Enhanced Selenium fetch"""
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
            
            # Scroll to load dynamic content
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            
            return self.driver.page_source
        except Exception as e:
            self.logger.debug(f"Selenium failed for {url}: {e}")
            return None

    def get_dutch_paths(self):
        """Enhanced list of paths to check"""
        return [
            '',  # Homepage
            '/contact', '/contactgegevens', '/contact-us', '/contact-info',
            '/over-ons', '/about', '/about-us', '/bedrijfsinfo', '/bedrijfsinformatie',
            '/algemene-voorwaarden', '/legal', '/juridisch', '/colofon', '/impressum',
            '/privacy', '/privacy-policy', '/privacybeleid',
            '/footer', '/imprint', '/company-info'
        ]

    def search_website_for_info(self, base_url, company_name):
        """Enhanced website search with better error handling"""
        found_legal_names = []
        found_codes = {'kvk': None, 'rsin': None, 'lei': None, 'btw': None}
        
        dutch_paths = self.get_dutch_paths()

        for path in dutch_paths:
            url = base_url.rstrip('/') + path
            self.logger.info(f"Checking: {url}")

            try:
                # Try requests first
                html_content = self.fetch_with_requests(url)
                
                # Fallback to Selenium if requests fails
                if not html_content:
                    self.logger.info(f"Requests failed, trying Selenium for: {url}")
                    html_content = self.fetch_with_selenium(url)

                if html_content:
                    # Extract codes
                    codes = self.extract_codes_from_html(html_content)
                    
                    # Update found codes (keep first valid one found)
                    for code_type, code_value in codes.items():
                        if code_value and not found_codes[code_type]:
                            found_codes[code_type] = code_value
                    
                    # Extract legal names
                    legal_names_with_sources = self.extract_legal_name_from_website(html_content, company_name)
                    found_legal_names.extend(legal_names_with_sources)
                    
                    # If we found KvK and legal names, we can break
                    if found_codes['kvk'] and found_legal_names:
                        break

            except Exception as e:
                self.logger.debug(f"Error processing {url}: {e}")

            time.sleep(0.5)

        return found_legal_names, found_codes

    def process_single_dutch_company(self, company):
        """Enhanced processing with improved search strategy"""
        self.logger.info(f"\n--- Processing Dutch company: {company['name']} ---")
        
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
            'legal_name_kvk': '',
            'kvk_number': '',
            'rsin': '',
            'lei': '',
            'btw': '',
            'status': 'not_found',
            'search_location': '',
            'search_attempts': [],
            'error': None
        }

        try:
            # Step 1: Direct KvK register search with original company name
            self.logger.info("Step 1: Searching KvK register with original name...")
            kvk_number, kvk_legal_name, search_results = self.search_kvk_register_enhanced(company['name'])
            
            if kvk_number and kvk_legal_name:
                result['kvk_number'] = kvk_number
                result['legal_name_kvk'] = kvk_legal_name
                result['status'] = 'found'
                result['search_location'] = 'kvk_register_direct'
                result['search_attempts'].append(f"KvK direct search successful: {kvk_number}")
                self.logger.info(f"✓ Found via KvK direct search: KvK {kvk_number}, Legal name: {kvk_legal_name}")
                return result
            else:
                result['search_attempts'].append("KvK direct search failed")
            
            # Step 2: Website scraping for legal names
            if not website:
                result['status'] = 'no_website'
                result['error'] = 'No website provided and not found in KvK register'
                result['search_attempts'].append("No website available for scraping")
                return result
            
            self.logger.info("Step 2: Scraping website for legal names...")
            legal_names_with_sources, website_codes = self.search_website_for_info(website, company['name'])
            
            if legal_names_with_sources:
                result['legal_name_website'] = legal_names_with_sources[0][0]
                result['legal_name_source'] = legal_names_with_sources[0][1]
                result['search_attempts'].append(f"Website scraping found legal name: {legal_names_with_sources[0][0]}")
                self.logger.info(f"✓ Found legal name on website: {legal_names_with_sources[0][0]} (source: {legal_names_with_sources[0][1]})")
            else:
                result['search_attempts'].append("Website scraping found no legal names")
            
            # Store website codes
            result['kvk_number'] = website_codes.get('kvk', '')
            result['rsin'] = website_codes.get('rsin', '')
            result['lei'] = website_codes.get('lei', '')
            result['btw'] = website_codes.get('btw', '')
            
            if website_codes.get('kvk'):
                result['status'] = 'found'
                result['search_location'] = 'website_direct'
                result['search_attempts'].append(f"Website direct extraction successful: {website_codes['kvk']}")
                self.logger.info(f"✓ Found KvK on website: {website_codes['kvk']}")
                return result
            
            # Step 3: KvK search with website legal names
            if legal_names_with_sources:
                for legal_name, source in legal_names_with_sources[:3]:  # Try top 3 legal names
                    self.logger.info(f"Step 3: Trying KvK search with legal name: {legal_name}")
                    kvk_number_from_legal, kvk_legal_name_found, _ = self.search_kvk_register_enhanced(legal_name)
                    
                    if kvk_number_from_legal:
                        result['kvk_number'] = kvk_number_from_legal
                        result['legal_name_kvk'] = kvk_legal_name_found
                        result['status'] = 'found'
                        result['search_location'] = 'kvk_via_website_name'
                        result['search_attempts'].append(f"KvK search via website legal name successful: {kvk_number_from_legal}")
                        self.logger.info(f"✓ Found KvK via legal name search: {kvk_number_from_legal}")
                        return result
                    else:
                        result['search_attempts'].append(f"KvK search with '{legal_name}' failed")
            
            # Step 4: LEI lookup as fallback
            self.logger.info("Step 4: Trying LEI lookup as fallback...")
            
            # Try LEI with original name first
            lei_code, lei_legal_name = self.search_lei_lookup(company['name'])
            if not lei_code and legal_names_with_sources:
                # Try with legal names from website
                for legal_name, source in legal_names_with_sources[:2]:
                    lei_code, lei_legal_name = self.search_lei_lookup(legal_name)
                    if lei_code:
                        break
            
            if lei_code:
                result['lei'] = lei_code
                if lei_legal_name:
                    result['legal_name_website'] = lei_legal_name
                    result['legal_name_source'] = 'lei_lookup'
                result['status'] = 'found_lei_only'
                result['search_location'] = 'lei_lookup'
                result['search_attempts'].append(f"LEI lookup successful: {lei_code}")
                self.logger.info(f"✓ Found LEI code: {lei_code}")
                return result
            else:
                result['search_attempts'].append("LEI lookup failed")
            
            # Step 5: Check if we have any other valid codes
            if any([website_codes.get('rsin'), website_codes.get('btw')]):
                result['status'] = 'found_partial'
                result['search_location'] = 'website_other_codes'
                found_codes = [k.upper() for k, v in website_codes.items() if v and k not in ['kvk', 'lei']]
                result['search_attempts'].append(f"Found other codes on website: {', '.join(found_codes)}")
                self.logger.info(f"✓ Found other codes on website: {', '.join(found_codes)}")
                return result
            
            # If we reach here, nothing was found
            result['status'] = 'not_found'
            result['search_location'] = 'exhausted_all_methods'
            result['search_attempts'].append("All search methods exhausted without success")
            self.logger.info("✗ No valid codes found through any method")

        except Exception as error:
            result['error'] = str(error)
            result['status'] = 'error'
            result['search_attempts'].append(f"Error occurred: {str(error)}")
            self.logger.error(f"✗ Error: {error}")

        return result

    def process_dutch_companies(self, companies):
        """Process multiple Dutch companies with enhanced reporting"""
        self.logger.info(f"Starting enhanced Dutch KvK extraction for {len(companies)} companies...")
        
        start_time = time.time()
        
        for i, company in enumerate(companies, 1):
            self.logger.info(f"\n[{i}/{len(companies)}] Processing: {company['name']}")
            result = self.process_single_dutch_company(company)
            self.results.append(result)
            
            # Progress reporting
            if i % 10 == 0 or i == len(companies):
                elapsed = time.time() - start_time
                avg_time = elapsed / i
                remaining = (len(companies) - i) * avg_time
                
                successful = len([r for r in self.results if r['status'] in ['found', 'found_lei_only', 'found_partial']])
                success_rate = (successful / i) * 100
                
                self.logger.info(f"Progress: {i}/{len(companies)} ({success_rate:.1f}% success rate, ~{remaining/60:.1f} min remaining)")
            
            # Respectful delay between companies
            time.sleep(2)

        return self.results

    def load_companies_from_excel(self, file_path):
        """Load Dutch companies from Excel file with enhanced column detection"""
        try:
            df = pd.read_excel(file_path)
            
            column_mappings = {
                'pe_name': ['PE NAME', 'PE_NAME', 'Private Equity', 'PE'],
                'pe_country': ['Country (HQ)', 'PE Country', 'Country'],
                'pe_website': ['Website', 'PE Website', 'PE_Website'],
                'company_name': ['Portfolio Companies', 'Company Name', 'Target Company', 'Company'],
                'target_website': ['Target Website', 'Company Website', 'Website', 'URL'],
                'target_geography': ['Target Geography', 'Geography', 'Location'],
                'target_industry': ['Target Industry', 'Industry', 'Sector'],
                'target_sub_industry': ['Target Sub-Industry', 'Sub Industry', 'Sub-Industry'],
                'entry_year': ['Entry', 'Entry Year', 'Year']
            }
            
            found_columns = {}
            for key, possible_names in column_mappings.items():
                for col in df.columns:
                    if any(col.strip().lower() == name.lower() for name in possible_names):
                        found_columns[key] = col
                        break
                    # Partial matching for similar column names
                    if not found_columns.get(key):
                        for name in possible_names:
                            if name.lower() in col.lower() or col.lower() in name.lower():
                                found_columns[key] = col
                                break
            
            required_columns = ['company_name']
            missing_required = [col for col in required_columns if col not in found_columns]
            if missing_required:
                self.logger.error(f"Missing required columns. Available columns: {list(df.columns)}")
                raise ValueError(f"Missing required columns for company names. Expected one of: {column_mappings['company_name']}")
            
            companies = []
            for idx, row in df.iterrows():
                company_name = row[found_columns['company_name']]
                
                if pd.isna(company_name) or str(company_name).strip() == '':
                    continue
                
                # Get target website
                target_website = ''
                if 'target_website' in found_columns and not pd.isna(row[found_columns['target_website']]):
                    target_website = str(row[found_columns['target_website']]).strip()
                
                # Get PE website
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
            
            self.logger.info(f"Loaded {len(companies)} Dutch companies from {file_path}")
            self.logger.info(f"Found column mappings: {found_columns}")
            return companies
            
        except Exception as e:
            self.logger.error(f"Error loading Excel file: {e}")
            raise

    def save_results_to_excel(self, filename='dutch_kvk_results_enhanced.xlsx'):
        """Save results to Excel with enhanced formatting"""
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
                    'Legal Name (KvK)': result['legal_name_kvk'],
                    'KvK Number': result['kvk_number'],
                    'RSIN': result['rsin'],
                    'LEI': result['lei'],
                    'BTW': result['btw'],
                    'Status': result['status'],
                    'Found At': result['search_location'],
                    'Search Attempts': '; '.join(result.get('search_attempts', [])),
                    'Error': result['error'] or ''
                })
            
            df = pd.DataFrame(data)
            
            # Create Excel writer with formatting
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Results', index=False)
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Results']
                
                # Auto-adjust column widths
                for column in df.columns:
                    column_length = max(df[column].astype(str).map(len).max(), len(column))
                    col_idx = df.columns.get_loc(column) + 1
                    worksheet.column_dimensions[worksheet.cell(row=1, column=col_idx).column_letter].width = min(column_length + 2, 50)
            
            self.logger.info(f"\nResults saved to {filename}")
            
        except Exception as e:
            self.logger.error(f"Error saving to Excel: {e}")
            # Fallback to JSON
            self.save_results_to_json()

    def save_results_to_json(self, filename='dutch_kvk_results_enhanced.json'):
        """Save results to JSON file"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)
        self.logger.info(f"Results saved to {filename}")

    def generate_summary_report(self):
        """Generate comprehensive enhanced summary report"""
        total = len(self.results)
        if total == 0:
            print("No results to summarize.")
            return
            
        found = len([r for r in self.results if r['status'] == 'found'])
        found_lei = len([r for r in self.results if r['status'] == 'found_lei_only'])
        found_partial = len([r for r in self.results if r['status'] == 'found_partial'])
        not_found = len([r for r in self.results if r['status'] == 'not_found'])
        errors = len([r for r in self.results if r['status'] == 'error'])
        no_website = len([r for r in self.results if r['status'] == 'no_website'])
        
        # Breakdown by search location
        kvk_direct = len([r for r in self.results if r['search_location'] == 'kvk_register_direct'])
        website_direct = len([r for r in self.results if r['search_location'] == 'website_direct'])
        kvk_via_website = len([r for r in self.results if r['search_location'] == 'kvk_via_website_name'])
        lei_found = len([r for r in self.results if r['search_location'] == 'lei_lookup'])
        
        print('\n' + '='*70)
        print('ENHANCED DUTCH KVK CODE EXTRACTION SUMMARY')
        print('='*70)
        print(f'Total companies processed: {total}')
        print(f'KvK codes successfully found: {found} ({(found/total*100):.1f}%)')
        print(f'  - Found via direct KvK search: {kvk_direct}')
        print(f'  - Found directly on websites: {website_direct}')
        print(f'  - Found via KvK search with website names: {kvk_via_website}')
        print(f'LEI codes found: {found_lei} ({(found_lei/total*100):.1f}%)')
        print(f'  - Found via LEI lookup: {lei_found}')
        print(f'Partial success (RSIN/BTW found): {found_partial} ({(found_partial/total*100):.1f}%)')
        print(f'Not found: {not_found} ({(not_found/total*100):.1f}%)')
        print(f'Errors: {errors} ({(errors/total*100):.1f}%)')
        print(f'No website provided: {no_website} ({(no_website/total*100):.1f}%)')
        
        total_success = found + found_lei + found_partial
        print(f'\nOverall success rate: {total_success}/{total} ({(total_success/total*100):.1f}%)')
        
        # Show some example successful results
        successful_results = [r for r in self.results if r['status'] in ['found', 'found_lei_only']][:3]
        if successful_results:
            print(f'\n{"-"*50}')
            print('SAMPLE SUCCESSFUL EXTRACTIONS:')
            print(f'{"-"*50}')
            for result in successful_results:
                print(f"\n• {result['original_company_name']}")
                if result['kvk_number']:
                    print(f"  KvK: {result['kvk_number']}")
                if result['lei']:
                    print(f"  LEI: {result['lei']}")
                print(f"  Method: {result['search_location']}")
                if result['legal_name_kvk']:
                    print(f"  Legal name: {result['legal_name_kvk']}")

    def close(self):
        """Clean up resources"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
        try:
            self.session.close()
        except:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


def main():
    excel_file = 'smalldbmachinery_nl2.xlsx'
    
    import sys
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found!")
        print("\nRequired Excel columns:")
        print("- Portfolio Companies (or similar for company names)")
        print("- Target Website (or similar for company websites) - optional")
        print(f"\nUsage: python {sys.argv[0]} your_file.xlsx")
        return

    with DutchKvKExtractor() as extractor:
        try:
            print("="*70)
            print("ENHANCED DUTCH KVK CODE EXTRACTOR")
            print("="*70)
            print("Loading Dutch companies from Excel...")
            companies = extractor.load_companies_from_excel(excel_file)
            
            if not companies:
                print("No valid companies found in the Excel file!")
                return
            
            print(f"\nProcessing {len(companies)} Dutch companies...")
            print("\nThis enhanced extractor will:")
            print("1. Search KvK register directly with company names")
            print("2. Scrape company websites for legal names with advanced extraction")
            print("3. Search KvK register with found legal names")
            print("4. Extract KvK numbers from detailed company pages")
            print("5. Fallback to LEI lookup at lei-lookup.com")
            print("6. Validate all codes according to Dutch standards")
            print("7. Provide detailed search attempt logs")
            print("=" * 70)
            
            input("\nPress Enter to start processing...")
            
            results = extractor.process_dutch_companies(companies)
            
            # Save results
            extractor.save_results_to_excel()
            extractor.save_results_to_json()
            
            # Generate summary
            extractor.generate_summary_report()
            
            # Show detailed results for successful extractions
            print('\n' + '='*70)
            print('DETAILED RESULTS - ALL SUCCESSFUL EXTRACTIONS')
            print('='*70)
            
            successful_results = [r for r in results if r['status'] in ['found', 'found_lei_only', 'found_partial']]
            for result in successful_results:
                print(f"\n{result['original_company_name']}:")
                if result['kvk_number']:
                    print(f"  ✓ KvK Number: {result['kvk_number']}")
                if result['rsin']:
                    print(f"  ✓ RSIN: {result['rsin']}")
                if result['lei']:
                    print(f"  ✓ LEI: {result['lei']}")
                if result['btw']:
                    print(f"  ✓ BTW: {result['btw']}")
                print(f"  Status: {result['status']}")
                print(f"  Method: {result['search_location']}")
                if result['legal_name_kvk']:
                    print(f"  Legal name (KvK): {result['legal_name_kvk']}")
                if result['legal_name_website']:
                    print(f"  Legal name (Website): {result['legal_name_website']} (source: {result['legal_name_source']})")
                if result['website']:
                    print(f"  Website: {result['website']}")
                if result['search_attempts']:
                    print(f"  Search path: {' → '.join(result['search_attempts'][-3:])}")  # Show last 3 attempts
            
            # Show failures for debugging
            failed_results = [r for r in results if r['status'] not in ['found', 'found_lei_only', 'found_partial']]
            if failed_results:
                print(f'\n{"="*70}')
                print(f'FAILED EXTRACTIONS ({len(failed_results)} companies)')
                print('='*70)
                for result in failed_results[:10]:  # Show first 10 failures
                    print(f"\n✗ {result['original_company_name']}: {result['status']}")
                    if result['error']:
                        print(f"  Error: {result['error']}")
                    if result['legal_name_website']:
                        print(f"  Found legal name on website: {result['legal_name_website']}")
                    if result['search_attempts']:
                        print(f"  Attempts: {' → '.join(result['search_attempts'])}")
                        
                if len(failed_results) > 10:
                    print(f"\n... and {len(failed_results) - 10} more failed extractions (see Excel file for complete list)")
                    
        except KeyboardInterrupt:
            print("\n\nProcess interrupted by user. Saving partial results...")
            if extractor.results:
                extractor.save_results_to_excel('partial_dutch_kvk_results.xlsx')
                extractor.generate_summary_report()
            
        except Exception as error:
            print(f'Fatal error: {error}')
            import traceback
            traceback.print_exc()
            
            # Save partial results if any
            if extractor.results:
                print("Saving partial results...")
                extractor.save_results_to_excel('error_dutch_kvk_results.xlsx')


if __name__ == "__main__":
    main()