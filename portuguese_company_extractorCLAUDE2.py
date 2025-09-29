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
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    ElementNotInteractableException, 
    TimeoutException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
    NoSuchElementException,
    WebDriverException
)
from urllib.parse import urljoin, urlparse, quote, urlencode
import logging
import os
from difflib import SequenceMatcher

class PortugueseCompanyExtractorFixed:
    def __init__(self):
        self.results = []
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept-Language': 'pt-PT,pt;q=0.9,en;q=0.8',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        })
        self.driver = None
        self.wait = None

        # Setup logging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        self.logger = logging.getLogger(__name__)

        # Portuguese NIF patterns
        self.nif_patterns = [
            r'NIF[\s:]*([0-9]{9})',
            r'N\.?I\.?F\.?[\s:]*([0-9]{9})',
            r'Contribuinte[\s:]*([0-9]{9})',
            r'NIPC[\s:]*([0-9]{9})',
            r'\b([0-9]{9})\b(?=\s*contribuinte)',
        ]

    def setup_driver(self, headless=True):
        """Setup Chrome driver with improved options"""
        if self.driver:
            return

        options = Options()
        
        # Essential options
        if headless:
            options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        
        # Anti-detection measures
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        # Performance
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-plugins')
        options.add_argument('--disable-images')
        options.add_argument('--disable-javascript')  # Try without JS first
        
        # User agent
        options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

        try:
            self.driver = webdriver.Chrome(options=options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.wait = WebDriverWait(self.driver, 30)  # Increased timeout
            self.logger.info("Chrome driver initialized successfully")
            
        except Exception as e:
            self.logger.error(f"Could not initialize Chrome driver: {str(e)}")
            raise

    def similarity(self, a, b):
        """Calculate similarity between two strings"""
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    def extract_nif_from_text(self, text):
        """Extract Portuguese NIF from text"""
        for pattern in self.nif_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                nif = match.group(1) if len(match.groups()) > 0 else match.group(0)
                if len(nif) == 9 and nif.isdigit() and nif[0] in '123456789':
                    return nif
        return None

    def search_einforma_direct_url(self, company_name):
        """Try direct URL approach for eInforma search"""
        try:
            # Method 1: Direct search URL
            encoded_name = quote(company_name.encode('utf-8'))
            search_urls = [
                f"https://www.einforma.pt/servlet/app/portal/ENTP/prod/LISTADO_EMPRESAS/criterio/denominacion/valor/{encoded_name}",
                f"https://www.einforma.pt/servlet/app/portal/ENTP/prod/LISTADO_EMPRESAS?denominacion={encoded_name}",
                f"https://www.einforma.pt/search?q={encoded_name}",
            ]
            
            for url in search_urls:
                try:
                    self.logger.info(f"Trying direct URL: {url}")
                    
                    response = self.session.get(url, timeout=15)
                    if response.status_code == 200:
                        soup = BeautifulSoup(response.content, 'html.parser')
                        results = self.parse_einforma_search_results(soup, company_name)
                        if results:
                            return results
                            
                except Exception as e:
                    self.logger.warning(f"Direct URL failed: {str(e)}")
                    continue
                    
            return None
            
        except Exception as e:
            self.logger.error(f"Error in direct URL search: {str(e)}")
            return None

    def search_einforma_with_requests(self, company_name):
        """Search using requests library with form submission"""
        try:
            self.logger.info(f"Searching eInforma with requests for: {company_name}")
            
            # First, get the main page to establish session
            main_page = self.session.get("https://www.einforma.pt/", timeout=15)
            if main_page.status_code != 200:
                return None
                
            soup = BeautifulSoup(main_page.content, 'html.parser')
            
            # Look for search forms
            forms = soup.find_all('form')
            for form in forms:
                form_action = form.get('action', '')
                
                # Try to submit search form
                if 'search' in form_action.lower() or 'empresa' in form_action.lower():
                    form_data = {}
                    
                    # Find input fields
                    inputs = form.find_all('input')
                    for inp in inputs:
                        name = inp.get('name')
                        if name:
                            if 'search' in name.lower() or 'nome' in name.lower() or 'empresa' in name.lower():
                                form_data[name] = company_name
                            elif inp.get('type') == 'hidden':
                                form_data[name] = inp.get('value', '')
                    
                    if form_data:
                        try:
                            # Submit form
                            if form_action.startswith('/'):
                                form_action = 'https://www.einforma.pt' + form_action
                            elif not form_action.startswith('http'):
                                form_action = 'https://www.einforma.pt/' + form_action
                                
                            result = self.session.post(form_action, data=form_data, timeout=15)
                            if result.status_code == 200:
                                soup = BeautifulSoup(result.content, 'html.parser')
                                results = self.parse_einforma_search_results(soup, company_name)
                                if results:
                                    return results
                                    
                        except Exception as e:
                            self.logger.warning(f"Form submission failed: {str(e)}")
                            continue
            
            return None
            
        except Exception as e:
            self.logger.error(f"Error in requests search: {str(e)}")
            return None

    def accept_cookies(self):
        """Accept cookies and privacy notices"""
        try:
            # Common cookie acceptance selectors
            cookie_selectors = [
                "button[id*='accept']",
                "button[class*='accept']",
                "button[id*='cookie']",
                "button[class*='cookie']",
                "button:contains('Aceitar')",
                "button:contains('Accept')",
                "button:contains('OK')",
                ".cookie-accept",
                "#cookie-accept",
                "[data-testid*='accept']",
                "[data-testid*='cookie']"
            ]
            
            for selector in cookie_selectors:
                try:
                    if ":contains(" in selector:
                        # Use XPath for text content
                        text = selector.split('("')[1].split('")')[0]
                        xpath = f"//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{text.lower()}')]"
                        elements = self.driver.find_elements(By.XPATH, xpath)
                    else:
                        elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    
                    for element in elements:
                        if element.is_displayed() and element.is_enabled():
                            self.safe_click_element(element)
                            self.logger.info("Cookies accepted")
                            time.sleep(2)
                            return True
                            
                except Exception as e:
                    continue
            
            return False
            
        except Exception as e:
            self.logger.warning(f"Error accepting cookies: {str(e)}")
            return False

    def wait_for_search_results(self):
        """Wait for search results to load"""
        try:
            # Wait for any of these indicators that results have loaded
            result_indicators = [
                (By.CSS_SELECTOR, "a[href*='/nif/']"),
                (By.CSS_SELECTOR, ".result"),
                (By.CSS_SELECTOR, ".company"),
                (By.CSS_SELECTOR, ".empresa"),
                (By.XPATH, "//a[contains(@href, 'nif')]"),
                (By.XPATH, "//div[contains(@class, 'result')]")
            ]
            
            for by, selector in result_indicators:
                try:
                    self.wait.until(EC.presence_of_element_located((by, selector)))
                    self.logger.info(f"Search results detected with selector: {selector}")
                    return True
                except TimeoutException:
                    continue
            
            # Fallback: wait for page change
            time.sleep(5)
            return True
            
        except Exception as e:
            self.logger.warning(f"Error waiting for results: {str(e)}")
            return False

    def click_company_result_and_extract(self, result_element, company_name):
        """Click on a company result and extract detailed information"""
        try:
            # Get the link before clicking (in case element becomes stale)
            company_link = result_element.get_attribute('href')
            company_text = result_element.text.strip()
            
            self.logger.info(f"Attempting to click on: {company_text}")
            
            # Method 1: Try regular click
            try:
                # Scroll element into view
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", result_element)
                time.sleep(1)
                
                # Wait for element to be clickable
                clickable_element = self.wait.until(EC.element_to_be_clickable(result_element))
                clickable_element.click()
                
                self.logger.info("Successfully clicked with regular click")
                
            except ElementClickInterceptedException as e:
                self.logger.warning(f"Click intercepted: {str(e)}")
                
                # Method 2: Remove overlaying elements and try again
                self.driver.execute_script("""
                    // Remove common overlay elements
                    var overlays = document.querySelectorAll(
                        '.modal, .popup, .overlay, .loading, .spinner, ' +
                        '[class*="modal"], [class*="popup"], [class*="overlay"], ' +
                        '[id*="modal"], [id*="popup"], [id*="overlay"]'
                    );
                    overlays.forEach(function(overlay) {
                        if (overlay.style) overlay.style.display = 'none';
                        overlay.remove();
                    });
                """)
                
                time.sleep(1)
                
                try:
                    result_element.click()
                    self.logger.info("Successfully clicked after removing overlays")
                except:
                    # Method 3: JavaScript click
                    self.driver.execute_script("arguments[0].click();", result_element)
                    self.logger.info("Successfully clicked with JavaScript")
            
            except Exception as e:
                self.logger.warning(f"Regular click failed: {str(e)}")
                # Method 4: Direct navigation if we have the link
                if company_link:
                    self.driver.get(company_link)
                    self.logger.info(f"Navigated directly to: {company_link}")
                else:
                    return None
            
            # Wait for company page to load
            time.sleep(3)
            self.wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            
            # Take screenshot of company page
            try:
                self.driver.save_screenshot(f"company_page_{company_name.replace(' ', '_')}.png")
            except:
                pass
            
            # Extract company details from the page
            return self.extract_company_details_from_page(company_name)
            
        except Exception as e:
            self.logger.error(f"Error clicking and extracting from result: {str(e)}")
            return None

    def extract_company_details_from_page(self, original_company_name):
        """Extract detailed company information from the current page"""
        try:
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            current_url = self.driver.current_url
            
            self.logger.info(f"Extracting details from: {current_url}")
            
            company_details = {
                'company_name': '',
                'legal_name': '',
                'nif': '',
                'einforma_url': current_url,
                'similarity': 0.0
            }
            
            # Strategy 1: Extract NIF from URL
            nif_from_url = re.search(r'/nif/([0-9]{9})', current_url)
            if nif_from_url:
                company_details['nif'] = nif_from_url.group(1)
                self.logger.info(f"Found NIF in URL: {company_details['nif']}")
            
            # Strategy 2: Find company name in page title or headers
            title_selectors = [
                'h1',
                'h2',
                '.title',
                '.company-name',
                '.empresa-nome',
                '[itemprop="name"]'
            ]
            
            for selector in title_selectors:
                try:
                    elements = soup.select(selector)
                    for element in elements:
                        text = element.get_text(strip=True)
                        
                        # Check if this looks like a company name
                        if any(suffix in text.upper() for suffix in ['LDA', 'LIMITADA', 'S.A.', 'SA', 'UNIPESSOAL']):
                            # Remove NIF from company name if present
                            clean_name = re.sub(r'\b[0-9]{9}\b', '', text).strip()
                            if len(clean_name) > 3:
                                company_details['company_name'] = clean_name
                                company_details['legal_name'] = clean_name
                                company_details['similarity'] = self.similarity(clean_name, original_company_name)
                                self.logger.info(f"Found company name: {clean_name}")
                                break
                    
                    if company_details['company_name']:
                        break
                        
                except Exception:
                    continue
            
            # Strategy 3: Extract NIF from page content if not found in URL
            if not company_details['nif']:
                page_text = soup.get_text()
                nif = self.extract_nif_from_text(page_text)
                if nif:
                    company_details['nif'] = nif
                    self.logger.info(f"Found NIF in page content: {nif}")
            
            # Strategy 4: Look for specific eInforma patterns
            # Check for "contribuinte de" pattern
            contribuinte_pattern = r'contribuinte de ([^<\n]+)'
            contribuinte_match = re.search(contribuinte_pattern, soup.get_text(), re.IGNORECASE)
            if contribuinte_match and not company_details['company_name']:
                company_name = contribuinte_match.group(1).strip()
                if len(company_name) > 3:
                    company_details['company_name'] = company_name
                    company_details['legal_name'] = company_name
                    company_details['similarity'] = self.similarity(company_name, original_company_name)
            
            # Ensure we have at least a company name
            if not company_details['company_name']:
                company_details['company_name'] = original_company_name
                company_details['legal_name'] = original_company_name
                company_details['similarity'] = 1.0
            
            return company_details
            
        except Exception as e:
            self.logger.error(f"Error extracting company details: {str(e)}")
            return None

    def search_einforma_selenium_improved(self, company_name):
        """Improved Selenium search with cookie handling and result clicking"""
        try:
            self.logger.info(f"Searching eInforma with Selenium for: {company_name}")
            
            # Setup driver if needed
            self.setup_driver(headless=False)  # Non-headless for debugging
            
            # Navigate to the site
            self.driver.get("https://www.einforma.pt/")
            
            # Wait for page to load
            self.wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            time.sleep(3)
            
            # Accept cookies
            self.accept_cookies()
            
            # Take initial screenshot
            try:
                self.driver.save_screenshot(f"einforma_initial_{company_name.replace(' ', '_')}.png")
                self.logger.info("Initial screenshot saved")
            except:
                pass
            
            # Find search input field
            search_input = None
            search_strategies = [
                (By.CSS_SELECTOR, "input[name*='nome']"),
                (By.CSS_SELECTOR, "input[name*='empresa']"),
                (By.CSS_SELECTOR, "input[name*='search']"),
                (By.CSS_SELECTOR, "input[type='text']"),
                (By.CSS_SELECTOR, "input[type='search']"),
                (By.CSS_SELECTOR, "input[placeholder*='empresa']"),
                (By.CSS_SELECTOR, "input[placeholder*='nome']"),
            ]
            
            for by, selector in search_strategies:
                try:
                    elements = self.driver.find_elements(by, selector)
                    for element in elements:
                        if element.is_displayed() and element.is_enabled():
                            search_input = element
                            self.logger.info(f"Found search input with: {selector}")
                            break
                    if search_input:
                        break
                except Exception as e:
                    self.logger.warning(f"Search strategy failed {selector}: {str(e)}")
                    continue
            
            if not search_input:
                self.logger.error("No search input found")
                return None
            
            # Perform search
            try:
                # Scroll into view and focus
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", search_input)
                time.sleep(1)
                
                search_input.click()
                search_input.clear()
                search_input.send_keys(company_name)
                time.sleep(2)
                
                # Submit search
                search_input.send_keys(Keys.RETURN)
                self.logger.info(f"Search submitted for: {company_name}")
                
                # Wait for results
                self.wait_for_search_results()
                
                # Take screenshot of search results
                try:
                    self.driver.save_screenshot(f"search_results_{company_name.replace(' ', '_')}.png")
                    self.logger.info("Search results screenshot saved")
                except:
                    pass
                
            except Exception as e:
                self.logger.error(f"Error performing search: {str(e)}")
                return None
            
            # Find and click on company results
            try:
                # Look for company result links
                result_selectors = [
                    "a[href*='/nif/']",
                    "a[href*='ETIQUETA_EMPRESA']",
                    "a[href*='contribuinte']",
                    ".result a",
                    ".company a",
                    ".empresa a"
                ]
                
                company_results = []
                
                for selector in result_selectors:
                    try:
                        elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        for element in elements:
                            if element.is_displayed():
                                href = element.get_attribute('href')
                                text = element.text.strip()
                                
                                # Check if this looks like a company result
                                if (href and text and len(text) > 3 and
                                    ('nif' in href or 'empresa' in href.lower())):
                                    
                                    similarity = self.similarity(text, company_name)
                                    company_results.append({
                                        'element': element,
                                        'text': text,
                                        'href': href,
                                        'similarity': similarity
                                    })
                        
                        if company_results:
                            break
                            
                    except Exception as e:
                        self.logger.warning(f"Error finding results with {selector}: {str(e)}")
                        continue
                
                if not company_results:
                    self.logger.warning("No company results found")
                    return None
                
                # Sort by similarity and try the best matches
                company_results.sort(key=lambda x: x['similarity'], reverse=True)
                
                self.logger.info(f"Found {len(company_results)} potential results")
                
                for i, result in enumerate(company_results[:3]):  # Try top 3 results
                    self.logger.info(f"Trying result {i+1}: {result['text']} (similarity: {result['similarity']:.2f})")
                    
                    try:
                        # Click and extract details
                        details = self.click_company_result_and_extract(result['element'], company_name)
                        
                        if details and details['nif']:
                            self.logger.info(f"Successfully extracted details for: {details['company_name']}")
                            return [details]  # Return as list for consistency
                        
                        # Navigate back to search results for next attempt
                        if i < len(company_results) - 1:
                            self.driver.back()
                            time.sleep(2)
                            self.wait_for_search_results()
                        
                    except Exception as e:
                        self.logger.warning(f"Error processing result {i+1}: {str(e)}")
                        
                        # Try to navigate back
                        try:
                            self.driver.back()
                            time.sleep(2)
                        except:
                            # Re-search if back doesn't work
                            self.driver.get("https://www.einforma.pt/")
                            time.sleep(3)
                            return self.search_einforma_selenium_improved(company_name)
                        continue
                
                return None
                
            except Exception as e:
                self.logger.error(f"Error processing search results: {str(e)}")
                return None
                
        except Exception as e:
            self.logger.error(f"Error in Selenium search: {str(e)}")
            return None

    def parse_einforma_search_results(self, soup, original_company_name):
        """Parse company search results from eInforma.pt"""
        try:
            companies_found = []
            
            # Strategy 1: Look for direct NIF links
            nif_links = soup.find_all('a', href=re.compile(r'nif/[0-9]{9}'))
            
            for link in nif_links:
                href = link.get('href', '')
                company_text = link.get_text(strip=True)
                
                nif_match = re.search(r'nif/([0-9]{9})', href)
                if nif_match:
                    nif = nif_match.group(1)
                    companies_found.append({
                        'company_name': company_text,
                        'nif': nif,
                        'einforma_url': urljoin('https://www.einforma.pt', href),
                        'similarity': self.similarity(company_text, original_company_name)
                    })
            
            # Strategy 2: Look for company names with legal suffixes
            legal_suffixes = ['LDA', 'LIMITADA', 'S.A.', 'SA', 'UNIPESSOAL']
            
            # Find all text containing company suffixes
            for suffix in legal_suffixes:
                suffix_elements = soup.find_all(string=re.compile(rf'\b{suffix}\b', re.IGNORECASE))
                
                for element in suffix_elements:
                    parent = element.parent
                    if parent:
                        # Look for NIF in nearby text
                        context_text = parent.get_text()
                        nif = self.extract_nif_from_text(context_text)
                        
                        if nif:
                            # Extract company name
                            lines = context_text.split('\n')
                            for line in lines:
                                if suffix in line.upper() and len(line.strip()) > 5:
                                    clean_name = re.sub(r'\b[0-9]{9}\b', '', line).strip()
                                    if clean_name:
                                        companies_found.append({
                                            'company_name': clean_name,
                                            'nif': nif,
                                            'einforma_url': f"https://www.einforma.pt/servlet/app/portal/ENTP/prod/ETIQUETA_EMPRESA_CONTRIBUINTE/nif/{nif}/contribuinte/{nif}",
                                            'similarity': self.similarity(clean_name, original_company_name)
                                        })
                                        break
            
            # Remove duplicates by NIF
            seen_nifs = set()
            unique_companies = []
            for company in companies_found:
                if company['nif'] not in seen_nifs:
                    seen_nifs.add(company['nif'])
                    unique_companies.append(company)
            
            # Sort by similarity
            if unique_companies:
                unique_companies.sort(key=lambda x: x['similarity'], reverse=True)
                self.logger.info(f"Found {len(unique_companies)} unique companies")
                return unique_companies
            
            return None
            
        except Exception as e:
            self.logger.error(f"Error parsing search results: {str(e)}")
            return None

    def search_alternative_sources(self, company_name):
        """Search alternative Portuguese business databases"""
        results = []
        
        # Try Racius.com (Portuguese business database)
        try:
            self.logger.info(f"Searching Racius.com for: {company_name}")
            
            search_url = f"https://www.racius.com/pesquisa/{quote(company_name)}"
            response = self.session.get(search_url, timeout=15)
            
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')
                
                # Look for company cards or results
                company_elements = soup.find_all(['div', 'a'], class_=re.compile(r'company|empresa|result'))
                
                for element in company_elements:
                    text = element.get_text()
                    nif = self.extract_nif_from_text(text)
                    
                    if nif and any(suffix in text.upper() for suffix in ['LDA', 'SA', 'LIMITADA']):
                        results.append({
                            'company_name': text.strip(),
                            'nif': nif,
                            'source': 'racius.com',
                            'similarity': self.similarity(text, company_name)
                        })
                        
        except Exception as e:
            self.logger.warning(f"Racius search failed: {str(e)}")
        
        return results

    def process_company_improved(self, company_name, company_url=None):
        """Process a single company with multiple search strategies"""
        result = {
            'portfolio_company': company_name,
            'company_url': company_url or '',
            'legal_name': '',
            'nif': '',
            'source': '',
            'search_method': '',
            'einforma_url': '',
            'similarity_score': 0.0,
            'status': ''
        }
        
        try:
            self.logger.info(f"Processing company: {company_name}")
            
            # Strategy 1: Direct URL approach
            search_results = self.search_einforma_direct_url(company_name)
            if search_results:
                result.update(self.format_result(search_results[0], 'einforma_direct'))
                return result
            
            # Strategy 2: Requests-based search
            search_results = self.search_einforma_with_requests(company_name)
            if search_results:
                result.update(self.format_result(search_results[0], 'einforma_requests'))
                return result
            
            # Strategy 3: Selenium search (last resort)
            search_results = self.search_einforma_selenium_improved(company_name)
            if search_results:
                result.update(self.format_result(search_results[0], 'einforma_selenium'))
                return result
            
            # Strategy 4: Alternative sources
            alt_results = self.search_alternative_sources(company_name)
            if alt_results:
                best_alt = max(alt_results, key=lambda x: x['similarity'])
                result.update({
                    'legal_name': best_alt['company_name'],
                    'nif': best_alt['nif'],
                    'source': best_alt['source'],
                    'search_method': 'alternative',
                    'similarity_score': best_alt['similarity'],
                    'status': 'Found (Alternative)'
                })
                return result
            
            # Strategy 5: Website extraction if URL provided
            if company_url:
                try:
                    response = self.session.get(company_url, timeout=10)
                    if response.status_code == 200:
                        nif = self.extract_nif_from_text(response.text)
                        if nif:
                            result.update({
                                'legal_name': company_name,
                                'nif': nif,
                                'source': 'company_website',
                                'search_method': 'website_extraction',
                                'similarity_score': 1.0,
                                'status': 'Found (Website)'
                            })
                            return result
                except Exception as e:
                    self.logger.warning(f"Website extraction failed for {company_url}: {str(e)}")
            
            # No results found
            result.update({
                'legal_name': company_name,
                'status': 'Not Found'
            })
            
        except Exception as e:
            self.logger.error(f"Error processing company {company_name}: {str(e)}")
            result.update({
                'legal_name': company_name,
                'status': 'Error'
            })
        
        return result
    
    def format_result(self, search_result, method):
        """Format search result for final output"""
        return {
            'legal_name': search_result['company_name'],
            'nif': search_result['nif'],
            'source': 'eInforma.pt',
            'search_method': method,
            'einforma_url': search_result['einforma_url'],
            'similarity_score': search_result['similarity'],
            'status': 'Found'
        }

    def process_portfolio_companies(self, companies_data):
        """Process a list of portfolio companies"""
        self.results = []
        
        for i, company_data in enumerate(companies_data):
            if isinstance(company_data, dict):
                company_name = company_data.get('name', '')
                company_url = company_data.get('url', '')
            else:
                company_name = str(company_data)
                company_url = ''
            
            if company_name:
                self.logger.info(f"Processing {i+1}/{len(companies_data)}: {company_name}")
                result = self.process_company_improved(company_name, company_url)
                self.results.append(result)
                
                # Respectful delay
                time.sleep(2)
        
        return self.results

    def save_results_to_csv(self, filename='portuguese_companies_fixed.csv'):
        """Save results to CSV file"""
        if self.results:
            df = pd.DataFrame(self.results)
            df.to_csv(filename, index=False)
            self.logger.info(f"Results saved to {filename}")
            return filename
        return None

    def get_results_summary(self):
        """Get a summary of extraction results"""
        if not self.results:
            return "No results available"
        
        total = len(self.results)
        with_nif = len([r for r in self.results if r['nif']])
        found_status = len([r for r in self.results if 'Found' in r['status']])
        
        return f"""
        Total companies processed: {total}
        Companies with data found: {found_status} ({found_status/total*100:.1f}%)
        Companies with NIF extracted: {with_nif} ({with_nif/total*100:.1f}%)
        Success rate: {found_status/total*100:.1f}%
        """

    def cleanup(self):
        """Clean up resources"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None

# Example usage
if __name__ == "__main__":
    # Test companies
    test_companies = [
        {'name': 'Addvolt', 'url': ''},
        {'name': 'Gosimac', 'url': ''},
        {'name': 'SKYPRO', 'url': ''},
    ]
    
    extractor = PortugueseCompanyExtractorFixed()
    
    try:
        # Process companies
        results = extractor.process_portfolio_companies(test_companies)
        
        # Print results
        for result in results:
            print(f"Company: {result['portfolio_company']}")
            print(f"Legal Name: {result['legal_name']}")
            print(f"NIF: {result['nif']}")
            print(f"Status: {result['status']}")
            print(f"Method: {result['search_method']}")
            print("-" * 50)
        
        # Save results
        extractor.save_results_to_csv()
        print(extractor.get_results_summary())
        
    finally:
        extractor.cleanup()