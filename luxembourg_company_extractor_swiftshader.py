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

class LuxembourgCompanyExtractor:
    def __init__(self):
        self.results = []
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept-Language': 'en-US,en;q=0.9,fr;q=0.8,de;q=0.7',
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

        # Luxembourg VAT patterns - format LU + 8 digits
        self.vat_patterns = [
            r'LU\s*([0-9]{8})',
            r'L\.?U\.?\s*([0-9]{8})',
            r'VAT[\s:]*LU\s*([0-9]{8})',
            r'TVA[\s:]*LU\s*([0-9]{8})',  # French term for VAT
            r'product-list-LU([0-9]{8})',  # Specific kompass pattern
            r'LUR([0-9]{6})',  # Alternative pattern mentioned in the query
            r'\b(LU[0-9]{8})\b',
            r'\b([0-9]{8})\b(?=.*luxemb)',  # 8 digits followed by luxembourg mention
        ]

        # Registration number patterns - B + 6 digits (Luxembourg format)
        self.registration_patterns = [
            r'B\s*([0-9]{6})',  # B165823 format
            r'B([0-9]{6})',      # Direct B format
            r'\b(B[0-9]{6})\b',  # Word boundary B format
            r'Registration[\s]+No[\.:]*\s*B\s*([0-9]{6})',
            r'Registr[\w]*[\s]+[Nn]°[\s]*:?\s*B\s*([0-9]{6})',
            r'Numéro[\s]+d['']?enregistrement[\s]*:?\s*B\s*([0-9]{6})',
        ]

    def setup_driver(self, headless=True):
        """Setup Chrome driver with improved options including SwiftShader"""
        if self.driver:
            return

        options = Options()

        # Essential options
        if headless:
            options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        # GPU and rendering options with SwiftShader
        options.add_argument('--enable-unsafe-swiftshader')  # Added SwiftShader flag
        options.add_argument('--use-gl=swiftshader')
        options.add_argument('--disable-gpu-sandbox')
        options.add_argument('--disable-software-rasterizer')
        options.add_argument('--disable-background-timer-throttling')
        options.add_argument('--disable-backgrounding-occluded-windows')
        options.add_argument('--disable-renderer-backgrounding')

        # Window and display
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--start-maximized')

        # Anti-detection measures
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        # Performance optimizations
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-plugins')
        options.add_argument('--disable-images')
        options.add_argument('--disable-javascript')  # Can be removed if JS is needed
        options.add_argument('--disable-css')

        # Popup and notification handling
        options.add_argument('--disable-popup-blocking')
        options.add_argument('--disable-notifications')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-translate')

        # Memory and process optimization
        options.add_argument('--memory-pressure-off')
        options.add_argument('--max_old_space_size=4096')

        # Additional stability flags
        options.add_argument('--no-first-run')
        options.add_argument('--no-default-browser-check')
        options.add_argument('--disable-default-apps')
        options.add_argument('--disable-component-update')

        # User agent
        options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

        # Additional experimental options for better performance
        options.add_experimental_option('prefs', {
            'profile.default_content_setting_values.notifications': 2,
            'profile.default_content_settings.popups': 0,
            'profile.managed_default_content_settings.images': 2,
            'profile.default_content_setting_values.cookies': 1
        })

        try:
            self.driver = webdriver.Chrome(options=options)

            # Additional JavaScript to mask automation
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.driver.execute_script("Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]})")
            self.driver.execute_script("Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']})")

            self.wait = WebDriverWait(self.driver, 30)
            self.logger.info("Chrome driver initialized successfully with SwiftShader")
        except Exception as e:
            self.logger.error(f"Could not initialize Chrome driver: {str(e)}")
            raise

    def similarity(self, a, b):
        """Calculate similarity between two strings"""
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    def extract_vat_from_text(self, text):
        """Extract Luxembourg VAT number from text"""
        for pattern in self.vat_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                if len(match.groups()) > 0:
                    vat = match.group(1)
                    # Validate Luxembourg VAT format (8 digits)
                    if len(vat) == 8 and vat.isdigit():
                        return f"LU{vat}"
                elif match.group(0):
                    vat = match.group(0)
                    # Clean and validate
                    vat_clean = re.sub(r'[^0-9A-Z]', '', vat.upper())
                    if vat_clean.startswith('LU') and len(vat_clean) == 10:
                        return vat_clean
        return None

    def extract_registration_number_from_text(self, text):
        """Extract Luxembourg Registration Number (B format) from text"""
        for pattern in self.registration_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                if len(match.groups()) > 0:
                    reg_num = match.group(1)
                    # Validate Luxembourg Registration format (6 digits)
                    if len(reg_num) == 6 and reg_num.isdigit():
                        return f"B{reg_num}"
                elif match.group(0):
                    reg_num = match.group(0)
                    # Clean and validate
                    reg_clean = re.sub(r'[^0-9B]', '', reg_num.upper())
                    if reg_clean.startswith('B') and len(reg_clean) == 7:
                        return reg_clean
        return None

    def extract_registration_from_blockinterieur(self, soup):
        """Extract registration number specifically from blockInterieur section"""
        try:
            # Look for blockInterieur elements (case insensitive)
            block_selectors = [
                '[class*="blockInterieur"]',
                '[class*="blockinterieur"]',
                '[class*="block-interieur"]',
                '[id*="blockInterieur"]',
                '[id*="blockinterieur"]',
                '.blockInterieur',
                '#blockInterieur'
            ]

            for selector in block_selectors:
                block_elements = soup.select(selector)
                for block in block_elements:
                    # Look for table cells (td) within this block
                    tds = block.find_all('td')
                    for td in tds:
                        td_text = td.get_text(strip=True)
                        # Check if it matches B + 6 digits pattern exactly
                        if re.match(r'^B[0-9]{6}$', td_text):
                            self.logger.info(f"Found registration number in blockInterieur: {td_text}")
                            return td_text

            # Fallback: look for any td with B + 6 digits pattern in the entire page
            all_tds = soup.find_all('td')
            for td in all_tds:
                td_text = td.get_text(strip=True)
                if re.match(r'^B[0-9]{6}$', td_text):
                    self.logger.info(f"Found registration number in td: {td_text}")
                    return td_text

            # Additional fallback: look in table rows that might contain registration info
            rows = soup.find_all('tr')
            for row in rows:
                row_text = row.get_text()
                if any(keyword in row_text.lower() for keyword in ['registration', 'registr', 'immatriculation']):
                    # Look for B number in this row
                    b_match = re.search(r'\b(B[0-9]{6})\b', row_text)
                    if b_match:
                        self.logger.info(f"Found registration number in table row: {b_match.group(1)}")
                        return b_match.group(1)

            return None

        except Exception as e:
            self.logger.warning(f"Error extracting from blockInterieur: {str(e)}")
            return None

    def search_kompass_direct_url(self, company_name):
        """Try direct URL approach for Kompass search"""
        try:
            # Method 1: Direct search URL
            encoded_name = quote(company_name.encode('utf-8'))
            search_urls = [
                f"https://lu.kompass.com/searchCompanies?text={encoded_name}",
                f"https://lu.kompass.com/search/{encoded_name}",
                f"https://lu.kompass.com/company/{encoded_name}",
            ]

            for url in search_urls:
                try:
                    self.logger.info(f"Trying direct URL: {url}")
                    response = self.session.get(url, timeout=15)
                    if response.status_code == 200:
                        soup = BeautifulSoup(response.content, 'html.parser')
                        results = self.parse_kompass_search_results(soup, company_name)
                        if results:
                            return results
                except Exception as e:
                    self.logger.warning(f"Direct URL failed: {str(e)}")
                    continue

            return None
        except Exception as e:
            self.logger.error(f"Error in direct URL search: {str(e)}")
            return None

    def search_kompass_with_selenium(self, company_name):
        """Search using Selenium for dynamic content"""
        try:
            self.logger.info(f"Searching Kompass with Selenium for: {company_name}")

            # Setup driver if needed
            self.setup_driver(headless=False)

            # Navigate to Luxembourg Kompass
            self.driver.get("https://lu.kompass.com/")

            # Wait for page to load
            self.wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            time.sleep(3)

            # Accept cookies and handle popups
            self.handle_cookies_and_popups()

            # Find search input field
            search_input = self.find_search_input()

            if not search_input:
                self.logger.error("No search input found")
                return None

            # Perform search
            return self.perform_search_and_get_results(search_input, company_name)

        except Exception as e:
            self.logger.error(f"Error in Selenium search: {str(e)}")
            return None

    def handle_cookies_and_popups(self):
        """Accept cookies and close popups automatically - Enhanced version"""
        try:
            # Wait a bit for popups to appear
            time.sleep(3)

            self.logger.info("Starting cookie and popup handling...")

            # Enhanced cookie acceptance selectors
            cookie_selectors = [
                "button[id*='accept']",
                "button[class*='accept']",
                "button[id*='cookie']",
                "button[class*='cookie']",
                "button[id*='consent']",
                "button[class*='consent']",
                ".cookie-accept",
                "#cookie-accept",
                ".accept-cookies",
                "#accept-cookies",
                "[data-testid*='accept']",
                "[aria-label*='accept']",
                "[data-cy*='accept']",
                "button[type='button'][class*='btn']",
                ".btn-accept",
                "#btn-accept"
            ]

            # XPath selectors for text-based buttons (multiple languages)
            cookie_text_xpaths = [
                "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]",
                "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accepter')]",
                "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'akzeptieren')]",
                "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'ok')]",
                "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'agree')]",
                "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'd\'accord')]",
                "//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]",
                "//span[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]/parent::button"
            ]

            # Try CSS selectors first
            for selector in cookie_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        if element.is_displayed() and element.is_enabled():
                            self.safe_click_element(element)
                            self.logger.info(f"Cookies accepted via CSS selector: {selector}")
                            time.sleep(2)
                            break
                    if elements:
                        break
                except Exception as e:
                    self.logger.debug(f"CSS selector failed: {selector} - {str(e)}")
                    continue

            # Try XPath selectors for text-based buttons
            for xpath in cookie_text_xpaths:
                try:
                    elements = self.driver.find_elements(By.XPATH, xpath)
                    for element in elements:
                        if element.is_displayed() and element.is_enabled():
                            self.safe_click_element(element)
                            self.logger.info(f"Cookies accepted via XPath: {xpath}")
                            time.sleep(2)
                            return True
                except Exception as e:
                    self.logger.debug(f"XPath selector failed: {xpath} - {str(e)}")
                    continue

            # Handle other popups (close buttons)
            popup_close_selectors = [
                "button[class*='close']",
                "button[id*='close']",
                ".close-button",
                "#close-button",
                "[data-dismiss='modal']",
                ".modal-close",
                "button[aria-label*='close']",
                "button[title*='close']",
                "button[title*='fermer']",
                "button[title*='schließen']",
                ".popup-close",
                ".dialog-close",
                "button[class*='dismiss']",
                ".btn-close",
                "#btn-close",
                "[onclick*='close']"
            ]

            for selector in popup_close_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        if element.is_displayed() and element.is_enabled():
                            self.safe_click_element(element)
                            self.logger.info(f"Popup closed via: {selector}")
                            time.sleep(1)
                            break
                except Exception as e:
                    self.logger.debug(f"Close selector failed: {selector} - {str(e)}")
                    continue

            # Additional step: Press ESC key to close any remaining modals
            try:
                from selenium.webdriver.common.keys import Keys
                self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                self.logger.info("Pressed ESC to close modals")
                time.sleep(1)
            except Exception:
                pass

            self.logger.info("Cookie and popup handling completed")
            return True

        except Exception as e:
            self.logger.warning(f"Error handling cookies/popups: {str(e)}")
            return False

    def find_search_input(self):
        """Find search input field using multiple strategies"""
        search_strategies = [
            (By.CSS_SELECTOR, "input[name*='search']"),
            (By.CSS_SELECTOR, "input[name*='text']"),
            (By.CSS_SELECTOR, "input[type='search']"),
            (By.CSS_SELECTOR, "input[type='text']"),
            (By.CSS_SELECTOR, "input[placeholder*='company']"),
            (By.CSS_SELECTOR, "input[placeholder*='search']"),
            (By.CSS_SELECTOR, "input[placeholder*='entreprise']"),
            (By.CSS_SELECTOR, "input[placeholder*='société']"),
            (By.CSS_SELECTOR, ".search-input"),
            (By.ID, "searchInput"),
            (By.NAME, "searchCompanies"),
            (By.NAME, "q"),
            (By.CSS_SELECTOR, "[role='searchbox']")
        ]

        for by, selector in search_strategies:
            try:
                elements = self.driver.find_elements(by, selector)
                for element in elements:
                    if element.is_displayed() and element.is_enabled():
                        self.logger.info(f"Found search input with: {selector}")
                        return element
            except Exception as e:
                self.logger.debug(f"Search input strategy failed: {selector} - {str(e)}")
                continue

        return None

    def perform_search_and_get_results(self, search_input, company_name):
        """Perform search and get results"""
        try:
            # Scroll into view and focus
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", search_input)
            time.sleep(1)

            # Clear and enter company name
            search_input.click()
            search_input.clear()
            time.sleep(1)
            search_input.send_keys(company_name)
            time.sleep(2)

            # Submit search
            search_input.send_keys(Keys.RETURN)
            self.logger.info(f"Search submitted for: {company_name}")

            # Wait for results
            time.sleep(5)

            # Handle any popups that might appear after search
            self.handle_cookies_and_popups()

            # Find and process results
            return self.find_and_process_results(company_name)

        except Exception as e:
            self.logger.error(f"Error performing search: {str(e)}")
            return None

    def safe_click_element(self, element):
        """Safely click an element with multiple strategies"""
        try:
            # Method 1: Regular click
            element.click()
        except ElementClickInterceptedException:
            try:
                # Method 2: JavaScript click
                self.driver.execute_script("arguments[0].click();", element)
            except Exception:
                try:
                    # Method 3: Action chains
                    ActionChains(self.driver).move_to_element(element).click().perform()
                except Exception:
                    # Method 4: Force click with JavaScript
                    self.driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('click', {bubbles: true}));", element)

    def find_and_process_results(self, company_name):
        """Find company results and extract information"""
        try:
            # Look for company result links
            result_selectors = [
                "a[href*='/c/']",  # Kompass company page links
                "a[href*='product-list-']",
                ".company-result a",
                ".search-result a",
                "a[class*='company']",
                ".result-item a",
                "table tr a",  # Table row links
                ".list-group-item a"
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
                            if href and text and len(text) > 3:
                                similarity = self.similarity(text, company_name)
                                if similarity > 0.1:  # Basic relevance filter
                                    company_results.append({
                                        'element': element,
                                        'text': text,
                                        'href': href,
                                        'similarity': similarity
                                    })

                    if company_results:
                        break
                except Exception as e:
                    self.logger.debug(f"Result selector failed: {selector} - {str(e)}")
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
                    details = self.click_company_result_and_extract_info(result['element'], company_name)
                    if details and (details['vat'] or details['registration_no']):
                        self.logger.info(f"Successfully extracted info for: {details['company_name']}")
                        return [details]

                    # Navigate back for next attempt
                    if i < len(company_results) - 1:
                        self.driver.back()
                        time.sleep(3)
                        # Handle any popups after navigation
                        self.handle_cookies_and_popups()

                except Exception as e:
                    self.logger.warning(f"Error processing result {i+1}: {str(e)}")
                    continue

            return None

        except Exception as e:
            self.logger.error(f"Error finding and processing results: {str(e)}")
            return None

    def click_company_result_and_extract_info(self, result_element, company_name):
        """Click on a company result and extract all information"""
        try:
            # Get the link before clicking
            company_link = result_element.get_attribute('href')
            company_text = result_element.text.strip()

            self.logger.info(f"Attempting to click on: {company_text}")

            # Try to click the element
            try:
                # Scroll element into view
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", result_element)
                time.sleep(1)

                # Wait for element to be clickable
                clickable_element = self.wait.until(EC.element_to_be_clickable(result_element))
                clickable_element.click()

            except Exception as e:
                self.logger.warning(f"Click failed, navigating directly: {str(e)}")
                if company_link:
                    self.driver.get(company_link)
                else:
                    return None

            # Wait for company page to load
            time.sleep(4)
            self.wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

            # Handle any popups on the company page
            self.handle_cookies_and_popups()

            # Extract company details from the page
            return self.extract_info_from_company_page(company_name)

        except Exception as e:
            self.logger.error(f"Error clicking and extracting from result: {str(e)}")
            return None

    def extract_info_from_company_page(self, original_company_name):
        """Extract VAT and Registration information from the current company page"""
        try:
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            current_url = self.driver.current_url

            self.logger.info(f"Extracting information from: {current_url}")

            company_details = {
                'company_name': '',
                'legal_name': '',
                'vat': '',
                'registration_no': '',
                'kompass_url': current_url,
                'similarity': 0.0
            }

            # Strategy 1: Extract Registration Number from blockInterieur (priority)
            registration_no = self.extract_registration_from_blockinterieur(soup)
            if registration_no:
                company_details['registration_no'] = registration_no
                self.logger.info(f"Found registration number: {registration_no}")

            # Strategy 2: Extract Registration Number from general text
            if not company_details['registration_no']:
                page_text = soup.get_text()
                registration_no = self.extract_registration_number_from_text(page_text)
                if registration_no:
                    company_details['registration_no'] = registration_no
                    self.logger.info(f"Found registration number in text: {registration_no}")

            # Strategy 3: Look for VAT number (keeping original functionality)
            page_html = str(soup)
            product_list_match = re.search(r'product-list-LU([0-9]{8})', page_html)
            if product_list_match:
                vat_number = f"LU{product_list_match.group(1)}"
                company_details['vat'] = vat_number
                self.logger.info(f"Found VAT in product-list pattern: {vat_number}")

            # Strategy 4: Look for LUR pattern
            lur_match = re.search(r'LUR([0-9]{6})', page_html)
            if lur_match and not company_details['vat']:
                lur_number = lur_match.group(1)
                if len(lur_number) == 6:
                    vat_number = f"LU{lur_number:0>8}"
                else:
                    vat_number = f"LU{lur_number}"
                company_details['vat'] = vat_number
                self.logger.info(f"Found VAT in LUR pattern: {vat_number}")

            # Strategy 5: Extract VAT from page content using patterns
            if not company_details['vat']:
                page_text = soup.get_text()
                vat = self.extract_vat_from_text(page_text)
                if vat:
                    company_details['vat'] = vat
                    self.logger.info(f"Found VAT in page content: {vat}")

            # Strategy 6: Find company name in page title or headers
            title_selectors = [
                'h1',
                'h2', 
                '.company-name',
                '.title',
                '[itemprop="name"]',
                '.company-title'
            ]

            for selector in title_selectors:
                try:
                    elements = soup.select(selector)
                    for element in elements:
                        text = element.get_text(strip=True)
                        # Check if this looks like a company name
                        if len(text) > 3 and not text.isdigit():
                            # Clean the company name
                            clean_name = re.sub(r'\bLU[0-9]{8}\b', '', text).strip()
                            clean_name = re.sub(r'\bB[0-9]{6}\b', '', clean_name).strip()
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

            # Ensure we have at least a company name
            if not company_details['company_name']:
                company_details['company_name'] = original_company_name
                company_details['legal_name'] = original_company_name
                company_details['similarity'] = 1.0

            return company_details

        except Exception as e:
            self.logger.error(f"Error extracting company details: {str(e)}")
            return None

    def parse_kompass_search_results(self, soup, original_company_name):
        """Parse company search results from lu.kompass.com"""
        try:
            companies_found = []

            # Look for company links and patterns
            company_links = soup.find_all('a', href=re.compile(r'/c/'))

            for link in company_links:
                href = link.get('href', '')
                company_text = link.get_text(strip=True)

                if company_text and len(company_text) > 3:
                    # Try to find VAT and Registration in the link or surrounding context
                    parent = link.parent
                    context_text = ""
                    if parent:
                        context_text = parent.get_text()

                    vat = self.extract_vat_from_text(context_text)
                    registration_no = self.extract_registration_number_from_text(context_text)

                    companies_found.append({
                        'company_name': company_text,
                        'vat': vat or '',
                        'registration_no': registration_no or '',
                        'kompass_url': urljoin('https://lu.kompass.com', href),
                        'similarity': self.similarity(company_text, original_company_name)
                    })

            # Remove duplicates and sort by similarity
            if companies_found:
                seen_companies = set()
                unique_companies = []
                for company in companies_found:
                    key = company['company_name'].lower()
                    if key not in seen_companies:
                        seen_companies.add(key)
                        unique_companies.append(company)

                unique_companies.sort(key=lambda x: x['similarity'], reverse=True)
                self.logger.info(f"Found {len(unique_companies)} unique companies")
                return unique_companies

            return None

        except Exception as e:
            self.logger.error(f"Error parsing search results: {str(e)}")
            return None

    def process_company(self, company_name, company_url=None):
        """Process a single company to find VAT and Registration information"""
        result = {
            'portfolio_company': company_name,
            'company_url': company_url or '',
            'legal_name': '',
            'vat': '',
            'registration_no': '',
            'source': '',
            'search_method': '',
            'kompass_url': '',
            'similarity_score': 0.0,
            'status': ''
        }

        try:
            self.logger.info(f"Processing company: {company_name}")

            # Strategy 1: Direct URL approach
            search_results = self.search_kompass_direct_url(company_name)
            if search_results and (search_results[0].get('vat') or search_results[0].get('registration_no')):
                result.update(self.format_result(search_results[0], 'kompass_direct'))
                return result

            # Strategy 2: Selenium search
            search_results = self.search_kompass_with_selenium(company_name)
            if search_results and (search_results[0].get('vat') or search_results[0].get('registration_no')):
                result.update(self.format_result(search_results[0], 'kompass_selenium'))
                return result

            # No information found
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
            'vat': search_result.get('vat', ''),
            'registration_no': search_result.get('registration_no', ''),
            'source': 'lu.kompass.com',
            'search_method': method,
            'kompass_url': search_result['kompass_url'],
            'similarity_score': search_result['similarity'],
            'status': 'Found' if (search_result.get('vat') or search_result.get('registration_no')) else 'Not Found'
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
                result = self.process_company(company_name, company_url)
                self.results.append(result)

                # Respectful delay
                time.sleep(3)

        return self.results

    def save_results_to_csv(self, filename='luxembourg_companies_info.csv'):
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
        with_vat = len([r for r in self.results if r['vat']])
        with_registration = len([r for r in self.results if r['registration_no']])
        found_status = len([r for r in self.results if 'Found' in r['status']])

        return f"""
Total companies processed: {total}
Companies with VAT found: {with_vat} ({with_vat/total*100:.1f}%)
Companies with Registration No. found: {with_registration} ({with_registration/total*100:.1f}%)
Companies with data found: {found_status} ({found_status/total*100:.1f}%)
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
    # Test companies (Luxembourg examples)
    test_companies = pd.read_excel('lux_companies.xlsx').to_dict(orient='records')

    extractor = LuxembourgCompanyExtractor()
    try:
        # Process companies
        results = extractor.process_portfolio_companies(test_companies)

        # Print results
        for result in results:
            print(f"Company: {result['portfolio_company']}")
            print(f"Legal Name: {result['legal_name']}")
            print(f"VAT: {result['vat']}")
            print(f"Registration No.: {result['registration_no']}")
            print(f"Status: {result['status']}")
            print(f"Method: {result['search_method']}")
            print("-" * 50)

        # Save results
        extractor.save_results_to_csv()
        print(extractor.get_results_summary())

    finally:
        extractor.cleanup()
