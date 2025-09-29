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

class AustrianCompanyExtractor:
    def __init__(self):
        self.results = []
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept-Language': 'de-AT,de;q=0.9,en;q=0.8'
        })
        self.driver = None

        # Setup logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)

        # Austrian VAT number patterns - ATU followed by 8 digits
        self.vat_patterns = [
            r'ATU\s*([0-9]{8})',  # ATU12345678
            r'VAT\s*ID[\s:\-]*ATU\s*([0-9]{8})',  # VAT ID: ATU12345678
            r'VAT\s*Number[\s:\-]*ATU\s*([0-9]{8})',  # VAT Number: ATU12345678
            r'Umsatzsteuer[\-\s]*(?:ID|nummer)[\s:\-]*ATU\s*([0-9]{8})',  # German: Umsatzsteuer-ID: ATU12345678
            r'UID[\s:\-]*ATU\s*([0-9]{8})',  # UID: ATU12345678
            r'Mehrwertsteuer[\-\s]*(?:ID|nummer)[\s:\-]*ATU\s*([0-9]{8})',  # Mehrwertsteuer-ID: ATU12345678
            r'FN\s*([0-9]{6}[a-z])',  # Commercial register number format: FN123456a
        ]

        # Pages to search for legal information
        self.legal_page_patterns = [
            'impressum', 'imprint', 'legal', 'legal-information', 'legal-info',
            'kontakt', 'contact', 'about', 'über-uns', 'ueber-uns',
            'datenschutz', 'privacy', 'terms', 'agb', 'rechtliches',
            'site-notice', 'disclaimer', 'rechtliche-hinweise'
        ]

    def similarity(self, a, b):
        """Calculate similarity between two strings"""
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    def extract_vat_from_text(self, text):
        """Extract Austrian VAT number from text"""
        for pattern in self.vat_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                vat_number = match.group(1)

                # Validate Austrian VAT format
                if pattern.startswith('ATU') and len(vat_number) == 8 and vat_number.isdigit():
                    return f"ATU{vat_number}"
                elif pattern.startswith('FN') and len(vat_number) == 7 and vat_number[:-1].isdigit() and vat_number[-1].isalpha():
                    return f"FN{vat_number}"

        return None

    def find_legal_pages(self, base_url):
        """Find potential legal information pages on a website"""
        legal_urls = []

        try:
            # First, try to get the homepage to look for legal page links
            response = self.session.get(base_url, timeout=10)
            if response.status_code != 200:
                return legal_urls

            soup = BeautifulSoup(response.content, 'html.parser')

            # Look for links that might contain legal information
            all_links = soup.find_all('a', href=True)

            for link in all_links:
                href = link.get('href', '').lower()
                link_text = link.get_text(strip=True).lower()

                # Check if the link or text matches legal page patterns
                for pattern in self.legal_page_patterns:
                    if pattern in href or pattern in link_text:
                        full_url = urljoin(base_url, link.get('href'))
                        if full_url not in legal_urls:
                            legal_urls.append(full_url)

            # Also try common legal page URLs directly
            common_legal_paths = [
                '/impressum', '/imprint', '/legal', '/legal-information',
                '/kontakt', '/contact', '/about', '/über-uns',
                '/datenschutz', '/privacy', '/terms', '/agb',
                '/en/legal-information', '/de/impressum', '/rechtliches'
            ]

            for path in common_legal_paths:
                potential_url = urljoin(base_url, path)
                if potential_url not in legal_urls:
                    legal_urls.append(potential_url)

        except Exception as e:
            self.logger.error(f"Error finding legal pages for {base_url}: {str(e)}")

        return legal_urls[:10]  # Limit to first 10 URLs to avoid overloading

    def extract_vat_from_page(self, url):
        """Extract VAT information from a specific page"""
        try:
            self.logger.info(f"Searching for VAT info on: {url}")

            response = self.session.get(url, timeout=10)
            if response.status_code != 200:
                return None

            # Try both raw text and parsed HTML
            vat_from_text = self.extract_vat_from_text(response.text)
            if vat_from_text:
                return {
                    'vat_number': vat_from_text,
                    'source_url': url,
                    'extraction_method': 'text_pattern'
                }

            # Parse HTML and look in specific sections
            soup = BeautifulSoup(response.content, 'html.parser')

            # Look for VAT in specific HTML structures
            vat_info = self.extract_vat_from_html_structure(soup, url)
            if vat_info:
                return vat_info

        except Exception as e:
            self.logger.error(f"Error extracting VAT from page {url}: {str(e)}")

        return None

    def extract_vat_from_html_structure(self, soup, source_url):
        """Extract VAT from structured HTML elements"""
        try:
            # Look for VAT in common HTML structures
            vat_selectors = [
                # Look for elements containing VAT-related keywords
                '[data-vat]', '[data-atu]', '[data-uid]',
                '.vat-number', '.atu-number', '.uid-number',
                '#vat-number', '#atu-number', '#uid-number',
                '.legal-info', '.company-info', '.impressum',
                # German-specific selectors
                '.umsatzsteuer', '.mehrwertsteuer', '.firmeninfo'
            ]

            for selector in vat_selectors:
                elements = soup.select(selector)
                for element in elements:
                    element_text = element.get_text(strip=True)
                    vat_number = self.extract_vat_from_text(element_text)
                    if vat_number:
                        return {
                            'vat_number': vat_number,
                            'source_url': source_url,
                            'extraction_method': 'html_structure'
                        }

            # Look for VAT in paragraph and div elements containing keywords
            keywords = ['vat', 'atu', 'uid', 'umsatzsteuer', 'mehrwertsteuer']
            for keyword in keywords:
                elements = soup.find_all(['p', 'div', 'span'], string=re.compile(keyword, re.IGNORECASE))
                for element in elements:
                    # Check the element and its parent
                    for check_element in [element, element.parent]:
                        if check_element:
                            element_text = check_element.get_text(strip=True)
                            vat_number = self.extract_vat_from_text(element_text)
                            if vat_number:
                                return {
                                    'vat_number': vat_number,
                                    'source_url': source_url,
                                    'extraction_method': 'keyword_search'
                                }

            # Look for table rows that might contain VAT information
            table_rows = soup.find_all('tr')
            for row in table_rows:
                row_text = row.get_text(strip=True)
                if any(keyword in row_text.lower() for keyword in ['vat', 'atu', 'uid', 'umsatzsteuer']):
                    vat_number = self.extract_vat_from_text(row_text)
                    if vat_number:
                        return {
                            'vat_number': vat_number,
                            'source_url': source_url,
                            'extraction_method': 'table_extraction'
                        }

        except Exception as e:
            self.logger.error(f"Error extracting VAT from HTML structure: {str(e)}")

        return None

    def extract_legal_name_with_austrian_suffixes(self, html_content, company_name):
        """Extract legal name using Austrian company suffixes"""
        soup = BeautifulSoup(html_content, 'html.parser')
        found_names = []

        # Austrian legal suffixes
        austrian_suffixes = [
            r'GmbH', r'AG', r'GesmbH',  # Standard Austrian forms
            r'Gesellschaft mit beschränkter Haftung',
            r'Aktiengesellschaft',
            r'Limited', r'Ltd', r'Ltée',  # International forms
            r'KG', r'OG', r'KEG',  # Partnership forms
            # Variations with punctuation
            r'G\.m\.b\.H\.', r'A\.G\.',
        ]

        suffix_pattern = '|'.join(austrian_suffixes)

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

    def process_company_website(self, company_name, company_url):
        """Process a single company website for VAT extraction"""
        result = {
            'portfolio_company': company_name,
            'company_url': company_url,
            'legal_name': company_name,
            'vat_number': '',
            'commercial_register': '',
            'source_url': '',
            'extraction_method': '',
            'pages_searched': 0,
            'status': 'Not Found'
        }

        if not company_url:
            result['status'] = 'No URL Provided'
            return result

        try:
            self.logger.info(f"Processing Austrian company: {company_name} - {company_url}")

            # Find potential legal information pages
            legal_pages = self.find_legal_pages(company_url)
            result['pages_searched'] = len(legal_pages)

            # Search each legal page for VAT information
            vat_info_found = None
            for page_url in legal_pages:
                vat_info = self.extract_vat_from_page(page_url)
                if vat_info:
                    vat_info_found = vat_info
                    break

                # Add small delay between requests
                time.sleep(0.5)

            # If no VAT found in legal pages, try homepage
            if not vat_info_found:
                vat_info_found = self.extract_vat_from_page(company_url)

            if vat_info_found:
                result.update({
                    'vat_number': vat_info_found['vat_number'],
                    'source_url': vat_info_found['source_url'],
                    'extraction_method': vat_info_found['extraction_method'],
                    'status': 'Found'
                })

                # Check if it's a commercial register number
                if vat_info_found['vat_number'].startswith('FN'):
                    result['commercial_register'] = vat_info_found['vat_number']
                    result['vat_number'] = ''

                self.logger.info(f"Found VAT info for {company_name}: {vat_info_found['vat_number']}")

            # Try to extract legal company name
            try:
                response = self.session.get(company_url, timeout=10)
                if response.status_code == 200:
                    legal_name = self.extract_legal_name_with_austrian_suffixes(response.text, company_name)
                    if legal_name != company_name:
                        result['legal_name'] = legal_name
            except Exception as e:
                self.logger.error(f"Error extracting legal name from {company_url}: {str(e)}")

        except Exception as e:
            self.logger.error(f"Error processing company {company_name}: {str(e)}")
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
                result = self.process_company_website(company_name, company_url)
                self.results.append(result)

                # Add delay between companies
                time.sleep(1)

        return self.results

    def save_results_to_csv(self, filename='austrian_company_extraction_results.csv'):
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
        with_vat = len([r for r in self.results if r['vat_number']])
        with_commercial_reg = len([r for r in self.results if r['commercial_register']])
        found_status = len([r for r in self.results if r['status'] == 'Found'])
        with_legal_name = len([r for r in self.results if r['legal_name'] != r['portfolio_company']])

        avg_pages_searched = sum([r['pages_searched'] for r in self.results]) / total if total > 0 else 0

        return f"""
        Total companies processed: {total}
        Companies with VAT number found: {with_vat} ({with_vat/total*100:.1f}%)
        Companies with commercial register found: {with_commercial_reg} ({with_commercial_reg/total*100:.1f}%)
        Companies with legal name extracted: {with_legal_name} ({with_legal_name/total*100:.1f}%)
        Success rate: {found_status/total*100:.1f}%
        Average pages searched per company: {avg_pages_searched:.1f}
        """

# Example usage
if __name__ == "__main__":
    # Example Austrian portfolio companies data
    example_companies = pd.read_excel('au_companies.xlsx').to_dict(orient='records')

    extractor = AustrianCompanyExtractor()
    results = extractor.process_portfolio_companies(example_companies)

    # Print results
    for result in results:
        print(f"Company: {result['portfolio_company']}")
        print(f"Legal Name: {result['legal_name']}")
        print(f"VAT Number: {result['vat_number']}")
        print(f"Commercial Register: {result['commercial_register']}")
        print(f"Source URL: {result['source_url']}")
        print(f"Pages Searched: {result['pages_searched']}")
        print(f"Status: {result['status']}")
        print("-" * 50)

    # Save to CSV
    extractor.save_results_to_csv()

    # Print summary
    print(extractor.get_results_summary())
