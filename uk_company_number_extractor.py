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

class UKCompanyNumberExtractor:
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

        # UK company number patterns - 8 digits, may start with 0 or have prefixes
        self.company_number_patterns = [
            r'Company\s+number[\s:]*([0-9]{8})',  # Direct company number match
            r'Company\s+No[\s.:]*([0-9]{8})',    # Company No. format
            r'Registration\s+number[\s:]*([0-9]{8})',  # Registration number
            r'Registered\s+number[\s:]*([0-9]{8})',    # Registered number
            r'([0-9]{8})(?=\s*(?:Company|Registration|Registered))',  # Number before keywords
            r'(?:SC|OC|SO)([0-9]{6})',  # Scottish/LLP prefixed numbers
        ]

    def similarity(self, a, b):
        """Calculate similarity between two strings"""
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    def extract_company_number_from_html(self, html_content):
        """Extract UK company number from HTML content"""
        for pattern in self.company_number_patterns:
            matches = re.finditer(pattern, html_content, re.IGNORECASE)
            for match in matches:
                code = match.group(1) if len(match.groups()) > 0 else match.group(0)
                clean_code = re.sub(r'[^0-9]', '', code)

                # UK company numbers are typically 8 digits
                if len(clean_code) == 8 and clean_code.isdigit():
                    return clean_code
                elif len(clean_code) == 6 and clean_code.isdigit():
                    # Handle Scottish/LLP format - add prefix back
                    if 'SC' in match.group(0).upper():
                        return f"SC{clean_code}"
                    elif 'OC' in match.group(0).upper():
                        return f"OC{clean_code}"
                    elif 'SO' in match.group(0).upper():
                        return f"SO{clean_code}"
        return None

    def search_companies_house_by_name(self, company_name):
        """Search Companies House for company by name and return company number"""
        try:
            # Clean company name for search
            clean_name = re.sub(r'[^\w\s]', '', company_name).strip()

            # Try direct web scraping of Companies House search
            search_url = "https://find-and-update.company-information.service.gov.uk/search/companies"
            params = {'q': clean_name}

            response = self.session.get(search_url, params=params, timeout=10)

            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')

                # Look for company links in search results
                company_links = soup.find_all('a', href=re.compile(r'/company/[0-9A-Z]{6,8}'))

                for link in company_links:
                    company_text = link.get_text(strip=True)
                    href = link.get('href', '')

                    # Extract company number from URL
                    match = re.search(r'/company/([0-9A-Z]{6,8})', href)
                    if match:
                        company_number = match.group(1)

                        # Check if this company name matches our search
                        if self.similarity(company_text.lower(), company_name.lower()) > 0.6:
                            return company_number

        except Exception as e:
            self.logger.error(f"Error searching Companies House for {company_name}: {str(e)}")

        return None

    def get_company_details_from_companies_house(self, company_number):
        """Get company details from Companies House company page"""
        try:
            url = f"https://find-and-update.company-information.service.gov.uk/company/{company_number}"
            response = self.session.get(url, timeout=10)

            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')

                # Extract company name
                company_name_elem = soup.find('h1', class_='heading-xlarge') or soup.find('h1')
                company_name = company_name_elem.get_text(strip=True) if company_name_elem else ""

                # Extract company status
                status_elem = soup.find('span', {'id': 'company-status'}) or soup.find('span', class_='status')
                status = status_elem.get_text(strip=True) if status_elem else ""

                # Extract company type
                type_elem = soup.find('dd', {'id': 'company-type'}) or soup.find(text=re.compile('Company type'))
                if type_elem and hasattr(type_elem, 'parent'):
                    type_text = type_elem.parent.find_next('dd').get_text(strip=True) if type_elem.parent.find_next('dd') else ""
                else:
                    type_text = ""

                return {
                    'company_number': company_number,
                    'company_name': company_name,
                    'status': status,
                    'company_type': type_text,
                    'source_url': url
                }

        except Exception as e:
            self.logger.error(f"Error getting details for company {company_number}: {str(e)}")

        return None

    def extract_legal_name_structured_approach(self, html_content, company_name):
        """Extract legal name using UK company suffixes"""
        soup = BeautifulSoup(html_content, 'html.parser')
        found_names = []

        # UK legal suffixes - comprehensive list
        uk_suffixes = [
            r'Ltd\.?', r'Limited', r'PLC', r'Plc', r'plc',
            r'LLP', r'LP', r'CIC', r'CIO', 
            r'Public Limited Company',
            r'Limited Liability Partnership',
            r'Community Interest Company',
            r'Charitable Incorporated Organisation',
            # Welsh equivalents
            r'Cyf\.?', r'Cyfyngedig',
            # Scottish variations
            r'Scottish Limited Partnership',
            # Other variations
            r'L\.L\.P\.?', r'P\.L\.C\.?', r'L\.T\.D\.?',
            r'\(UK\)', r'\(England\)', r'\(Wales\)', r'\(Scotland\)',
        ]

        suffix_pattern = '|'.join(uk_suffixes)

        # 1. Extract from title tags
        title_tags = soup.find_all(['title', 'h1', 'h2', 'h3'])
        for tag in title_tags:
            text = tag.get_text(strip=True)
            if re.search(rf'\b({suffix_pattern})\b', text, re.IGNORECASE):
                found_names.append(('title', text.strip()))

        # 2. Look for company information sections
        info_sections = soup.find_all(['div', 'section', 'p'], 
                                    class_=re.compile(r'company|corporate|legal|entity', re.I))
        for section in info_sections:
            text = section.get_text(strip=True)
            if re.search(rf'\b({suffix_pattern})\b', text, re.IGNORECASE):
                sentences = re.split(r'[.!?]', text)
                for sentence in sentences:
                    if re.search(rf'\b({suffix_pattern})\b', sentence, re.IGNORECASE):
                        found_names.append(('info_section', sentence.strip()))

        # 3. Search for patterns with company keywords
        company_patterns = [
            rf'([A-Z][A-Za-z\s&\-\.]+(?:{suffix_pattern}))\b',
            rf'\b([A-Z][A-Za-z\s&\-\.]*\s+(?:{suffix_pattern}))\b',
            rf'([A-Za-z\s&\-\.]+\s+(?:{suffix_pattern}))(?=\s|$|[.,;])',
        ]

        for pattern in company_patterns:
            matches = re.finditer(pattern, html_content, re.IGNORECASE | re.MULTILINE)
            for match in matches:
                candidate = match.group(1).strip()
                if len(candidate) > 3 and self.similarity(candidate, company_name) > 0.3:
                    found_names.append(('pattern_match', candidate))

        # 4. Rank and select best match
        if found_names:
            scored_names = []
            for source, name in found_names:
                score = self.similarity(name, company_name)
                # Boost score for names with UK suffixes
                if re.search(rf'\b({suffix_pattern})\b', name, re.IGNORECASE):
                    score += 0.2
                scored_names.append((score, source, name))

            scored_names.sort(reverse=True)
            return scored_names[0][2] if scored_names else company_name

        return company_name

    def process_company_with_companies_house(self, company_name, company_url=None):
        """Process a single company using Companies House search"""
        result = {
            'portfolio_company': company_name,
            'company_url': company_url or '',
            'legal_name': '',
            'company_number': '',
            'status': '',
            'company_type': '',
            'source': 'Companies House',
            'search_method': 'companies_house_web',
            'companies_house_url': ''
        }

        try:
            self.logger.info(f"Processing company: {company_name}")

            # First, try to search Companies House by company name
            company_number = self.search_companies_house_by_name(company_name)

            if company_number:
                self.logger.info(f"Found company number: {company_number}")

                # Get detailed information from Companies House
                company_details = self.get_company_details_from_companies_house(company_number)

                if company_details:
                    result.update({
                        'legal_name': company_details['company_name'],
                        'company_number': company_details['company_number'],
                        'status': company_details['status'],
                        'company_type': company_details['company_type'],
                        'companies_house_url': company_details['source_url']
                    })

                    self.logger.info(f"Successfully extracted company details for {company_name}")
                else:
                    result['legal_name'] = company_name
                    result['company_number'] = company_number
                    result['companies_house_url'] = f"https://find-and-update.company-information.service.gov.uk/company/{company_number}"
            else:
                self.logger.warning(f"Could not find company number for {company_name}")
                result['legal_name'] = company_name

            # If we have the company URL, try to extract additional info from their website
            if company_url:
                try:
                    response = self.session.get(company_url, timeout=10)
                    if response.status_code == 200:
                        # Try to extract company number from company website
                        website_company_number = self.extract_company_number_from_html(response.text)
                        if website_company_number:
                            result['company_number'] = website_company_number
                            result['search_method'] = 'website_extraction'

                        # Extract legal name from website using UK suffixes
                        legal_name = self.extract_legal_name_structured_approach(response.text, company_name)
                        if legal_name and legal_name != company_name:
                            result['legal_name'] = legal_name

                except Exception as e:
                    self.logger.error(f"Error processing company website {company_url}: {str(e)}")

        except Exception as e:
            self.logger.error(f"Error processing company {company_name}: {str(e)}")
            result['legal_name'] = company_name

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
                result = self.process_company_with_companies_house(company_name, company_url)
                self.results.append(result)

                # Add delay to be respectful to Companies House
                time.sleep(1)

        return self.results

    def save_results_to_csv(self, filename='uk_company_extraction_results.csv'):
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
        with_company_number = len([r for r in self.results if r['company_number']])
        with_legal_name = len([r for r in self.results if r['legal_name'] and r['legal_name'] != r['portfolio_company']])

        return f"""
        Total companies processed: {total}
        Companies with company number found: {with_company_number} ({with_company_number/total*100:.1f}%)
        Companies with legal name extracted: {with_legal_name} ({with_legal_name/total*100:.1f}%)
        """

# Example usage
if __name__ == "__main__":
    # Example portfolio companies data
    example_companies = pd.read_excel('uk_companies.xlsx').to_dict(orient='records')

    extractor = UKCompanyNumberExtractor()
    results = extractor.process_portfolio_companies(example_companies)

    # Print results
    for result in results:
        print(f"Company: {result['portfolio_company']}")
        print(f"Legal Name: {result['legal_name']}")
        print(f"Company Number: {result['company_number']}")
        print(f"Status: {result['status']}")
        print(f"Companies House URL: {result['companies_house_url']}")
        print("-" * 50)

    # Save to CSV
    extractor.save_results_to_csv()

    # Print summary
    print(extractor.get_results_summary())
