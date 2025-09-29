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

class ItalyVATExtractor:
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
        
        # Italy-specific VAT/Tax code patterns
        self.vat_patterns = [
            r'P\.?\s*IVA[\s#:]*([0-9]{11})',  # P.IVA with prefix
            r'Partita\s+IVA[\s#:]*([0-9]{11})',  # Partita IVA
            r'IVA[\s#:]*([0-9]{11})',  # IVA with prefix
            r'VAT[\s#:]*IT([0-9]{11})',  # VAT with IT prefix
            r'Codice\s+Fiscale[\s#:]*([0-9]{11})',  # Codice Fiscale (company)
            r'C\.?\s*F\.?[\s#:]*([0-9]{11})',  # C.F. abbreviation
            r'IT[\s]*([0-9]{11})',  # IT prefix for VAT
            r'P\.?\s*IVA\s*:?\s*([0-9]{2}[\s\-]?[0-9]{3}[\s\-]?[0-9]{3}[\s\-]?[0-9]{3})',  # Formatted P.IVA
        ]

    def similarity(self, a, b):
        """Calculate similarity between two strings"""
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    def extract_vat_from_html(self, html_content):
        """Extract VAT/Tax code from HTML content"""
        for pattern in self.vat_patterns:
            matches = re.finditer(pattern, html_content, re.IGNORECASE)
            for match in matches:
                code = match.group(1) if len(match.groups()) > 0 else match.group(0)
                clean_code = re.sub(r'[^0-9]', '', code)
                if len(clean_code) == 11 and clean_code.isdigit():
                    return clean_code
        return None

    def extract_legal_name_italian_approach(self, html_content, company_name):
        """Extract legal name using Italian website structure"""
        soup = BeautifulSoup(html_content, 'html.parser')
        found_names = []
        
        # Italian legal suffixes
        italian_suffixes = [
            r'SRL', r'S\.R\.L\.?', r'SPA', r'S\.P\.A\.?', r'SAPA', r'S\.A\.P\.A\.?', 
            r'SAS', r'S\.A\.S\.?', r'SNC', r'S\.N\.C\.?', r'SS', r'S\.S\.?',
            r'SCARL', r'S\.C\.A\.R\.L\.?', r'SCRL', r'S\.C\.R\.L\.?',
            r'Società\s+a\s+Responsabilità\s+Limitata', r'Società\s+per\s+Azioni',
            r'Società\s+in\s+Nome\s+Collettivo', r'Società\s+Semplice'
        ]
        suffix_pattern = '|'.join(italian_suffixes)
        
        # 1. Extract from <title> tag
        title_tag = soup.find('title')
        if title_tag:
            title_text = title_tag.get_text().strip()
            title_matches = re.findall(rf'([A-Za-z][A-Za-z\s&.\'-]+\s+(?:{suffix_pattern}))', title_text, re.I)
            for match in title_matches:
                clean_name = re.sub(r'\s+', ' ', match.strip())
                if 10 <= len(clean_name) <= 100 and self.similarity(clean_name, company_name) > 0.3:
                    found_names.append((clean_name, 'title'))
        
        # 2. Extract from <meta name="og:site_name">
        og_site_name = soup.find('meta', attrs={'name': 'og:site_name'}) or soup.find('meta', attrs={'property': 'og:site_name'})
        if og_site_name and og_site_name.get('content'):
            site_name = og_site_name.get('content').strip()
            if re.search(rf'(?:{suffix_pattern})', site_name, re.I):
                if 10 <= len(site_name) <= 100 and self.similarity(site_name, company_name) > 0.3:
                    found_names.append((site_name, 'og:site_name'))
        
        # 3. Parse JSON-LD structured data
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
                        if obj_type in ['organization', 'corporation', 'localbusiness']:
                            name = obj.get('name', '').strip()
                            legal_name = obj.get('legalName', '').strip()
                            
                            candidate_name = legal_name if legal_name else name
                            if candidate_name:
                                if re.search(rf'(?:{suffix_pattern})', candidate_name, re.I):
                                    if 10 <= len(candidate_name) <= 100 and self.similarity(candidate_name, company_name) > 0.3:
                                        found_names.append((candidate_name, 'json-ld'))
            except json.JSONDecodeError:
                continue
        
        # 4. Footer extraction (priority for Italian sites)
        footers = soup.find_all(['footer', 'div'], class_=re.compile(r'footer|piede|bottom', re.I))
        for footer in footers:
            footer_text = footer.get_text()
            
            # Look for VAT codes in footer first
            vat_code = self.extract_vat_from_html(footer_text)
            if vat_code:
                # If VAT found in footer, look for company name nearby
                vat_patterns_with_context = [
                    rf'([A-Za-z][A-Za-z\s&.\'-]+(?:{suffix_pattern}))\s*[^\w]*P\.?\s*IVA[\s#:]*{vat_code}',
                    rf'P\.?\s*IVA[\s#:]*{vat_code}\s*[^\w]*([A-Za-z][A-Za-z\s&.\'-]+(?:{suffix_pattern}))'
                ]
                for pattern in vat_patterns_with_context:
                    match = re.search(pattern, footer_text, re.I)
                    if match:
                        name = match.group(1).strip()
                        clean_name = re.sub(r'\s+', ' ', name)
                        if 10 <= len(clean_name) <= 100:
                            found_names.append((clean_name, 'footer_with_vat'))
            
            # Copyright patterns for Italian sites
            copyright_patterns = [
                rf'(?:©|Copyright|Tutti\s+i\s+diritti\s+riservati)\s*(?:20[0-9]{{2}})?\s*([A-Za-z][A-Za-z\s&.\'-]+(?:{suffix_pattern}))',
                rf'©\s*([A-Za-z][A-Za-z\s&.\'-]+(?:{suffix_pattern}))'
            ]
            
            for pattern in copyright_patterns:
                matches = re.finditer(pattern, footer_text, re.IGNORECASE)
                for match in matches:
                    name = match.group(1).strip()
                    clean_name = re.sub(r'\s+', ' ', name)
                    if 10 <= len(clean_name) <= 100 and self.similarity(clean_name, company_name) > 0.3:
                        found_names.append((clean_name, 'copyright'))
        
        # 5. End of page content (common in Italian sites)
        page_text = soup.get_text()
        page_sections = page_text.split('\n')
        last_sections = page_sections[-20:]  # Last 20 lines
        end_text = '\n'.join(last_sections)
        
        legal_name_matches = re.findall(rf'([A-Z][A-Za-z\s&.\'-]+\s+(?:{suffix_pattern}))', end_text, re.I)
        for match in legal_name_matches:
            clean_name = re.sub(r'\s+', ' ', match.strip())
            clean_name = re.sub(r'^[^\w]+|[^\w]+$', '', clean_name)
            if (10 <= len(clean_name) <= 150 and 
                not re.search(r'cookie|privacy|termini|contatti|home|menu', clean_name, re.I) and
                len(clean_name.split()) >= 2 and
                self.similarity(clean_name, company_name) > 0.3):
                found_names.append((clean_name, 'end_of_page'))
        
        # Sort by similarity to company name
        unique_names = {}
        for name, source in found_names:
            if name not in unique_names:
                unique_names[name] = source
        
        scored_names = []
        for name, source in unique_names.items():
            score = self.similarity(name, company_name)
            scored_names.append((name, source, score))
        
        scored_names.sort(key=lambda x: x[2], reverse=True)
        return [(name, source) for name, source, score in scored_names]

    def search_google_for_vat(self, company_name):
        """Search Google for company + 'partita iva' to find VAT information"""
        self.logger.info(f"Searching Google for VAT info: {company_name}")
        
        # Use a simple approach to search for VAT info
        search_query = f'"{company_name}" "partita iva"'
        
        try:
            # We'll use the website search approach instead of direct Google API
            # This is more practical for the scope of this tool
            return None, None
        except Exception as e:
            self.logger.debug(f"Google search error: {e}")
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

    def get_italian_paths(self):
        """Return common paths for Italian websites"""
        return [
            '',  # Main page (homepage)
            '/contatti', '/contattaci', '/contact', '/chi-siamo',
            '/about', '/chi-siamo', '/la-nostra-azienda', '/azienda',
            '/note-legali', '/informazioni-legali', '/privacy', '/cookie'
        ]

    def search_website_for_info(self, base_url, company_name):
        """Search website for VAT code and legal name"""
        found_legal_names = []
        found_vat = None
        
        italian_paths = self.get_italian_paths()

        for path in italian_paths:
            url = base_url.rstrip('/') + path
            self.logger.info(f"Checking: {url}")

            try:
                # Try requests first
                html_content = self.fetch_with_requests(url)
                
                # Fallback to Selenium
                if not html_content:
                    html_content = self.fetch_with_selenium(url)

                if html_content:
                    # Look for VAT code first (priority)
                    if not found_vat:
                        vat = self.extract_vat_from_html(html_content)
                        if vat:
                            found_vat = vat
                            self.logger.info(f"Found VAT code: {vat}")
                    
                    # Extract legal names using Italian approach
                    legal_names_with_sources = self.extract_legal_name_italian_approach(html_content, company_name)
                    found_legal_names.extend(legal_names_with_sources)
                    
                    # Priority to footer and end-of-page findings for Italian sites
                    if found_vat and any(source in ['footer_with_vat', 'footer', 'end_of_page'] for _, source in legal_names_with_sources):
                        break

            except Exception as e:
                self.logger.debug(f"Error processing {url}: {e}")

            time.sleep(0.5)

        return found_legal_names, found_vat

    def process_single_italian_company(self, company):
        """Process a single Italian company to extract VAT code"""
        self.logger.info(f"\n--- Processing Italian company: {company['name']} ---")
        
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
            'vat_code': '',
            'status': 'not_found',
            'search_location': '',
            'error': None
        }

        try:
            # If no website provided, skip
            if not website:
                result['status'] = 'no_website'
                result['error'] = 'No website provided'
                return result
            
            self.logger.info("Searching company website...")
            legal_names_with_sources, website_vat = self.search_website_for_info(website, company['name'])
            
            if legal_names_with_sources:
                result['legal_name_website'] = legal_names_with_sources[0][0]
                result['legal_name_source'] = legal_names_with_sources[0][1]
                self.logger.info(f"Found legal name on website: {legal_names_with_sources[0][0]} (source: {legal_names_with_sources[0][1]})")
            
            if website_vat:
                result['vat_code'] = website_vat
                result['status'] = 'found'
                result['search_location'] = 'website'
                self.logger.info(f"Found VAT code on website: {website_vat}")
            else:
                result['status'] = 'vat_not_found'
                if legal_names_with_sources:
                    result['error'] = 'Legal name found but no VAT code'
                else:
                    result['error'] = 'No VAT code or legal name found on website'

        except Exception as error:
            result['error'] = str(error)
            result['status'] = 'error'
            self.logger.error(f"Error: {error}")

        return result

    def process_italian_companies(self, companies):
        """Process multiple Italian companies"""
        self.logger.info(f"Starting Italian VAT extraction for {len(companies)} companies...")
        
        for i, company in enumerate(companies, 1):
            self.logger.info(f"\n[{i}/{len(companies)}]")
            result = self.process_single_italian_company(company)
            self.results.append(result)
            
            # Respectful delay between companies
            time.sleep(2)

        return self.results

    def load_companies_from_excel(self, file_path):
        """Load Italian companies from Excel file"""
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
            
            self.logger.info(f"Loaded {len(companies)} Italian companies from {file_path}")
            return companies
            
        except Exception as e:
            self.logger.error(f"Error loading Excel file: {e}")
            raise

    def save_results_to_excel(self, filename='italy_vat_results.xlsx'):
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
                    'VAT Code (P.IVA)': result['vat_code'],
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

    def save_results_to_json(self, filename='italy_vat_results.json'):
        """Save results to JSON file"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)
        self.logger.info(f"Results saved to {filename}")

    def generate_summary_report(self):
        """Generate comprehensive summary report"""
        total = len(self.results)
        found = len([r for r in self.results if r['status'] == 'found'])
        not_found = len([r for r in self.results if r['status'] in ['not_found', 'vat_not_found']])
        errors = len([r for r in self.results if r['status'] == 'error'])
        no_website = len([r for r in self.results if r['status'] == 'no_website'])
        
        website_found = len([r for r in self.results if r['status'] == 'found' and r['search_location'] == 'website'])
        
        print('\n' + '='*60)
        print('ITALY VAT CODE (PARTITA IVA) EXTRACTION SUMMARY')
        print('='*60)
        print(f'Total companies processed: {total}')
        print(f'VAT codes successfully found: {found} ({(found/total*100):.1f}%)')
        print(f'  - Found on company websites: {website_found}')
        print(f'Not found: {not_found}')
        print(f'Errors: {errors}')
        print(f'No website: {no_website}')

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
    excel_file = 'gis_websitesonly.xlsx'
    
    import sys
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found!")
        print("\nRequired Excel columns:")
        print("- Portfolio Companies (company name)")
        print("- Target Website (company website) - required")
        print(f"\nUsage: python {sys.argv[0]} your_file.xlsx")
        return

    with ItalyVATExtractor() as extractor:
        try:
            print("Loading Italian companies from Excel...")
            companies = extractor.load_companies_from_excel(excel_file)
            
            if not companies:
                print("No valid companies found in the Excel file!")
                return
            
            print(f"\nProcessing {len(companies)} Italian companies...")
            print("This extractor will:")
            print("1. Search company websites for VAT codes (Partita IVA, P.IVA, Codice Fiscale)")
            print("2. Focus on Italian website structure:")
            print("   - Footer sections (priority)")
            print("   - End of page content")
            print("   - About/Contact pages")
            print("   - JSON-LD structured data")
            print("3. Extract legal names with Italian suffixes (SRL, SPA, etc.)")
            print("=" * 60)
            
            results = extractor.process_italian_companies(companies)
            
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
                print(f"  VAT Code (P.IVA): {result['vat_code']}")
                print(f"  Found at: {result['search_location']}")
                if result['legal_name_website']:
                    print(f"  Legal name: {result['legal_name_website']} (source: {result['legal_name_source']})")
                if result['website']:
                    print(f"  Website: {result['website']}")
            
            # Show failures for debugging
            failed_results = [r for r in results if r['status'] not in ['found']]
            if failed_results:
                print(f'\n{"="*60}')
                print(f'FAILED EXTRACTIONS ({len(failed_results)} companies)')
                print('='*60)
                for result in failed_results:
                    print(f"\n{result['original_company_name']}: {result['status']}")
                    if result['error']:
                        print(f"  Error: {result['error']}")
                    if result['legal_name_website']:
                        print(f"  Found legal name: {result['legal_name_website']}")
                        print(f"  (but no VAT code found)")
                    
        except Exception as error:
            print(f'Fatal error: {error}')
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()