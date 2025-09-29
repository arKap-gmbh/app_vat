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

class FranceSIRENExtractor:
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
        
        # France-specific SIREN patterns
        self.siren_patterns = [
            r'SIREN[\s#:]*([0-9]{9})',  # SIREN with prefix
            r'(?:N°\s*SIREN|Numéro\s*SIREN)[\s#:]*([0-9]{9})',  # French SIREN format
            r'SIREN\s*:?\s*([0-9]{3}[\s\-]?[0-9]{3}[\s\-]?[0-9]{3})',  # Formatted SIREN
        ]

    def similarity(self, a, b):
        """Calculate similarity between two strings"""
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    def extract_siren_from_html(self, html_content):
        """Extract SIREN code from HTML content"""
        for pattern in self.siren_patterns:
            matches = re.finditer(pattern, html_content, re.IGNORECASE)
            for match in matches:
                code = match.group(1) if len(match.groups()) > 0 else match.group(0)
                clean_code = re.sub(r'[^0-9]', '', code)
                if len(clean_code) == 9 and clean_code.isdigit():
                    return clean_code
        return None

    def extract_legal_name_structured_approach(self, html_content, company_name):
        """Extract legal name using the specified structured approach"""
        soup = BeautifulSoup(html_content, 'html.parser')
        found_names = []
        
        # French legal suffixes
        french_suffixes = [
            r'SARL', r'SA', r'SAS', r'EURL', r'SCA', r'SNC', r'SCS', r'SEL', r'SELARL', r'SELAFA', r'SELCA',
            r'S\.A\.R\.L\.?', r'S\.A\.S\.?', r'S\.A\.', r'E\.U\.R\.L\.?'
        ]
        suffix_pattern = '|'.join(french_suffixes)
        
        # 1. Extract from <title> tag
        title_tag = soup.find('title')
        if title_tag:
            title_text = title_tag.get_text().strip()
            # Look for legal suffixes in title
            title_matches = re.findall(rf'([A-Za-z][A-Za-z\s&.\'-]+\s+(?:{suffix_pattern}))', title_text, re.I)
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
        footers = soup.find_all(['footer', 'div'], class_=re.compile(r'footer|pied', re.I))
        for footer in footers:
            footer_text = footer.get_text()
            # Look for copyright patterns
            copyright_patterns = [
                rf'(?:©|Copyright|Tous\s+droits\s+réservés)\s*(?:20[0-9]{{2}})?\s*([A-Za-z][A-Za-z\s&.\'-]+(?:{suffix_pattern}))',
                rf'©\s*([A-Za-z][A-Za-z\s&.\'-]+(?:{suffix_pattern}))'
            ]
            
            for pattern in copyright_patterns:
                matches = re.finditer(pattern, footer_text, re.IGNORECASE)
                for match in matches:
                    name = match.group(1).strip()
                    clean_name = re.sub(r'\s+', ' ', name)
                    if 10 <= len(clean_name) <= 100 and self.similarity(clean_name, company_name) > 0.3:
                        found_names.append((clean_name, 'copyright'))
        
        # 5. General legal suffix search in HTML
        legal_name_matches = re.findall(rf'([A-Z][A-Za-z\s&.\'-]+\s+(?:{suffix_pattern}))', html_content, re.I)
        for match in legal_name_matches:
            clean_name = re.sub(r'\s+', ' ', match.strip())
            clean_name = re.sub(r'^[^\w]+|[^\w]+$', '', clean_name)
            if (10 <= len(clean_name) <= 150 and 
                not re.search(r'cookie|privacy|terms|contact|home|menu', clean_name, re.I) and
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

    def search_annuaire_entreprises(self, company_name):
        """Search for company on Annuaire Entreprises and return SIREN and legal name"""
        self.logger.info(f"Searching Annuaire Entreprises for: {company_name}")
        
        api_url = "https://recherche-entreprises.api.gouv.fr/search"
        
        try:
            params = {
                'q': company_name,
                'limite': 10,
                'per_page': 10
            }
            
            response = self.session.get(api_url, params=params, timeout=10)
            if response.status_code == 200:
                data = response.json()
                
                if 'results' in data and data['results']:
                    results = data['results']
                    
                    # If only one result, return it
                    if len(results) == 1:
                        result = results[0]
                        siren = result.get('siren')
                        legal_name = result.get('nom_complet', result.get('nom_raison_sociale', ''))
                        
                        self.logger.info(f"Found single match: {legal_name}")
                        return siren, legal_name
                    
                    # Multiple results - focus on full legal name similarity
                    self.logger.info(f"Found {len(results)} matches, selecting best legal name match")
                    best_match = None
                    best_score = 0
                    
                    for result in results:
                        legal_name = result.get('nom_complet', result.get('nom_raison_sociale', ''))
                        
                        # Calculate similarity with the search term
                        score = self.similarity(legal_name, company_name)
                        
                        self.logger.info(f"  Candidate: {legal_name} (score: {score:.2f})")
                        
                        if score > best_score and score > 0.5:  # Minimum threshold
                            best_match = result
                            best_score = score
                    
                    if best_match:
                        siren = best_match.get('siren')
                        legal_name = best_match.get('nom_complet', best_match.get('nom_raison_sociale', ''))
                        
                        self.logger.info(f"Selected best match: {legal_name} (score: {best_score:.2f})")
                        return siren, legal_name
                    else:
                        self.logger.info("No sufficiently similar match found")
                
            time.sleep(1)  # Respectful delay
                
        except Exception as e:
            self.logger.debug(f"Annuaire Entreprises search error: {e}")
        
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

    def get_french_paths(self):
        """Return common paths for French websites"""
        return [
            '',  # Main page (homepage)
            '/contact', '/contacts', '/nous-contacter',
            '/a-propos', '/qui-sommes-nous', '/notre-societe', '/about',
            '/mentions-legales', '/legal', '/informations-legales'
        ]

    def search_website_for_info(self, base_url, company_name):
        """Search website for SIREN code and legal name"""
        found_legal_names = []
        found_siren = None
        
        french_paths = self.get_french_paths()

        for path in french_paths:
            url = base_url.rstrip('/') + path
            self.logger.info(f"Checking: {url}")

            try:
                # Try requests first
                html_content = self.fetch_with_requests(url)
                
                # Fallback to Selenium
                if not html_content:
                    html_content = self.fetch_with_selenium(url)

                if html_content:
                    # Look for SIREN code
                    if not found_siren:
                        siren = self.extract_siren_from_html(html_content)
                        if siren:
                            found_siren = siren
                    
                    # Extract legal names using structured approach
                    legal_names_with_sources = self.extract_legal_name_structured_approach(html_content, company_name)
                    found_legal_names.extend(legal_names_with_sources)
                    
                    # If we found SIREN and legal names, we can break
                    if found_siren and found_legal_names:
                        break

            except Exception as e:
                self.logger.debug(f"Error processing {url}: {e}")

            time.sleep(0.5)

        return found_legal_names, found_siren

    def process_single_french_company(self, company):
        """Process a single French company to extract SIREN code"""
        self.logger.info(f"\n--- Processing French company: {company['name']} ---")
        
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
            'legal_name_annuaire': '',
            'siren': '',
            'status': 'not_found',
            'search_location': '',
            'error': None
        }

        try:
            # Step 1: Search Annuaire Entreprises first
            self.logger.info("Searching Annuaire Entreprises...")
            annuaire_siren, annuaire_legal_name = self.search_annuaire_entreprises(company['name'])
            
            if annuaire_siren and annuaire_legal_name:
                result['siren'] = annuaire_siren
                result['legal_name_annuaire'] = annuaire_legal_name
                result['status'] = 'found'
                result['search_location'] = 'annuaire_entreprises'
                self.logger.info(f"Found via Annuaire: SIREN {annuaire_siren}, Legal name: {annuaire_legal_name}")
                return result
            
            # Step 2: If not found in Annuaire or no website, try website scraping
            if not website:
                result['status'] = 'no_website'
                result['error'] = 'No website provided and not found in Annuaire Entreprises'
                return result
            
            self.logger.info("Not found in Annuaire Entreprises. Searching website...")
            legal_names_with_sources, website_siren = self.search_website_for_info(website, company['name'])
            
            if legal_names_with_sources:
                result['legal_name_website'] = legal_names_with_sources[0][0]
                result['legal_name_source'] = legal_names_with_sources[0][1]
                self.logger.info(f"Found legal name on website: {legal_names_with_sources[0][0]} (source: {legal_names_with_sources[0][1]})")
            
            if website_siren:
                result['siren'] = website_siren
                result['status'] = 'found'
                result['search_location'] = 'website'
                self.logger.info(f"Found SIREN on website: {website_siren}")
            else:
                # If we found legal name on website but no SIREN, try searching Annuaire with the legal name
                if legal_names_with_sources:
                    self.logger.info(f"Trying Annuaire search with website legal name: {legal_names_with_sources[0][0]}")
                    legal_name_siren, _ = self.search_annuaire_entreprises(legal_names_with_sources[0][0])
                    if legal_name_siren:
                        result['siren'] = legal_name_siren
                        result['status'] = 'found'
                        result['search_location'] = 'annuaire_via_website_name'
                        self.logger.info(f"Found SIREN via Annuaire using website legal name: {legal_name_siren}")

        except Exception as error:
            result['error'] = str(error)
            result['status'] = 'error'
            self.logger.error(f"Error: {error}")

        return result

    def process_french_companies(self, companies):
        """Process multiple French companies"""
        self.logger.info(f"Starting French SIREN extraction for {len(companies)} companies...")
        
        for i, company in enumerate(companies, 1):
            self.logger.info(f"\n[{i}/{len(companies)}]")
            result = self.process_single_french_company(company)
            self.results.append(result)
            
            # Respectful delay between companies
            time.sleep(2)

        return self.results

    def load_companies_from_excel(self, file_path):
        """Load French companies from Excel file"""
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
            
            self.logger.info(f"Loaded {len(companies)} French companies from {file_path}")
            return companies
            
        except Exception as e:
            self.logger.error(f"Error loading Excel file: {e}")
            raise

    def save_results_to_excel(self, filename='france_siren_results.xlsx'):
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
                    'Legal Name (Annuaire)': result['legal_name_annuaire'],
                    'SIREN': result['siren'],
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

    def save_results_to_json(self, filename='france_siren_results.json'):
        """Save results to JSON file"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)
        self.logger.info(f"Results saved to {filename}")

    def generate_summary_report(self):
        """Generate comprehensive summary report"""
        total = len(self.results)
        found = len([r for r in self.results if r['status'] == 'found'])
        not_found = len([r for r in self.results if r['status'] == 'not_found'])
        errors = len([r for r in self.results if r['status'] == 'error'])
        no_website = len([r for r in self.results if r['status'] == 'no_website'])
        
        annuaire_found = len([r for r in self.results if r['status'] == 'found' and 'annuaire' in r['search_location']])
        website_found = len([r for r in self.results if r['status'] == 'found' and r['search_location'] == 'website'])
        
        print('\n' + '='*60)
        print('FRANCE SIREN CODE EXTRACTION SUMMARY')
        print('='*60)
        print(f'Total companies processed: {total}')
        print(f'SIREN codes successfully found: {found} ({(found/total*100):.1f}%)')
        print(f'  - Found via Annuaire Entreprises: {annuaire_found}')
        print(f'  - Found on company websites: {website_found}')
        print(f'Not found: {not_found}')
        print(f'Errors: {errors}')
        print(f'No website (and not in Annuaire): {no_website}')

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
    excel_file = 'smalldbmachinery_fr.xlsx'
    
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

    with FranceSIRENExtractor() as extractor:
        try:
            print("Loading French companies from Excel...")
            companies = extractor.load_companies_from_excel(excel_file)
            
            if not companies:
                print("No valid companies found in the Excel file!")
                return
            
            print(f"\nProcessing {len(companies)} French companies...")
            print("This extractor will:")
            print("1. Search Annuaire Entreprises API for company names from the Excel file")
            print("2. If not found, extract legal names from company websites using structured approach:")
            print("   - <title> tags, <meta og:site_name>, JSON-LD structured data")
            print("   - Footer copyright sections and legal suffixes")
            print("3. Focus on SIREN codes only (no SIRET or VAT conversion)")
            print("=" * 60)
            
            results = extractor.process_french_companies(companies)
            
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
                print(f"  SIREN: {result['siren']}")
                print(f"  Found at: {result['search_location']}")
                if result['legal_name_annuaire']:
                    print(f"  Legal name (Annuaire): {result['legal_name_annuaire']}")
                if result['legal_name_website']:
                    print(f"  Legal name (Website): {result['legal_name_website']} (source: {result['legal_name_source']})")
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
                        print(f"  Found legal name on website: {result['legal_name_website']}")
                        print(f"  (but no SIREN code found)")
                    
        except Exception as error:
            print(f'Fatal error: {error}')
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()