# Update the code to show clearer instructions and add file path configuration
with open('arkap_vat_extractor_integrated.py', 'w', encoding='utf-8') as f:
    f.write('''import streamlit as st
import pandas as pd
import io, re, time, requests, random, string, os
from typing import Dict, List, Optional
from difflib import SequenceMatcher
from datetime import datetime, timedelta

# ============================================
# CONFIGURATION - MODIFY THIS PATH
# ============================================
DATABASE_FILE = "databaseaziendearkap.xlsx"
# Alternative: Use absolute path like:
# DATABASE_FILE = "C:/Users/YourName/Documents/databaseaziendearkap.xlsx"
# or on Mac/Linux:
# DATABASE_FILE = "/home/username/documents/databaseaziendearkap.xlsx"
# ============================================

ALLOWED_DOMAIN = "@arkap.ch"
CODE_EXPIRY_MINUTES = 10
SESSION_TIMEOUT_MINUTES = 60

COUNTRY_CODES = {
    'AT': 'Austria', 'CH': 'Switzerland', 'DE': 'Germany',
    'FR': 'France', 'GB': 'United Kingdom', 'IT': 'Italy',
    'LU': 'Luxembourg', 'NL': 'Netherlands', 'PT': 'Portugal'
}

def safe_format(value, fmt="{:,.0f}", pre="", suf="", default="N/A"):
    if pd.isna(value) or value is None or value == '': return default
    try:
        if isinstance(value, str):
            v = value.replace(',', '').replace(' ', '').replace('€', '').replace('k', '').strip()
            if not v or v == '-': return default
            value = float(v)
        return f"{pre}{fmt.format(float(value))}{suf}"
    except: return str(value) if value else default

class CompanyDatabase:
    def __init__(self, df=None):
        self.db, self.name_idx, self.vat_idx, self.country_idx = df, {}, {}, {}
        if df is not None: self._init()
    
    def _init(self):
        mapping = {}
        for col in self.db.columns:
            c = col.lower()
            if 'company' in c and 'name' in c: mapping[col] = 'Company Name'
            elif 'vat' in c and 'code' in c: mapping[col] = 'VAT Code'
            elif 'national' in c and 'id' in c: mapping[col] = 'National ID'
            elif 'fiscal' in c: mapping[col] = 'Fiscal Code'
            elif 'country' in c and 'code' in c: mapping[col] = 'Country Code'
            elif 'nace' in c: mapping[col] = 'Nace Code'
            elif 'last' in c and 'yr' in c: mapping[col] = 'Last Yr'
            elif 'production' in c: mapping[col] = 'Value of production (th)'
            elif 'employee' in c: mapping[col] = 'Employees'
            elif 'ebitda' in c: mapping[col] = 'Ebitda (th)'
            elif 'pfn' in c: mapping[col] = 'PFN (th)'
        self.db = self.db.rename(columns=mapping)
        
        for idx, row in self.db.iterrows():
            if 'Company Name' in self.db.columns and pd.notna(row.get('Company Name')):
                k = str(row['Company Name']).lower().strip()
                self.name_idx.setdefault(k, []).append(idx)
            if 'VAT Code' in self.db.columns and pd.notna(row.get('VAT Code')):
                k = str(row['VAT Code']).upper().replace(' ', '').replace('-', '').replace('.', '')
                self.vat_idx.setdefault(k, []).append(idx)
        
        if 'Country Code' in self.db.columns:
            for cc in self.db['Country Code'].unique():
                if pd.notna(cc):
                    self.country_idx[str(cc).upper()] = self.db[self.db['Country Code'] == cc].index.tolist()
        st.success(f"✅ Database loaded: {len(self.db)} companies, {len(self.name_idx)} unique names")
    
    def search_name(self, name, country=None):
        k = name.lower().strip()
        if k in self.name_idx:
            idxs = self.name_idx[k]
            if country and country in self.country_idx:
                idxs = [i for i in idxs if i in self.country_idx[country]]
            return self._extract(self.db.iloc[idxs[0]]) if idxs else None
        return None
    
    def search_vat(self, vat, country=None):
        k = str(vat).upper().replace(' ', '').replace('-', '').replace('.', '')
        if k in self.vat_idx:
            idxs = self.vat_idx[k]
            if country and country in self.country_idx:
                idxs = [i for i in idxs if i in self.country_idx[country]]
            return self._extract(self.db.iloc[idxs[0]]) if idxs else None
        return None
    
    def _extract(self, row):
        d = {'source': 'database'}
        for f in ['Company Name', 'National ID', 'Fiscal Code', 'VAT Code', 'Country Code', 'Nace Code', 'Last Yr', 'Value of production (th)', 'Employees', 'Ebitda (th)', 'PFN (th)']:
            if f in row.index and pd.notna(row[f]):
                d[f.lower().replace(' ', '_').replace('(', '').replace(')', '')] = row[f]
        return d

class AuthenticationManager:
    def __init__(self):
        for k in ['auth_codes', 'authenticated', 'user_email', 'auth_time', 'company_db', 'search_mode', 'db_file_path']:
            if k not in st.session_state:
                st.session_state[k] = {} if k == 'auth_codes' else (False if k == 'authenticated' else ("" if k in ['user_email', 'db_file_path'] else None))
    
    def is_valid_email(self, e):
        return re.match(r'^[\\w.+-]+@[\\w.-]+\\.[\\w]+$', e) and e.lower().endswith(ALLOWED_DOMAIN.lower())
    
    def gen_code(self): return ''.join(random.choices(string.digits, k=6))
    
    def store_code(self, e, c):
        st.session_state.auth_codes[e] = {'code': c, 'timestamp': datetime.now(), 'attempts': 0}
    
    def verify(self, e, c):
        if e not in st.session_state.auth_codes: return False, "No code"
        d = st.session_state.auth_codes[e]
        if datetime.now() - d['timestamp'] > timedelta(minutes=CODE_EXPIRY_MINUTES):
            del st.session_state.auth_codes[e]
            return False, "Expired"
        if d['attempts'] >= 3:
            del st.session_state.auth_codes[e]
            return False, "Too many"
        if d['code'] == c:
            st.session_state.authenticated = True
            st.session_state.user_email = e
            st.session_state.auth_time = datetime.now()
            del st.session_state.auth_codes[e]
            return True, "Success"
        d['attempts'] += 1
        return False, f"{3-d['attempts']} left"
    
    def is_valid(self):
        return (st.session_state.authenticated and st.session_state.auth_time and 
                datetime.now() - st.session_state.auth_time <= timedelta(minutes=SESSION_TIMEOUT_MINUTES))
    
    def logout(self):
        st.session_state.authenticated = False
        st.session_state.user_email = ""
        st.session_state.auth_time = None

class SimpleUKExtractor:
    def __init__(self):
        self.s = requests.Session()
        self.s.headers.update({'User-Agent': 'Mozilla/5.0'})
    
    def process(self, name, url=None):
        r = {'company_name': name, 'website': url or '', 'status': 'Not Found', 'source': 'web'}
        if url:
            try:
                resp = self.s.get(url, timeout=10)
                if resp.status_code == 200:
                    for p in [r'Company\\s+number[\\s:]*([0-9]{8})']:
                        for m in re.finditer(p, resp.text, re.I):
                            c = re.sub(r'[^0-9]', '', m.group(1))
                            if len(c) in [6, 8]:
                                r['company_number'], r['status'] = c, 'Found'
                                return r
            except: pass
        return r

class MultiModeExtractor:
    def __init__(self, db=None, use_db=True):
        self.db, self.use_db = db, use_db
        self.extractors = {'GB': SimpleUKExtractor()}
        self.patterns = {
            'DE': [r'Steuernummer[\\s#:]*([0-9]{2,3}/[0-9]{3,4}/[0-9]{4,5})'],
            'FR': [r'SIREN[\\s#:]*([0-9]{9})'],
            'IT': [r'P\\.?\\s*IVA[\\s#:]*([0-9]{11})'],
            'PT': [r'NIF[\\s:]*([0-9]{9})'],
            'NL': [r'KvK[\\s#:]*([0-9]{8})'],
            'AT': [r'ATU\\s*([0-9]{8})'],
            'CH': [r'CHE[\\s-]?([0-9]{3})'],
            'LU': [r'LU\\s*([0-9]{8})']
        }
    
    def process_single(self, name, web, country, vat=None):
        if self.use_db and self.db:
            r = self.db.search_name(name, country)
            if r: return {**r, 'search_method': 'DB-Name', 'status': 'Found'}
            if vat:
                r = self.db.search_vat(vat, country)
                if r: return {**r, 'search_method': 'DB-VAT', 'status': 'Found'}
            w = self._web(name, web, country)
            w['search_method'] = 'DB failed-Web'
            return w
        w = self._web(name, web, country)
        w['search_method'] = 'Web only'
        return w
    
    def _web(self, name, web, country):
        if country in self.extractors:
            return self.extractors[country].process(name, web)
        r = {'company_name': name, 'website': web, 'country_code': country, 'status': 'Not Found', 'source': 'web'}
        if web and country in self.patterns:
            try:
                resp = requests.get(web, timeout=10)
                if resp.status_code == 200:
                    for p in self.patterns[country]:
                        for m in re.finditer(p, resp.text, re.I):
                            r[f'{country.lower()}_code'], r['status'] = m.group(1), 'Found'
                            break
            except: pass
        return r
    
    def process_list(self, df, prog=None):
        results = []
        nc = [c for c in df.columns if 'company' in c.lower() or 'name' in c.lower()]
        name_col = nc[0] if nc else df.columns[0]
        wc = [c for c in df.columns if 'website' in c.lower() or 'url' in c.lower()]
        web_col = wc[0] if wc else None
        cc = [c for c in df.columns if 'country' in c.lower()]
        country_col = cc[0] if cc else None
        vc = [c for c in df.columns if 'vat' in c.lower() or 'fiscal' in c.lower()]
        vat_col = vc[0] if vc else None
        
        for idx, row in df.iterrows():
            if prog: prog(idx+1, len(df))
            name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
            web = str(row[web_col]).strip() if web_col and pd.notna(row[web_col]) else ''
            vat = str(row[vat_col]).strip() if vat_col and pd.notna(row[vat_col]) else None
            country = 'GB'
            if country_col and pd.notna(row[country_col]):
                cv = str(row[country_col]).strip().upper()
                if len(cv) == 2 and cv in COUNTRY_CODES: country = cv
            results.append(self.process_single(name, web, country, vat))
            time.sleep(0.2)
        return results

def show_auth(auth):
    st.title("🔐 arKap VAT Extractor")
    st.info("🏢 @arkap.ch only")
    t1, t2 = st.tabs(["📧 Email", "🔑 Code"])
    with t1:
        e = st.text_input("Email")
        if st.button("Send Code", type="primary"):
            if auth.is_valid_email(e):
                c = auth.gen_code()
                auth.store_code(e, c)
                st.success(f"Code: {c}")
            else: st.error("Invalid")
    with t2:
        e = st.text_input("Email", key="e2")
        c = st.text_input("Code", max_chars=6)
        if st.button("Verify", type="primary"):
            ok, msg = auth.verify(e, c)
            if ok: st.success(msg); st.balloons(); time.sleep(1); st.rerun()
            else: st.error(msg)

def show_main():
    st.title("🌍 arKap VAT Extractor")
    c1, c2 = st.columns([3,1])
    with c1: st.markdown(f"**User:** {st.session_state.user_email}")
    with c2:
        if st.button("Logout"): AuthenticationManager().logout(); st.rerun()
    st.markdown("---")
    
    # Database loading with better UI
    if st.session_state.company_db is None:
        st.header("📊 Database Configuration")
        
        with st.expander("ℹ️ Database Setup Instructions", expanded=True):
            st.markdown(f"""
            **Current database file path:** `{DATABASE_FILE}`
            
            **To fix "database not found" error:**
            
            1. **Place your file in the correct location:**
               - Put `databaseaziendearkap.xlsx` in the same folder as this script
               - Current working directory: `{os.getcwd()}`
            
            2. **OR modify the path at the top of the script:**
               - Open `arkap_vat_extractor_integrated.py`
               - Change line 9: `DATABASE_FILE = "your/full/path/here.xlsx"`
               - Example Windows: `DATABASE_FILE = "C:/Users/Andrea/Documents/databaseaziendearkap.xlsx"`
               - Example Mac/Linux: `DATABASE_FILE = "/home/andrea/documents/databaseaziendearkap.xlsx"`
            
            3. **OR upload manually below**
            """)
        
        st.subheader("Option 1: Auto-load from configured path")
        col1, col2 = st.columns([2, 1])
        with col1:
            st.code(f"Looking for: {DATABASE_FILE}")
            if os.path.exists(DATABASE_FILE):
                st.success(f"✅ File found at: {os.path.abspath(DATABASE_FILE)}")
            else:
                st.error(f"❌ File not found at: {os.path.abspath(DATABASE_FILE)}")
        with col2:
            if st.button("🔄 Try Load", type="primary"):
                try:
                    if os.path.exists(DATABASE_FILE):
                        df = pd.read_excel(DATABASE_FILE)
                        st.session_state.company_db = CompanyDatabase(df)
                        st.session_state.db_file_path = DATABASE_FILE
                        st.rerun()
                    else:
                        st.error("File not found!")
                except Exception as e:
                    st.error(f"Error loading: {e}")
        
        st.markdown("---")
        st.subheader("Option 2: Manual upload")
        uploaded = st.file_uploader("Upload Database File", type=['xlsx', 'xls', 'csv'], key="db_upload")
        if uploaded:
            try:
                if uploaded.name.endswith('.csv'):
                    df = pd.read_csv(uploaded)
                else:
                    df = pd.read_excel(uploaded)
                st.session_state.company_db = CompanyDatabase(df)
                st.session_state.db_file_path = f"Uploaded: {uploaded.name}"
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
        
        st.markdown("---")
        st.subheader("Option 3: Continue without database (Web Only mode)")
        if st.button("⏭️ Skip Database - Use Web Only", type="secondary"):
            st.session_state.company_db = None
            st.session_state.search_mode = 'web'
            st.rerun()
        
        return
    
    # Show DB info
    if st.session_state.company_db:
        st.success(f"📊 Database active: {st.session_state.db_file_path}")
    
    if st.session_state.search_mode is None:
        st.header("🔍 Select Search Mode")
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("🗄️ DB + Web")
            st.write("• DB by name → DB by VAT → Web")
            st.write("• Fast, enriched data")
            if st.button("Use DB+Web", type="primary", use_container_width=True):
                if st.session_state.company_db:
                    st.session_state.search_mode = 'db'; st.rerun()
                else: st.error("DB not available")
        with c2:
            st.subheader("🌐 Web Only")
            st.write("• Direct scraping")
            st.write("• No DB lookup")
            if st.button("Use Web Only", use_container_width=True):
                st.session_state.search_mode = 'web'; st.rerun()
        return
    
    st.info(f"**Mode:** {st.session_state.search_mode.upper()}")
    if st.button("Change"): st.session_state.search_mode = None; st.rerun()
    st.markdown("---")
    
    t1, t2 = st.tabs(["Bulk", "Single"])
    with t1:
        f = st.file_uploader("Company List", type=['csv','xlsx'])
        if f:
            df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
            st.dataframe(df.head())
            if st.button("Process"):
                ext = MultiModeExtractor(st.session_state.company_db, st.session_state.search_mode=='db')
                p = st.progress(0)
                res = ext.process_list(df, lambda c,t: p.progress(c/t))
                st.success("Done")
                rdf = pd.DataFrame(res)
                st.dataframe(rdf)
                c1,c2,c3 = st.columns(3)
                with c1: st.metric("Total", len(res))
                with c2: st.metric("Found", len([r for r in res if r['status']=='Found']))
                with c3: st.metric("Rate%", f"{len([r for r in res if r['status']=='Found'])/len(res)*100:.1f}")
                csv = io.StringIO()
                rdf.to_csv(csv, index=False)
                st.download_button("Download", csv.getvalue(), f"res_{datetime.now().strftime('%Y%m%d_%H%M')}.csv")
    
    with t2:
        c1,c2 = st.columns(2)
        with c1:
            n = st.text_input("Name")
            w = st.text_input("Website")
            v = st.text_input("VAT")
        with c2:
            co = st.selectbox("Country", list(COUNTRY_CODES.keys()), format_func=lambda x:f"{COUNTRY_CODES[x]} ({x})")
        if st.button("Search") and n:
            ext = MultiModeExtractor(st.session_state.company_db, st.session_state.search_mode=='db')
            r = ext.process_single(n, w, co, v)
            if r['status']=='Found':
                st.success(f"✅ {r.get('search_method')}")
                if r.get('source')=='database':
                    st.subheader("Info")
                    c1,c2=st.columns(2)
                    with c1:
                        for k in ['company_name','vat_code']: 
                            if k in r: st.write(f"**{k}:** {r[k]}")
                    with c2:
                        for k in ['country_code','nace_code']: 
                            if k in r: st.write(f"**{k}:** {r[k]}")
                    st.subheader("Financial")
                    c1,c2,c3=st.columns(3)
                    with c1:
                        if 'last_yr' in r: st.metric("Year",r['last_yr'])
                        if 'employees' in r: st.metric("Emp",safe_format(r.get('employees')))
                    with c2:
                        if 'value_of_production_th' in r: st.metric("Prod",safe_format(r.get('value_of_production_th'),pre="€",suf="k"))
                        if 'ebitda_th' in r: st.metric("EBITDA",safe_format(r.get('ebitda_th'),pre="€",suf="k"))
                    with c3:
                        if 'pfn_th' in r: st.metric("PFN",safe_format(r.get('pfn_th'),pre="€",suf="k"))
                else:
                    st.info("Web data")
                    for k,v in r.items():
                        if k not in ['company_name','website','status','source','search_method','country_code']:
                            st.write(f"{k}: {v}")
            else: st.warning("Not found")
            with st.expander("Raw"): st.json(r)

def main():
    st.set_page_config(page_title="arKap", page_icon="⚡", layout="wide")
    auth = AuthenticationManager()
    if auth.is_valid(): show_main()
    else: show_auth(auth)

if __name__ == "__main__": main()
''')

print("✅ UPDATED VERSION WITH DATABASE FIX!")
print("\n" + "="*70)
print("DATABASE CONFIGURATION OPTIONS:")
print("="*70)
print("\n🔧 The app now provides 3 ways to load the database:")
print("\n1️⃣ AUTO-LOAD (Recommended):")
print("   • Place 'databaseaziendearkap.xlsx' in the same folder as the script")
print("   • Click 'Try Load' button")
print("\n2️⃣ MANUAL UPLOAD:")
print("   • Use the file uploader in the app")
print("   • Upload your .xlsx/.xls/.csv file")
print("\n3️⃣ SKIP DATABASE:")
print("   • Click 'Skip Database - Use Web Only'")
print("   • App will work in web-only mode")
print("\n" + "="*70)
print("TO FIX 'DATABASE NOT FOUND':")
print("="*70)
print("\n📁 Option A: Put file in script folder")
print("   1. Find where arkap_vat_extractor_integrated.py is saved")
print("   2. Put databaseaziendearkap.xlsx in the SAME folder")
print("   3. Run app and click 'Try Load'")
print("\n📝 Option B: Edit the file path")
print("   1. Open arkap_vat_extractor_integrated.py")
print("   2. Line 9: Change DATABASE_FILE to full path")
print("   3. Example: DATABASE_FILE = 'C:/Users/Andrea/Documents/databaseaziendearkap.xlsx'")
print("   4. Save and run")
print("\n📤 Option C: Upload in app")
print("   1. Run the app")
print("   2. Use 'Manual upload' section")
print("   3. Select your database file")
print("\n🚀 The app will now show you:")
print("   • Current working directory")
print("   • Where it's looking for the file")
print("   • Clear instructions to fix the issue")
)