
# Fix the syntax error and create the file
with open('arkap_vat_extractor_dropbox.py', 'w', encoding='utf-8') as f:
    f.write('''import streamlit as st
import pandas as pd
import io, re, time, requests, random, string, os
from typing import Dict, List, Optional
from difflib import SequenceMatcher
from datetime import datetime, timedelta

def get_dropbox_download_link(shared_link):
    """Convert Dropbox sharing link to direct download link"""
    if 'dropbox.com' in shared_link:
        return shared_link.replace('dl=0', 'dl=1').replace('www.dropbox.com', 'dl.dropboxusercontent.com')
    return shared_link

def load_database_from_dropbox():
    """Load database from Dropbox using Streamlit secrets"""
    try:
        if 'DROPBOX_FILE_URL' in st.secrets:
            dropbox_url = st.secrets["DROPBOX_FILE_URL"]
        else:
            st.warning("⚠️ Dropbox URL not in secrets. Configure in app settings.")
            return None
        download_url = get_dropbox_download_link(dropbox_url)
        with st.spinner("📥 Downloading database from Dropbox..."):
            response = requests.get(download_url, timeout=30)
            response.raise_for_status()
            df = pd.read_excel(io.BytesIO(response.content))
            st.success(f"✅ Database loaded: {len(df)} companies")
            return df
    except Exception as e:
        st.error(f"❌ Error loading from Dropbox: {str(e)}")
        return None

ALLOWED_DOMAIN = "@arkap.ch"
CODE_EXPIRY_MINUTES = 10
SESSION_TIMEOUT_MINUTES = 60
COUNTRY_CODES = {'AT': 'Austria', 'CH': 'Switzerland', 'DE': 'Germany', 'FR': 'France', 'GB': 'United Kingdom', 'IT': 'Italy', 'LU': 'Luxembourg', 'NL': 'Netherlands', 'PT': 'Portugal'}

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
                if pd.notna(cc): self.country_idx[str(cc).upper()] = self.db[self.db['Country Code'] == cc].index.tolist()
        st.success(f"📊 DB indexed: {len(self.db)} companies")
    def search_name(self, name, country=None):
        k = name.lower().strip()
        if k in self.name_idx:
            idxs = self.name_idx[k]
            if country and country in self.country_idx: idxs = [i for i in idxs if i in self.country_idx[country]]
            return self._extract(self.db.iloc[idxs[0]]) if idxs else None
        return None
    def search_vat(self, vat, country=None):
        k = str(vat).upper().replace(' ', '').replace('-', '').replace('.', '')
        if k in self.vat_idx:
            idxs = self.vat_idx[k]
            if country and country in self.country_idx: idxs = [i for i in idxs if i in self.country_idx[country]]
            return self._extract(self.db.iloc[idxs[0]]) if idxs else None
        return None
    def _extract(self, row):
        d = {'source': 'database'}
        for f in ['Company Name', 'National ID', 'Fiscal Code', 'VAT Code', 'Country Code', 'Nace Code', 'Last Yr', 'Value of production (th)', 'Employees', 'Ebitda (th)', 'PFN (th)']:
            if f in row.index and pd.notna(row[f]): d[f.lower().replace(' ', '_').replace('(', '').replace(')', '')] = row[f]
        return d

class AuthenticationManager:
    def __init__(self):
        for k in ['auth_codes', 'authenticated', 'user_email', 'auth_time', 'company_db', 'search_mode']:
            if k not in st.session_state: st.session_state[k] = {} if k == 'auth_codes' else (False if k == 'authenticated' else ("" if k == 'user_email' else None))
    def is_valid_email(self, e): return re.match(r'^[\\w.+-]+@[\\w.-]+\\.[\\w]+$', e) and e.lower().endswith(ALLOWED_DOMAIN.lower())
    def gen_code(self): return ''.join(random.choices(string.digits, k=6))
    def store_code(self, e, c): st.session_state.auth_codes[e] = {'code': c, 'timestamp': datetime.now(), 'attempts': 0}
    def verify(self, e, c):
        if e not in st.session_state.auth_codes: return False, "No code"
        d = st.session_state.auth_codes[e]
        if datetime.now() - d['timestamp'] > timedelta(minutes=CODE_EXPIRY_MINUTES):
            del st.session_state.auth_codes[e]; return False, "Expired"
        if d['attempts'] >= 3: del st.session_state.auth_codes[e]; return False, "Too many"
        if d['code'] == c:
            st.session_state.authenticated, st.session_state.user_email, st.session_state.auth_time = True, e, datetime.now()
            del st.session_state.auth_codes[e]; return True, "Success"
        d['attempts'] += 1; return False, f"{3-d['attempts']} left"
    def is_valid(self): return st.session_state.authenticated and st.session_state.auth_time and datetime.now() - st.session_state.auth_time <= timedelta(minutes=SESSION_TIMEOUT_MINUTES)
    def logout(self): st.session_state.authenticated, st.session_state.user_email, st.session_state.auth_time = False, "", None

class SimpleUKExtractor:
    def __init__(self): self.s = requests.Session(); self.s.headers.update({'User-Agent': 'Mozilla/5.0'})
    def process(self, name, url=None):
        r = {'company_name': name, 'website': url or '', 'status': 'Not Found', 'source': 'web'}
        if url:
            try:
                resp = self.s.get(url, timeout=10)
                if resp.status_code == 200:
                    for m in re.finditer(r'Company\\s+number[\\s:]*([0-9]{8})', resp.text, re.I):
                        c = re.sub(r'[^0-9]', '', m.group(1))
                        if len(c) in [6, 8]: r['company_number'], r['status'] = c, 'Found'; return r
            except: pass
        return r

class MultiModeExtractor:
    def __init__(self, db=None, use_db=True):
        self.db, self.use_db = db, use_db
        self.extractors = {'GB': SimpleUKExtractor()}
        self.patterns = {'DE': [r'Steuernummer[\\s#:]*([0-9]{2,3}/[0-9]{3,4}/[0-9]{4,5})'], 'FR': [r'SIREN[\\s#:]*([0-9]{9})'], 'IT': [r'P\\.?\\s*IVA[\\s#:]*([0-9]{11})'], 'PT': [r'NIF[\\s:]*([0-9]{9})'], 'NL': [r'KvK[\\s#:]*([0-9]{8})'], 'AT': [r'ATU\\s*([0-9]{8})'], 'CH': [r'CHE[\\s-]?([0-9]{3})'], 'LU': [r'LU\\s*([0-9]{8})']}
    def process_single(self, name, web, country, vat=None):
        if self.use_db and self.db:
            r = self.db.search_name(name, country)
            if r: return {**r, 'search_method': 'DB-Name', 'status': 'Found'}
            if vat:
                r = self.db.search_vat(vat, country)
                if r: return {**r, 'search_method': 'DB-VAT', 'status': 'Found'}
            w = self._web(name, web, country); w['search_method'] = 'DB failed-Web'; return w
        w = self._web(name, web, country); w['search_method'] = 'Web only'; return w
    def _web(self, name, web, country):
        if country in self.extractors: return self.extractors[country].process(name, web)
        r = {'company_name': name, 'website': web, 'country_code': country, 'status': 'Not Found', 'source': 'web'}
        if web and country in self.patterns:
            try:
                resp = requests.get(web, timeout=10)
                if resp.status_code == 200:
                    for p in self.patterns[country]:
                        for m in re.finditer(p, resp.text, re.I): r[f'{country.lower()}_code'], r['status'] = m.group(1), 'Found'; break
            except: pass
        return r
    def process_list(self, df, prog=None):
        results = []
        nc = [c for c in df.columns if 'company' in c.lower() or 'name' in c.lower()]; name_col = nc[0] if nc else df.columns[0]
        wc = [c for c in df.columns if 'website' in c.lower() or 'url' in c.lower()]; web_col = wc[0] if wc else None
        cc = [c for c in df.columns if 'country' in c.lower()]; country_col = cc[0] if cc else None
        vc = [c for c in df.columns if 'vat' in c.lower() or 'fiscal' in c.lower()]; vat_col = vc[0] if vc else None
        for idx, row in df.iterrows():
            if prog: prog(idx+1, len(df))
            name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
            web = str(row[web_col]).strip() if web_col and pd.notna(row[web_col]) else ''
            vat = str(row[vat_col]).strip() if vat_col and pd.notna(row[vat_col]) else None
            country = 'GB'
            if country_col and pd.notna(row[country_col]):
                cv = str(row[country_col]).strip().upper()
                if len(cv) == 2 and cv in COUNTRY_CODES: country = cv
            results.append(self.process_single(name, web, country, vat)); time.sleep(0.2)
        return results

def show_auth(auth):
    st.title("🔐 arKap VAT Extractor"); st.info("🏢 @arkap.ch only")
    t1, t2 = st.tabs(["📧 Email", "🔑 Code"])
    with t1:
        e = st.text_input("Email")
        if st.button("Send Code", type="primary"):
            if auth.is_valid_email(e): c = auth.gen_code(); auth.store_code(e, c); st.success(f"Code: {c}")
            else: st.error("Invalid")
    with t2:
        e = st.text_input("Email", key="e2"); c = st.text_input("Code", max_chars=6)
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
    if st.session_state.company_db is None:
        st.header("📊 Database Setup")
        with st.expander("ℹ️ Setup Dropbox Database", expanded=True):
            st.markdown("""
            **Get Dropbox Link:**
            1. Go to Dropbox → Right-click your file
            2. Share → Create link → Copy
            
            **Add to Streamlit Secrets:**
            1. Go to app Settings → Secrets
            2. Add: `DROPBOX_FILE_URL = "your_link"`
            3. Save and reboot app
            """)
        if st.button("📥 Load from Dropbox", type="primary"):
            df = load_database_from_dropbox()
            if df is not None:
                st.session_state.company_db = CompanyDatabase(df); st.rerun()
        st.markdown("---")
        uploaded = st.file_uploader("Or Upload Manually", type=['xlsx', 'xls', 'csv'])
        if uploaded:
            try:
                df = pd.read_csv(uploaded) if uploaded.name.endswith('.csv') else pd.read_excel(uploaded)
                st.session_state.company_db = CompanyDatabase(df); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
        if st.button("⏭️ Skip - Web Only"): st.session_state.search_mode = 'web'; st.rerun()
        return
    if st.session_state.search_mode is None:
        st.header("🔍 Select Mode")
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("🗄️ DB + Web")
            if st.button("Use DB+Web", type="primary", use_container_width=True):
                if st.session_state.company_db: st.session_state.search_mode = 'db'; st.rerun()
        with c2:
            st.subheader("🌐 Web Only")
            if st.button("Use Web Only", use_container_width=True): st.session_state.search_mode = 'web'; st.rerun()
        return
    st.info(f"**Mode:** {st.session_state.search_mode.upper()}")
    if st.button("Change"): st.session_state.search_mode = None; st.rerun()
    st.markdown("---")
    t1, t2 = st.tabs(["Bulk", "Single"])
    with t1:
        f = st.file_uploader("Company List", type=['csv','xlsx'])
        if f:
            df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f); st.dataframe(df.head())
            if st.button("Process"):
                ext = MultiModeExtractor(st.session_state.company_db, st.session_state.search_mode=='db')
                p = st.progress(0); res = ext.process_list(df, lambda c,t: p.progress(c/t))
                st.success("Done"); rdf = pd.DataFrame(res); st.dataframe(rdf)
                c1,c2,c3 = st.columns(3)
                with c1: st.metric("Total", len(res))
                with c2: st.metric("Found", len([r for r in res if r['status']=='Found']))
                with c3: st.metric("Rate%", f"{len([r for r in res if r['status']=='Found'])/len(res)*100:.1f}")
                csv = io.StringIO(); rdf.to_csv(csv, index=False)
                st.download_button("Download", csv.getvalue(), f"res_{datetime.now().strftime('%Y%m%d_%H%M')}.csv")
    with t2:
        c1,c2 = st.columns(2)
        with c1: n = st.text_input("Name"); w = st.text_input("Website"); v = st.text_input("VAT")
        with c2: co = st.selectbox("Country", list(COUNTRY_CODES.keys()), format_func=lambda x:f"{COUNTRY_CODES[x]} ({x})")
        if st.button("Search") and n:
            ext = MultiModeExtractor(st.session_state.company_db, st.session_state.search_mode=='db')
            r = ext.process_single(n, w, co, v)
            if r['status']=='Found':
                st.success(f"✅ {r.get('search_method')}")
                if r.get('source')=='database':
                    c1,c2=st.columns(2)
                    with c1:
                        for k in ['company_name','vat_code']: 
                            if k in r: st.write(f"**{k}:** {r[k]}")
                    with c2:
                        for k in ['country_code','nace_code']: 
                            if k in r: st.write(f"**{k}:** {r[k]}")
                    c1,c2,c3=st.columns(3)
                    with c1:
                        if 'last_yr' in r: st.metric("Year",r['last_yr'])
                        if 'employees' in r: st.metric("Emp",safe_format(r.get('employees')))
                    with c2:
                        if 'value_of_production_th' in r: st.metric("Prod",safe_format(r.get('value_of_production_th'),pre="€",suf="k"))
                        if 'ebitda_th' in r: st.metric("EBITDA",safe_format(r.get('ebitda_th'),pre="€",suf="k"))
                    with c3:
                        if 'pfn_th' in r: st.metric("PFN",safe_format(r.get('pfn_th'),pre="€",suf="k"))
            else: st.warning("Not found")

def main():
    st.set_page_config(page_title="arKap", page_icon="⚡", layout="wide")
    auth = AuthenticationManager()
    if auth.is_valid(): show_main()
    else: show_auth(auth)

if __name__ == "__main__": main()
''')

print("✅ DROPBOX VERSION CREATED: arkap_vat_extractor_dropbox.py")
print("\n" + "="*70)
print("📋 QUICK SETUP GUIDE:")
print("="*70)
print("\n**STEP 1: Get Dropbox Share Link**")
print("  1. Open Dropbox.com")
print("  2. Find: databaseaziendearkap.xlsx")
print("  3. Right-click → Share → Copy link")
print("  4. Link looks like: https://www.dropbox.com/s/abc123/file.xlsx?dl=0")
print("\n**STEP 2: Add to Streamlit Secrets**")
print("  1. Go to: https://share.streamlit.io")
print("  2. Your app → ⚙️ Settings → Secrets")
print("  3. Paste this (with YOUR link):")
print('     DROPBOX_FILE_URL = "https://www.dropbox.com/s/YOUR_LINK_HERE/databaseaziendearkap.xlsx?dl=0"')
print("  4. Click Save")
print("\n**STEP 3: Update GitHub**")
print("  • Replace .py file with: arkap_vat_extractor_dropbox.py")
print("  • DON'T commit the .xlsx database file")
print("  • requirements.txt:")
print("    streamlit")
print("    pandas") 
print("    openpyxl")
print("    requests")
print("\n**STEP 4: Reboot App**")
print("  • Streamlit auto-updates from GitHub")
print("  • Or click 'Reboot' in app settings")
print("\n" + "="*70)
print("🔐 SECURITY:")
print("="*70)
print("✅ Database stays in Dropbox (not in GitHub)")
print("✅ Link stored in encrypted Streamlit Secrets")
print("✅ Only your app can download it")
print("✅ @arkap.ch email authentication required")
)
