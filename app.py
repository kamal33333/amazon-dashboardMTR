import streamlit as st
import pandas as pd
import zipfile
import io
import re
import numpy as np
import calendar
from datetime import datetime
from scipy.stats import linregress

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Amazon MTR Analytics Engine",
    page_icon="📦",
    layout="wide"
)

# ==========================================
# CORE ENGINE CLASS (Adapted for Streamlit)
# ==========================================
class AmazonAnalyticsEngine:
    def __init__(self, uploaded_zips, inv_file):
        self.uploaded_zips = uploaded_zips # List of UploadedFile objects
        self.inv_file = inv_file           # UploadedFile object or None
        self.monthly_files = []
        self.master_df = pd.DataFrame()
        self.inventory_df = pd.DataFrame()
        self.wh_region_map = {} 
        self.wh_state_map = {}
        self.projection_factor = 1.0
        self.latest_month_name = ""
        self.quarter_map = {} 

    # --- UTILITIES ---
    def _parse_file_info(self, filename):
        fn = filename.lower()
        if not fn.endswith('.zip'): return None, "Not .zip"
        channel = 'B2B' if 'b2b' in fn else 'B2C' if 'b2c' in fn else None
        if not channel: return None, "No Channel"
        months = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
        m_num = next((val for key, val in months.items() if key in fn), None)
        if not m_num: return None, "No Month"
        year_match = re.search(r'202[0-9]', fn)
        year = int(year_match.group(0)) if year_match else datetime.now().year
        return (datetime(year, m_num, 1), channel), "Success"

    def _extract_brand(self, description):
        desc = str(description).upper()
        if "KRESSA" in desc: return "Kressa"
        if "TLC365" in desc or "TLC 365" in desc: return "TLC365"
        if re.search(r"K[-\s]+ONE", desc): return "K One"
        if "KUANTUM" in desc: return "Kuantum Bond"
        return "Other"

    def _determine_region(self, state):
        st = str(state).upper().strip()
        alias = {'JAMMU & KASHMIR': 'JAMMU AND KASHMIR', 'ORISSA': 'ODISHA', 'ANDAMAN & NICOBAR': 'ANDAMAN AND NICOBAR ISLANDS', 'ANDAMAN & NICOBAR ISLANDS': 'ANDAMAN AND NICOBAR ISLANDS', 'DADRA & NAGAR HAVELI': 'DADRA AND NAGAR HAVELI AND DAMAN AND DIU', 'PONDICHERRY': 'PUDUCHERRY', 'TELENGANA': 'TELANGANA'}
        st = alias.get(st, st)
        r1 = ['DELHI','HARYANA','HIMACHAL PRADESH','JAMMU AND KASHMIR','LADAKH','PUNJAB','RAJASTHAN','UTTAR PRADESH','UTTARAKHAND','CHANDIGARH']
        r2 = ['GUJARAT','MAHARASHTRA','GOA','MADHYA PRADESH','DADRA AND NAGAR HAVELI AND DAMAN AND DIU']
        r3 = ['ANDHRA PRADESH','KARNATAKA','KERALA','TAMIL NADU','TELANGANA','PUDUCHERRY','LAKSHADWEEP','ANDAMAN AND NICOBAR ISLANDS']
        r4 = ['WEST BENGAL','BIHAR','JHARKHAND','ODISHA','CHHATTISGARH','ASSAM','ARUNACHAL PRADESH','MANIPUR','MEGHALAYA','MIZORAM','NAGALAND','SIKKIM','TRIPURA']
        if st in r1: return 'Region 1'
        if st in r2: return 'Region 2'
        if st in r3: return 'Region 3'
        if st in r4: return 'Region 4'
        return 'Unmapped'

    def _apply_regional_logic(self, row):
        state = str(row['Ship To State']).upper().strip()
        city = str(row['Ship To City']).upper().strip()
        up_b = ['AYODHYA','AZAMGARH','BASTI','DEVIPATAN','GORAKHPUR','MIRZAPUR','PRAYAGRAJ','VARANASI','ALLAHABAD']
        region = self._determine_region(state)
        final_state = state
        if state == 'UTTAR PRADESH':
            if any(c in city for c in up_b): final_state, region = 'Uttar Pradesh Zone B', 'Region 4'
        return pd.Series([final_state, region])

    def _get_fiscal_quarter(self, date_obj):
        m, y = date_obj.month, date_obj.year
        q_key = f"{y}-Q1" if 4<=m<=6 else f"{y}-Q2" if 7<=m<=9 else f"{y}-Q3" if 10<=m<=12 else f"{y}-Q4"
        q_name = f"Q1 FY{str(y+1)[-2:]}" if 4<=m<=6 else f"Q2 FY{str(y+1)[-2:]}" if 7<=m<=9 else f"Q3 FY{str(y+1)[-2:]}" if 10<=m<=12 else f"Q4 FY{str(y)[-2:]}"
        return q_key, q_name

    # --- DATA LOADING ---
    def load_data(self):
        if not self.uploaded_zips: 
            return False, "No zip files provided."
        
        # 1. Map files from Streamlit in-memory uploads
        f_map = {}
        for uploaded_file in self.uploaded_zips:
            res, msg = self._parse_file_info(uploaded_file.name)
            if res:
                dt, ch = res
                if dt not in f_map: f_map[dt] = {'date': dt, 'B2B': None, 'B2C': None}
                # Store the uploaded file object instead of path
                f_map[dt][ch] = uploaded_file 
        
        self.monthly_files = sorted(f_map.values(), key=lambda x: x['date'])
        if not self.monthly_files: 
            return False, "No valid MTR zip files found matching the naming convention (e.g., must contain 'B2B'/'B2C' and month like 'Jan')."
        
        self.latest_month_name = self.monthly_files[-1]['date'].strftime('%b-%Y')
        dfs, now, q_days = [], datetime.now(), {}
        
        # 2. Process Files in memory
        for i, m in enumerate(self.monthly_files):
            m_name, (qk, qn) = m['date'].strftime('%b-%Y'), self._get_fiscal_quarter(m['date'])
            c_dfs = []
            for ch in ['B2B', 'B2C']:
                file_obj = m[ch]
                if file_obj:
                    # Read zip from memory
                    with zipfile.ZipFile(file_obj, 'r') as z:
                        csv_files = [f for f in z.namelist() if f.endswith('.csv')]
                        if csv_files:
                            csv = csv_files[0]
                            with z.open(csv) as f_in:
                                tmp = pd.read_csv(f_in)
                                tmp['Channel'] = ch
                                c_dfs.append(tmp)
            
            if not c_dfs: continue
            df = pd.concat(c_dfs, ignore_index=True)
            df.columns = df.columns.str.strip()
            
            # Key fixes
            if 'Tax Exclusive Gross' in df.columns:
                df['Revenue'] = pd.to_numeric(df['Tax Exclusive Gross'], errors='coerce').fillna(0)
            else:
                df['Revenue'] = 0.0 # Fallback
                
            if 'Quantity' in df.columns:
                df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
            else:
                df['Quantity'] = 0.0

            if 'Transaction Type' in df.columns:
                df['Transaction Type'] = df['Transaction Type'].astype(str).fillna('')
                df.loc[df['Transaction Type'].str.contains('Cancel', case=False), 'Quantity'] = 0
                df.loc[df['Transaction Type'].str.contains('Refund', case=False), 'Quantity'] *= -1

            df['Month_Year'], df['Fiscal_Quarter'], df['Date_Obj'] = m_name, qn, m['date']
            if 'Sku' in df.columns:
                df['Sku'] = df['Sku'].astype(str).str.strip().str.upper()
            
            # Warehouse Logic
            if 'Warehouse Id' in df.columns and 'Ship From State' in df.columns:
                wh_pairs = df[['Warehouse Id', 'Ship From State']].dropna().drop_duplicates()
                for _, row in wh_pairs.iterrows():
                    wh_id, state = row['Warehouse Id'], row['Ship From State'].upper().strip()
                    if wh_id not in self.wh_region_map: self.wh_region_map[wh_id] = self._determine_region(state)
                    if wh_id not in self.wh_state_map: self.wh_state_map[wh_id] = state

            # Projection Logic
            _, d_in_m = calendar.monthrange(m['date'].year, m['date'].month)
            if i == len(self.monthly_files) - 1 and (m['date'].year == now.year and m['date'].month == now.month):
                d_col = 'Invoice Date' if 'Invoice Date' in df.columns else 'Order Date' if 'Order Date' in df.columns else None
                if d_col:
                    dates = pd.to_datetime(df[d_col], dayfirst=True, errors='coerce')
                    if not dates.isnull().all():
                        d_passed = dates.max().day
                        if d_passed < d_in_m: self.projection_factor = d_in_m / d_passed
            
            q_days[qn] = q_days.get(qn, 0) + d_in_m
            dfs.append(df)
            
        if not dfs:
            return False, "Failed to read CSVs from the provided zip files."
            
        self.master_df = pd.concat(dfs, ignore_index=True)
        
        if 'Ship To State' in self.master_df.columns and 'Ship To City' in self.master_df.columns:
            self.master_df[['Final State', 'Region']] = self.master_df.apply(self._apply_regional_logic, axis=1)
            self.master_df['Ship To City'] = self.master_df['Ship To City'].astype(str).str.upper().str.strip()
        else:
            self.master_df['Final State'] = 'Unknown'
            self.master_df['Region'] = 'Unknown'
            
        self.quarter_map = q_days
        return True, "Data loaded successfully!"

    def load_inventory(self):
        if not self.inv_file: return False
        try:
            inv = pd.read_csv(self.inv_file)
            if 'Disposition' in inv.columns and 'Location' in inv.columns:
                inv = inv[(inv['Disposition'] == 'SELLABLE') & (inv['Location'] != 'VNDV')]
                inv['Region'] = inv['Location'].map(self.wh_region_map).fillna('Unmapped')
                inv['State'] = inv['Location'].map(self.wh_state_map).fillna('Unmapped')
                self.inventory_df = inv
                return True
            return False
        except Exception as e:
            st.warning(f"Failed to load inventory: {e}")
            return False

    # --- ANALYTICS ENGINE ---
    def _clean_and_group(self):
        df = self.master_df.copy()
        if 'Sku' in df.columns and 'Item Description' in df.columns:
            sku_desc = df.groupby('Sku')['Item Description'].agg(lambda x: max(x.astype(str), key=len)).to_dict()
            df['Item Description'] = df['Sku'].map(sku_desc)
            df['Brand'] = df['Item Description'].apply(self._extract_brand)
        return df

    def _calc_stats(self, row, cols):
        baseline_cols = cols[:-1] if len(cols) > 1 else cols
        vals = row[baseline_cols].values.astype(float)
        nz = np.nonzero(vals > 0)[0]
        if len(nz) == 0: return 0.0, 0.0, 0.0
        win = vals[nz[0]:]
        avg, std = np.mean(win), np.std(win)
        slope = linregress(np.arange(len(win)), win)[0] if len(win) > 1 else 0.0
        return avg, std, slope

    def _safe_growth(self, cur, prev):
        return (cur - prev) / prev if prev > 0 else np.nan

    def _gen_metric(self, groupby, metric):
        df = self._clean_and_group()
        
        # Check if necessary columns exist
        missing = [col for col in groupby + ['Month_Year', metric] if col not in df.columns]
        if missing:
            return pd.DataFrame()
            
        piv = df.groupby(groupby + ['Month_Year'])[metric].sum().reset_index().pivot_table(index=groupby, columns='Month_Year', values=metric, aggfunc='sum').fillna(0)
        cols = [m['date'].strftime('%b-%Y') for m in self.monthly_files if m['date'].strftime('%b-%Y') in piv.columns]
        if not cols: return pd.DataFrame()
        
        piv = piv[cols]; an_df = piv.copy()
        if self.projection_factor > 1.05: an_df[cols[-1]] *= self.projection_factor
        stats = an_df.apply(lambda r: self._calc_stats(r, cols), axis=1)
        suf = "Revenue" if metric == "Revenue" else "Units"
        
        piv[f'Hist Avg {suf}'] = stats.apply(lambda x: x[0])
        piv[f'{suf} Std Dev'] = stats.apply(lambda x: x[1])
        piv[f'{suf} Vol %'] = (piv[f'{suf} Std Dev'] / piv[f'Hist Avg {suf}'].replace(0, 1)) * 100
        piv[f'{suf} Trend'] = stats.apply(lambda x: x[2])
        
        std_safe = piv[f'{suf} Std Dev'].replace(0, 1) 
        raw_z = (an_df[cols[-1]] - piv[f'Hist Avg {suf}']) / std_safe
        piv[f'{suf} Z-Score'] = raw_z.apply(lambda z: 0 if np.isinf(z) else min(max(z, -10), 10)) 
        
        prev_col = cols[-2] if len(cols) > 1 else cols[-1]
        piv[f'{suf} Growth %'] = an_df.apply(lambda r: self._safe_growth(r[cols[-1]], r[prev_col]), axis=1)
        piv[f'Total {suf}'] = an_df.sum(axis=1) 
        
        if self.projection_factor > 1.0: 
            piv[f'Projected {cols[-1]} ({suf})'] = piv[cols[-1]] * self.projection_factor
            
        return piv.rename(columns={c: f"{c} {suf}" for c in cols})

    def generate_combined(self, groupby):
        rev = self._gen_metric(groupby, 'Revenue')
        qty = self._gen_metric(groupby, 'Quantity')
        
        if rev.empty or qty.empty: return pd.DataFrame()
        
        res = pd.merge(rev, qty, left_index=True, right_index=True, how='outer').fillna(0)
        c_rev = f"{self.latest_month_name} Revenue"
        if c_rev in res.columns:
            total = res[c_rev].sum()
            if total > 0:
                res['ABC Class'] = res.sort_values(c_rev, ascending=False)[c_rev].cumsum().apply(lambda x: "A" if x/total <= 0.8 else "B" if x/total <= 0.95 else "C")
            else:
                res['ABC Class'] = "C"
            res['Volatility Class'] = res['Revenue Vol %'].apply(lambda v: "🟢 Stable" if v < 20 else "🟡 Variable" if v < 50 else "🔴 Erratic")
            return res.sort_values(c_rev, ascending=False)
        return res

    def generate_supply_demand(self):
        if self.inventory_df.empty: return pd.DataFrame(), pd.DataFrame()
        latest = self._clean_and_group()
        if 'Month_Year' not in latest.columns: return pd.DataFrame(), pd.DataFrame()
        
        latest = latest[latest['Month_Year'] == self.latest_month_name]
        
        r_comp = pd.DataFrame()
        s_comp = pd.DataFrame()
        
        # Region
        if all(c in latest.columns for c in ['Brand', 'Sku', 'Item Description', 'Region', 'Quantity']) and all(c in self.inventory_df.columns for c in ['MSKU', 'Region', 'Ending Warehouse Balance']):
            r_sales = latest.groupby(['Brand', 'Sku', 'Item Description', 'Region'])['Quantity'].sum().unstack(fill_value=0).add_suffix(' Sales')
            r_stock = self.inventory_df.groupby(['MSKU', 'Region'])['Ending Warehouse Balance'].sum().unstack(fill_value=0).add_suffix(' Stock')
            r_stock.index.name = 'Sku'
            r_comp = pd.merge(r_sales.reset_index(), r_stock, on='Sku', how='outer').fillna(0)
            r_comp['Grand Total Sales'] = r_comp[[c for c in r_comp.columns if 'Sales' in c]].sum(axis=1)
            r_comp['Grand Total Stock'] = r_comp[[c for c in r_comp.columns if 'Stock' in c]].sum(axis=1)
            r_comp = r_comp.sort_values('Grand Total Sales', ascending=False)
            
        # State
        if all(c in latest.columns for c in ['Brand', 'Sku', 'Item Description', 'Final State', 'Quantity']) and all(c in self.inventory_df.columns for c in ['MSKU', 'State', 'Ending Warehouse Balance']):
            wh_sts = [s for s in self.inventory_df['State'].unique() if s != 'Unmapped']
            s_sales = latest[latest['Final State'].isin(wh_sts)].groupby(['Brand', 'Sku', 'Item Description', 'Final State'])['Quantity'].sum().unstack(fill_value=0).add_suffix(' Sales')
            s_stock = self.inventory_df.groupby(['MSKU', 'State'])['Ending Warehouse Balance'].sum().unstack(fill_value=0).add_suffix(' Stock')
            s_stock.index.name = 'Sku'
            s_comp = pd.merge(s_sales.reset_index(), s_stock, on='Sku', how='outer').fillna(0)
            s_comp['Grand Total Sales'] = s_comp[[c for c in s_comp.columns if 'Sales' in c]].sum(axis=1)
            s_comp['Grand Total Stock'] = s_comp[[c for c in s_comp.columns if 'Stock' in c]].sum(axis=1)
            s_comp = s_comp.sort_values('Grand Total Sales', ascending=False)
            
        return r_comp, s_comp

    def generate_excel_bytes(self):
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        
        reg_s, st_s = self.generate_combined(['Region']), self.generate_combined(['Final State'])
        city_s = self.generate_combined(['Final State', 'Ship To City'])
        prod_s = self.generate_combined(['Brand', 'Sku', 'Item Description'])
        
        insights = []
        if not prod_s.empty and 'Revenue Z-Score' in prod_s.columns and 'Revenue Vol %' in prod_s.columns:
            for idx, row in prod_s.iterrows():
                # Ensure idx has at least 3 elements before accessing idx[2]
                if isinstance(idx, tuple) and len(idx) >= 3:
                    ent = str(idx[2])
                    if row['Revenue Z-Score'] > 2.0: insights.append({'Type': '🔥 Breakout', 'Entity': ent[:30], 'Note': f"Z: {round(row['Revenue Z-Score'], 1)}"})
                    if row['Revenue Vol %'] > 80: insights.append({'Type': '📦 Erratic Demand', 'Entity': ent[:30], 'Note': f"Vol: {int(row['Revenue Vol %'])}%"})

        reg_comp, st_comp = self.generate_supply_demand()

        # Formatting
        wb = writer.book
        f_h = wb.add_format({'bold':True, 'fg_color':'#203764', 'font_color':'white', 'border':1})
        f_c = wb.add_format({'num_format':'[$₹-4009]#,##0', 'border':1})
        f_i = wb.add_format({'num_format':'#,##0', 'border':1})
        f_p = wb.add_format({'num_format':'0.0%', 'border':1})
        f_z = wb.add_format({'num_format':'0.00', 'border':1, 'font_color':'#0000FF'})

        def write_ws(df, name):
            if df is None or df.empty: return
            d_df = df.copy()
            for c in d_df.columns:
                if 'Growth' in str(c): d_df[c] = d_df[c].apply(lambda x: "★ New" if pd.isna(x) else x)
            if isinstance(d_df.index, pd.Index) and d_df.index.name: d_df.reset_index(inplace=True)
            elif isinstance(d_df.index, pd.MultiIndex): d_df.reset_index(inplace=True)
            
            # Avoid invalid sheet names
            name = name[:31].replace(':', '').replace('/', '').replace('\\', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
            
            d_df.to_excel(writer, sheet_name=name, startrow=1, header=False, index=False)
            ws = writer.sheets[name]
            for i, col in enumerate(d_df.columns):
                ws.write(0, i, col, f_h); ws.set_column(i, i, max(len(str(col))+2, 12))
                c_str = str(col)
                if 'Z-Score' in c_str: ws.set_column(i, i, 10, f_z)
                elif any(x in c_str for x in ['%', 'Growth', 'Share', 'Vol', 'Concentration']): ws.set_column(i, i, 10, f_p)
                elif any(x in c_str for x in ['Revenue', 'Avg', 'Price']): ws.set_column(i, i, 15, f_c)
                else: ws.set_column(i, i, 10, f_i)

        # Executive Summary
        if 'Month_Year' in self.master_df.columns and 'Revenue' in self.master_df.columns and 'Quantity' in self.master_df.columns:
            latest_rev = self.master_df[self.master_df['Month_Year'] == self.latest_month_name]['Revenue'].sum()
            latest_qty = self.master_df[self.master_df['Month_Year'] == self.latest_month_name]['Quantity'].sum()
            summ = pd.DataFrame([
                {'Metric':'Total Revenue (YTD)', 'Val':self.master_df['Revenue'].sum()},
                {'Metric':'Total Units (YTD)', 'Val':self.master_df['Quantity'].sum()},
                {'Metric':f'Revenue ({self.latest_month_name})', 'Val':latest_rev},
                {'Metric':f'Units ({self.latest_month_name})', 'Val':latest_qty},
                {'Metric':'Projected Monthly Run Rate', 'Val':latest_rev * self.projection_factor}
            ])
            summ.to_excel(writer, 'EXECUTIVE SUMMARY', index=False)
        
        write_ws(pd.DataFrame(insights), 'AI Insights')
        write_ws(reg_comp, 'Regional Supply vs Demand')
        write_ws(st_comp, 'Local State Distribution')
        write_ws(reg_s, 'Regional Dashboard')
        write_ws(st_s, 'State Insights')
        write_ws(city_s, 'City Insights')
        write_ws(prod_s, 'All Products')
        
        if not self.master_df.empty:
            # Drop very large columns or limit rows if needed for performance, but writing all for now
            self.master_df.to_excel(writer, 'Combined Data', index=False)
            
        writer.close()
        processed_data = output.getvalue()
        return processed_data


# ==========================================
# STREAMLIT UI
# ==========================================
def main():
    st.title("📦 Amazon MTR Analytics Engine")
    st.markdown("Upload your Amazon MTR ZIP files and Inventory CSV to generate the Ultimate Command Excel report.")

    with st.sidebar:
        st.header("1. Upload Data")
        
        uploaded_zips = st.file_uploader(
            "Upload MTR Zip Files (B2B/B2C)", 
            type=["zip"], 
            accept_multiple_files=True,
            help="Select all the monthly B2B and B2C zip files you want to analyze."
        )
        
        inv_file = st.file_uploader(
            "Upload Inventory Report (Optional)", 
            type=["csv"],
            help="Upload the Amazoninventoryrep.csv file to unlock Supply vs Demand insights."
        )
        
        st.divider()
        process_btn = st.button("🚀 Process Data", use_container_width=True, type="primary")

    # Main Area
    if process_btn:
        if not uploaded_zips:
            st.error("Please upload at least one MTR Zip file in the sidebar to proceed.")
            return

        with st.spinner("Analyzing data and generating report. This may take a minute..."):
            # Initialize engine with uploaded files
            engine = AmazonAnalyticsEngine(uploaded_zips, inv_file)
            
            # Load Sales Data
            success, msg = engine.load_data()
            if not success:
                st.error(msg)
                return
                
            # Load Inventory
            if inv_file:
                inv_success = engine.load_inventory()
                if inv_success:
                    st.toast("Inventory data loaded successfully!", icon="✅")
                else:
                    st.toast("Could not process inventory file. Proceeding with sales data only.", icon="⚠️")

            # Generate Excel file in memory
            try:
                excel_data = engine.generate_excel_bytes()
                
                st.success("✅ Analysis Complete! Your report is ready to download.")
                
                # Show some quick stats on screen
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Files Processed", len(engine.monthly_files) * 2)
                col2.metric("Total Records", f"{len(engine.master_df):,}")
                col3.metric("Latest Month", engine.latest_month_name)
                
                st.download_button(
                    label="⬇️ Download Ultimate Command Report (Excel)",
                    data=excel_data,
                    file_name=f"MTR_Ultimate_Command_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
            except Exception as e:
                st.error(f"An error occurred during report generation: {str(e)}")
                st.exception(e) # Remove this in production if you want to hide traceback

if __name__ == "__main__":
    main()
