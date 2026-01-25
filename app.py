import streamlit as st
import pandas as pd
import zipfile
import io
import re
import numpy as np
import calendar
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# 1. CONFIG & AUTH
# ==========================================
st.set_page_config(page_title="Amazon Management HQ", page_icon="📈", layout="wide")

PASSWORD = "kressa_admin" 

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def check_password():
    if st.session_state.authenticated:
        return True
    
    c1, c2, c3 = st.columns([1,2,1])
    with c2:
        st.title("🔒 Amazon HQ Login")
        pwd = st.text_input("Enter Password", type="password")
        if st.button("Login"):
            if pwd == PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Incorrect Password")
    return False

if not check_password():
    st.stop()

# ==========================================
# 2. DATA PROCESSING ENGINE
# ==========================================

def parse_file_info(filename):
    filename_clean = filename.lower()
    if 'b2b' in filename_clean: channel = 'B2B'
    elif 'b2c' in filename_clean: channel = 'B2C'
    else: return None, None 

    months = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
    }
    
    month_num = None
    for m_str, m_int in months.items():
        if m_str in filename_clean:
            month_num = m_int
            break
            
    year_match = re.search(r'202[0-9]', filename_clean)
    year = int(year_match.group(0)) if year_match else datetime.now().year
    
    if month_num:
        return datetime(year, month_num, 1), channel
    return None, None

def read_csv_from_uploaded_zip(uploaded_file):
    try:
        with zipfile.ZipFile(uploaded_file) as z:
            csv_files = [f for f in z.namelist() if f.endswith('.csv')]
            if not csv_files: return None
            with z.open(csv_files[0]) as f:
                return pd.read_csv(f)
    except Exception: return None

def extract_brand(description):
    desc = str(description).upper().strip()
    if "KRESSA" in desc: return "Kressa"
    if "TLC365" in desc or "TLC 365" in desc: return "TLC365"
    if re.search(r"K[-\s]+ONE", desc): return "K One"
    if "KUANTUM" in desc: return "Kuantum Bond"
    return "Other"

def apply_regional_logic(row):
    state_raw = str(row['Ship To State']).upper().strip()
    city = str(row['Ship To City']).upper().strip()
    
    alias_map = {
        'JAMMU & KASHMIR': 'JAMMU AND KASHMIR', 'ORISSA': 'ODISHA',
        'ANDAMAN & NICOBAR': 'ANDAMAN AND NICOBAR ISLANDS',
        'ANDAMAN & NICOBAR ISLANDS': 'ANDAMAN AND NICOBAR ISLANDS', 
        'DADRA & NAGAR HAVELI': 'DADRA AND NAGAR HAVELI AND DAMAN AND DIU',
        'PONDICHERRY': 'PUDUCHERRY', 'TELENGANA': 'TELANGANA'
    }
    state = alias_map.get(state_raw, state_raw)
    up_zone_b_cities = ['AYODHYA', 'AZAMGARH', 'BASTI', 'DEVIPATAN', 'GORAKHPUR', 'MIRZAPUR', 'PRAYAGRAJ', 'VARANASI', 'ALLAHABAD']
    r1 = ['DELHI', 'HARYANA', 'HIMACHAL PRADESH', 'JAMMU AND KASHMIR', 'LADAKH', 'PUNJAB', 'RAJASTHAN', 'UTTAR PRADESH', 'UTTARAKHAND', 'CHANDIGARH']
    r2 = ['GUJARAT', 'MAHARASHTRA', 'GOA', 'MADHYA PRADESH', 'DADRA AND NAGAR HAVELI AND DAMAN AND DIU']
    r3 = ['ANDHRA PRADESH', 'KARNATAKA', 'KERALA', 'TAMIL NADU', 'TELANGANA', 'PUDUCHERRY', 'LAKSHADWEEP', 'ANDAMAN AND NICOBAR ISLANDS']
    r4 = ['WEST BENGAL', 'BIHAR', 'JHARKHAND', 'ODISHA', 'CHHATTISGARH', 'ASSAM', 'ARUNACHAL PRADESH', 'MANIPUR', 'MEGHALAYA', 'MIZORAM', 'NAGALAND', 'SIKKIM', 'TRIPURA']

    region = 'Unmapped'; final_state = state
    if state == 'UTTAR PRADESH':
        if any(up_city in city for up_city in up_zone_b_cities):
            final_state = 'Uttar Pradesh Zone B'; region = 'Region 4'
        else: region = 'Region 1'
    elif state in r1: region = 'Region 1'
    elif state in r2: region = 'Region 2'
    elif state in r3: region = 'Region 3'
    elif state in r4: region = 'Region 4'
    return final_state, region

# ==========================================
# 3. UI & FILE HANDLING
# ==========================================
st.title("📈 Amazon Management HQ")
st.markdown("Upload monthly ZIP files to generate the **Executive Business Review**.")

uploaded_files = st.file_uploader("Drop All Monthly ZIPs Here", type=['zip'], accept_multiple_files=True)

if uploaded_files and st.button("🚀 Generate Management Report"):
    
    # --- 1. FILE ORG ---
    files_map = {}
    with st.spinner("Organizing timeline..."):
        for uploaded_file in uploaded_files:
            date_obj, channel = parse_file_info(uploaded_file.name)
            if date_obj and channel:
                if date_obj not in files_map: 
                    files_map[date_obj] = {'date': date_obj, 'B2B': None, 'B2C': None}
                files_map[date_obj][channel] = uploaded_file

        sorted_files = sorted(files_map.values(), key=lambda x: x['date'])
        if not sorted_files:
            st.error("No valid dated files found.")
            st.stop()

    # --- 2. LOAD & STITCH ---
    all_dfs = []
    now = datetime.now()
    proj_factors = {} # Store projection factor per month

    progress = st.progress(0)
    for i, m_data in enumerate(sorted_files):
        month_name = m_data['date'].strftime('%b-%Y')
        b2b = read_csv_from_uploaded_zip(m_data['B2B']) if m_data['B2B'] else None
        b2c = read_csv_from_uploaded_zip(m_data['B2C']) if m_data['B2C'] else None
        
        if b2b is not None and b2c is not None:
            b2b['Channel'] = 'B2B'; b2c['Channel'] = 'B2C'
            curr_df = pd.concat([b2b, b2c], ignore_index=True)
            curr_df.columns = curr_df.columns.str.strip()
            
            # Clean
            curr_df['Revenue'] = pd.to_numeric(curr_df['Tax Exclusive Gross'], errors='coerce').fillna(0)
            curr_df['Quantity'] = pd.to_numeric(curr_df['Quantity'], errors='coerce').fillna(0)
            curr_df.loc[curr_df['Transaction Type'].str.contains('Cancel', case=False, na=False), 'Quantity'] = 0
            
            # Partial Month Logic
            factor = 1.0
            if i == len(sorted_files) - 1: # Only check last month
                is_curr = (m_data['date'].year == now.year) and (m_data['date'].month == now.month)
                if is_curr:
                    date_col = 'Invoice Date' if 'Invoice Date' in curr_df.columns else 'Order Date'
                    if date_col in curr_df.columns:
                        curr_df['Parsed_Date'] = pd.to_datetime(curr_df[date_col], dayfirst=True, errors='coerce')
                        max_d = curr_df['Parsed_Date'].max()
                        if pd.notnull(max_d):
                            days_in_data = max_d.day
                            _, days_in_month = calendar.monthrange(max_d.year, max_d.month)
                            if days_in_data < days_in_month and days_in_data > 0:
                                factor = days_in_month / days_in_data
            
            proj_factors[month_name] = factor
            curr_df['Revenue_Projected'] = curr_df['Revenue'] * factor
            curr_df['Month_Year'] = month_name
            curr_df['Sort_Date'] = m_data['date']
            curr_df['Brand'] = curr_df['Item Description'].apply(extract_brand)
            all_dfs.append(curr_df)
        progress.progress((i + 1) / len(sorted_files))
    
    if not all_dfs:
        st.error("Processing failed.")
        st.stop()

    master_df = pd.concat(all_dfs, ignore_index=True)
    master_df[['Final State', 'Region']] = master_df.apply(lambda row: pd.Series(apply_regional_logic(row)), axis=1)

    # ==========================================
    # 4. DASHBOARD TABS
    # ==========================================
    
    st.divider()
    
    # --- TAB 1: EXECUTIVE TRENDS ---
    st.header(f"📊 Executive Summary ({len(sorted_files)} Months)")
    
    # 1. B2B vs B2C Trend Stacked Bar
    trend_data = master_df.groupby(['Month_Year', 'Channel', 'Sort_Date'])['Revenue_Projected'].sum().reset_index()
    trend_data = trend_data.sort_values('Sort_Date')
    
    fig_trend = px.bar(
        trend_data, x='Month_Year', y='Revenue_Projected', color='Channel', 
        title="Monthly Revenue Trend (B2B vs B2C)", text_auto='.2s',
        color_discrete_map={'B2B': '#1f77b4', 'B2C': '#ff7f0e'}
    )
    st.plotly_chart(fig_trend, use_container_width=True)

    # 2. Key Metrics Table (MoM)
    monthly_metrics = master_df.groupby(['Sort_Date', 'Month_Year']).agg(
        Total_Revenue=('Revenue', 'sum'),
        Projected_Revenue=('Revenue_Projected', 'sum'),
        Total_Units=('Quantity', 'sum'),
        Orders=('Revenue', 'count')
    ).reset_index()
    
    monthly_metrics['AOV'] = monthly_metrics['Total_Revenue'] / monthly_metrics['Orders']
    monthly_metrics['MoM Growth'] = monthly_metrics['Total_Revenue'].pct_change() * 100
    
    # Format for display
    display_metrics = monthly_metrics[['Month_Year', 'Total_Revenue', 'Projected_Revenue', 'Total_Units', 'AOV', 'MoM Growth']].copy()
    display_metrics['Total_Revenue'] = display_metrics['Total_Revenue'].map('₹{:,.0f}'.format)
    display_metrics['Projected_Revenue'] = display_metrics['Projected_Revenue'].map('₹{:,.0f}'.format)
    display_metrics['AOV'] = display_metrics['AOV'].map('₹{:,.0f}'.format)
    display_metrics['MoM Growth'] = display_metrics['MoM Growth'].map('{:+.1f}%'.format).replace('nan%', '-')

    with st.expander("📄 View Detailed Monthly Metrics Table", expanded=True):
        st.dataframe(display_metrics, use_container_width=True)


    # --- TAB 2: MONTHLY HEROES ---
    st.divider()
    st.subheader("🏆 Hero of the Month")
    
    hero_list = []
    for m in monthly_metrics['Month_Year']:
        m_df = master_df[master_df['Month_Year'] == m]
        
        # Top Product
        top_prod = m_df.groupby('Item Description')['Revenue'].sum().idxmax()
        top_prod_rev = m_df.groupby('Item Description')['Revenue'].sum().max()
        
        # Top State
        top_state = m_df.groupby('Final State')['Revenue'].sum().idxmax()
        
        # Top Channel Split
        b2b_share = m_df[m_df['Channel'] == 'B2B']['Revenue'].sum() / m_df['Revenue'].sum()
        
        hero_list.append({
            'Month': m,
            '👑 Top Product': top_prod[:40] + '...',
            'Product Rev': f"₹{top_prod_rev:,.0f}",
            '🌍 Top State': top_state,
            '🏢 B2B Share': f"{b2b_share:.0%}"
        })
        
    st.dataframe(pd.DataFrame(hero_list), use_container_width=True)


    # --- TAB 3: REGIONAL & BRAND PERFORMANCE ---
    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("🌍 Revenue Heatmap by Region")
        latest_month = sorted_files[-1]['date'].strftime('%b-%Y')
        latest_df = master_df[master_df['Month_Year'] == latest_month]
        
        reg_fig = px.sunburst(
            latest_df, path=['Region', 'Final State'], values='Revenue',
            title=f"Regional Breakdown ({latest_month})"
        )
        st.plotly_chart(reg_fig, use_container_width=True)
        
    with c2:
        st.subheader("📦 Pareto Analysis (80/20 Rule)")
        # Calculate Product Contribution
        prod_perf = master_df.groupby('Item Description')['Revenue'].sum().sort_values(ascending=False).reset_index()
        prod_perf['Cumulative %'] = 100 * prod_perf['Revenue'].cumsum() / prod_perf['Revenue'].sum()
        
        # Cutoff for top 20 products
        top_20_prods = prod_perf.head(20)
        
        pareto_fig = px.bar(
            top_20_prods, x='Item Description', y='Revenue',
            title="Top 20 Products Contributing to Revenue (All Time)",
            text_auto='.2s'
        )
        # Add cumulative line
        pareto_fig.add_scatter(x=top_20_prods['Item Description'], y=top_20_prods['Cumulative %'], yaxis='y2', name='Cumulative %', line=dict(color='red'))
        pareto_fig.update_layout(yaxis2=dict(overlaying='y', side='right', range=[0, 100]))
        pareto_fig.update_xaxes(showticklabels=False) # Hide messy labels
        st.plotly_chart(pareto_fig, use_container_width=True)


    # ==========================================
    # 5. EXCEL EXPORT
    # ==========================================
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    
    # Formats
    header_fmt = workbook.add_format({'bold': True, 'fg_color': '#203764', 'font_color': '#FFFFFF', 'border': 1})
    money_fmt = workbook.add_format({'num_format': '₹#,##0'})
    
    # 1. Summary Sheet
    monthly_metrics.to_excel(writer, sheet_name='Executive Summary', index=False)
    
    # 2. Regional Data
    reg_pivot = master_df.pivot_table(index='Region', columns='Month_Year', values='Revenue', aggfunc='sum').fillna(0)
    reg_pivot.to_excel(writer, sheet_name='Regional Trends')
    
    # 3. Product Data
    prod_pivot = master_df.pivot_table(index='Item Description', columns='Month_Year', values='Revenue', aggfunc='sum').fillna(0)
    prod_pivot['Total'] = prod_pivot.sum(axis=1)
    prod_pivot.sort_values('Total', ascending=False).head(100).to_excel(writer, sheet_name='Top 100 Products')
    
    writer.close()
    output.seek(0)
    
    st.download_button(
        "📥 Download Management Report", output, 
        f"Amazon_Executive_Report_{latest_month}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
