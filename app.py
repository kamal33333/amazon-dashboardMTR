import streamlit as st
import pandas as pd
import zipfile
import io
import re
import numpy as np
import calendar
from datetime import datetime
from scipy.stats import linregress  # <--- Essential for AI Logic
import plotly.express as px

# ==========================================
# 1. CONFIG & AUTH
# ==========================================
st.set_page_config(page_title="Amazon HQ + AI Engine", page_icon="🚀", layout="wide")

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
# 2. DATA ENGINE (AI + LOGIC)
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

# --- AI STATS LOGIC (Restored for Excel) ---
def get_trend_slope(series):
    y = series.values; mask = ~np.isnan(y); y = y[mask]
    if len(y) < 2: return 0
    slope, _, _, _, _ = linregress(np.arange(len(y)), y)
    return slope

def calculate_advanced_stats(df_pivot, latest_month_factor=1.0):
    month_cols = df_pivot.columns.tolist()
    curr_col = month_cols[-1]
    analysis_df = df_pivot[month_cols].copy()
    if latest_month_factor > 1.05:
        analysis_df[curr_col] = analysis_df[curr_col] * latest_month_factor
        
    df_pivot['Average'] = analysis_df.mean(axis=1)
    df_pivot['Std_Dev'] = analysis_df.std(axis=1).fillna(0)
    df_pivot['Trend_Score'] = analysis_df.apply(get_trend_slope, axis=1)
    
    prev_col = month_cols[-2] if len(month_cols) > 1 else month_cols[-1]
    df_pivot['Growth %'] = ((analysis_df[curr_col] - df_pivot[prev_col]) / df_pivot[prev_col].replace(0, 1))
    df_pivot['Z_Score'] = ((analysis_df[curr_col] - df_pivot['Average']) / df_pivot['Std_Dev']).replace([np.inf, -np.inf], 0).fillna(0)
    if latest_month_factor > 1.0:
        df_pivot[f'Projected {curr_col}'] = df_pivot[curr_col] * latest_month_factor
    return df_pivot, month_cols

# ==========================================
# 3. UI & EXECUTION
# ==========================================
st.title("📈 Amazon Management HQ")
st.markdown("Upload monthly ZIP files. The Dashboard shows **Trends**, the Download contains **AI Analysis**.")

uploaded_files = st.file_uploader("Drop All Monthly ZIPs Here", type=['zip'], accept_multiple_files=True)

if uploaded_files and st.button("🚀 Run Management Report"):
    
    # --- 1. FILE ORG ---
    files_map = {}
    with st.spinner("Processing Timeline..."):
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

    # --- 2. LOAD DATA ---
    all_dfs = []
    projection_factor = 1.0
    now = datetime.now()
    
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
            
            # Projection Logic for AI Report
            if i == len(sorted_files) - 1:
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
                                projection_factor = days_in_month / days_in_data

            curr_df['Month_Year'] = month_name
            curr_df['Sort_Date'] = m_data['date']
            curr_df['Brand'] = curr_df['Item Description'].apply(extract_brand)
            all_dfs.append(curr_df)
        progress.progress((i + 1) / len(sorted_files))
    
    if not all_dfs:
        st.stop()

    master_df = pd.concat(all_dfs, ignore_index=True)
    master_df[['Final State', 'Region']] = master_df.apply(lambda row: pd.Series(apply_regional_logic(row)), axis=1)
    master_df['Ship To City'] = master_df['Ship To City'].astype(str).str.upper().str.strip()

    # ==========================================
    # 4. DASHBOARD VISUALS (MANAGEMENT STYLE)
    # ==========================================
    st.divider()
    st.header(f"📊 Executive Summary ({len(sorted_files)} Months)")
    
    # --- A. Stacked Bar Chart ---
    trend_data = master_df.groupby(['Month_Year', 'Channel', 'Sort_Date'])['Revenue'].sum().reset_index()
    trend_data = trend_data.sort_values('Sort_Date')
    fig_trend = px.bar(
        trend_data, x='Month_Year', y='Revenue', color='Channel', 
        title="Monthly Revenue Trend (B2B vs B2C)", text_auto='.2s',
        color_discrete_map={'B2B': '#1f77b4', 'B2C': '#ff7f0e'}
    )
    st.plotly_chart(fig_trend, use_container_width=True)

    # --- B. Hero of the Month ---
    st.subheader("🏆 Hero of the Month")
    hero_list = []
    # Metrics Calculation
    monthly_metrics = master_df.groupby(['Sort_Date', 'Month_Year']).agg(
        Total_Revenue=('Revenue', 'sum'),
        Total_Units=('Quantity', 'sum')
    ).reset_index()

    for m in monthly_metrics['Month_Year']:
        m_df = master_df[master_df['Month_Year'] == m]
        top_prod = m_df.groupby('Item Description')['Revenue'].sum().idxmax()
        top_prod_rev = m_df.groupby('Item Description')['Revenue'].sum().max()
        top_state = m_df.groupby('Final State')['Revenue'].sum().idxmax()
        b2b_share = m_df[m_df['Channel'] == 'B2B']['Revenue'].sum() / m_df['Revenue'].sum()
        
        hero_list.append({
            'Month': m,
            '👑 Top Product': top_prod[:40] + '...',
            'Product Rev': f"₹{top_prod_rev:,.0f}",
            '🌍 Top State': top_state,
            '🏢 B2B Share': f"{b2b_share:.0%}"
        })
    st.dataframe(pd.DataFrame(hero_list), use_container_width=True)

    # --- C. Pareto ---
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📦 Top 20 Products (Pareto)")
        prod_perf = master_df.groupby('Item Description')['Revenue'].sum().sort_values(ascending=False).reset_index().head(20)
        fig_pareto = px.bar(prod_perf, x='Item Description', y='Revenue', title="Top 20 Revenue Drivers")
        fig_pareto.update_xaxes(showticklabels=False)
        st.plotly_chart(fig_pareto, use_container_width=True)
    with c2:
        st.subheader("🌍 Regional Heatmap")
        latest_month = sorted_files[-1]['date'].strftime('%b-%Y')
        latest_df = master_df[master_df['Month_Year'] == latest_month]
        fig_sun = px.sunburst(latest_df, path=['Region', 'Final State'], values='Revenue', title=f"Mix ({latest_month})")
        st.plotly_chart(fig_sun, use_container_width=True)

    # ==========================================
    # 5. EXCEL EXPORT (AI ENGINE RESTORED)
    # ==========================================
    
    # --- Prepare AI Dataframes ---
    def create_stat_sheet(groupby_cols):
        grouped = master_df.groupby(groupby_cols + ['Month_Year'])['Revenue'].sum().reset_index()
        pivoted = grouped.pivot_table(index=groupby_cols, columns='Month_Year', values='Revenue', aggfunc='sum').fillna(0)
        sorted_cols = [m['date'].strftime('%b-%Y') for m in sorted_files if m['date'].strftime('%b-%Y') in pivoted.columns]
        pivoted = pivoted[sorted_cols]
        stats_df, _ = calculate_advanced_stats(pivoted, latest_month_factor=projection_factor)
        return stats_df.sort_values(sorted_cols[-1], ascending=False)

    region_stats = create_stat_sheet(['Region'])
    state_stats = create_stat_sheet(['Final State'])
    city_stats = create_stat_sheet(['Final State', 'Ship To City'])
    product_stats = create_stat_sheet(['Brand', 'Sku', 'Item Description'])

    # AI Insights Generation
    insights_list = []
    latest_m = sorted_files[-1]['date'].strftime('%b-%Y')
    
    # Lost Territory Logic
    lost = city_stats[(city_stats[latest_m] == 0) & (city_stats['Average'] > 1000)]
    for idx, row in lost.iterrows():
        insights_list.append({'Type': '🚨 Lost Territory', 'Entity': str(idx[1]), 'Note': f"Avg ₹{int(row['Average'])} -> 0"})
    
    # Breakout Logic
    breakouts = product_stats[product_stats['Z_Score'] > 2.0]
    for idx, row in breakouts.head(5).iterrows():
        insights_list.append({'Type': '🔥 Breakout', 'Entity': str(idx[2])[:40], 'Note': f"Z-Score: {round(row['Z_Score'], 2)}"})
    
    ai_df = pd.DataFrame(insights_list)

    # --- Write Excel ---
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    
    # Formats
    header_fmt = workbook.add_format({'bold': True, 'fg_color': '#203764', 'font_color': '#FFFFFF', 'border': 1})
    currency_fmt = workbook.add_format({'num_format': '[$₹-4009]#,##0', 'border': 1})
    pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1})
    green = workbook.add_format({'font_color': '#006100', 'bg_color': '#C6EFCE', 'num_format': '0.0%'})
    red = workbook.add_format({'font_color': '#9C0006', 'bg_color': '#FFC7CE', 'num_format': '0.0%'})

    def write_sheet(df, name):
        df = df.copy()
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [' '.join(map(str, col)).strip() for col in df.columns.values]
        df = df.reset_index()
        df.to_excel(writer, sheet_name=name, startrow=1, header=False, index=False)
        ws = writer.sheets[name]
        for i, col in enumerate(df.columns):
            ws.write(0, i, col, header_fmt)
            ws.set_column(i, i, 20)
            if 'Revenue' in col or 'Avg' in col or 'Projected' in col: ws.set_column(i, i, 18, currency_fmt)
            if 'Growth' in col: 
                ws.set_column(i, i, 12, pct_fmt)
                ws.conditional_format(1, i, len(df), i, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green})
                ws.conditional_format(1, i, len(df), i, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red})

    # Write All Sheets (Restoring previous report structure)
    write_sheet(region_stats, 'Regional Dashboard')
    write_sheet(state_stats, 'State Insights')
    write_sheet(city_stats, 'City Insights')
    write_sheet(product_stats.head(50), 'Top 50 Products')
    write_sheet(product_stats, 'All Products')
    if not ai_df.empty: ai_df.to_excel(writer, sheet_name='AI Insights', index=False)

    writer.close()
    output.seek(0)
    
    st.download_button(
        "📥 Download AI Analysis Report (.xlsx)", output, 
        f"Amazon_Report_{latest_m}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
