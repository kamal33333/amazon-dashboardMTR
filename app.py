import streamlit as st
import pandas as pd
import zipfile
import io
import re
import numpy as np
import calendar
from datetime import datetime
from scipy.stats import linregress
import plotly.express as px

# ==========================================
# 1. CONFIG & AUTH
# ==========================================
st.set_page_config(page_title="Amazon AI Analyst", page_icon="🤖", layout="wide")

PASSWORD = "kressa_admin" 

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def check_password():
    if st.session_state.authenticated:
        return True
    c1, c2, c3 = st.columns([1,2,1])
    with col2:
        st.title("🔒 Amazon AI Analyst Login")
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
# 2. HELPER FUNCTIONS (Logic from your script)
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
# 3. MAIN APP INTERFACE
# ==========================================
st.title("🤖 Amazon AI Report Generator")
st.markdown("""
### Instructions:
1. Select **ALL** your ZIP files from your folder (Jan, Feb, Mar... etc).
2. Drag and drop them into the box below.
3. The AI will sort them by date automatically.
""")

uploaded_files = st.file_uploader("Upload Monthly ZIPs (B2B & B2C)", type=['zip'], accept_multiple_files=True)

if uploaded_files and st.button("🚀 Run AI Analysis"):
    
    # --- 1. ORGANIZE FILES ---
    files_map = {}
    with st.spinner("Sorting and organizing files..."):
        for uploaded_file in uploaded_files:
            date_obj, channel = parse_file_info(uploaded_file.name)
            if date_obj and channel:
                if date_obj not in files_map: 
                    files_map[date_obj] = {'date': date_obj, 'B2B': None, 'B2C': None}
                files_map[date_obj][channel] = uploaded_file

        sorted_files = sorted(files_map.values(), key=lambda x: x['date'])
        
        if not sorted_files:
            st.error("No valid B2B/B2C dated files found.")
            st.stop()
        
        st.success(f"📅 Identified {len(sorted_files)} Months of data from {sorted_files[0]['date'].strftime('%b-%Y')} to {sorted_files[-1]['date'].strftime('%b-%Y')}")

    # --- 2. LOAD DATA ---
    all_dfs = []
    projection_factor = 1.0
    now = datetime.now()
    
    progress_bar = st.progress(0)
    
    with st.spinner("Parsing CSVs and stitching data..."):
        for i, m_data in enumerate(sorted_files):
            month_name = m_data['date'].strftime('%b-%Y')
            b2b = read_csv_from_uploaded_zip(m_data['B2B']) if m_data['B2B'] else None
            b2c = read_csv_from_uploaded_zip(m_data['B2C']) if m_data['B2C'] else None
            
            if b2b is not None and b2c is not None:
                b2b['Channel'] = 'B2B'; b2c['Channel'] = 'B2C'
                curr_df = pd.concat([b2b, b2c], ignore_index=True)
                curr_df.columns = curr_df.columns.str.strip()
                
                # Cleanup
                curr_df['Revenue'] = pd.to_numeric(curr_df['Tax Exclusive Gross'], errors='coerce').fillna(0)
                curr_df['Quantity'] = pd.to_numeric(curr_df['Quantity'], errors='coerce').fillna(0)
                curr_df.loc[curr_df['Transaction Type'].str.contains('Cancel', case=False, na=False), 'Quantity'] = 0
                
                # Projection Logic
                if i == len(sorted_files) - 1:
                    is_current_month = (m_data['date'].year == now.year) and (m_data['date'].month == now.month)
                    if is_current_month:
                        date_col = 'Invoice Date' if 'Invoice Date' in curr_df.columns else 'Order Date'
                        if date_col in curr_df.columns:
                            curr_df['Parsed_Date'] = pd.to_datetime(curr_df[date_col], dayfirst=True, errors='coerce')
                            max_date = curr_df['Parsed_Date'].max()
                            if pd.notnull(max_date):
                                days_in_data = max_date.day
                                _, days_in_month = calendar.monthrange(max_date.year, max_date.month)
                                if days_in_data < days_in_month and days_in_data > 0:
                                    projection_factor = days_in_month / days_in_data
                                    st.info(f"⚠️ Partial Month Detected ({month_name}): Extrapolating x{projection_factor:.2f}")

                curr_df['Month_Year'] = month_name
                curr_df['Brand'] = curr_df['Item Description'].apply(extract_brand)
                all_dfs.append(curr_df)
            
            progress_bar.progress((i + 1) / len(sorted_files))

    if not all_dfs:
        st.error("No valid data could be processed.")
        st.stop()
        
    master_df = pd.concat(all_dfs, ignore_index=True)
    master_df[['Final State', 'Region']] = master_df.apply(lambda row: pd.Series(apply_regional_logic(row)), axis=1)
    master_df['Ship To City'] = master_df['Ship To City'].astype(str).str.upper().str.strip()

    # --- 3. CALCULATIONS ---
    def create_stat_sheet(groupby_cols, value_col='Revenue'):
        grouped = master_df.groupby(groupby_cols + ['Month_Year'])[value_col].sum().reset_index()
        pivoted = grouped.pivot_table(index=groupby_cols, columns='Month_Year', values=value_col, aggfunc='sum').fillna(0)
        # Ensure proper time sorting
        sorted_cols = [m['date'].strftime('%b-%Y') for m in sorted_files if m['date'].strftime('%b-%Y') in pivoted.columns]
        pivoted = pivoted[sorted_cols]
        stats_df, _ = calculate_advanced_stats(pivoted, latest_month_factor=projection_factor)
        return stats_df.sort_values(sorted_cols[-1], ascending=False)

    region_stats = create_stat_sheet(['Region'])
    state_stats = create_stat_sheet(['Final State'])
    city_stats = create_stat_sheet(['Final State', 'Ship To City'])
    product_stats = create_stat_sheet(['Brand', 'Sku', 'Item Description'])
    
    # AI Insights
    latest_month = sorted_files[-1]['date'].strftime('%b-%Y')
    prev_month = sorted_files[-2]['date'].strftime('%b-%Y') if len(sorted_files) > 1 else latest_month
    
    insights_list = []
    # 1. Lost Territory
    lost = city_stats[(city_stats[latest_month] == 0) & (city_stats['Average'] > 1000)]
    for idx, row in lost.iterrows():
        insights_list.append({'Type': '🚨 Lost Territory', 'Entity': str(idx[1]), 'Note': f"Avg ₹{int(row['Average'])} -> 0"})
    # 2. Breakouts
    breakouts = product_stats[product_stats['Z_Score'] > 2.0]
    for idx, row in breakouts.head(5).iterrows():
        insights_list.append({'Type': '🔥 Breakout', 'Entity': str(idx[2])[:40], 'Note': f"Z-Score: {round(row['Z_Score'], 2)}"})
    
    ai_df = pd.DataFrame(insights_list)

    # --- 4. VISUALS ---
    st.divider()
    st.header("📊 Executive Overview")
    
    # Time Series Chart
    time_data = master_df.groupby('Month_Year')['Revenue'].sum().reindex(
        [m['date'].strftime('%b-%Y') for m in sorted_files]
    ).reset_index()
    
    fig_trend = px.line(time_data, x='Month_Year', y='Revenue', markers=True, title="Revenue Trend (YTD)")
    st.plotly_chart(fig_trend, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Regional Performance")
        st.dataframe(region_stats.style.format("₹{0:,.0f}"), use_container_width=True)
    with c2:
        if not ai_df.empty:
            st.subheader("🤖 AI Insights")
            st.dataframe(ai_df, use_container_width=True)
        else:
            st.info("No anomalies detected by AI this month.")

    # --- 5. EXCEL EXPORT ---
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    
    # Styles
    header_fmt = workbook.add_format({'bold': True, 'fg_color': '#203764', 'font_color': '#FFFFFF', 'border': 1})
    currency_fmt = workbook.add_format({'num_format': '[$₹-4009]#,##0', 'border': 1})
    pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1})
    green_text = workbook.add_format({'font_color': '#006100', 'bg_color': '#C6EFCE', 'num_format': '0.0%', 'border': 1})
    red_text = workbook.add_format({'font_color': '#9C0006', 'bg_color': '#FFC7CE', 'num_format': '0.0%', 'border': 1})

    def write_sheet_beautified(df, sheet_name):
        df_export = df.copy()
        if isinstance(df_export.columns, pd.MultiIndex):
            df_export.columns = [' - '.join(map(str, col)).strip() for col in df_export.columns.values]
        df_export = df_export.reset_index()
        
        df_export.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
        ws = writer.sheets[sheet_name]
        
        # Headers
        for col_num, value in enumerate(df_export.columns.values):
            ws.write(0, col_num, value, header_fmt)
            ws.set_column(col_num, col_num, 20)
            
        # Formats
        for i, col in enumerate(df_export.columns):
            if any(x in col for x in ['Revenue', 'Average', 'Projected', 'Std_Dev']):
                ws.set_column(i, i, 18, currency_fmt)
            elif 'Growth' in col or 'Share' in col:
                ws.set_column(i, i, 12, pct_fmt)
                ws.conditional_format(1, i, len(df_export), i, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_text})
                ws.conditional_format(1, i, len(df_export), i, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_text})
    
    write_sheet_beautified(region_stats, 'Regional Dashboard')
    write_sheet_beautified(state_stats, 'State Insights')
    write_sheet_beautified(city_stats, 'City Insights')
    write_sheet_beautified(product_stats.head(50), 'Top 50 Products')
    write_sheet_beautified(product_stats, 'All Products')
    if not ai_df.empty:
        ai_df.to_excel(writer, sheet_name='AI Insights', index=False)

    writer.close()
    output.seek(0)
    
    st.download_button(
        label="📥 Download AI Analysis Report (.xlsx)",
        data=output,
        file_name=f"Amazon_AI_Report_{latest_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
