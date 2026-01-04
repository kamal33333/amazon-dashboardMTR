import streamlit as st
import pandas as pd
import zipfile
import io
import plotly.express as px

# ==========================================
# 1. CONFIG & PASSWORD PROTECTION
# ==========================================
st.set_page_config(page_title="Amazon MTR Master", page_icon="📦", layout="wide")

PASSWORD = "kamal_mtramazon" # <--- Change password here

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def check_password():
    if st.session_state.authenticated:
        return True
    
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.title("🔒 Amazon MTR Login")
        password = st.text_input("Enter Password", type="password")
        if st.button("Login"):
            if password == PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Incorrect Password")
    return False

if not check_password():
    st.stop()

# ==========================================
# 2. LOGIC FUNCTIONS
# ==========================================

def read_csv_from_zip_buffer(uploaded_file):
    """Extracts the first CSV found in a ZIP file uploaded to Streamlit."""
    try:
        with zipfile.ZipFile(uploaded_file) as z:
            csv_files = [f for f in z.namelist() if f.endswith('.csv')]
            if not csv_files:
                return None
            with z.open(csv_files[0]) as f:
                return pd.read_csv(f)
    except Exception as e:
        st.error(f"Error extracting ZIP: {e}")
        return None

def extract_brand(description):
    desc = str(description).upper().strip()
    if "KRESSA" in desc: return "Kressa"
    if "TLC365" in desc or "TLC 365" in desc: return "TLC365"
    if "K ONE" in desc or "K-ONE" in desc: return "K One"
    return "Other"

def apply_regional_logic(row):
    state_raw = str(row['Ship To State']).upper().strip()
    city = str(row['Ship To City']).upper().strip()

    alias_map = {
        'JAMMU & KASHMIR': 'JAMMU AND KASHMIR',
        'ORISSA': 'ODISHA',
        'ANDAMAN & NICOBAR': 'ANDAMAN AND NICOBAR ISLANDS',
        'DADRA & NAGAR HAVELI': 'DADRA AND NAGAR HAVELI AND DAMAN AND DIU',
        'PONDICHERRY': 'PUDUCHERRY',
        'TELENGANA': 'TELANGANA'
    }
    state = alias_map.get(state_raw, state_raw)

    up_zone_b_cities = ['AYODHYA', 'AZAMGARH', 'BASTI', 'DEVIPATAN', 'GORAKHPUR', 'MIRZAPUR', 'PRAYAGRAJ', 'VARANASI', 'ALLAHABAD']

    r1_north = ['DELHI', 'HARYANA', 'HIMACHAL PRADESH', 'JAMMU AND KASHMIR', 'LADAKH', 'PUNJAB', 'RAJASTHAN', 'UTTAR PRADESH', 'UTTARAKHAND', 'CHANDIGARH']
    r2_west = ['GUJARAT', 'MAHARASHTRA', 'GOA', 'MADHYA PRADESH', 'DADRA AND NAGAR HAVELI AND DAMAN AND DIU']
    r3_south = ['ANDHRA PRADESH', 'KARNATAKA', 'KERALA', 'TAMIL NADU', 'TELANGANA', 'PUDUCHERRY', 'LAKSHADWEEP', 'ANDAMAN AND NICOBAR ISLANDS']
    r4_east = ['WEST BENGAL', 'BIHAR', 'JHARKHAND', 'ODISHA', 'CHHATTISGARH', 'ASSAM', 'ARUNACHAL PRADESH', 'MANIPUR', 'MEGHALAYA', 'MIZORAM', 'NAGALAND', 'SIKKIM', 'TRIPURA']

    region = 'Unmapped'
    final_state = state

    if state == 'UTTAR PRADESH':
        if any(up_city in city for up_city in up_zone_b_cities):
            final_state = 'Uttar Pradesh Zone B'
            region = 'Region 4'
        else:
            region = 'Region 1'
    elif state in r1_north: region = 'Region 1'
    elif state in r2_west: region = 'Region 2'
    elif state in r3_south: region = 'Region 3'
    elif state in r4_east: region = 'Region 4'
    
    return final_state, region

def process_month_data(b2b_file, b2c_file):
    """Helper to process a pair of ZIP files (B2B + B2C) into a clean DataFrame."""
    if not b2b_file or not b2c_file:
        return None
    
    b2b_df = read_csv_from_zip_buffer(b2b_file)
    b2c_df = read_csv_from_zip_buffer(b2c_file)
    
    if b2b_df is None or b2c_df is None:
        return None
        
    b2b_df['Channel'] = 'B2B'
    b2c_df['Channel'] = 'B2C'
    df = pd.concat([b2b_df, b2c_df], ignore_index=True)
    df.columns = df.columns.str.strip()

    # Cleaning
    df['Revenue'] = pd.to_numeric(df['Tax Exclusive Gross'], errors='coerce').fillna(0)
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
    df.loc[df['Transaction Type'].str.contains('Cancel', case=False, na=False), 'Quantity'] = 0
    
    # Logic
    df['Brand'] = df['Item Description'].apply(extract_brand)
    df[['Final State', 'Region']] = df.apply(lambda row: pd.Series(apply_regional_logic(row)), axis=1)
    
    return df

# ==========================================
# 3. APP UI & EXECUTION
# ==========================================
st.title("📦 Amazon MTR Consolidated Master")
st.markdown("Upload raw ZIP files for both **Current Month** and **Previous Month** to generate a comparison report.")

# Sidebar Uploads
st.sidebar.header("📅 Current Month Data")
curr_b2b = st.sidebar.file_uploader("Current B2B (ZIP)", type=['zip'], key="curr_b2b")
curr_b2c = st.sidebar.file_uploader("Current B2C (ZIP)", type=['zip'], key="curr_b2c")

st.sidebar.header("⏮️ Previous Month Data")
prev_b2b = st.sidebar.file_uploader("Previous B2B (ZIP)", type=['zip'], key="prev_b2b")
prev_b2c = st.sidebar.file_uploader("Previous B2C (ZIP)", type=['zip'], key="prev_b2c")

if st.sidebar.button("🔴 Logout"):
    st.session_state.authenticated = False
    st.rerun()

if curr_b2b and curr_b2c:
    with st.spinner('Processing Data...'):
        
        # 1. Process Current Month
        curr_df = process_month_data(curr_b2b, curr_b2c)
        
        # 2. Process Previous Month (If uploaded)
        prev_df = process_month_data(prev_b2b, prev_b2c)

        if curr_df is not None:
            # Reorder Columns for Final Output
            cols = list(curr_df.columns)
            if 'Ship To State' in cols and 'Region' in cols:
                cols.remove('Region')
                state_idx = cols.index('Ship To State')
                cols.insert(state_idx + 1, 'Region')
                curr_df = curr_df[cols]

            # 3. Comparison Logic
            curr_prods = curr_df.groupby('Item Description').agg({
                'Revenue': 'sum', 'Quantity': 'sum'
            }).rename(columns={'Revenue': 'Curr Revenue', 'Quantity': 'Curr Units'})

            comparison_available = False
            top_products_data = curr_prods # Default

            if prev_df is not None:
                prev_prods = prev_df.groupby('Item Description').agg({
                    'Revenue': 'sum', 'Quantity': 'sum'
                }).rename(columns={'Revenue': 'Prev Revenue', 'Quantity': 'Prev Units'})
                
                comparison = curr_prods.join(prev_prods, how='left').fillna(0)
                
                # Trend Logic
                def get_trend(curr, prev):
                    if prev == 0: return "★ New"
                    if curr > prev: return "▲ Up"
                    if curr < prev: return "▼ Down"
                    return "▬ Neutral"

                comparison['Revenue Trend'] = comparison.apply(lambda x: get_trend(x['Curr Revenue'], x['Prev Revenue']), axis=1)
                comparison['Units Trend'] = comparison.apply(lambda x: get_trend(x['Curr Units'], x['Prev Units']), axis=1)
                
                top_products_data = comparison[['Curr Revenue', 'Prev Revenue', 'Revenue Trend', 'Curr Units', 'Prev Units', 'Units Trend']]
                comparison_available = True
                st.success("✅ Comparison Generated Successfully!")
            else:
                st.info("ℹ️ Previous month files not uploaded. Generating Single Month Report.")

            # ==========================================
            # 4. DASHBOARD VISUALS
            # ==========================================
            total_rev = curr_df['Revenue'].sum()
            total_qty = curr_df['Quantity'].sum()
            
            # Growth Metrics
            growth_delta = None
            if prev_df is not None:
                prev_rev = prev_df['Revenue'].sum()
                growth_delta = f"{(total_rev - prev_rev):,.0f} vs Last Month"

            st.markdown("### 📊 Executive Summary")
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Revenue", f"₹{total_rev:,.0f}", delta=growth_delta)
            c2.metric("Total Units", f"{total_qty:,.0f}")
            c3.metric("B2B vs B2C Split", f"{len(curr_df[curr_df['Channel']=='B2B'])} / {len(curr_df[curr_df['Channel']=='B2C'])} Orders")
            
            st.divider()
            
            # Graphs
            g1, g2 = st.columns(2)
            
            with g1:
                st.subheader("🌍 Revenue by Region")
                reg_data = curr_df.groupby('Region')['Revenue'].sum().reset_index()
                fig_reg = px.bar(reg_data, x='Region', y='Revenue', color='Region', text_auto='.2s')
                st.plotly_chart(fig_reg, use_container_width=True)
                
            with g2:
                st.subheader("🥧 Brand Share")
                brand_data = curr_df.groupby('Brand')['Revenue'].sum().reset_index()
                fig_brand = px.pie(brand_data, values='Revenue', names='Brand', hole=0.4)
                st.plotly_chart(fig_brand, use_container_width=True)

            st.subheader("🏆 Top 10 Products")
            top_10 = curr_df.groupby('Item Description')['Revenue'].sum().nlargest(10).reset_index()
            fig_top = px.bar(top_10, x='Revenue', y='Item Description', orientation='h', text_auto='.2s')
            fig_top.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_top, use_container_width=True)

            # ==========================================
            # 5. EXCEL GENERATION
            # ==========================================
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            workbook = writer.book

            # Formats
            money_fmt = workbook.add_format({'num_format': '₹#,##0'})
            green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True})

            # Sheet 1: Combined Data (Current)
            curr_df.to_excel(writer, sheet_name='Combined Data', index=False)
            
            # Sheet 2: Regional Dashboard
            curr_df.groupby('Region').agg({'Revenue': 'sum', 'Quantity': 'sum', 'Sku': 'nunique'}).to_excel(writer, sheet_name='Regional Dashboard')
            
            # Sheet 3: Regional Brand Share
            curr_df.groupby(['Region', 'Brand']).agg({'Revenue': 'sum', 'Quantity': 'sum'}).unstack(fill_value=0).to_excel(writer, sheet_name='Regional Brand Share')

            # Sheet 4: Top Products
            top_products_data.sort_values('Curr Revenue', ascending=False).head(20).to_excel(writer, sheet_name='Top Products')
            if comparison_available:
                writer.sheets['Top Products'].set_column('B:C', 15, money_fmt)

            # Sheet 5: All Products
            all_prods = curr_df.groupby(['Brand', 'Sku', 'Item Description']).agg({'Revenue': 'sum', 'Quantity': 'sum'}).reset_index()
            all_prods = all_prods.sort_values(['Brand', 'Revenue'], ascending=[True, False])
            all_prods.to_excel(writer, sheet_name='All Products', index=False)
            
            # Highlight Logic (Over 50k)
            writer.sheets['All Products'].conditional_format(1, 3, len(all_prods), 3, {'type': 'cell', 'criteria': '>', 'value': 50000, 'format': green_fmt})

            writer.close()
            output.seek(0)

            st.download_button(
                label="📥 Download MTR Master Excel",
                data=output,
                file_name="Amazon_MTR_Master_Processed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

else:
    st.info("👈 Please upload at least the **Current Month** ZIP files to begin.")
