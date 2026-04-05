import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime
from io import BytesIO
import os
import errno

# --- 1. PAGE SETUP ---
st.set_page_config(page_title="Asset Lifecycle Analytics", layout="wide")

# =====================================================================
# ⚡ FILE-SPECIFIC SKIPROWS CONFIG
# Root cause fix: 100400_A.xlsx has ONE blank row at top (header at row 1),
# while 100500 and PV files have TWO rows before header (header at row 2).
# The old code used skiprows=2 for ALL files — causing 100400 to read data
# rows AS headers (numeric column names like 1008, 40002000, etc.),
# which broke column detection, Gross value parsing, and all downstream logic.
# =====================================================================
FILE_SKIPROWS = {
    "100400_A.xlsx": 1,       # Header is on row 1 (only 1 blank row above)
    "100500_A.xlsx": 2,       # Header is on row 2
    "PV Pending IT_Admin Rajan Kapoor.xlsx": 2,
}
DEFAULT_SKIPROWS = 2

def get_skiprows(file_path):
    """Return the correct skiprows for a known file, or detect dynamically."""
    fname = os.path.basename(str(file_path))
    if fname in FILE_SKIPROWS:
        return FILE_SKIPROWS[fname]
    # Dynamic detection for uploaded files: scan first 10 rows for known header keywords
    try:
        xl = pd.ExcelFile(file_path)
        for i in range(6):
            row = xl.parse(xl.sheet_names[0], skiprows=i, nrows=1)
            cols = " ".join(str(c) for c in row.columns).lower()
            if "cost center" in cols and ("asset" in cols or "gross" in cols):
                return i
    except Exception:
        pass
    return DEFAULT_SKIPROWS

# =====================================================================
# ⚡ MEMORY OPTIMIZATION
# =====================================================================
def optimize_df_memory(df):
    if df.empty:
        return df
    df.columns = [str(c).strip() if isinstance(c, str) else str(c) for c in df.columns]
    df = df.rename(columns=lambda x: f"Col_{x}" if x.isdigit() else x)
    for col in df.columns:
        if df[col].dtype == 'object':
            num_unique = df[col].nunique()
            if num_unique > 0 and num_unique < len(df) * 0.4:
                df[col] = df[col].astype('category')
    return df

def smart_read_file(file_path):
    """Read a file using the correct skiprows for its structure."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), file_path)

    skip = get_skiprows(file_path)

    if str(file_path).endswith('.csv'):
        try:
            return pd.read_csv(file_path, skiprows=skip, low_memory=False)
        except Exception:
            return pd.read_csv(file_path, skiprows=DEFAULT_SKIPROWS, low_memory=False)
    else:
        try:
            xl = pd.ExcelFile(file_path)
            return xl.parse(xl.sheet_names[0], skiprows=skip)
        except Exception:
            return pd.DataFrame()

# =====================================================================
# 🔐 LOGIN SYSTEM
# =====================================================================
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None

if not st.session_state.logged_in:
    st.title("🔐 Asset Lifecycle Analytics - Login")
    tab1, tab2 = st.tabs(["Admin Login", "Continue as Guest"])
    with tab1:
        st.markdown("### Admin Access")
        user = st.text_input("Username", key="admin_user")
        pwd = st.text_input("Password", type="password", key="admin_pwd")
        if st.button("Login"):
            if user == "admin" and pwd == "admin123":
                st.session_state.logged_in = True
                st.session_state.role = "admin"
                st.rerun()
            else:
                st.error("❌ Invalid credentials")
    with tab2:
        st.markdown("### Guest Access")
        st.write("View the dashboard in read-only mode. Uploads and editing are disabled.")
        if st.button("Continue as Guest"):
            st.session_state.logged_in = True
            st.session_state.role = "guest"
            st.rerun()
    st.stop()

# =====================================================================
# 🧠 DATA PERSISTENCE
# =====================================================================
UPLOAD_DIR = "./uploaded_files/"
os.makedirs(UPLOAD_DIR, exist_ok=True)

def _standardize_columns(df):
    """Strip whitespace from column names and avoid digit-only names."""
    df.columns = [str(c).strip() if isinstance(c, str) else str(c) for c in df.columns]
    df = df.rename(columns=lambda x: f"Col_{x}" if x.isdigit() else x)
    return df

def _fix_gross_value(df):
    """Find and standardize the gross value column."""
    # Already correct name
    if 'Gross value' in df.columns:
        df['Gross value'] = pd.to_numeric(df['Gross value'], errors='coerce').fillna(0)
        return df
    # Search for it
    gross_col = next((c for c in df.columns if 'gross' in c.lower() and
                      ('amt' in c.lower() or 'value' in c.lower())), None)
    if gross_col:
        df.rename(columns={gross_col: 'Gross value'}, inplace=True)
        df['Gross value'] = pd.to_numeric(df['Gross value'], errors='coerce').fillna(0)
    else:
        df['Gross value'] = 0
    return df

def _fix_cap_date_and_age(df):
    if 'Capitalization Date' in df.columns:
        df['Capitalization Date'] = pd.to_datetime(
            df['Capitalization Date'].astype(str),
            errors='coerce', format='mixed', dayfirst=True
        )
        df['Age_Years'] = (pd.to_datetime("today") - df['Capitalization Date']).dt.days / 365.25
        df['Age_Years'] = df['Age_Years'].clip(lower=0)
    else:
        df['Age_Years'] = np.nan
    return df

def _fix_asset_category(df):
    if 'Asset Category Description' in df.columns:
        df['Asset Category Description'] = df['Asset Category Description'].fillna('Uncategorized')
    else:
        df['Asset Category Description'] = 'Uncategorized'
    return df

def clean_uploaded_data(df):
    df = _standardize_columns(df)
    df = _fix_gross_value(df)
    df = _fix_cap_date_and_age(df)
    df = _fix_asset_category(df)
    return optimize_df_memory(df)

def load_uploaded_files():
    uploaded_dfs = {}
    for fname in os.listdir(UPLOAD_DIR):
        fpath = os.path.join(UPLOAD_DIR, fname)
        try:
            df = smart_read_file(fpath)
            if not df.empty:
                uploaded_dfs[fname] = clean_uploaded_data(df)
        except Exception:
            pass
    return uploaded_dfs

if 'uploaded_dfs' not in st.session_state:
    st.session_state.uploaded_dfs = load_uploaded_files()

# --- CUSTOM INDIAN CURRENCY FORMATTER ---
def format_curr(num):
    if pd.isna(num): return "₹ 0.00"
    return "₹ {:,.2f}".format(num)

# =====================================================================
# 📁 FILE PATHS
# =====================================================================
FILE_1 = r"100400_A.xlsx"
FILE_2 = r"100500_A.xlsx"
FILE_3 = "PV Pending IT_Admin Rajan Kapoor.xlsx"

# --- 2. OPTIMIZED DATA LOADER ---
@st.cache_data(show_spinner=False)
def optimized_load_and_clean_data():
    df1 = smart_read_file(FILE_1)
    df2 = smart_read_file(FILE_2)
    df3 = smart_read_file(FILE_3)

    if not df1.empty: df1['Dataset_Source'] = 'Office Equipment'
    if not df2.empty: df2['Dataset_Source'] = 'Furniture & Fittings'
    if not df3.empty: df3['Dataset_Source'] = 'IT & Admin (Pending PV)'

    def clean_data(df):
        if df.empty: return df
        df = _standardize_columns(df)
        df = _fix_gross_value(df)
        df = _fix_cap_date_and_age(df)
        df = _fix_asset_category(df)
        return optimize_df_memory(df)

    return clean_data(df1), clean_data(df2), clean_data(df3)

# ---------------------------------------------------------------------
# MAIN APP EXECUTION
# ---------------------------------------------------------------------
try:
    with st.spinner("Auto-loading local data files..."):
        df1, df2, df3 = optimized_load_and_clean_data()

    master_df = pd.concat([df1, df2, df3], ignore_index=True)
    if 'Age_Years' not in master_df.columns:
        master_df['Age_Years'] = np.nan
    if 'Gross value' not in master_df.columns:
        master_df['Gross value'] = 0

    if st.session_state.uploaded_dfs:
        uploaded_combined = pd.concat(list(st.session_state.uploaded_dfs.values()), ignore_index=True)
        master_df = pd.concat([master_df, uploaded_combined], ignore_index=True)

    unique_categories = master_df['Asset Category Description'].dropna().unique()
    if 'custom_rates' not in st.session_state:
        default_rates = pd.DataFrame({
            'Asset Category Description': unique_categories,
            'Lifecycle (Years)': 15,
            'Value @ 5-10 Yrs (%)': 40.0,
            'Value @ 10-15 Yrs (%)': 20.0,
            'Value @ 15-20 Yrs (%)': 10.0,
            'Value @ 20+ Yrs (%)': 5.0
        })
        st.session_state.custom_rates = default_rates

    st.sidebar.success("✅ Files Loaded Automatically!")

    st.sidebar.markdown("---")
    st.sidebar.write(f"👤 Logged in as: **{st.session_state.role.title()}**")
    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.role = None
        st.rerun()

    # FILE UPLOAD (Admin only)
    if st.session_state.role == "admin":
        st.sidebar.markdown("---")
        st.sidebar.header("📁 Upload New Dataset")
        uploaded_file = st.sidebar.file_uploader("Accepts .xlsx or .csv", type=["csv", "xlsx"])
        if uploaded_file is not None:
            file_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            try:
                new_df = smart_read_file(file_path)
                if not new_df.empty:
                    cleaned_df = clean_uploaded_data(new_df)
                    st.session_state.uploaded_dfs[uploaded_file.name] = cleaned_df
                    st.sidebar.success(f"✅ {uploaded_file.name} uploaded successfully!")
                    st.rerun()
                else:
                    st.sidebar.error("❌ Failed to parse data from file.")
            except Exception as e:
                st.sidebar.error(f"❌ Error processing file: {e}")

    st.sidebar.markdown("---")
    st.sidebar.header("🎛️ Select Dashboard View")

    base_options = [
        "🏢 Office Equipment Only",
        "🪑 Furniture & Fittings Only",
        "💻 Pending IT Admin Only",
        "🌍 Master Data"
    ]
    uploaded_options = [f"📂 {fname} (Uploaded)" for fname in st.session_state.uploaded_dfs.keys()]
    all_options = base_options + uploaded_options + ["⚙️ Advanced Category Modeler"]
    view_option = st.sidebar.radio("Choose which view to analyze:", all_options)

    # =========================================================================
    # 🔥 DASHBOARD: CUSTOM CATEGORY MODELER
    # =========================================================================
    if view_option == "⚙️ Advanced Category Modeler":
        st.title("⚙️ Advanced Category-Wise Lifecycle Modeler")
        st.markdown("Set custom lifecycle expectations and depreciation (salvage) rates for **each specific appliance category**. ⚠️ **Rules set here will automatically apply to ALL other tabs.**")

        st.markdown("### 1. 🎛️ Edit Category Rules (Live Table)")
        st.info("💡 Tip: Click inside any cell below to change the percentage or lifecycle years. The charts will update instantly.")

        is_guest = st.session_state.role == "guest"
        if is_guest:
            st.warning("🔒 You are viewing this in Guest mode. Editing rates is disabled.")

        edited_rates = st.data_editor(
            st.session_state.custom_rates,
            hide_index=True,
            disabled=is_guest,
            column_config={
                "Asset Category Description": st.column_config.TextColumn("Appliance Category", disabled=True),
                "Lifecycle (Years)": st.column_config.NumberColumn("Lifecycle (Yrs)", min_value=1, max_value=50, step=1),
                "Value @ 5-10 Yrs (%)": st.column_config.NumberColumn("Value @ 5-10 Yrs (%)", min_value=0.0, max_value=100.0, format="%.1f %%"),
                "Value @ 10-15 Yrs (%)": st.column_config.NumberColumn("Value @ 10-15 Yrs (%)", min_value=0.0, max_value=100.0, format="%.1f %%"),
                "Value @ 15-20 Yrs (%)": st.column_config.NumberColumn("Value @ 15-20 Yrs (%)", min_value=0.0, max_value=100.0, format="%.1f %%"),
                "Value @ 20+ Yrs (%)": st.column_config.NumberColumn("Value @ 20+ Yrs (%)", min_value=0.0, max_value=100.0, format="%.1f %%"),
            }
        )
        st.session_state.custom_rates = edited_rates

        custom_df = pd.merge(master_df, edited_rates, on='Asset Category Description', how='left')

        def calculate_custom_impact(row):
            age = row['Age_Years']
            gv = row['Gross value']
            if pd.isna(age): return gv * 0.10
            if age < 5: return gv * 0.80
            elif age < 10: return gv * (row['Value @ 5-10 Yrs (%)'] / 100)
            elif age < 15: return gv * (row['Value @ 10-15 Yrs (%)'] / 100)
            elif age < 20: return gv * (row['Value @ 15-20 Yrs (%)'] / 100)
            else: return gv * (row['Value @ 20+ Yrs (%)'] / 100)

        custom_df['Est. Current Value'] = custom_df.apply(calculate_custom_impact, axis=1)
        custom_df['Financial Impact'] = custom_df['Gross value'] - custom_df['Est. Current Value']
        custom_df['Is_EOL'] = custom_df['Age_Years'] >= custom_df['Lifecycle (Years)']
        eol_df = custom_df[custom_df['Is_EOL']]

        st.markdown("---")
        st.markdown("### 2. 📊 Master Financial Impact (Based on Custom Rules)")

        total_assets = len(custom_df)
        eol_count = len(eol_df)
        total_gross = custom_df['Gross value'].sum()
        total_current_val = custom_df['Est. Current Value'].sum()
        total_impact = custom_df['Financial Impact'].sum()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Assets Exceeding Lifecycle", f"{eol_count:,} / {total_assets:,}", "Requires Review", delta_color="off")
        col2.metric("Total Portfolio Value", format_curr(total_gross))
        col3.metric("Est. Current Value (Custom)", format_curr(total_current_val))
        col4.metric("📉 Total Financial Impact", format_curr(total_impact), delta="Depreciation Applied", delta_color="inverse")

        st.markdown("---")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Financial Impact by Category (Top 10)")
            cat_impact = custom_df.groupby('Asset Category Description')['Financial Impact'].sum().reset_index()
            cat_impact = cat_impact.nlargest(10, 'Financial Impact')
            fig_custom_pie = px.pie(cat_impact, names='Asset Category Description', values='Financial Impact', hole=0.5)
            st.plotly_chart(fig_custom_pie, use_container_width=True)
        with c2:
            st.markdown("#### Impact by Location (Top 10)")
            loc_impact = custom_df.groupby('Cost Center desc.')['Financial Impact'].sum().reset_index()
            loc_impact = loc_impact.nlargest(10, 'Financial Impact')
            fig_custom_bar = px.bar(loc_impact, x='Cost Center desc.', y='Financial Impact', color='Cost Center desc.')
            fig_custom_bar.update_layout(showlegend=False, xaxis_title="Location", yaxis_title="Impact (Exact ₹)")
            st.plotly_chart(fig_custom_bar, use_container_width=True)

        st.markdown("### 📥 Download Custom Evaluation Report")
        dl_cols = ['Asset No.', 'Serial No.', 'Asset Category Description', 'Cost Center desc.', 'Gross value', 'Age_Years', 'Lifecycle (Years)', 'Est. Current Value', 'Financial Impact', 'Is_EOL']
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            custom_df[[c for c in dl_cols if c in custom_df.columns]].to_excel(writer, index=False, sheet_name='Custom_Evaluation')
        st.download_button(
            label="📥 Download Custom Rule Excel Report",
            data=output.getvalue(),
            file_name="Custom_Asset_Evaluation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    # =========================================================================
    # STANDARD DASHBOARD VIEWS
    # =========================================================================
    else:
        if view_option == "🏢 Office Equipment Only":
            active_df = df1
            st.title("🏢 Dashboard: Office Equipment")
            file_tag = "OfficeEq"
        elif view_option == "🪑 Furniture & Fittings Only":
            active_df = df2
            st.title("🪑 Dashboard: Furniture & Fittings")
            file_tag = "Furniture"
        elif view_option == "💻 Pending IT Admin Only":
            active_df = df3
            st.title("💻 Dashboard: Pending IT Admin")
            file_tag = "IT_Admin"
        elif view_option.startswith("📂"):
            fname = view_option.replace("📂 ", "").replace(" (Uploaded)", "")
            active_df = st.session_state.uploaded_dfs[fname]
            st.title(f"📂 Dashboard: {fname}")
            file_tag = fname.replace(".", "_")
        else:
            active_df = master_df
            st.title("🌍 CARO DETAILS & ASSET LIFECYCLE")
            file_tag = "Master"

        st.markdown("---")

        merged_active_df = pd.merge(active_df, st.session_state.custom_rates, on='Asset Category Description', how='left')

        def calculate_global_impact(row):
            age = row['Age_Years']
            gv = row['Gross value']
            if pd.isna(age): return gv * 0.10
            if age < 5: return gv * 0.80
            elif age < 10: return gv * (row['Value @ 5-10 Yrs (%)'] / 100)
            elif age < 15: return gv * (row['Value @ 10-15 Yrs (%)'] / 100)
            elif age < 20: return gv * (row['Value @ 15-20 Yrs (%)'] / 100)
            else: return gv * (row['Value @ 20+ Yrs (%)'] / 100)

        merged_active_df['Est. Salvage Value'] = merged_active_df.apply(calculate_global_impact, axis=1)
        merged_active_df['Financial Impact'] = merged_active_df['Gross value'] - merged_active_df['Est. Salvage Value']

        def determine_risk(row):
            age = row['Age_Years']
            lifecycle = row['Lifecycle (Years)']
            if pd.isna(age) or pd.isna(lifecycle): return "⚪ Unknown"
            if age >= lifecycle: return "🔴 High Risk"
            elif (lifecycle - age) <= 2: return "🟡 Medium Risk"
            else: return "🟢 Low Risk"

        merged_active_df['Risk_Level'] = merged_active_df.apply(determine_risk, axis=1)

        st.markdown("### 🚦 Portfolio Risk Summary")
        r_high = len(merged_active_df[merged_active_df['Risk_Level'] == '🔴 High Risk'])
        r_med = len(merged_active_df[merged_active_df['Risk_Level'] == '🟡 Medium Risk'])
        r_low = len(merged_active_df[merged_active_df['Risk_Level'] == '🟢 Low Risk'])

        rk1, rk2, rk3 = st.columns(3)
        rk1.metric("🔴 High Risk (Age ≥ Lifecycle)", f"{r_high:,} Assets")
        rk2.metric("🟡 Medium Risk (Within 2 Years)", f"{r_med:,} Assets")
        rk3.metric("🟢 Low Risk (Safe)", f"{r_low:,} Assets")

        total_assets_risk = len(merged_active_df)
        if total_assets_risk > 0:
            high_risk_pct = r_high / total_assets_risk
            if high_risk_pct > 0.30:
                st.warning(f"🚨 **CRITICAL ALERT:** High Risk assets constitute **{high_risk_pct:.1%}** of this portfolio. Immediate replacement planning is strongly recommended.")
            elif high_risk_pct > 0.15:
                st.warning(f"⚠️ **NOTICE:** High Risk assets constitute **{high_risk_pct:.1%}** of this portfolio. Start planning phased replacements.")

        st.markdown("---")

        st.markdown("### 🔮 Predictive Insights: Cost of Inaction & Loss Curve")

        future_df = merged_active_df.copy()
        future_df['Age_Years'] = future_df['Age_Years'] + 2
        future_df['Est. Salvage Value'] = future_df.apply(calculate_global_impact, axis=1)
        future_df['Financial Impact'] = future_df['Gross value'] - future_df['Est. Salvage Value']

        current_loss = merged_active_df['Financial Impact'].sum()
        future_loss = future_df['Financial Impact'].sum()
        delayed_loss = future_loss - current_loss

        coi_col1, coi_col2 = st.columns([1, 1.5])
        with coi_col1:
            st.info(
                f"💡 **2-Year Delay Projection:**\n\n"
                f"• If replaced **NOW**, Impact: **{format_curr(current_loss)}**\n\n"
                f"• If delayed **2 YEARS**, Impact: **{format_curr(future_loss)}**\n\n"
                f"🔻 **Additional Loss from delay: {format_curr(delayed_loss)}**"
            )
        with coi_col2:
            curve_df = merged_active_df[merged_active_df['Age_Years'].notna()]
            fig_curve = px.scatter(curve_df, x='Age_Years', y='Financial Impact', color='Asset Category Description',
                                   title="Loss Curve (Impact vs Asset Age)",
                                   labels={'Age_Years': 'Age (Years)', 'Financial Impact': 'Financial Impact (₹)'})
            fig_curve.update_layout(margin=dict(l=20, r=20, t=40, b=20), height=300)
            st.plotly_chart(fig_curve, use_container_width=True)

        st.markdown("---")

        analysis_mode = st.radio("⚙️ Choose Filtering Mode:", ["📑 Age Slabs (Buckets)", "🎯 Custom Slider", "🎯 Exact Year"], horizontal=True)

        if analysis_mode == "📑 Age Slabs (Buckets)":
            st.markdown("### ⏱️ Select Asset Age Range (Buckets)")
            tab_preview, tab0, tab1, tab2, tab3, tab4 = st.tabs(["📊 Data Preview", "🟢 0 to 5 Years", "5 to 10 Years", "10 to 15 Years", "15 to 20 Years", "20+ Years"])

            with tab_preview:
                st.markdown("### 📊 Raw Data Explorer")
                f_col1, f_col2, f_col3 = st.columns(3)
                with f_col1:
                    cat_options = ["All"] + list(active_df['Asset Category Description'].dropna().unique())
                    preview_cat = st.selectbox("Filter Category:", cat_options, key=f"prev_cat_{file_tag}")
                with f_col2:
                    loc_options = ["All"] + list(active_df['Cost Center desc.'].dropna().unique())
                    preview_loc = st.selectbox("Filter Location:", loc_options, key=f"prev_loc_{file_tag}")
                with f_col3:
                    preview_search = st.text_input("Search Asset No. / Serial No.:", key=f"prev_search_{file_tag}")

                preview_df = merged_active_df.copy()
                if preview_cat != "All":
                    preview_df = preview_df[preview_df['Asset Category Description'] == preview_cat]
                if preview_loc != "All":
                    preview_df = preview_df[preview_df['Cost Center desc.'] == preview_loc]
                if preview_search.strip():
                    s = preview_search.strip()
                    preview_df = preview_df[
                        preview_df['Asset No.'].astype(str).str.contains(s, case=False, na=False) |
                        preview_df['Serial No.'].astype(str).str.contains(s, case=False, na=False)
                    ]

                st.markdown("---")
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Assets", f"{len(preview_df):,}")
                m2.metric("Total Gross Value", format_curr(preview_df['Gross value'].sum()))
                avg_age = preview_df['Age_Years'].mean()
                m3.metric("Average Asset Age", f"{avg_age:.2f} Years" if pd.notna(avg_age) else "N/A")
                st.markdown("---")

                disp_preview = preview_df.copy()
                if 'Gross value' in disp_preview.columns:
                    disp_preview['Gross value'] = disp_preview['Gross value'].apply(format_curr)
                display_cols_prev = ['Asset No.', 'Serial No.', 'Asset Category Description', 'Cost Center desc.', 'Gross value', 'Age_Years', 'Lifecycle (Years)', 'Risk_Level']
                st.dataframe(disp_preview[[c for c in display_cols_prev if c in disp_preview.columns]], use_container_width=True)

            with tab0:
                st.markdown("### 🟢 Baseline Validation: Assets (0 to 5 Years)")
                valid_df0 = merged_active_df[merged_active_df['Age_Years'].notna()]
                df_0_5 = valid_df0[(valid_df0['Age_Years'] >= 0) & (valid_df0['Age_Years'] < 5)].copy()
                df_0_5['Financial Impact'] = 0.00
                df_0_5['Est. Salvage Value'] = df_0_5['Gross value']
                st.info("💡 **Baseline View:** Assets in this bucket are considered 'New'. Depreciation and Financial Impact logic is bypassed (Evaluated at ₹0).")
                b1, b2, b3 = st.columns(3)
                b1.metric("Total Assets (0-5 Yrs)", f"{len(df_0_5):,}")
                b2.metric("Total Gross Value", format_curr(df_0_5['Gross value'].sum()))
                avg_0_5 = df_0_5['Age_Years'].mean()
                b3.metric("Average Age", f"{avg_0_5:.2f} Years" if pd.notna(avg_0_5) else "N/A")
                st.markdown("---")
                disp_0_5 = df_0_5.copy()
                for col in ['Gross value', 'Est. Salvage Value', 'Financial Impact']:
                    if col in disp_0_5.columns: disp_0_5[col] = disp_0_5[col].apply(format_curr)
                cols_to_show = ['Asset No.', 'Asset Category Description', 'Cost Center desc.', 'Gross value', 'Age_Years', 'Risk_Level']
                st.dataframe(disp_0_5[[c for c in cols_to_show if c in disp_0_5.columns]], use_container_width=True)
                if not df_0_5.empty:
                    out0 = BytesIO()
                    with pd.ExcelWriter(out0, engine='openpyxl') as w0:
                        df_0_5.to_excel(w0, index=False, sheet_name='0_5_Baseline')
                    st.download_button("📥 Download Baseline Report (0-5 Yrs)", out0.getvalue(), f"{file_tag}_0_5_Baseline.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            def render_scenario(age_range, range_label, current_df, tab_context):
                with tab_context:
                    valid_df = current_df[current_df['Age_Years'].notna()].copy()
                    if age_range == "20+":
                        eol_df = valid_df[(valid_df['Age_Years'] >= 20) & (valid_df['Age_Years'] >= valid_df['Lifecycle (Years)'])].copy()
                    else:
                        min_age, max_age = age_range
                        eol_df = valid_df[(valid_df['Age_Years'] >= min_age) & (valid_df['Age_Years'] < max_age) & (valid_df['Age_Years'] >= valid_df['Lifecycle (Years)'])].copy()

                    eol_assets_count = len(eol_df)
                    gross_value_eol = eol_df['Gross value'].sum()
                    residual_recovered = eol_df['Est. Salvage Value'].sum()
                    financial_impact = eol_df['Financial Impact'].sum()

                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("End-of-Life (EOL) Assets", f"{eol_assets_count:,}")
                    col2.metric("Total Original Valuation", format_curr(gross_value_eol))
                    col3.metric("Est. Salvage Value", format_curr(residual_recovered))
                    col4.metric("📉 Net Financial Impact", format_curr(financial_impact), delta="Depreciated", delta_color="inverse")

                    st.markdown("---")

                    if not eol_df.empty:
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"#### 1. Category-Wise Impact ({range_label})")
                            appliance_impact = eol_df.groupby('Asset Category Description')['Financial Impact'].sum().reset_index()
                            appliance_impact = appliance_impact.nlargest(10, 'Financial Impact')
                            fig_app = px.pie(appliance_impact, names='Asset Category Description', values='Financial Impact', hole=0.5)
                            st.plotly_chart(fig_app, use_container_width=True)
                        with c2:
                            st.markdown(f"#### 2. Location-Wise Impact ({range_label})")
                            location_impact = eol_df.groupby('Cost Center desc.')['Financial Impact'].sum().reset_index()
                            location_impact = location_impact.nlargest(10, 'Financial Impact')
                            fig_loc = px.bar(location_impact, x='Cost Center desc.', y='Financial Impact', color='Cost Center desc.')
                            fig_loc.update_layout(showlegend=False, xaxis_title="Location / Cost Center", yaxis_title="Financial Impact (Exact ₹)")
                            st.plotly_chart(fig_loc, use_container_width=True)

                        st.markdown("---")
                        st.markdown("### 🔍 Deep Dive: Asset Report for this Range")

                        filter_col1, filter_col2 = st.columns(2)
                        with filter_col1:
                            locations = ["All"] + list(eol_df['Cost Center desc.'].dropna().unique())
                            selected_loc = st.selectbox("Select Location:", locations, key=f"loc_{range_label}_{file_tag}")
                        with filter_col2:
                            devices = ["All"] + list(eol_df['Asset Category Description'].dropna().unique())
                            selected_dev = st.selectbox("Select Category:", devices, key=f"dev_{range_label}_{file_tag}")

                        filtered_df = eol_df.copy()
                        if selected_loc != "All":
                            filtered_df = filtered_df[filtered_df['Cost Center desc.'] == selected_loc]
                        if selected_dev != "All":
                            filtered_df = filtered_df[filtered_df['Asset Category Description'] == selected_dev]

                        desired_cols = ['Asset No.', 'Serial No.', 'Dataset_Source', 'Asset Category Description', 'Cost Center desc.', 'Description', 'Gross value', 'Age_Years', 'Lifecycle (Years)', 'Risk_Level', 'Est. Salvage Value', 'Financial Impact']
                        display_cols = [col for col in desired_cols if col in filtered_df.columns]
                        disp_drilldown = filtered_df[display_cols].copy()
                        for col in ['Gross value', 'Est. Salvage Value', 'Financial Impact']:
                            if col in disp_drilldown.columns:
                                disp_drilldown[col] = disp_drilldown[col].apply(format_curr)

                        st.dataframe(disp_drilldown, use_container_width=True)

                        if not filtered_df.empty:
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                filtered_df[display_cols].to_excel(writer, index=False, sheet_name='EOL_Asset_List')
                            safe_loc = selected_loc.replace('/', '_').replace(' ', '_')[:15]
                            safe_dev = selected_dev.replace('/', '_').replace(' ', '_')[:15]
                            safe_range = range_label.replace(' ', '_')
                            st.download_button(
                                label=f"📥 Download Excel Report ({range_label})",
                                data=output.getvalue(),
                                file_name=f"{file_tag}_{safe_range}_{safe_loc}_{safe_dev}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                    else:
                        st.success("No assets fall into this age range (or exceed lifecycle) for the selected view!")

            render_scenario((5, 10), "5 to 10 Years", merged_active_df, tab1)
            render_scenario((10, 15), "10 to 15 Years", merged_active_df, tab2)
            render_scenario((15, 20), "15 to 20 Years", merged_active_df, tab3)
            render_scenario("20+", "20+ Years", merged_active_df, tab4)

        elif analysis_mode == "🎯 Custom Slider":
            st.markdown("### 🎯 Select Custom Age")

            if 'sync_age_slider' not in st.session_state: st.session_state.sync_age_slider = 10
            if 'sync_age_num' not in st.session_state: st.session_state.sync_age_num = 10

            def update_from_slider(): st.session_state.sync_age_num = st.session_state.sync_age_slider
            def update_from_num(): st.session_state.sync_age_slider = st.session_state.sync_age_num

            age_col1, age_col2 = st.columns([4, 1])
            with age_col1:
                st.slider("Slide to select Target Lifecycle (Years)", min_value=0, max_value=25, step=1,
                          key='sync_age_slider', on_change=update_from_slider)
            with age_col2:
                st.number_input("Exact Age", min_value=0, max_value=25, step=1,
                                key='sync_age_num', on_change=update_from_num)

            selected_age = st.session_state.sync_age_num
            st.markdown(f"### 🔮 Scenario Simulation: Evaluating Assets with Target Lifecycle ≤ {selected_age} Years")

            valid_df = merged_active_df[merged_active_df['Age_Years'].notna()]
            custom_age_df = valid_df[(valid_df['Lifecycle (Years)'] <= selected_age) & (valid_df['Age_Years'] >= selected_age)].copy()
            baseline_df = valid_df[valid_df['Age_Years'] >= valid_df['Lifecycle (Years)']].copy()

            diff_count = len(custom_age_df) - len(baseline_df)
            diff_impact = custom_age_df['Financial Impact'].sum() - baseline_df['Financial Impact'].sum()
            word_count = "more" if diff_count >= 0 else "fewer"
            word_impact = "additional" if diff_impact >= 0 else "reduced"

            st.info(f"💡 **Scenario vs Current Global Rules:** If evaluated for assets with **Target Lifecycle ≤ {selected_age} years**, you would target **{abs(diff_count):,} {word_count}** assets compared to the current Modeler baseline. Resulting in **{format_curr(abs(diff_impact))}** {word_impact} financial impact.")

            sc1, sc2, sc3, sc4 = st.columns(4)
            sc1.metric("Assets Eligible for Replacement", f"{len(custom_age_df):,}")
            sc2.metric("Total Gross Value", format_curr(custom_age_df['Gross value'].sum()))
            sc3.metric("Estimated Salvage Value", format_curr(custom_age_df['Est. Salvage Value'].sum()))
            sc4.metric("Financial Impact", format_curr(custom_age_df['Financial Impact'].sum()), delta="Scenario Impact", delta_color="inverse")

            if not custom_age_df.empty:
                sc_c1, sc_c2 = st.columns(2)
                with sc_c1:
                    sc_pie = px.pie(custom_age_df.groupby('Asset Category Description')['Financial Impact'].sum().reset_index(),
                                    names='Asset Category Description', values='Financial Impact', hole=0.5,
                                    title="Category-wise Financial Impact")
                    st.plotly_chart(sc_pie, use_container_width=True)
                with sc_c2:
                    sc_bar = px.bar(custom_age_df.groupby('Cost Center desc.')['Financial Impact'].sum().nlargest(10).reset_index(),
                                    x='Cost Center desc.', y='Financial Impact', color='Cost Center desc.',
                                    title="Top 10 Locations by Impact")
                    sc_bar.update_layout(showlegend=False, xaxis_title="Location", yaxis_title="Impact (Exact ₹)")
                    st.plotly_chart(sc_bar, use_container_width=True)
            else:
                st.info("No assets fall into this custom scenario.")

            st.markdown("---")
            tab_preview_custom, tab_custom = st.tabs(["📊 Data Preview", f"Scenario Data (Target Lifecycle ≤ {selected_age} Yrs)"])

            with tab_preview_custom:
                st.markdown("### 📊 Raw Data Explorer")
                f_col1, f_col2, f_col3 = st.columns(3)
                with f_col1:
                    cat_options = ["All"] + list(merged_active_df['Asset Category Description'].dropna().unique())
                    preview_cat = st.selectbox("Filter Category:", cat_options, key=f"prev_cat_{file_tag}_custom")
                with f_col2:
                    loc_options = ["All"] + list(merged_active_df['Cost Center desc.'].dropna().unique())
                    preview_loc = st.selectbox("Filter Location:", loc_options, key=f"prev_loc_{file_tag}_custom")
                with f_col3:
                    preview_search = st.text_input("Search Asset No. / Serial No.:", key=f"prev_search_{file_tag}_custom")
                preview_df = merged_active_df.copy()
                if preview_cat != "All": preview_df = preview_df[preview_df['Asset Category Description'] == preview_cat]
                if preview_loc != "All": preview_df = preview_df[preview_df['Cost Center desc.'] == preview_loc]
                if preview_search.strip():
                    s = preview_search.strip()
                    preview_df = preview_df[
                        preview_df['Asset No.'].astype(str).str.contains(s, case=False, na=False) |
                        preview_df['Serial No.'].astype(str).str.contains(s, case=False, na=False)
                    ]
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Assets", f"{len(preview_df):,}")
                m2.metric("Total Gross Value", format_curr(preview_df['Gross value'].sum()))
                avg_age = preview_df['Age_Years'].mean()
                m3.metric("Average Asset Age", f"{avg_age:.2f} Years" if pd.notna(avg_age) else "N/A")
                disp_preview = preview_df.copy()
                if 'Gross value' in disp_preview.columns:
                    disp_preview['Gross value'] = disp_preview['Gross value'].apply(format_curr)
                show_cols_p = ['Asset No.', 'Serial No.', 'Asset Category Description', 'Cost Center desc.', 'Gross value', 'Age_Years', 'Lifecycle (Years)', 'Risk_Level']
                st.dataframe(disp_preview[[c for c in show_cols_p if c in disp_preview.columns]], use_container_width=True)

            with tab_custom:
                st.markdown("### 🔍 Deep Dive: Scenario Asset Report")
                filter_col1, filter_col2 = st.columns(2)
                with filter_col1:
                    locations = ["All"] + list(custom_age_df['Cost Center desc.'].dropna().unique())
                    selected_loc = st.selectbox("Select Location:", locations, key=f"loc_custom_{file_tag}_tab")
                with filter_col2:
                    devices = ["All"] + list(custom_age_df['Asset Category Description'].dropna().unique())
                    selected_dev = st.selectbox("Select Category:", devices, key=f"dev_custom_{file_tag}_tab")
                filtered_df = custom_age_df.copy()
                if selected_loc != "All": filtered_df = filtered_df[filtered_df['Cost Center desc.'] == selected_loc]
                if selected_dev != "All": filtered_df = filtered_df[filtered_df['Asset Category Description'] == selected_dev]
                desired_cols = ['Asset No.', 'Serial No.', 'Dataset_Source', 'Asset Category Description', 'Cost Center desc.', 'Description', 'Gross value', 'Age_Years', 'Lifecycle (Years)', 'Risk_Level', 'Est. Salvage Value', 'Financial Impact']
                display_cols = [col for col in desired_cols if col in filtered_df.columns]
                disp_drilldown = filtered_df[display_cols].copy()
                for col in ['Gross value', 'Est. Salvage Value', 'Financial Impact']:
                    if col in disp_drilldown.columns: disp_drilldown[col] = disp_drilldown[col].apply(format_curr)
                st.dataframe(disp_drilldown, use_container_width=True)
                if not filtered_df.empty:
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        filtered_df[display_cols].to_excel(writer, index=False, sheet_name='Scenario_List')
                    safe_loc = selected_loc.replace('/', '_').replace(' ', '_')[:15]
                    safe_dev = selected_dev.replace('/', '_').replace(' ', '_')[:15]
                    st.download_button(
                        label="📥 Download Scenario Report",
                        data=output.getvalue(),
                        file_name=f"{file_tag}_Scenario_{selected_age}Yrs_{safe_loc}_{safe_dev}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

        else:  # Exact Year
            st.markdown("### 🎯 Select Exact Year")

            if 'sync_age_slider' not in st.session_state: st.session_state.sync_age_slider = 10
            if 'sync_age_num' not in st.session_state: st.session_state.sync_age_num = 10

            def update_from_slider_exact(): st.session_state.sync_age_num = st.session_state.sync_age_slider
            def update_from_num_exact(): st.session_state.sync_age_slider = st.session_state.sync_age_num

            age_col1, age_col2 = st.columns([4, 1])
            with age_col1:
                st.slider("Slide to select Exact Age (Years)", min_value=0, max_value=25, step=1,
                          key='sync_age_slider', on_change=update_from_slider_exact)
            with age_col2:
                st.number_input("Exact Age", min_value=0, max_value=25, step=1,
                                key='sync_age_num', on_change=update_from_num_exact)

            selected_age = st.session_state.sync_age_num
            st.markdown(f"### 📊 Exact Year Analysis: Assets strictly between {selected_age} and {selected_age + 1} Years")

            valid_df = merged_active_df[merged_active_df['Age_Years'].notna()]
            exact_year_df = valid_df[(valid_df['Age_Years'] >= selected_age) &
                                     (valid_df['Age_Years'] < selected_age + 1) &
                                     (valid_df['Age_Years'] >= valid_df['Lifecycle (Years)'])].copy()
            cumulative_df = valid_df[(valid_df['Lifecycle (Years)'] <= selected_age) & (valid_df['Age_Years'] >= selected_age)].copy()

            diff_count = len(cumulative_df) - len(exact_year_df)
            diff_impact = cumulative_df['Financial Impact'].sum() - exact_year_df['Financial Impact'].sum()
            st.info(f"💡 **Exact Year vs Cumulative Scenario:** By focusing strictly on assets aged **{selected_age} to {selected_age+1} years**, you isolate **{len(exact_year_df):,}** assets. A cumulative target (≥ {selected_age} Yrs) would encompass **{abs(diff_count):,} more** assets, carrying an additional **{format_curr(abs(diff_impact))}** in financial impact.")

            sc1, sc2, sc3, sc4 = st.columns(4)
            sc1.metric("Assets in Exact Year", f"{len(exact_year_df):,}")
            sc2.metric("Total Gross Value", format_curr(exact_year_df['Gross value'].sum()))
            sc3.metric("Estimated Salvage Value", format_curr(exact_year_df['Est. Salvage Value'].sum()))
            sc4.metric("Financial Impact", format_curr(exact_year_df['Financial Impact'].sum()), delta="Exact Year Impact", delta_color="inverse")

            if not exact_year_df.empty:
                sc_c1, sc_c2 = st.columns(2)
                with sc_c1:
                    sc_pie = px.pie(exact_year_df.groupby('Asset Category Description')['Financial Impact'].sum().reset_index(),
                                    names='Asset Category Description', values='Financial Impact', hole=0.5,
                                    title="Category-wise Financial Impact")
                    st.plotly_chart(sc_pie, use_container_width=True)
                with sc_c2:
                    sc_bar = px.bar(exact_year_df.groupby('Cost Center desc.')['Financial Impact'].sum().nlargest(10).reset_index(),
                                    x='Cost Center desc.', y='Financial Impact', color='Cost Center desc.',
                                    title="Top 10 Locations by Impact")
                    sc_bar.update_layout(showlegend=False, xaxis_title="Location", yaxis_title="Impact (Exact ₹)")
                    st.plotly_chart(sc_bar, use_container_width=True)
            else:
                st.info(f"No assets fall exactly in the {selected_age} to {selected_age+1} year range that have exceeded their lifecycle.")

            st.markdown("---")
            tab_preview_exact, tab_exact = st.tabs(["📊 Data Preview", f"Exact Year Data ({selected_age}-{selected_age+1} Yrs)"])

            with tab_preview_exact:
                st.markdown("### 📊 Raw Data Explorer")
                f_col1, f_col2, f_col3 = st.columns(3)
                with f_col1:
                    cat_options = ["All"] + list(merged_active_df['Asset Category Description'].dropna().unique())
                    preview_cat = st.selectbox("Filter Category:", cat_options, key=f"prev_cat_{file_tag}_exact")
                with f_col2:
                    loc_options = ["All"] + list(merged_active_df['Cost Center desc.'].dropna().unique())
                    preview_loc = st.selectbox("Filter Location:", loc_options, key=f"prev_loc_{file_tag}_exact")
                with f_col3:
                    preview_search = st.text_input("Search Asset No. / Serial No.:", key=f"prev_search_{file_tag}_exact")
                preview_df = merged_active_df.copy()
                if preview_cat != "All": preview_df = preview_df[preview_df['Asset Category Description'] == preview_cat]
                if preview_loc != "All": preview_df = preview_df[preview_df['Cost Center desc.'] == preview_loc]
                if preview_search.strip():
                    s = preview_search.strip()
                    preview_df = preview_df[
                        preview_df['Asset No.'].astype(str).str.contains(s, case=False, na=False) |
                        preview_df['Serial No.'].astype(str).str.contains(s, case=False, na=False)
                    ]
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Assets", f"{len(preview_df):,}")
                m2.metric("Total Gross Value", format_curr(preview_df['Gross value'].sum()))
                avg_age = preview_df['Age_Years'].mean()
                m3.metric("Average Asset Age", f"{avg_age:.2f} Years" if pd.notna(avg_age) else "N/A")
                disp_preview = preview_df.copy()
                if 'Gross value' in disp_preview.columns:
                    disp_preview['Gross value'] = disp_preview['Gross value'].apply(format_curr)
                show_cols_p = ['Asset No.', 'Serial No.', 'Asset Category Description', 'Cost Center desc.', 'Gross value', 'Age_Years', 'Lifecycle (Years)', 'Risk_Level']
                st.dataframe(disp_preview[[c for c in show_cols_p if c in disp_preview.columns]], use_container_width=True)

            with tab_exact:
                st.markdown("### 🔍 Deep Dive: Exact Year Asset Report")
                filter_col1, filter_col2 = st.columns(2)
                with filter_col1:
                    locations = ["All"] + list(exact_year_df['Cost Center desc.'].dropna().unique())
                    selected_loc = st.selectbox("Select Location:", locations, key=f"loc_exact_{file_tag}_tab")
                with filter_col2:
                    devices = ["All"] + list(exact_year_df['Asset Category Description'].dropna().unique())
                    selected_dev = st.selectbox("Select Category:", devices, key=f"dev_exact_{file_tag}_tab")
                filtered_df = exact_year_df.copy()
                if selected_loc != "All": filtered_df = filtered_df[filtered_df['Cost Center desc.'] == selected_loc]
                if selected_dev != "All": filtered_df = filtered_df[filtered_df['Asset Category Description'] == selected_dev]
                desired_cols = ['Asset No.', 'Serial No.', 'Dataset_Source', 'Asset Category Description', 'Cost Center desc.', 'Description', 'Gross value', 'Age_Years', 'Lifecycle (Years)', 'Risk_Level', 'Est. Salvage Value', 'Financial Impact']
                display_cols = [col for col in desired_cols if col in filtered_df.columns]
                disp_drilldown = filtered_df[display_cols].copy()
                for col in ['Gross value', 'Est. Salvage Value', 'Financial Impact']:
                    if col in disp_drilldown.columns: disp_drilldown[col] = disp_drilldown[col].apply(format_curr)
                st.dataframe(disp_drilldown, use_container_width=True)
                if not filtered_df.empty:
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        filtered_df[display_cols].to_excel(writer, index=False, sheet_name='Exact_Year_List')
                    safe_loc = selected_loc.replace('/', '_').replace(' ', '_')[:15]
                    safe_dev = selected_dev.replace('/', '_').replace(' ', '_')[:15]
                    st.download_button(
                        label="📥 Download Exact Year Report",
                        data=output.getvalue(),
                        file_name=f"{file_tag}_ExactYear_{selected_age}Yrs_{safe_loc}_{safe_dev}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

except FileNotFoundError as e:
    st.error(f"❌ Error: Could not find the file. Please ensure '{e.filename}' is in the same folder as app.py!")
except Exception as e:
    st.error(f"❌ An error occurred: {str(e)}")
    import traceback
    st.code(traceback.format_exc())

# =========================================================================
# 🟢 UNIFIED STATUS ANALYTICS
# =========================================================================
st.markdown("---")
st.header("🎯 Unified Status Analytics")
st.write("File-specific extraction: Uses ONLY 'Approver Comments' for Office Eq & F&F, and Merges 'Status' + 'Comments' for IT Admin.")

try:
    analyze_df = merged_active_df.copy()

    status_cols = [c for c in analyze_df.columns if 'status' in str(c).lower()]
    comment_cols = [c for c in analyze_df.columns if 'approver comment' in str(c).lower()]

    def exact_file_specific_extractor(row):
        source = str(row.get('Dataset_Source', ''))
        status_vals = [str(row[c]).strip().upper() for c in status_cols]
        comment_vals = [str(row[c]).strip().upper() for c in comment_cols]
        categories = ['AVAILABLE', 'OBSOLETE', 'FAULTY', 'NOTPERTAIN', 'REJECT', 'THEFT', 'MISSING', 'NOTTRACED']

        def find_category(text_list):
            for text in text_list:
                if text in ['NAN', 'NONE', 'NULL', 'NAT', '']: continue
                clean_text = text.replace(' ', '')
                for cat in categories:
                    if cat in clean_text or cat in text:
                        return cat
            return None

        if source == 'IT & Admin (Pending PV)':
            found = find_category(status_vals)
            if found: return found
            found = find_category(comment_vals)
            if found: return found
            for text in status_vals + comment_vals:
                if text not in ['NAN', 'NONE', 'NULL', 'NAT', '']: return "OTHER"
        else:
            found = find_category(comment_vals)
            if found: return found
            for text in comment_vals:
                if text not in ['NAN', 'NONE', 'NULL', 'NAT', '']: return "OTHER"
        return "BLANK / UNSPECIFIED"

    analyze_df['Final_Unified_Status'] = analyze_df.apply(exact_file_specific_extractor, axis=1)
    status_list = sorted([s for s in analyze_df['Final_Unified_Status'].unique() if "BLANK" not in s])

    if status_list:
        selected_status = st.selectbox(
            "🔽 Select Asset Status to Analyze:",
            ["All"] + status_list,
            key=f"unified_status_dropdown_{file_tag}"
        )

        if selected_status != "All":
            filtered_df = analyze_df[analyze_df['Final_Unified_Status'] == selected_status].copy()
            st.markdown(f"### 📊 Analytics for: **{selected_status}**")

            st_count = len(filtered_df)
            st_gross = pd.to_numeric(filtered_df['Gross value'], errors='coerce').fillna(0).sum()
            st_salvage = filtered_df['Est. Salvage Value'].sum() if 'Est. Salvage Value' in filtered_df.columns else 0
            st_impact = filtered_df['Financial Impact'].sum() if 'Financial Impact' in filtered_df.columns else 0

            sm1, sm2, sm3, sm4 = st.columns(4)
            sm1.metric("Total Assets", f"{st_count:,}")
            sm2.metric("Total Gross Value", format_curr(st_gross))
            sm3.metric("Est. Salvage Value", format_curr(st_salvage))
            sm4.metric("📉 Financial Impact", format_curr(st_impact), delta_color="inverse")

            st.markdown("---")
            chart_col1, chart_col2 = st.columns(2)
            with chart_col1:
                st.markdown("#### Category-wise Distribution")
                if st_impact > 0:
                    fig_st_pie = px.pie(filtered_df, names='Asset Category Description', values='Financial Impact', hole=0.4)
                else:
                    fig_st_pie = px.pie(filtered_df, names='Asset Category Description', title="By Asset Count", hole=0.4)
                fig_st_pie.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_st_pie, use_container_width=True)
            with chart_col2:
                st.markdown("#### Location-wise Impact")
                if st_impact > 0:
                    loc_st_df = filtered_df.groupby('Cost Center desc.')['Financial Impact'].sum().nlargest(10).reset_index()
                    fig_st_bar = px.bar(loc_st_df, x='Cost Center desc.', y='Financial Impact', color='Cost Center desc.')
                else:
                    loc_st_df = filtered_df['Cost Center desc.'].value_counts().nlargest(10).reset_index()
                    loc_st_df.columns = ['Cost Center desc.', 'Count']
                    fig_st_bar = px.bar(loc_st_df, x='Cost Center desc.', y='Count', color='Cost Center desc.')
                fig_st_bar.update_layout(showlegend=False, xaxis_title="Location", yaxis_title="Impact / Count")
                st.plotly_chart(fig_st_bar, use_container_width=True)

            st.markdown("### 📋 Detailed Asset List")
            base_cols = ['Asset No.', 'Description', 'Dataset_Source', 'Asset Category Description', 'Cost Center desc.', 'Age_Years', 'Gross value', 'Financial Impact', 'Final_Unified_Status']
            cols_to_show = [c for c in base_cols if c in filtered_df.columns]
            for sc in status_cols:
                if sc in filtered_df.columns and sc not in cols_to_show: cols_to_show.append(sc)
            for cc in comment_cols:
                if cc in filtered_df.columns and cc not in cols_to_show: cols_to_show.append(cc)

            disp_df = filtered_df[cols_to_show].copy()
            for c in ['Gross value', 'Financial Impact']:
                if c in disp_df.columns: disp_df[c] = disp_df[c].apply(format_curr)
            st.dataframe(disp_df, use_container_width=True)

            output_st = BytesIO()
            with pd.ExcelWriter(output_st, engine='openpyxl') as writer:
                filtered_df[cols_to_show].to_excel(writer, index=False, sheet_name="Unified_Assets")
            safe_stat = str(selected_status)[:15].replace('/', '_')
            st.download_button(
                label=f"📥 Download {selected_status} Report",
                data=output_st.getvalue(),
                file_name=f"{file_tag}_{safe_stat}_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
    else:
        st.info("⚠️ No status or comment data found to analyze in this view.")

except Exception as e:
    st.warning(f"⚠️ Status analytics unavailable for this view: {e}")