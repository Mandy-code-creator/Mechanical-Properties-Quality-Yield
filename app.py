import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from matplotlib.patches import Patch

# --- PAGE CONFIG ---
st.set_page_config(page_title="Quality Dashboard", layout="wide")
st.title("📊 Production Quality Yield & Period Comparison")
st.markdown("---")

# --- GLOBAL SPECS (Fixed as per request) ---
GLOBAL_SPECS = {
    'YS': {'min': 400, 'max': 460, 'target': 430}, # 430 +/- 30
    'TS': {'min': 410, 'max': 470, 'target': 440}, # 440 +/- 30
    'EL': {'min': 25, 'max': None, 'target': None},
    'YPE': {'min': 4, 'max': None, 'target': None}
}

uploaded_file = st.file_uploader("Upload Excel data (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.astype(str).str.strip()

    # --- 1. DATA EXTRACTION ---
    if 'Thickness' in df.columns:
        df.rename(columns={'Thickness': 'Actual_Thickness'}, inplace=True)
    else:
        for i, c in enumerate(df.columns):
            if '型式' in c and i > 0:
                df.rename(columns={df.columns[i-1]: 'Actual_Thickness'}, inplace=True)
                break
    
    if 'Actual_Thickness' in df.columns:
        df['Actual_Thickness'] = pd.to_numeric(df['Actual_Thickness'], errors='coerce').round(3)

    if '烤三生產日期' in df.columns:
        d_str = df['烤三生產日期'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df['烤三生產日期'] = pd.to_datetime(d_str, format='%Y%m%d', errors='coerce').fillna(pd.to_datetime(d_str, errors='coerce'))

    # --- 2. GRADE & MECH FEATURES ---
    base_grades = ['A-B+', 'A-B', 'A-B-', 'B+', 'B']
    for g in base_grades:
        match_cols = [c for c in df.columns if g == c or str(c).startswith(f"{g}.")]
        df[g] = df[match_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) if match_cols else 0
    df['Total_Qty'] = df[base_grades].sum(axis=1)

    df.rename(columns={'烤漆降伏強度': 'YS', '烤漆抗拉強度': 'TS', '伸長率': 'EL'}, inplace=True)
    mech_features = ['YS', 'TS', 'EL', 'YPE', 'HARDNESS']
    for f in mech_features:
        if f in df.columns: df[f] = pd.to_numeric(df[f], errors='coerce')

    if '熱軋材質' in df.columns:
        df['HR_Material'] = df['熱軋材質'].astype(str).str.strip().replace(['nan', ''], 'Unknown')
    else:
        df['HR_Material'] = 'Unknown'

    # --- 3. TIME PERIOD LOGIC ---
    def categorize_period(d):
        if pd.isnull(d): return "Unknown"
        y = d.year
        q3_s, q3_e = pd.Timestamp(2025, 6, 29), pd.Timestamp(2025, 9, 30)
        if y == 2024: return "2024 (Full Year)"
        if y == 2025:
            if d < q3_s: return "2025 H1 (Until 06/28)"
            if q3_s <= d <= q3_e: return "2025 Q3 (06/29 - 09/30)"
            return "2025 Q4"
        if y == 2026: return "2026 Q1"
        return "Other"

    if '烤三生產日期' in df.columns:
        df['Time_Group'] = df['烤三生產日期'].apply(categorize_period)
        df = df[df['Time_Group'] != "Other"]
    else:
        df['Time_Group'] = "Unknown"

    # --- KHAI BÁO TABS (QUAN TRỌNG: Phải có tab3 ở đây mới hiện được Task 3) ---
    tab0, tab1, tab2, tab3 = st.tabs(["0. Raw Check", "1. Yield Summary", "2. Distribution Analysis", "3. Period Comparison"])

    with tab0:
        st.header("0. System Health Check")
        st.write(f"**Valid Rows:** {len(df)}")
        st.dataframe(df.head(5), use_container_width=True)

    # --- FILTERING ---
    df = df[df['Actual_Thickness'].isin([0.5, 0.6, 0.75, 0.8])]
    st.sidebar.header("🔎 Dashboard Filters")
    all_periods = sorted(df['Time_Group'].unique())
    ui_selection = st.sidebar.multiselect("📅 Select Period(s):", options=["All"] + all_periods, default=["All"])
    selected_periods = all_periods if ("All" in ui_selection or not ui_selection) else ui_selection
    df_filtered = df[df['Time_Group'].isin(selected_periods)]
    thickness_list = sorted(df['Actual_Thickness'].dropna().unique())

    # --- TAB 1: YIELD ---
    with tab1:
        st.header("1. Quality Yield Summary")
        g_cols = ['Time_Group', 'Actual_Thickness', 'HR_Material']
        sum_df = df_filtered.groupby(g_cols, dropna=False)[base_grades].sum().reset_index()
        sum_df['Total_Qty'] = sum_df[base_grades].sum(axis=1)
        for col in base_grades: 
            sum_df[f"% {col}"] = ((sum_df[col] / sum_df['Total_Qty'].replace(0, np.nan)) * 100).fillna(0).round(1)
        sum_df.rename(columns={'Time_Group': 'Period', 'Actual_Thickness': 'Thickness'}, inplace=True)
        for period in selected_periods:
            p_data = sum_df[sum_df['Period'] == period]
            if not p_data.empty:
                st.markdown(f"### 📅 Period: **{period}**")
                st.dataframe(p_data.drop(columns=['Period']), use_container_width=True, hide_index=True)

    # --- TAB 2: DISTRIBUTION ---
    def plot_dist(ax, data, feat, title, y_lim):
        c_map = {'A-B+': '#2ca02c', 'A-B': '#1f77b4', 'A-B-': '#ff7f0e', 'B+': '#9467bd', 'B': '#d62728'}
        v_l, w_l, clrs, m_info = [], [], [], []
        for g in base_grades:
            td = data[[feat, g]].dropna()
            if not td[td[g] > 0].empty:
                v_l.append(td[feat].values); w_l.append(td[g].values); clrs.append(c_map[g])
                m_info.append({'v': np.average(td[feat].values, weights=td[g].values), 'c': c_map[g]})
        if v_l:
            ax.hist(v_l, bins=15, weights=w_l, color=clrs, stacked=True, edgecolor='white', alpha=0.7)
            # Add Spec lines
            if feat in GLOBAL_SPECS:
                s = GLOBAL_SPECS[feat]
                if s['min']: ax.axvline(s['min'], color='red', ls='-', lw=1.5, label='Limit')
                if s['max']: ax.axvline(s['max'], color='red', ls='-', lw=1.5)
        ax.set_title(title, fontsize=10, fontweight='bold')
        ax.legend(fontsize=7)

    with tab2:
        st.header("2. Detailed Distribution Analysis")
        for period in selected_periods:
            df_p = df_filtered[df_filtered['Time_Group'] == period]
            if df_p.empty: continue
            st.markdown(f"### 📅 Period: **{period}**")
            cols = st.columns(2)
            for idx, f in enumerate([x for x in ['YS', 'TS', 'EL'] if x in df_p.columns]):
                with cols[idx%2]:
                    fig, ax = plt.subplots(figsize=(8, 4))
                    plot_dist(ax, df_p, f, f"{f} Distribution", 100)
                    st.pyplot(fig); plt.close(fig)

    # --- TAB 3: SIDE-BY-SIDE COMPARISON (FIXED) ---
    with tab3:
        st.header("3. Side-by-Side Period Comparison")
        st.info("Identify shifts in mechanical stability across production periods.")
        
        for thick in thickness_list:
            df_t = df_filtered[df_filtered['Actual_Thickness'] == thick]
            if df_t.empty: continue
            
            st.markdown(f"---")
            st.subheader(f"📏 Thickness: {thick} mm")
            
            comp_cols = st.columns(2)
            for idx, f in enumerate(['YS', 'TS', 'EL']):
                if f in df_t.columns:
                    with comp_cols[idx % 2]:
                        fig, ax = plt.subplots(figsize=(10, 6))
                        sns.boxplot(data=df_t, x='Time_Group', y=f, palette="Set2", ax=ax)
                        
                        # Add Spec Zone
                        if f in GLOBAL_SPECS:
                            s = GLOBAL_SPECS[f]
                            if s['min'] and s['max']:
                                ax.axhspan(s['min'], s['max'], color='green', alpha=0.1, label='Target Zone')
                            elif s['min']:
                                ax.axhline(s['min'], color='red', ls='--', alpha=0.5)
                        
                        ax.set_title(f"{f} Comparison (Thick: {thick})", fontsize=12, fontweight='bold')
                        plt.xticks(rotation=15)
                        st.pyplot(fig); plt.close(fig)

        # Quality Yield Trend
        st.markdown("### 📈 Quality Yield Trend (%)")
        for thick in thickness_list:
            df_y_all = df_filtered[df_filtered['Actual_Thickness'] == thick]
            if df_y_all.empty: continue
            
            yield_trend = df_y_all.groupby('Time_Group')[base_grades].sum()
            yield_trend = yield_trend.div(yield_trend.sum(axis=1), axis=0) * 100
            
            fig, ax = plt.subplots(figsize=(12, 5))
            yield_trend.plot(kind='bar', stacked=True, ax=ax, color=['#2ca02c', '#1f77b4', '#ff7f0e', '#9467bd', '#d62728'])
            ax.set_title(f"Quality Yield % Trend - Thickness: {thick}", fontsize=12, fontweight='bold')
            ax.set_ylabel("Percentage (%)")
            plt.xticks(rotation=0)
            st.pyplot(fig); plt.close(fig)
    # --- EXPORT SECTION ---
    st.sidebar.header("📥 Export Options")
    if st.sidebar.button("Download Detailed Excel"):
        towrite = io.BytesIO()
        pd.DataFrame(all_export_data).to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.sidebar.download_button(label="Click to Download Excel", data=towrite, file_name="QC_Detailed_Optimization.xlsx")

    # --- PDF EXPORT ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("🖨️ PDF Reports")
    def clean(t): return str(t).replace('±', '+/-').replace('–', '-').encode('latin-1', 'ignore').decode('latin-1')

    if st.sidebar.button("Generate PDF Report"):
        pdf = FPDF(orientation='L')
        if 'display_df' in locals() and not display_df.empty:
            pdf.add_page(); pdf.set_font('Arial', 'B', 16); pdf.cell(0, 10, "1. YIELD SUMMARY", ln=True, align="C"); pdf.ln(5)
            pdf.set_font('Arial', 'B', 8)
            cw_tab1 = [25, 15, 25, 15] + [12]*len(base_grades) + [12]*len(p_cols)
            for i, col in enumerate(display_df.columns): pdf.cell(cw_tab1[i] if i < len(cw_tab1) else 20, 8, clean(col), border=1, align='C')
            pdf.ln(); pdf.set_font('Arial', '', 8)
            for _, r in display_df.head(25).iterrows(): 
                for i, v in enumerate(r): pdf.cell(cw_tab1[i] if i < len(cw_tab1) else 20, 8, clean(v), border=1, align='C')
                pdf.ln()

        heads = ["Feature", "Method", "Limit", "Segment Dist", "TARGET", "TOL", f"MILL {sigma_mill}σ", f"RELEASE {sigma_release}σ"]
        c_w3 = [16, 22, 24, 60, 15, 12, 28, 30] 

        for period in selected_periods:
            safe_p = "".join([c if c.isalnum() else "_" for c in period])
            pdf.add_page(); pdf.set_font('Arial', 'B', 16); pdf.cell(0, 10, f"--- PERIOD: {period} ---", ln=True, align="C"); pdf.ln(5)
            
            pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, "Overall Distribution", ln=True); ys = pdf.get_y()
            for idx, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
                path = f"overall_{f}_{safe_p}.png"
                if os.path.exists(path): pdf.image(path, x=(10 if idx%2==0 else 150), y=(ys if idx<2 else ys+75), w=135)
            
            pdf.add_page(); pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, "Overall Goals", ln=True)
            pdf.set_font('Arial', 'B', 8)
            for i, h in enumerate(heads): pdf.cell(c_w3[i], 7, clean(h), border=1, align='C')
            pdf.ln(); pdf.set_font('Arial', '', 7)
            
            last_feat = ""
            for row in [r for r in overall_export_data if r['Period'] == period]:
                is_first = (row["Feature"] != last_feat)
                if is_first: last_feat = row["Feature"]
                v_list = [row["Feature"] if is_first else "", row["Method"], row["Limit"], row["Segment Dist"] if is_first else "", str(row["TARGET GOAL"]), str(row["TOLERANCE"]), row[f"MILL {sigma_mill}σ"], row[f"RELEASE {sigma_release}σ"]]
                for i, v in enumerate(v_list): pdf.cell(c_w3[i], 7, clean(v), border=1, align='C')
                pdf.ln()

            for thick in thickness_list:
                period_thick_data = [r for r in all_export_data if r['Period'] == period and r['Thickness'] == thick]
                if not period_thick_data: continue
                
                pdf.add_page(); pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, f"Distribution - Thick: {thick}", ln=True); ys = pdf.get_y()
                for idx, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
                    path = f"dist_{f}_{thick}_{safe_p}.png"
                    if os.path.exists(path): pdf.image(path, x=(10 if idx%2==0 else 150), y=(ys if idx<2 else ys+75), w=135)
                
                pdf.add_page(); pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, f"Control Limits - Thick: {thick}", ln=True)
                pdf.set_font('Arial', 'B', 8)
                for i, h in enumerate(heads): pdf.cell(c_w3[i], 7, clean(h), border=1, align='C')
                pdf.ln(); pdf.set_font('Arial', '', 7)
                
                last_feat_t = ""
                for row in period_thick_data:
                    is_first = (row["Feature"] != last_feat_t)
                    if is_first: last_feat_t = row["Feature"]
                    v_list = [row["Feature"] if is_first else "", row["Method"], row["Limit"] if is_first else "", row["Segment Dist"] if is_first else "", str(row["TARGET GOAL"]), str(row["TOLERANCE"]), row[f"MILL {sigma_mill}σ"], row[f"RELEASE {sigma_release}σ"]]
                    for i, v in enumerate(v_list): pdf.cell(c_w3[i], 7, clean(v), border=1, align='C')
                    pdf.ln()
                
                pdf.add_page(); pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, f"I-MR Charts - Thick: {thick}", ln=True)
                y_imr = pdf.get_y() + 2 
                for idx, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
                    path = f"imr_{f}_{thick}_{safe_p}.png"
                    if os.path.exists(path): pdf.image(path, x=(10 if idx%2==0 else 150), y=(y_imr if idx<2 else y_imr+90), w=130)

        pdf.output("QC_Report.pdf")
        with open("QC_Report.pdf", "rb") as f:
            st.sidebar.download_button("📥 Download PDF", f.read(), "QC_Report.pdf", "application/pdf")
