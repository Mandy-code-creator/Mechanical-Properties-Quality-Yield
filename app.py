import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import io
from matplotlib.patches import Patch
from fpdf import FPDF
import os

# --- PAGE CONFIG ---
st.set_page_config(page_title="Quality Dashboard", layout="wide")
st.title("📊 Production Quality Yield & Distribution")
st.markdown("---")

GLOBAL_SPECS = {
    'YS': {'min': 400, 'max': 460, 'target': 430},
    'TS': {'min': 410, 'max': 470, 'target': 440},
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
    good_grades = ['A-B+', 'A-B']
    bad_grades = ['A-B-', 'B+', 'B']
    
    for g in base_grades:
        match_cols = [c for c in df.columns if c == g or str(c).startswith(f"{g}.")]
        df[g] = df[match_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) if match_cols else 0
    
    df['Total_Qty'] = df[base_grades].sum(axis=1)
    df['Good_Qty'] = df[good_grades].sum(axis=1)
    df['Bad_Qty'] = df[bad_grades].sum(axis=1)

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
        df_25 = df[df['烤三生產日期'].dt.year == 2025].copy()
        if not df_25.empty:
            df_25['Time_Group'] = "2025 (Full Year)"
            df = pd.concat([df, df_25], ignore_index=True)
    else:
        df['Time_Group'] = "Unknown"

    # --- 4. NAVIGATION MENU ---
    st.sidebar.header("📂 Menu Điều Hướng")
    menu_options = ["0. Raw Check", "1. Yield Summary", "2. Distribution Analysis", "3. Executive Auto-Insight", "4. I-MR Stability Tracking", "5. I-MR (Acceptable Grades)"]
    current_tab = st.sidebar.radio("Chọn Task muốn xem:", menu_options)
    st.sidebar.markdown("---")

    # --- GLOBAL FILTERING ---
    df = df[df['Actual_Thickness'].isin([0.5, 0.6, 0.75, 0.8])]
    st.sidebar.header("🔎 Filters")
    all_periods = sorted(df['Time_Group'].unique())
    ui_selection = st.sidebar.multiselect("📅 Select Period(s):", options=["All"] + all_periods, default=["All"])
    selected_periods = all_periods if ("All" in ui_selection or not ui_selection) else ui_selection
    df_filtered = df[df['Time_Group'].isin(selected_periods)].copy()
    thickness_list = sorted(df['Actual_Thickness'].dropna().unique())

    # --- TASK 0: RAW CHECK ---
    if current_tab == "0. Raw Check":
        st.dataframe(df_filtered.head(10), use_container_width=True)

    # --- TASK 1: YIELD SUMMARY ---
    elif current_tab == "1. Yield Summary":
        st.header("1. Quality Yield Summary & Worst Offenders")
        severe_grades = ['B+', 'B']
        df_filtered['Severe_Bad_Qty'] = df_filtered[severe_grades].sum(axis=1)
        df_filtered['Acceptable_Qty'] = df_filtered['Total_Qty'] - df_filtered['Severe_Bad_Qty']
        
        period_summary = df_filtered.groupby(['Time_Group', 'Actual_Thickness', 'HR_Material'])[['Total_Qty', 'Acceptable_Qty', 'Severe_Bad_Qty']].sum().reset_index()
        period_summary = period_summary[period_summary['Total_Qty'] > 0]
        period_summary['Yield (%)'] = (period_summary['Acceptable_Qty'] / period_summary['Total_Qty'] * 100).fillna(0).round(2)
        period_summary['Defect_Rate (%)'] = (period_summary['Severe_Bad_Qty'] / period_summary['Total_Qty'] * 100).fillna(0).round(2)
        period_summary = period_summary.sort_values(by=['Defect_Rate (%)', 'Severe_Bad_Qty'], ascending=[False, False])
        period_summary.rename(columns={'Severe_Bad_Qty': 'Bad_Qty (B+, B)'}, inplace=True)

        if not period_summary.empty and period_summary['Bad_Qty (B+, B)'].sum() > 0:
            worst_row = period_summary.iloc[0]
            st.error(f"⚠️ **Executive Insight:** The highest risk item is **{worst_row['Actual_Thickness']:.2f}mm ({worst_row['HR_Material']})** during **{worst_row['Time_Group']}**, hitting a severe defect rate of **{worst_row['Defect_Rate (%)']:.2f}%**.")
            st.dataframe(period_summary.style.background_gradient(subset=['Defect_Rate (%)'], cmap='Reds').background_gradient(subset=['Yield (%)'], cmap='Greens').format({'Actual_Thickness': '{:.2f}', 'Total_Qty': '{:.0f}', 'Acceptable_Qty': '{:.0f}', 'Bad_Qty (B+, B)': '{:.0f}', 'Yield (%)': '{:.2f}%', 'Defect_Rate (%)': '{:.2f}%'}), use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("📑 Detailed Yield by Period (All Grades)")
        sum_df = df_filtered.groupby(['Time_Group', 'Actual_Thickness', 'HR_Material'])[base_grades].sum().reset_index()
        sum_df['Total_Qty'] = sum_df[base_grades].sum(axis=1)
        for col in base_grades: sum_df[f"% {col}"] = ((sum_df[col] / sum_df['Total_Qty'].replace(0, np.nan)) * 100).fillna(0).round(1)
        
        for period in selected_periods:
            p_data = sum_df[sum_df['Time_Group'] == period]
            if not p_data.empty:
                st.markdown(f"#### 📅 Period: **{period}**")
                format_dict = {'Actual_Thickness': '{:.2f}', 'Total_Qty': '{:.0f}'}
                for col in base_grades: format_dict[col] = '{:.0f}'; format_dict[f"% {col}"] = '{:.1f}%'
                st.dataframe(p_data.style.format(format_dict), use_container_width=True, hide_index=True)

    # --- TASK 2: DISTRIBUTION ---
    elif current_tab == "2. Distribution Analysis":
        global_x_bounds = {}
        for feat in mech_features:
            if feat in df.columns:
                vd = df[[feat, 'Total_Qty']].dropna().copy()
                if not vd.empty:
                    q1, q99 = np.percentile(vd[feat], 1), np.percentile(vd[feat], 99)
                    global_x_bounds[feat] = (q1 - (q99-q1)*0.25, q99 + (q99-q1)*0.25)

        def get_shared_y(data, features):
            max_y = 0
            for feat in features:
                if feat in data.columns:
                    vd = data.dropna(subset=[feat])
                    if not vd.empty:
                        cnts, _ = np.histogram(vd[feat], bins=15, weights=vd['Total_Qty'])
                        max_y = max(max_y, cnts.max())
            return max_y * 1.35 if max_y > 0 else 50

        def plot_dist(ax, data, feat, title, y_lim):
            c_map = {'A-B+': '#2ca02c', 'A-B': '#1f77b4', 'A-B-': '#ff7f0e', 'B+': '#9467bd', 'B': '#d62728'}
            fmin, fmax = global_x_bounds.get(feat, (0, 100))
            v_l, w_l, clrs, m_info = [], [], [], []
            for g in base_grades:
                td = data[[feat, g]].dropna()
                td = td[td[g] > 0]
                if not td.empty:
                    v_l.append(td[feat].values); w_l.append(td[g].values); clrs.append(c_map[g])
                    m = np.average(td[feat].values, weights=td[g].values)
                    ax.axvline(m, color=c_map[g], ls='--', lw=1.2)
                    m_info.append({'v': m, 'c': c_map[g]})
            if v_l:
                ax.hist(v_l, bins=np.linspace(fmin, fmax, 16), weights=w_l, color=clrs, stacked=True, edgecolor='white', alpha=0.7)
                for i, info in enumerate(sorted(m_info, key=lambda x: x['v'])):
                    h = y_lim * (0.85 - (i % 3) * 0.12)
                    ax.text(info['v'], h, f"{info['v']:.1f}", color='white', fontweight='bold', fontsize=8, ha='center', bbox=dict(facecolor=info['c'], alpha=0.8, boxstyle='round,pad=0.2'))
            ax.set_xlim(fmin, fmax); ax.set_ylim(0, y_lim); ax.set_title(title, fontsize=10, fontweight='bold')

        for period in selected_periods:
            df_p = df_filtered[df_filtered['Time_Group'] == period]
            if df_p.empty: continue
            st.markdown(f"## 📅 Period: **{period}**")
            ov_y = get_shared_y(df_p, ['YS', 'TS', 'EL', 'YPE'])
            cols = st.columns(2)
            for idx, f in enumerate([x for x in ['YS', 'TS', 'EL', 'YPE'] if x in df_p.columns]):
                with cols[idx%2]:
                    fig, ax = plt.subplots(figsize=(8, 4.5))
                    plot_dist(ax, df_p, f, f"{f} (Overall)", ov_y)
                    st.pyplot(fig); plt.close(fig)

    # --- TASK 3: EXECUTIVE AUTO-INSIGHT ---
    elif current_tab == "3. Executive Auto-Insight":
        st.header("🧠 Executive Auto-Insight & Root Cause")
        df_filtered['Severe_Bad_Qty'] = df_filtered[['B+', 'B']].sum(axis=1)
        df_filtered['Spec_Label'] = df_filtered['Actual_Thickness'].astype(str) + "mm (" + df_filtered['HR_Material'] + ")"

        heat_data = df_filtered.groupby(['Spec_Label', 'Time_Group']).apply(lambda x: (x['Severe_Bad_Qty'].sum() / x['Total_Qty'].sum() * 100) if x['Total_Qty'].sum() > 0 else 0)
        
        if not heat_data.empty and heat_data.max() > 0:
            heatmap_long = heat_data.reset_index(); heatmap_long.columns = ['Spec', 'Period', 'Defect_Rate']
            top_issues = heatmap_long[heatmap_long['Defect_Rate'] > 0].sort_values('Defect_Rate', ascending=False).head(5)
            
            rc_results = {}
            for f in ['YS', 'TS', 'EL', 'YPE']:
                if f in df_filtered.columns:
                    good_m = df_filtered[df_filtered['Total_Qty'] > df_filtered['Severe_Bad_Qty']][f].mean()
                    bad_m = df_filtered[df_filtered['Severe_Bad_Qty'] > 0][f].mean()
                    if pd.notnull(good_m) and pd.notnull(bad_m): rc_results[f] = bad_m - good_m
            
            rc_s = pd.Series(rc_results).dropna().sort_values(key=abs, ascending=False)
            if not top_issues.empty and not rc_s.empty:
                direction = "HIGHER ⬆️" if rc_s.iloc[0] > 0 else "LOWER ⬇️"
                st.success(f"### 🎯 EXECUTIVE CONCLUSION:\n* 🚨 **Hotspot:** {top_issues.iloc[0]['Spec']} in {top_issues.iloc[0]['Period']}\n* 🧠 **Driver:** {rc_s.index[0]} is {abs(rc_s.iloc[0]):.1f} {direction} in bad coils.")

            fig, ax = plt.subplots(figsize=(12, 5))
            sns.heatmap(heat_data.unstack(), annot=True, fmt=".1f", cmap="YlOrRd", ax=ax)
            st.pyplot(fig); plt.close(fig)

            # GLOBAL I-MR FOR B+/B
            st.markdown("---")
            st.header("📈 Global I-MR Stability Tracking (Severe Defects: B+ and Below)")
            df_sev_global = df_filtered[(df_filtered['B+'] > 0) | (df_filtered['B'] > 0)].sort_values(by='烤三生產日期').reset_index(drop=True)
            if not df_sev_global.empty:
                for feat in ['YS', 'TS', 'EL', 'YPE']:
                    if feat in df_sev_global.columns:
                        valid_data = df_sev_global.dropna(subset=[feat]).reset_index(drop=True)
                        if len(valid_data) > 1:
                            vals = valid_data[feat].values
                            x_seq = np.arange(len(vals)); mean_v = np.mean(vals)
                            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 8), gridspec_kw={'height_ratios': [2, 1]})
                            ax1.plot(x_seq, vals, marker='o', ms=4, lw=1, color='#1f77b4', alpha=0.6)
                            ax1.axhline(mean_v, color='green', ls='--')
                            ax1.text(x_seq[-1], mean_v, f" Mean: {mean_v:.1f}", color='green', fontweight='bold')
                            if feat in GLOBAL_SPECS:
                                s = GLOBAL_SPECS[feat]
                                if s['min']: ax1.axhline(s['min'], color='red', lw=2); ax1.text(x_seq[-1], s['min'], f" Min: {s['min']}", color='red')
                                if s['max']: ax1.axhline(s['max'], color='red', lw=2); ax1.text(x_seq[-1], s['max'], f" Max: {s['max']}", color='red')
                            st.pyplot(fig); plt.close(fig)

    # --- TASK 4: I-MR (B+/B) ---
    elif current_tab == "4. I-MR Stability Tracking":
        st.header("📈 Task 4: I-MR Stability Tracking (Severe Defects: B+/B)")
        df_sev = df_filtered[(df_filtered['B+'] > 0) | (df_filtered['B'] > 0)].copy()
        if not df_sev.empty:
            c1, c2, c3 = st.columns(3)
            sel_p = c1.selectbox("Filter Period:", sorted(df_sev['Time_Group'].unique()), key="t4_p")
            sel_t = c2.selectbox("Filter Thickness:", sorted(df_sev['Actual_Thickness'].unique()), key="t4_t")
            sel_m = c3.selectbox("Filter Material:", sorted(df_sev['HR_Material'].unique()), key="t4_m")
            imr_df = df_sev[(df_sev['Time_Group'] == sel_p) & (df_sev['Actual_Thickness'] == sel_t) & (df_sev['HR_Material'] == sel_m)].sort_values('烤三生產日期')
            if not imr_df.empty:
                for feat in ['YS', 'TS', 'EL', 'YPE']:
                    vals = imr_df[feat].dropna().values
                    if len(vals) > 1:
                        x_seq = np.arange(len(vals))
                        fig, ax = plt.subplots(figsize=(12, 4))
                        ax.plot(x_seq, vals, marker='o', color='red', alpha=0.6)
                        ax.axhline(np.mean(vals), color='green', ls='--')
                        ax.set_title(f"I-Chart {feat}: {sel_p} | {sel_t}mm")
                        st.pyplot(fig); plt.close(fig)
            else: st.warning("No data for this filter.")
        else: st.success("No severe defects found.")

    # --- TASK 5: I-MR (ACCEPTABLE) ---
    elif current_tab == "5. I-MR (Acceptable Grades)":
        st.header("📈 Task 5: I-MR Stability Tracking (Acceptable Grades)")
        df_acc = df_filtered[df_filtered['Total_Qty'] > (df_filtered['B+'] + df_filtered['B'])].copy()
        if not df_acc.empty:
            c1, c2, c3 = st.columns(3)
            sel_p = c1.selectbox("Filter Period:", sorted(df_acc['Time_Group'].unique()), key="t5_p")
            sel_t = c2.selectbox("Filter Thickness:", sorted(df_acc['Actual_Thickness'].unique()), key="t5_t")
            sel_m = c3.selectbox("Filter Material:", sorted(df_acc['HR_Material'].unique()), key="t5_m")
            imr_df = df_acc[(df_acc['Time_Group'] == sel_p) & (df_acc['Actual_Thickness'] == sel_t) & (df_acc['HR_Material'] == sel_m)].sort_values('烤三生產日期')
            if not imr_df.empty:
                for feat in ['YS', 'TS', 'EL', 'YPE']:
                    vals = imr_df[feat].dropna().values
                    if len(vals) > 1:
                        x_seq = np.arange(len(vals))
                        fig, ax = plt.subplots(figsize=(12, 4))
                        ax.plot(x_seq, vals, marker='o', color='green', alpha=0.6)
                        ax.axhline(np.mean(vals), color='blue', ls='--')
                        ax.set_title(f"I-Chart {feat} (Acceptable): {sel_p}")
                        st.pyplot(fig); plt.close(fig)

    # --- EXPORT PDF VISUAL REPORT (Bản gốc của Mandy) ---
    st.sidebar.header("📥 Export PDF Report")
    if st.sidebar.button("🖨️ Generate PDF containing Charts"):
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=15)
        if os.path.exists("export_heatmap.png"):
            pdf.add_page(); pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, "1. DEFECT HOTSPOT DIAGNOSTIC MAP", ln=True, align='C')
            pdf.image("export_heatmap.png", x=15, y=25, w=260)
        
        pdf.add_page(); pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, "3. I-MR PROCESS STABILITY TRACKING", ln=True, align='C')
        y_pos = 25; chart_count = 0
        for feat in ['YS', 'TS', 'EL', 'YPE']:
            img_path = f"export_imr_global_{feat}.png"
            if os.path.exists(img_path):
                if chart_count == 2: pdf.add_page(); y_pos = 20; chart_count = 0
                pdf.image(img_path, x=20, y=y_pos, w=250); y_pos += 90; chart_count += 1
        
        pdf.output("Quality_Visual_Report.pdf")
        with open("Quality_Visual_Report.pdf", "rb") as f:
            st.sidebar.download_button(label="✅ Click to Download PDF", data=f.read(), file_name="Quality_Visual_Report.pdf", mime="application/pdf")
        st.sidebar.success("PDF Generated Successfully!")
