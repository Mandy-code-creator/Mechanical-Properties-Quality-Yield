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

    # --- TẠO ID DUY NHẤT ĐỂ CHỐNG TRÙNG LẶP (ANTI-DOUBLE COUNTING) ---
    df['Row_ID'] = np.arange(len(df))

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

    # --- ĐỊNH NGHĨA TABS ---
    tab0, tab1, tab2, tab3, tab4 = st.tabs(["0. Raw Check", "1. Yield Summary", "2. Distribution Analysis", "3.🔍 Root Cause & Diagnostic", "4. I-MR Analysis"])

    with tab0:
        st.dataframe(df.head(10), use_container_width=True)

    # --- GLOBAL FILTERING ---
    df = df[df['Actual_Thickness'].isin([0.5, 0.6, 0.75, 0.8])]
    st.sidebar.header("🔎 Filters")
    all_periods = sorted(df['Time_Group'].unique())
    ui_selection = st.sidebar.multiselect("📅 Select Period(s):", options=["All"] + all_periods, default=["All"])
    selected_periods = all_periods if ("All" in ui_selection or not ui_selection) else ui_selection
    df_filtered = df[df['Time_Group'].isin(selected_periods)].copy()
    thickness_list = sorted(df['Actual_Thickness'].dropna().unique())

    # --- TAB 1: YIELD SUMMARY ---
    with tab1:
        st.header("1. Quality Yield Summary & Worst Offenders")
        st.info("Overview of production yield. Chronologically sorted from 2024 onwards.")

        st.subheader("📊 Executive Summary: Production Quality Timeline")
        
        severe_grades = ['B+', 'B']
        df_filtered['Severe_Bad_Qty'] = df_filtered[severe_grades].sum(axis=1)
        df_filtered['Acceptable_Qty'] = df_filtered['Total_Qty'] - df_filtered['Severe_Bad_Qty']
        
        period_summary = df_filtered.groupby(['Time_Group', 'Actual_Thickness', 'HR_Material'])[['Total_Qty', 'Acceptable_Qty', 'Severe_Bad_Qty']].sum().reset_index()
        period_summary = period_summary[period_summary['Total_Qty'] > 0]
        
        period_summary['Yield (%)'] = (period_summary['Acceptable_Qty'] / period_summary['Total_Qty'] * 100).fillna(0).round(2)
        period_summary['Defect_Rate (%)'] = (period_summary['Severe_Bad_Qty'] / period_summary['Total_Qty'] * 100).fillna(0).round(2)
        
        time_order_map = {
            "2024 (Full Year)": 1,
            "2025 H1 (Until 06/28)": 2,
            "2025 Q3 (06/29 - 09/30)": 3,
            "2025 Q4": 4,
            "2025 (Full Year)": 5,
            "Unknown": 99
        }
        period_summary['Sort_Key'] = period_summary['Time_Group'].map(time_order_map).fillna(90)
        period_summary = period_summary.sort_values(by=['Sort_Key', 'Actual_Thickness'], ascending=[True, True])
        period_summary = period_summary.drop(columns=['Sort_Key'])
        period_summary.rename(columns={'Severe_Bad_Qty': 'Bad_Qty (B+, B)'}, inplace=True)

        if not period_summary.empty:
            st.dataframe(
                period_summary.style.background_gradient(subset=['Defect_Rate (%)'], cmap='Reds')
                                    .background_gradient(subset=['Yield (%)'], cmap='Greens')
                                    .format({
                                        'Actual_Thickness': '{:.2f}', 'Total_Qty': '{:.0f}', 
                                        'Acceptable_Qty': '{:.0f}', 'Bad_Qty (B+, B)': '{:.0f}', 
                                        'Yield (%)': '{:.2f}%', 'Defect_Rate (%)': '{:.2f}%'
                                    }),
                use_container_width=True, hide_index=True
            )

        st.markdown("---")
        st.subheader("📑 Detailed Yield by Period (All Grades)")
        sum_df = df_filtered.groupby(['Time_Group', 'Actual_Thickness', 'HR_Material'], dropna=False)[base_grades].sum().reset_index()
        sum_df['Total_Qty'] = sum_df[base_grades].sum(axis=1)
        for col in base_grades: sum_df[f"% {col}"] = ((sum_df[col] / sum_df['Total_Qty'].replace(0, np.nan)) * 100).fillna(0).round(1)
        
        sum_df['Sort_Key'] = sum_df['Time_Group'].map(time_order_map).fillna(90)
        sum_df = sum_df.sort_values(by=['Sort_Key', 'Actual_Thickness']).drop(columns=['Sort_Key'])
        sum_df.rename(columns={'Time_Group': 'Period', 'Actual_Thickness': 'Thickness'}, inplace=True)
        ordered_periods = sorted(sum_df['Period'].unique(), key=lambda x: time_order_map.get(x, 99))
        
        for period in ordered_periods:
            p_data = sum_df[sum_df['Period'] == period]
            if not p_data.empty:
                st.markdown(f"#### 📅 Period: **{period}**")
                format_dict = {'Thickness': '{:.2f}', 'Total_Qty': '{:.0f}'}
                for col in base_grades: format_dict[col] = '{:.0f}'; format_dict[f"% {col}"] = '{:.1f}%'
                st.dataframe(p_data.drop(columns=['Period']).style.format(format_dict), use_container_width=True, hide_index=True)

        st.markdown("---")
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                period_summary.to_excel(writer, index=False, sheet_name='Executive_Summary')
                sum_df.to_excel(writer, index=False, sheet_name='Detailed_Yield')
                workbook  = writer.book
                worksheet = writer.sheets['Executive_Summary']
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                num_fmt = workbook.add_format({'align': 'center', 'border': 1})
                pct_fmt = workbook.add_format({'num_format': '0.00"%"', 'align': 'center', 'border': 1})
                for col_num, value in enumerate(period_summary.columns.values):
                    worksheet.write(0, col_num, value, header_fmt)
                    worksheet.set_column(col_num, col_num, 15)
                worksheet.conditional_format(1, 6, len(period_summary), 6, {'type': '2_color_scale', 'min_color': "#F7FCF5", 'max_color': "#41AB5D"})
                worksheet.conditional_format(1, 7, len(period_summary), 7, {'type': '2_color_scale', 'min_color': "#FFF5F0", 'max_color': "#EF3B2C"})
                for row in range(1, len(period_summary) + 1):
                    for col in range(len(period_summary.columns)):
                        val = period_summary.iloc[row-1, col]
                        if col >= 6 and isinstance(val, (int, float)): worksheet.write(row, col, val/100, pct_fmt)
                        else: worksheet.write(row, col, val, num_fmt)
            st.download_button(label="📥 Download Formatted Excel", data=output.getvalue(), file_name="Colored_Yield.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except:
            st.download_button(label="📥 Download Basic CSV", data=period_summary.to_csv(index=False).encode('utf-8'), file_name="Yield_Summary.csv", mime="text/csv")

    # --- TAB 2: DISTRIBUTION ---
    with tab2:
        global_x_bounds = {}
        for feat in mech_features:
            if feat in df.columns:
                vd = df[[feat, 'Total_Qty']].dropna().copy()
                vd = vd[vd['Total_Qty'] > 0]
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
            fmin, fmax = global_x_bounds.get(feat, (data[feat].min() if not data.empty else 0, data[feat].max() if not data.empty else 100))
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
                m_info.sort(key=lambda x: x['v'])
                for i, info in enumerate(m_info):
                    h = y_lim * (0.85 - (i % 3) * 0.12)
                    ax.text(info['v'], h, f"{info['v']:.1f}", color='white', fontweight='bold', fontsize=8, ha='center', bbox=dict(facecolor=info['c'], alpha=0.8, boxstyle='round,pad=0.2'))
            ax.legend(handles=[Patch(facecolor=c_map[g], label=g) for g in base_grades if g in data.columns], loc='upper right', fontsize=7)
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
            for thick in thickness_list:
                df_t = df_p[df_p['Actual_Thickness'] == thick]
                if df_t.empty: continue
                st.markdown(f"**📏 Thickness: {thick}**")
                ly = get_shared_y(df_t, ['YS', 'TS', 'EL', 'YPE'])
                tcols = st.columns(2)
                for idx, f in enumerate([x for x in ['YS', 'TS', 'EL', 'YPE'] if x in df_t.columns]):
                    with tcols[idx%2]:
                        fig, ax = plt.subplots(figsize=(8, 4.5))
                        plot_dist(ax, df_t, f, f"{f} (Thick: {thick})", ly)
                        st.pyplot(fig); plt.close(fig)

    # --- TAB 3: EXECUTIVE AUTO-INSIGHT & ROOT CAUSE ---
    with tab3:
        st.header("🧠 Executive Auto-Insight & Root Cause")
        st.info("Automated diagnostic engine: Quantifying impact based on severe defects (B+, B).")

        severe_grades = ['B+', 'B']
        df_filtered['Severe_Bad_Qty'] = df_filtered[severe_grades].sum(axis=1)
        df_filtered['Spec_Label'] = df_filtered['Actual_Thickness'].astype(str) + "mm (" + df_filtered['HR_Material'] + ")"

        heat_data = df_filtered.groupby(['Spec_Label', 'Time_Group']).apply(
            lambda x: (x['Severe_Bad_Qty'].sum() / x['Total_Qty'].sum() * 100) if x['Total_Qty'].sum() > 0 else 0
        )
        
        if not heat_data.empty and heat_data.max() > 0:
            heatmap_long = heat_data.reset_index()
            heatmap_long.columns = ['Spec', 'Period', 'Defect_Rate']
            top_issues = heatmap_long[heatmap_long['Defect_Rate'] > 0].sort_values('Defect_Rate', ascending=False).head(5)

            top_3_specs = top_issues.head(3)['Spec'].tolist()
            top_3_periods = top_issues.head(3)['Period'].tolist()
            
            top_3_subsets = []
            for _, row in top_issues.head(3).iterrows():
                subset = df_filtered[(df_filtered['Spec_Label'] == row['Spec']) & (df_filtered['Time_Group'] == row['Period'])]
                top_3_subsets.append(subset)
            
            if top_3_subsets:
                df_top3 = pd.concat(top_3_subsets).drop_duplicates(subset=['Row_ID'])
            else:
                df_top3 = pd.DataFrame(columns=df_filtered.columns)

            df_unique_global = df_filtered.drop_duplicates(subset=['Row_ID'])

            rc_results = {}
            for f in ['YS', 'TS', 'EL', 'YPE']:
                if f in df_top3.columns:
                    good_mean_global = df_unique_global[df_unique_global['Good_Qty'] > 0][f].mean()
                    bad_mean_top3 = df_top3[df_top3['Severe_Bad_Qty'] > 0][f].mean()
                    if pd.notnull(good_mean_global) and pd.notnull(bad_mean_top3):
                        rc_results[f] = bad_mean_top3 - good_mean_global

            rc_s = pd.Series(rc_results).dropna().sort_values(key=abs, ascending=False)

            if not top_issues.empty and not rc_s.empty:
                top_issue = top_issues.iloc[0]
                top_driver = rc_s.index[0]
                gap_val = rc_s.iloc[0]
                direction = "HIGHER ⬆️" if gap_val > 0 else "LOWER ⬇️"

                st.success(f"""
                ### 🎯 EXECUTIVE CONCLUSION & ACTION PLAN:
                * 🚨 **Biggest Hotspot:** Specification **{top_issue['Spec']}** during **{top_issue['Period']}** (Severe Defect Rate hits **{top_issue['Defect_Rate']:.1f}%**).
                * 🧠 **Main Root Cause Driver:** **{top_driver}** is the primary culprit.
                * 📊 **Quantified Impact:** Defective coils have a {top_driver} that is on average **{abs(gap_val):.1f} {direction}** than good coils.
                """)

            st.markdown("---")
            col1, col2 = st.columns(2)

            with col1:
                st.error("🔥 TOP 5 PROBLEM SEGMENTS (Where to fix)")
                st.dataframe(
                    top_issues.style.background_gradient(subset=['Defect_Rate'], cmap='Reds').format({'Defect_Rate': '{:.1f}%'}), 
                    use_container_width=True, hide_index=True
                )

            with col2:
                st.warning("🧠 ROOT CAUSE DRIVER (Top 3 Hotspots)")
                st.info("Analysis: Mean difference of Bad Coils (in Top 3 problem areas) vs. Global Good Coils (Deduplicated).")
                rc_df = rc_s.reset_index()
                rc_df.columns = ['Mechanical Feature', 'Impact Gap (Top 3 Bad vs Good)']
                def color_gap(val):
                    color = '#d62728' if abs(val) > 5 else ('#ff7f0e' if abs(val) > 0 else 'black')
                    return f'color: {color}; font-weight: bold'
                st.dataframe(rc_df.style.map(color_gap, subset=['Impact Gap (Top 3 Bad vs Good)']).format({'Impact Gap (Top 3 Bad vs Good)': '{:+.2f}'}), use_container_width=True, hide_index=True)

            if not rc_s.empty:
                st.markdown("---")
                top_driver = rc_s.index[0]
                st.info(f"📏 DRILL DOWN: {top_driver} Shift by Thickness (Isolating the issue)")
                drill_data = []
                for th in df_unique_global['Actual_Thickness'].dropna().unique():
                    th_df = df_unique_global[df_unique_global['Actual_Thickness'] == th]
                    g_val = th_df[th_df['Good_Qty'] > 0][top_driver].mean()
                    b_val = th_df[th_df['Severe_Bad_Qty'] > 0][top_driver].mean()
                    if pd.notnull(g_val) or pd.notnull(b_val):
                        drill_data.append({
                            'Thickness': th, 'GOOD Coils (Mean)': g_val, 'BAD Coils (Mean)': b_val,
                            'Impact Gap': (b_val - g_val) if (pd.notnull(g_val) and pd.notnull(b_val)) else None
                        })
                drill_df = pd.DataFrame(drill_data).sort_values('Impact Gap', key=abs, ascending=False)
                st.dataframe(drill_df.style.map(color_gap, subset=['Impact Gap']).format({'Thickness': '{:.2f}mm', 'GOOD Coils (Mean)': '{:.1f}', 'BAD Coils (Mean)': '{:.1f}', 'Impact Gap': '{:+.1f}'}), use_container_width=True, hide_index=True)
                
            st.markdown("---")
            st.subheader("🗺️ Evidence: Visual Hotspot Map (Grades B+ and Below)")
            heat_pivot = heat_data.unstack()
            fig, ax = plt.subplots(figsize=(12, 5))
            vmax_threshold = 30.0 if heat_pivot.max().max() > 30 else heat_pivot.max().max()
            sns.heatmap(heat_pivot, annot=True, fmt=".1f", cmap="YlOrRd", linewidths=.5, vmax=vmax_threshold, ax=ax)
            ax.set_title("SEVERE DEFECT RATE (%)", fontweight='bold', color='#d62728')
            ax.set_ylabel(""); ax.set_xlabel("")
            fig.tight_layout()
            plt.savefig("export_heatmap.png", bbox_inches='tight', dpi=150)
            st.pyplot(fig); plt.close(fig)

        else:
            st.success("✅ Process is completely stable. No severe defect patterns detected to analyze.")

        st.markdown("---")
        st.header("📈 Global I-MR Stability Tracking (Severe Defects: B+ and Below)")
        df_severe_global = df_unique_global[(df_unique_global['B+'] > 0) | (df_unique_global['B'] > 0)].sort_values(by='烤三生產日期').reset_index(drop=True)

        if not df_severe_global.empty:
            for feat in ['YS', 'TS', 'EL', 'YPE']:
                if feat in df_severe_global.columns:
                    valid_data = df_severe_global.dropna(subset=[feat, '烤三生產日期']).reset_index(drop=True)
                    if len(valid_data) > 1:
                        st.markdown(f"#### 🛡️ Global Stability: **{feat}**")
                        dates = valid_data['烤三生產日期']
                        vals = valid_data[feat].values
                        x_seq = np.arange(len(vals)) 
                        mean_v = np.mean(vals)
                        
                        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 8), gridspec_kw={'height_ratios': [2, 1]})
                        
                        # --- I-Chart Updated Colors ---
                        ax1.plot(x_seq, vals, marker='o', ms=5, lw=1.5, color='#004C99', alpha=0.9, label=f"Value ({feat})")
                        ax1.axhline(mean_v, color='black', ls='--', lw=1.5, label=f'Mean: {mean_v:.1f}')
                        ax1.text(x_seq[-1], mean_v, f" Mean: {mean_v:.1f}", va='bottom', color='black', fontweight='bold')
                        
                        v_x, v_y = [], []
                        if feat in GLOBAL_SPECS:
                            s = GLOBAL_SPECS[feat]
                            if s['min']: ax1.axhline(s['min'], color='red', lw=2); ax1.text(x_seq[-1], s['min'], f" Min: {s['min']}", va='bottom', color='red', fontweight='bold')
                            if s['max']: ax1.axhline(s['max'], color='red', lw=2); ax1.text(x_seq[-1], s['max'], f" Max: {s['max']}", va='bottom', color='red', fontweight='bold')
                            for i, v in enumerate(vals):
                                if (s['min'] and v < s['min']) or (s['max'] and v > s['max']):
                                    v_x.append(i); v_y.append(v)
                            if v_x: ax1.scatter(v_x, v_y, color='red', s=60, zorder=5)
                        
                        for i in range(1, len(dates)):
                            if dates.iloc[i].year != dates.iloc[i-1].year:
                                ax1.axvline(i, color='gray', ls=':', alpha=0.5)
                                ax1.text(i, ax1.get_ylim()[1], f" {dates.iloc[i].year}", fontsize=10, va='top')

                        ax1.set_title(f"Individual Chart (I) - {feat}", fontweight='bold')
                        ax1.legend(loc='upper right', fontsize=9, bbox_to_anchor=(1.15, 1))
                        ax1.set_xticks([]) 

                        # --- MR-Chart Updated Colors ---
                        mr = np.abs(np.diff(vals))
                        mr_mean = np.mean(mr)
                        ucl_mr = 3.267 * mr_mean
                        
                        ax2.plot(x_seq[1:], mr, marker='o', ms=5, lw=1.5, color='#4B0082', alpha=0.9, label="Moving Range")
                        ax2.axhline(mr_mean, color='black', ls='--', lw=1.5, label=f'MR Mean: {mr_mean:.1f}')
                        ax2.axhline(ucl_mr, color='red', ls=':', lw=1.5, label=f'UCL: {ucl_mr:.1f}')
                        
                        hv_x, hv_y = [], []
                        for i, m_val in enumerate(mr):
                            if m_val > ucl_mr: hv_x.append(i+1); hv_y.append(m_val)
                        if hv_x: ax2.scatter(hv_x, hv_y, color='red', s=40, zorder=5)

                        ax2.set_title("Moving Range Chart (MR)", fontweight='bold')
                        ax2.legend(loc='upper right', fontsize=9, bbox_to_anchor=(1.15, 1))
                        
                        step = max(1, len(x_seq) // 12) 
                        ax2.set_xticks(x_seq[::step])
                        ax2.set_xticklabels(dates.dt.strftime('%Y-%m-%d').iloc[::step], rotation=45, ha='right')

                        fig.tight_layout()
                        plt.savefig(f"export_imr_global_{feat}.png", bbox_inches='tight', dpi=150)
                        st.pyplot(fig); plt.close(fig)

    # --- TAB 4: I-MR CHART (TIMELINE STABILITY) ---
    with tab4:
        st.header("📈 Task 4: I-MR Stability Tracking (Chronological)")
        st.info("Analysis based on production sequence from 2024 to 2026. Red dots = Out of Spec.")

        @st.fragment
        def render_tab4():
            imr_periods = ["All Periods"] + sorted(df_filtered['Time_Group'].dropna().unique().tolist())
            imr_thicks = sorted(df_filtered['Actual_Thickness'].dropna().unique())
            imr_mats = sorted(df_filtered['HR_Material'].astype(str).unique())
            
            c1, c2, c3 = st.columns(3)
            sel_p = c1.selectbox("Filter Period:", imr_periods, key="t4_p")
            sel_t = c2.selectbox("Filter Thickness:", imr_thicks, key="t4_t")
            sel_m = c3.selectbox("Filter Material:", imr_mats, key="t4_m")

            if sel_p == "All Periods":
                imr_df = df_filtered[(df_filtered['Actual_Thickness'] == sel_t) & (df_filtered['HR_Material'] == sel_m)]
                imr_df = imr_df.drop_duplicates(subset=['Row_ID'])
            else:
                imr_df = df_filtered[(df_filtered['Time_Group'] == sel_p) & (df_filtered['Actual_Thickness'] == sel_t) & (df_filtered['HR_Material'] == sel_m)]

            imr_df = imr_df.sort_values(by='烤三生產日期').reset_index(drop=True)

            if not imr_df.empty:
                for feat in ['YS', 'TS', 'EL', 'YPE']:
                    if feat in imr_df.columns:
                        valid_data = imr_df.dropna(subset=[feat, '烤三生產日期']).copy()
                        if len(valid_data) > 1:
                            st.markdown(f"### 🛡️ Stability: **{feat}**")
                            valid_data = valid_data.reset_index(drop=True)
                            dates = valid_data['烤三生產日期']
                            vals = valid_data[feat].values
                            
                            x_seq = np.arange(len(vals)) 
                            mean_v = np.mean(vals)
                            
                            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 7), gridspec_kw={'height_ratios': [2, 1]})
                            
                            # --- I-Chart Updated Colors ---
                            ax1.plot(x_seq, vals, marker='o', ms=5, lw=1.5, color='#004C99', alpha=0.9, label=feat)
                            ax1.axhline(mean_v, color='black', ls='--', lw=1.5, label='Mean')
                            
                            if feat in GLOBAL_SPECS:
                                s = GLOBAL_SPECS[feat]
                                if s['min']: ax1.axhline(s['min'], color='red', lw=2)
                                if s['max']: ax1.axhline(s['max'], color='red', lw=2)
                                
                                v_x, v_y = [], []
                                for i, v in enumerate(vals):
                                    if (s['min'] and v < s['min']) or (s['max'] and v > s['max']):
                                        v_x.append(i); v_y.append(v)
                                if v_x: ax1.scatter(v_x, v_y, color='red', s=60, zorder=5)
                            
                            if sel_p == "All Periods":
                                for i in range(1, len(dates)):
                                    if dates.iloc[i].year != dates.iloc[i-1].year:
                                        ax1.axvline(i, color='gray', ls=':', alpha=0.5)
                                        ax1.text(i, ax1.get_ylim()[1], f" {dates.iloc[i].year}", fontsize=10, va='top')

                            ax1.set_title(f"Individual Chart (I) - {feat}", fontweight='bold')
                            ax1.legend(loc='upper right', fontsize=8)
                            ax1.set_xticks([]) 

                            # --- MR-Chart Updated Colors ---
                            mr = np.abs(np.diff(vals))
                            mr_mean = np.mean(mr)
                            ucl_mr = 3.267 * mr_mean
                            
                            ax2.plot(x_seq[1:], mr, marker='o', ms=5, lw=1.5, color='#4B0082', alpha=0.9)
                            ax2.axhline(mr_mean, color='black', ls='--', lw=1.5)
                            ax2.axhline(ucl_mr, color='red', ls=':', lw=1.5)
                            
                            hv_x, hv_y = [], []
                            for i, m_val in enumerate(mr):
                                if m_val > ucl_mr: hv_x.append(i+1); hv_y.append(m_val)
                            if hv_x: ax2.scatter(hv_x, hv_y, color='red', s=40, zorder=5)

                            ax2.set_title("Moving Range Chart (MR)", fontweight='bold')
                            
                            step = max(1, len(x_seq) // 12) 
                            ax2.set_xticks(x_seq[::step])
                            ax2.set_xticklabels(dates.dt.strftime('%Y-%m-%d').iloc[::step], rotation=45, ha='right')

                            fig.tight_layout()
                            safe_p_imr = "".join([c if c.isalnum() else "_" for c in sel_p])
                            plt.savefig(f"export_imr_{feat}.png", bbox_inches='tight', dpi=150)
                            st.pyplot(fig); plt.close(fig)
            else:
                st.warning("No data found for the selected combination.")

        render_tab4()                

# --- EXPORT SECTION ---
st.sidebar.header("📥 Export PDF Report")
st.sidebar.info("💡 Tip: Make sure to click through Tab 3 and Tab 4 so the app generates all the charts before exporting the PDF.")

if st.sidebar.button("🖨️ Generate PDF Report"):
    try:
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # --- PART 1: HEATMAP DIAGNOSTIC ---
        if os.path.exists("export_heatmap.png"):
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, "1. DEFECT HOTSPOT DIAGNOSTIC MAP (SEVERE DEFECTS)", ln=True, align='C')
            pdf.image("export_heatmap.png", x=15, y=25, w=260)
            
        # --- PART 2: GLOBAL I-MR (FROM TAB 3) ---
        global_imr_files = [f"export_imr_global_{feat}.png" for feat in ['YS', 'TS', 'EL', 'YPE'] if os.path.exists(f"export_imr_global_{feat}.png")]
        
        if global_imr_files:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, "2. GLOBAL PROCESS STABILITY (2024-2025 ALL SEVERE DEFECTS)", ln=True, align='C')
            
            y_pos = 25
            chart_count = 0
            for img_path in global_imr_files:
                if chart_count == 2: 
                    pdf.add_page()
                    y_pos = 20
                    chart_count = 0
                pdf.image(img_path, x=20, y=y_pos, w=250)
                y_pos += 90 
                chart_count += 1

        # --- PART 3: FILTERED I-MR (FROM TAB 4) ---
        filtered_imr_files = [f"export_imr_{feat}.png" for feat in ['YS', 'TS', 'EL', 'YPE'] if os.path.exists(f"export_imr_{feat}.png")]
        
        if filtered_imr_files:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, "3. SPECIFIC I-MR TRACKING (FILTERED SEGMENT)", ln=True, align='C')
            
            y_pos = 25
            chart_count = 0
            for img_path in filtered_imr_files:
                if chart_count == 2:
                    pdf.add_page()
                    y_pos = 20
                    chart_count = 0
                pdf.image(img_path, x=20, y=y_pos, w=250)
                y_pos += 90 
                chart_count += 1

        pdf.output("Quality_Visual_Report.pdf")
        
        with open("Quality_Visual_Report.pdf", "rb") as f:
            st.sidebar.download_button(
                label="✅ Click to Download your PDF Report", 
                data=f.read(), 
                file_name="Quality_Visual_Report.pdf", 
                mime="application/pdf"
            )
        st.sidebar.success("🎉 PDF Generated Successfully! File is ready to download.")
        
    except Exception as e:
        st.sidebar.error(f"⚠️ Error generating PDF: {e}")
