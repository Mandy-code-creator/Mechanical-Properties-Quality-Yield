import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import scipy.stats as stats
import io
from fpdf import FPDF
import os
from matplotlib.patches import Patch

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="QC Yield & Control Limit Optimizer", layout="wide")

st.title("📊 Quality Yield & Control Limit Optimizer")
st.markdown("---")

# --- 1. FILE UPLOAD ---
uploaded_file = st.file_uploader("Upload Excel data (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip() 

    # --- COLUMN TRANSLATION ---
    rename_map = {
        '烤漆降伏強度': 'YS', '烤漆抗拉強度': 'TS', '伸長率': 'EL',
        'A-B+': 'A-B+數', 'A-B': 'A-B數', 'A-B-': 'A-B-數', 'B+': 'B+數', 'B': 'B數'
    }
    df.rename(columns=rename_map, inplace=True)
    
    # --- 2. DATA PREPROCESSING & TYPE CASTING ---
    if '年度' in df.columns:
        df['年度'] = df['年度'].astype(str).str.replace(r'\.0$', '', regex=True)
        
    # FIX DATE PARSING (YYYYMMDD)
    if '烤三生產日期' in df.columns:
        date_str = df['烤三生產日期'].astype(str).str.replace(r'\.0$', '', regex=True)
        df['烤三生產日期'] = pd.to_datetime(date_str, format='%Y%m%d', errors='coerce').fillna(pd.to_datetime(date_str, errors='coerce'))

    # STRICTLY USE '厚度' COLUMN FOR THICKNESS AND AVOID FLOATING POINT ERRORS
    if '厚度' in df.columns:
        df['厚度'] = pd.to_numeric(df['厚度'], errors='coerce').round(3)
        
    if '熱軋材質' in df.columns:
        df['熱軋材質'] = df['熱軋材質'].astype(str).str.strip()

    # --- PERIOD CATEGORIZATION LOGIC ---
    def categorize_period(date_val):
        if pd.isnull(date_val): return "Unknown"
        
        y = date_val.year
        q3_start = pd.Timestamp(2025, 6, 29)
        q3_end = pd.Timestamp(2025, 9, 30)

        if y == 2024:
            return "2024 (Full Year)"
        elif y == 2025:
            if date_val < q3_start:
                return "2025 H1 (Until 06/28)"
            elif q3_start <= date_val <= q3_end:
                return "2025 Q3 (06/29 - 09/30)"
            else:
                return "2025 Q4"
        elif y == 2026:
            return "2026 Q1"
        return "Other"

    if '烤三生產日期' in df.columns:
        df['Time_Group'] = df['烤三生產日期'].apply(categorize_period)
        df = df[~df['Time_Group'].isin(["Other", "Unknown"])]
    else:
        df['Time_Group'] = "Unknown"

    # --- SIDEBAR FILTERS ---
    st.sidebar.header("🔎 Dashboard Filters")
    all_periods = sorted(df['Time_Group'].unique())
    options_list = ["All"] + all_periods
    
    ui_selection = st.sidebar.multiselect("📅 Select Period(s):", options=options_list, default=["All"])
    
    if "All" in ui_selection or len(ui_selection) == 0:
        selected_periods = all_periods
    else:
        selected_periods = ui_selection
        df = df[df['Time_Group'].isin(selected_periods)]

    # --- GET UNIQUE THICKNESS VALUES FROM '厚度' ---
    if '厚度' in df.columns:
        thickness_list = sorted(df['厚度'].dropna().unique(), key=lambda x: float(x))
    else:
        thickness_list = []

    count_cols = [col for col in ['A-B+數', 'A-B數', 'A-B-數', 'B+數', 'B數'] if col in df.columns]
    for col in count_cols: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    df['Total_Count'] = df[count_cols].sum(axis=1)

    mech_features = [feat for feat in ['YS', 'TS', 'EL', 'YPE', 'HARDNESS'] if feat in df.columns]
    for feat in mech_features: df[feat] = pd.to_numeric(df[feat], errors='coerce')

    # --- GLOBAL X-AXIS FIXATION ---
    global_x_bounds = {}
    for feat in mech_features:
        vd = df[[feat] + count_cols].dropna(subset=[feat]).copy()
        vd['T_Count'] = vd[count_cols].sum(axis=1)
        vd = vd[vd['T_Count'] > 0]
        if not vd.empty:
            q1, q99 = np.percentile(vd[feat], 1), np.percentile(vd[feat], 99)
            iqr = q99 - q1
            fmin, fmax = max(vd[feat].min(), q1 - 0.5*iqr), min(vd[feat].max(), q99 + 0.5*iqr)
            if fmin >= fmax: fmin -= 5; fmax += 5
            global_x_bounds[feat] = (fmin - (fmax-fmin)*0.05, fmax + (fmax-fmin)*0.05)

    def get_shared_y(data, features):
        max_y = 0
        for feat in features:
            if feat in data.columns:
                vd = data.dropna(subset=[feat]).copy()
                vd['T_Count'] = vd[count_cols].sum(axis=1)
                if not vd.empty:
                    fmin, fmax = global_x_bounds.get(feat, (0, 100))
                    cnts, _ = np.histogram(vd[feat], bins=np.linspace(fmin, fmax, 16), weights=vd['T_Count'])
                    max_y = max(max_y, cnts.max())
        return max_y * 1.35 if max_y > 0 else 50 

    overall_export_data = [] 
    all_export_data = []

    tab1, tab2, tab3 = st.tabs(["1. Yield Summary", "2. Distribution Analysis", "3. Control Limits & I-MR"])

    # --- TAB 1: YIELD SUMMARY ---
    with tab1:
        st.header("1. Quality Yield Summary (Thickness ➔ Material)")
        group_cols = ['Time_Group', '厚度', '熱軋材質']
        existing_group_cols = [col for col in group_cols if col in df.columns]
        
        if existing_group_cols:
            summary_df = df.groupby(existing_group_cols)[count_cols].sum().reset_index()
            summary_df['Total_Qty'] = summary_df[count_cols].sum(axis=1)
            summary_df = summary_df[summary_df['Total_Qty'] > 0]
            
            for col in count_cols: summary_df[f"% {col.replace('數','')}"] = (summary_df[col]/summary_df['Total_Qty']*100).fillna(0).round(1)
            
            display_df = summary_df.rename(columns={'Time_Group': 'Period', '厚度': 'Thickness', '熱軋材質': 'HR Material'})
            for col in count_cols: display_df.rename(columns={col: col.replace('數','')}, inplace=True)
            
            base_cols = [c for c in ['Period', 'Thickness', 'HR Material', 'Total_Qty'] if c in display_df.columns]
            grade_cols = [c.replace('數','') for c in count_cols]
            pct_cols = [f"% {c}" for c in grade_cols]
            
            for c in grade_cols + ['Total_Qty']:
                if c in display_df.columns: display_df[c] = display_df[c].astype(int)
            
            display_df = display_df[base_cols + grade_cols + pct_cols]
            if 'Period' in display_df.columns: display_df = display_df.sort_values(by=['Period', 'Thickness'])
            
            for period in selected_periods:
                period_data = display_df[display_df['Period'] == period]
                if not period_data.empty:
                    st.markdown(f"### 📅 Period: **{period}**")
                    st.dataframe(period_data.drop(columns=['Period'], errors='ignore'), use_container_width=True, hide_index=True)
            
            st.markdown("---")
            towrite_summary = io.BytesIO()
            display_df.to_excel(towrite_summary, index=False, engine='openpyxl')
            towrite_summary.seek(0)
            st.download_button("📥 Download Pivot Summary (Excel)", data=towrite_summary, file_name="Quality_Yield_Summary.xlsx")

    # --- SHARED STATS FUNCTIONS ---
    def calculate_stats(v_arr, w_arr, k_factor):
        m_std = np.average(v_arr, weights=w_arr)
        s_std = np.sqrt(np.average((v_arr - m_std)**2, weights=w_arr))
        try:
            expanded_v = np.repeat(v_arr, w_arr.astype(int))
            q1, q3 = np.percentile(expanded_v, 25), np.percentile(expanded_v, 75)
            iqr = q3 - q1
            lower_iqr, upper_iqr = q1 - k_factor * iqr, q3 + k_factor * iqr
            mask = (v_arr >= lower_iqr) & (v_arr <= upper_iqr)
            vf, wf = v_arr[mask], w_arr[mask]
            if len(vf) > 0 and sum(wf) > 0:
                m_iqr = np.average(vf, weights=wf)
                s_iqr = np.sqrt(np.average((vf - m_iqr)**2, weights=wf))
            else: m_iqr, s_iqr = m_std, s_std
        except: m_iqr, s_iqr = m_std, s_std
        return (m_std, s_std), (m_iqr, s_iqr)

    def plot_qc_dist(ax, data, feat, title, custom_y_limit, is_right=False):
        k_b = 15
        color_map = {'A-B+數': '#2ca02c', 'A-B數': '#1f77b4', 'A-B-數': '#ff7f0e', 'B+數': '#9467bd', 'B數': '#d62728'}
        mean_inf = []
        ax.grid(axis='y', linestyle=':', alpha=0.6, zorder=0)
        f_min, f_max = global_x_bounds.get(feat, (0, 100))
        bins_arr = np.linspace(f_min, f_max, k_b + 1)
        vals_list, wgts_list, colors_list = [], [], []
        
        for col_n in count_cols:
            temp_d = data[[feat, col_n]].dropna()
            temp_d = temp_d[temp_d[col_n] > 0]
            if not temp_d.empty:
                vals, wgts = temp_d[feat].values, temp_d[col_n].values
                color = color_map.get(col_n, '#7f7f7f')
                vals_list.append(vals); wgts_list.append(wgts); colors_list.append(color)
                
                m = np.average(vals, weights=wgts)
                s = np.sqrt(np.average((vals - m)**2, weights=wgts))
                ax.axvline(m, color=color, ls='--', lw=1.5, zorder=3)
                mean_inf.append({'val': m, 'color': color})

                if len(vals) > 2 and s > 0:
                    x_r = np.linspace(f_min, f_max, 100)
                    ax.plot(x_r, stats.norm.pdf(x_r, m, s) * wgts.sum() * ((f_max - f_min) / k_b), color=color, lw=2, zorder=4)

        if vals_list:
            ax.hist(vals_list, bins=bins_arr, weights=wgts_list, color=colors_list, stacked=True, edgecolor='white', alpha=0.8, zorder=2)

        ax.set_xlim(f_min, f_max)
        ax.set_ylim(0, custom_y_limit)

        if mean_inf:
            mean_inf.sort(key=lambda x: x['val'])
            levels = [0.90, 0.82, 0.74, 0.66, 0.58]
            for i, info in enumerate(mean_inf):
                y_pos = custom_y_limit * levels[i % len(levels)]
                ax.text(info['val'], y_pos, f"{int(round(info['val']))}", color=info['color'], fontsize=9, fontweight='bold', ha='center', va='center', zorder=5, bbox=dict(facecolor='white', alpha=0.9, edgecolor=info['color'], boxstyle='round,pad=0.3'))

        ax.set_title(title, fontsize=11, fontweight='bold', pad=10)
        ax.set_ylabel("Count", fontsize=10)
        
        if is_right:
            legend_elements = [Patch(facecolor=color_map[k], edgecolor='white', label=k.replace('數',''), alpha=0.5) for k in color_map]
            ax.legend(handles=legend_elements, title="Grade", title_fontsize='9', fontsize='8', bbox_to_anchor=(1.02, 1), loc='upper left')

    # --- TAB 3 SETTINGS ---
    with tab3:
        st.markdown("##### ⚙️ Production Configuration")
        c1, c2, c3, c4 = st.columns(4)
        sigma_release = c1.number_input("Release Range (Sigma)", min_value=1.0, max_value=6.0, value=2.0, step=0.1)
        sigma_mill = c2.number_input("Mill Range (Sigma)", min_value=0.5, max_value=4.0, value=1.0, step=0.1)
        iqr_k = c3.number_input("IQR Filter Factor (k)", min_value=1.0, max_value=4.0, value=1.5, step=0.1)
        chart_method = c4.radio("I-MR Chart Limits Based On:", ["Standard Method", "IQR Filtered Method"])

    spec_limits = {"YS": (405, 500), "TS": (415, 550), "EL": (25, None), "YPE": (4, None)}
    good_cols = [c for c in ['A-B+數', 'A-B數'] if c in df.columns]

    # --- ITERATE THROUGH PERIODS ---
    for period in selected_periods:
        df_p = df[df['Time_Group'] == period]
        if df_p.empty: continue
        safe_p = "".join([c if c.isalnum() else "_" for c in period])

        with tab2:
            st.markdown(f"## 📅 Time Period: **{period}**")
            st.subheader(f"🌐 Overall Factory Distribution ({period})")
            ov_y = get_shared_y(df_p, ['YS', 'TS', 'EL', 'YPE'])
            cols = st.columns(2)
            for i, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
                if f in mech_features:
                    with cols[i%2]:
                        fig, ax = plt.subplots(figsize=(10, 5))
                        plot_qc_dist(ax, df_p, f, f"{f} (Overall - {period})", ov_y, is_right=(i%2!=0))
                        st.pyplot(fig)
                        fig.savefig(f"overall_{f}_{safe_p}.png", bbox_inches='tight')
            
            st.subheader(f"🔍 Distribution by Thickness ({period})")
            for thick in thickness_list:
                df_thick = df_p[df_p['厚度'] == thick]
                if df_thick.empty: continue
                st.markdown(f"**📏 Thick: {thick}**")
                local_y = get_shared_y(df_thick, ['YS', 'TS', 'EL', 'YPE']) 
                cols_dist = st.columns(2)
                for i, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
                    if f in mech_features:
                        with cols_dist[i%2]:
                            fig, ax = plt.subplots(figsize=(10, 5))
                            plot_qc_dist(ax, df_thick, f, f"{f} (Thick: {thick})", local_y, is_right=(i%2!=0))
                            st.pyplot(fig)
                            fig.savefig(f"dist_{f}_{thick}_{safe_p}.png", bbox_inches='tight')
            st.markdown("---")

        with tab3:
            st.markdown(f"## 📅 Time Period: **{period}**")
            st.subheader(f"🌐 Overall Factory Goals ({period})")
            
            period_overall_data = []
            total_n_overall = df_p[count_cols].sum().sum()
            seg_dist_overall = "N/A" if total_n_overall == 0 else ", ".join([f"{k.replace('數','')}:{int(round(df_p[k].sum()/total_n_overall*100))}%" for k in count_cols])

            for f in mech_features:
                if good_cols:
                    df_ov = df_p[[f] + good_cols].dropna(subset=[f]).copy()
                    df_ov['Good_Qty'] = df_ov[good_cols].sum(axis=1)
                    df_ov = df_ov[df_ov['Good_Qty'] > 0]
                    spec_str_ov = f"{int(spec_limits[f][0])}-{int(spec_limits[f][1])}" if f in spec_limits and spec_limits[f][1] else (f">={int(spec_limits[f][0])}" if f in spec_limits and spec_limits[f][0] else "N/A")

                    if not df_ov.empty:
                        v, w = df_ov[f].values, df_ov['Good_Qty'].values
                        (m_s, s_s), (m_i, s_i) = calculate_stats(v, w, iqr_k)
                        for m_name, m_val, s_val in [("Standard", m_s, s_s), (f"IQR (k={iqr_k})", m_i, s_i)]:
                            is_std = (m_name == "Standard")
                            row = {
                                "Period": period,
                                "Feature": f if is_std else "", 
                                "Method": m_name,
                                "Limit": spec_str_ov if is_std else "", 
                                "Segment Dist": seg_dist_overall if is_std else "",
                                "TARGET GOAL": int(round(m_val)), "TOLERANCE": int(round(s_val)),
                                f"MILL {sigma_mill}σ": f"{max(0, int(round(m_val - sigma_mill*s_val)))}-{int(round(m_val + sigma_mill*s_val))}",
                                f"RELEASE {sigma_release}σ": f"{max(0, int(round(m_val - sigma_release*s_val)))}-{int(round(m_val + sigma_release*s_val))}"
                            }
                            period_overall_data.append(row)
                            
                            exp_ov_row = row.copy()
                            exp_ov_row['Feature'] = f; exp_ov_row['Limit'] = spec_str_ov; exp_ov_row['Segment Dist'] = seg_dist_overall
                            overall_export_data.append(exp_ov_row)

            display_period_df = pd.DataFrame(period_overall_data).drop(columns=['Period'], errors='ignore')
            st.dataframe(display_period_df, use_container_width=True, hide_index=True)

            st.subheader(f"🔍 Local Control Limits & I-MR ({period})")
            for thick in thickness_list:
                df_t = df_p[df_p['厚度'] == thick]
                if df_t.empty: continue
                
                st.markdown(f"**📏 Thick: {thick}**")
                period_thick_data = []
                plot_data_dict = {}
                
                for f in mech_features:
                    temp_calc = df_t[[f] + good_cols].dropna(subset=[f]).copy() if good_cols else pd.DataFrame()
                    if not temp_calc.empty:
                        temp_calc['Good_Qty'] = temp_calc[good_cols].sum(axis=1)
                        temp_calc = temp_calc[temp_calc['Good_Qty'] > 0]
                    spec_str = f"{int(spec_limits[f][0])}-{int(spec_limits[f][1])}" if f in spec_limits and spec_limits[f][1] else (f">={int(spec_limits[f][0])}" if f in spec_limits and spec_limits[f][0] else "N/A")

                    if not temp_calc.empty:
                        v, w = temp_calc[f].values, temp_calc['Good_Qty'].values
                        (m_s, s_s), (m_i, s_i) = calculate_stats(v, w, iqr_k)
                        plot_data_dict[f] = {'values': v, 'mean_std': m_s, 'std_std': s_s, 'mean_iqr': m_i, 'std_iqr': s_i}
                        
                        total_n = df_t[count_cols].sum().sum()
                        seg_dist = "N/A" if total_n == 0 else ", ".join([f"{k.replace('數','')}:{int(round(df_t[k].sum()/total_n*100))}%" for k in count_cols])
                        
                        for m_name, m_val, s_val in [("Standard", m_s, s_s), (f"IQR (k={iqr_k})", m_i, s_i)]:
                            is_std = (m_name == "Standard")
                            row = {
                                "Period": period, "Thickness": thick,
                                "Feature": f if is_std else "", "Method": m_name,
                                "Limit": spec_str if is_std else "", "Segment Dist": seg_dist if is_std else "",
                                "TARGET GOAL": int(round(m_val)), "TOLERANCE": int(round(s_val)),
                                f"MILL {sigma_mill}σ": f"{max(0, int(round(m_val - sigma_mill*s_val)))}-{int(round(m_val + sigma_mill*s_val))}",
                                f"RELEASE {sigma_release}σ": f"{max(0, int(round(m_val - sigma_release*s_val)))}-{int(round(m_val + sigma_release*s_val))}"
                            }
                            period_thick_data.append(row)
                            
                            exp_row = row.copy()
                            exp_row['Feature'] = f; exp_row['Limit'] = spec_str; exp_row['Segment Dist'] = seg_dist
                            all_export_data.append(exp_row)

                display_thick_df = pd.DataFrame(period_thick_data).drop(columns=['Period', 'Thickness'], errors='ignore')
                st.dataframe(display_thick_df, use_container_width=True, hide_index=True)
                
                # --- I-MR CHARTS ---
                cols_imr = st.columns(2)
                top4 = [f_name for f_name in ['YS', 'TS', 'EL', 'YPE'] if f_name in plot_data_dict]
                for idx, f in enumerate(top4):
                    with cols_imr[idx % 2]:
                        d = plot_data_dict[f]
                        v = d['values']
                        mv, sv = (d['mean_std'], d['std_std']) if chart_method == "Standard Method" else (d['mean_iqr'], d['std_iqr'])
                            
                        if len(v) > 1:
                            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 7), gridspec_kw={'height_ratios': [2, 1]})
                            ucl, lcl = mv + sigma_release*sv, max(0, mv - sigma_release*sv)
                            
                            ax1.plot(v, marker='o', color='#1f77b4', ms=4, lw=1, zorder=1)
                            outs = np.where((v > ucl) | (v < lcl))[0]
                            if len(outs) > 0:
                                ax1.scatter(outs, v[outs], color='red', s=60, zorder=2, label=f'Out of Control ({sigma_release}σ)')
                                ax1.legend(loc='upper left', fontsize=8)
                            
                            ax1.axhline(mv, color='green', ls='--', lw=1.5)
                            ax1.axhline(ucl, color='red', ls='--', lw=1.2)
                            ax1.axhline(lcl, color='red', ls='--', lw=1.2)
                            
                            v_max, v_min = np.max(v), np.min(v)
                            ax1.axhline(v_max, color='gray', ls=':', lw=1, alpha=0.7)
                            ax1.axhline(v_min, color='gray', ls=':', lw=1, alpha=0.7)
                            trans1 = ax1.get_yaxis_transform()
                            ax1.text(1.02, mv, f"Mean: {mv:.1f}", color='green', transform=trans1, va='center', fontweight='bold')
                            ax1.text(1.02, ucl, f"UCL: {ucl:.1f}", color='red', transform=trans1, va='center', fontweight='bold')
                            ax1.text(1.02, lcl, f"LCL: {lcl:.1f}", color='red', transform=trans1, va='center', fontweight='bold')
                            
                            ax1.set_title(f"I-Chart: {f} ({chart_method})", fontsize=11, fontweight='bold')
                            
                            mr = np.abs(np.diff(v))
                            mrm = np.mean(mr)
                            mru = 3.267 * mrm
                            
                            ax2.plot(mr, marker='o', color='orange', ms=4, lw=1, zorder=1)
                            mr_outs = np.where(mr > mru)[0]
                            if len(mr_outs) > 0: ax2.scatter(mr_outs, mr[mr_outs], color='red', s=60, zorder=2)
                            ax2.axhline(mrm, color='green', ls='--', lw=1.5)
                            ax2.axhline(mru, color='red', ls='--', lw=1.2)
                            
                            trans2 = ax2.get_yaxis_transform()
                            ax2.text(1.02, mrm, f"Mean: {mrm:.1f}", color='green', transform=trans2, va='center', fontweight='bold')
                            ax2.text(1.02, mru, f"UCL: {mru:.1f}", color='red', transform=trans2, va='center', fontweight='bold')
                            ax2.set_title("Moving Range", fontsize=10)
                            
                            fig.tight_layout(); fig.subplots_adjust(right=0.85)
                            st.pyplot(fig)
                            fig.savefig(f"imr_{f}_{thick}_{safe_p}.png", bbox_inches='tight')
            st.markdown("---")

    # --- EXPORT SECTION ---
    st.sidebar.header("📥 Full Export Options")
    if st.sidebar.button("Download Detailed Excel"):
        towrite = io.BytesIO()
        pd.DataFrame(all_export_data).to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.sidebar.download_button(label="Click to Download Excel", data=towrite, file_name="QC_Detailed_Optimization.xlsx")

    # --- PDF EXPORT ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("🖨️ Generate PDF Reports")
    def clean(t): return str(t).replace('±', '+/-').replace('–', '-').encode('latin-1', 'ignore').decode('latin-1')

    if st.sidebar.button("Generate FULL PDF (Detailed)"):
        pdf = FPDF(orientation='L')
        
        # 1. PIVOT SUMMARY
        if 'display_df' in locals() and not display_df.empty:
            pdf.add_page(); pdf.set_font('Arial', 'B', 16); pdf.cell(0, 10, "1. QUALITY YIELD SUMMARY", ln=True, align="C"); pdf.ln(5)
            pdf.set_font('Arial', 'B', 8)
            cw_tab1 = [25, 15, 25, 15] + [12]*len(grade_cols) + [12]*len(pct_cols)
            for i, col in enumerate(display_df.columns): pdf.cell(cw_tab1[i] if i < len(cw_tab1) else 20, 8, clean(col), border=1, align='C')
            pdf.ln(); pdf.set_font('Arial', '', 8)
            for _, r in display_df.head(25).iterrows(): 
                for i, v in enumerate(r): pdf.cell(cw_tab1[i] if i < len(cw_tab1) else 20, 8, clean(v), border=1, align='C')
                pdf.ln()
            if len(display_df) > 25: pdf.cell(0, 10, "...(See Excel for full pivot data)", ln=True)

        # 2. CHRONOLOGICAL PERIODS
        heads = ["Feature", "Method", "Limit", "Segment Dist", "TARGET", "TOL", f"MILL {sigma_mill}σ", f"RELEASE {sigma_release}σ"]
        c_w3 = [16, 22, 24, 60, 15, 12, 28, 30] 

        for period in selected_periods:
            safe_p = "".join([c if c.isalnum() else "_" for c in period])
            pdf.add_page(); pdf.set_font('Arial', 'B', 16); pdf.cell(0, 10, f"--- PERIOD: {period} ---", ln=True, align="C"); pdf.ln(5)
            
            pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, "Overall Factory Distribution", ln=True); ys = pdf.get_y()
            for idx, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
                path = f"overall_{f}_{safe_p}.png"
                if os.path.exists(path): pdf.image(path, x=(10 if idx%2==0 else 150), y=(ys if idx<2 else ys+75), w=135)
            
            pdf.add_page(); pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, "Overall Factory Goals", ln=True)
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

        pdf.output("QC_Full_Report_by_Period.pdf")
        with open("QC_Full_Report_by_Period.pdf", "rb") as f:
            st.sidebar.download_button("📥 Download FULL PDF", f.read(), "QC_Full_Report_by_Period.pdf", "application/pdf")
