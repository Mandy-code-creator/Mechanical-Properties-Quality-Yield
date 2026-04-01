import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import scipy.stats as stats
import io
from fpdf import FPDF
import os
from matplotlib.patches import Patch

st.set_page_config(page_title="QC Yield & Control Limit", layout="wide")
st.title("📊 Quality Yield & Control Limit Optimizer")
st.markdown("---")

uploaded_file = st.file_uploader("Upload Excel data (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.astype(str).str.strip()

    # --- 1. THICKNESS EXTRACTION ---
    # Prioritize the explicitly renamed 'Thickness' column
    if 'Thickness' in df.columns:
        df.rename(columns={'Thickness': 'Actual_Thickness'}, inplace=True)
    else:
        # Fallback if 'Thickness' is not found
        for i, c in enumerate(df.columns):
            if '型式' in c and i > 0:
                df.rename(columns={df.columns[i-1]: 'Actual_Thickness'}, inplace=True)
                break
        if 'Actual_Thickness' not in df.columns and '厚度' in df.columns:
            df.rename(columns={'厚度': 'Actual_Thickness'}, inplace=True)

    if 'Actual_Thickness' in df.columns:
        df['Actual_Thickness'] = pd.to_numeric(df['Actual_Thickness'], errors='coerce').round(3)
    else:
        df['Actual_Thickness'] = np.nan

    # --- 2. DATE PARSING ---
    if '烤三生產日期' in df.columns:
        d_str = df['烤三生產日期'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df['烤三生產日期'] = pd.to_datetime(d_str, format='%Y%m%d', errors='coerce').fillna(pd.to_datetime(d_str, errors='coerce'))

    # --- 3. GRADE COLUMNS MERGING ---
    base_grades = ['A-B+', 'A-B', 'A-B-', 'B+', 'B']
    for g in base_grades:
        match_cols = [c for c in df.columns if c == g or str(c).startswith(f"{g}.")]
        if match_cols:
            df[g] = df[match_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
        else:
            df[g] = 0
    df['Total_Qty'] = df[base_grades].sum(axis=1)

    # --- 4. MECH FEATURES & MATERIAL ---
    df.rename(columns={'烤漆降伏強度': 'YS', '烤漆抗拉強度': 'TS', '伸長率': 'EL'}, inplace=True)
    mech_features = ['YS', 'TS', 'EL', 'YPE', 'HARDNESS']
    for f in mech_features:
        if f in df.columns: df[f] = pd.to_numeric(df[f], errors='coerce')

    if '熱軋材質' in df.columns:
        df['HR_Material'] = df['熱軋材質'].astype(str).str.strip().replace(['nan', ''], 'Unknown')
    else:
        df['HR_Material'] = 'Unknown'

    # --- 5. TIME GROUPS ---
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
    else:
        df['Time_Group'] = "Unknown"

    # ==========================================
    # TAB 0: RAW DATA CHECK (BEFORE FILTERING)
    # ==========================================
    tab0, tab1, tab2, tab3 = st.tabs(["0. Raw Data Check", "1. Yield Summary", "2. Distribution", "3. Control Limits & I-MR"])

    with tab0:
        st.header("0. System Health & Raw Data Check")
        st.info("Showing parsed data BEFORE removing any invalid thicknesses or periods.")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Rows Read", len(df))
        c2.write("**Thickness Found:**")
        c2.write(df['Actual_Thickness'].value_counts().dropna().to_dict())
        c3.write("**Periods Found:**")
        c3.write(df['Time_Group'].value_counts().to_dict())
        
        preview_cols = ['烤三生產日期', 'Time_Group', 'Actual_Thickness', 'HR_Material', 'Total_Qty'] + base_grades + [f for f in mech_features if f in df.columns]
        exist_preview = [c for c in preview_cols if c in df.columns]
        st.dataframe(df[exist_preview].head(30), use_container_width=True)

    # ==========================================
    # APPLY STRICT FILTERS & VIRTUAL DATA
    # ==========================================
    df = df[df['Actual_Thickness'].isin([0.5, 0.6, 0.75, 0.8])]
    df = df[df['Time_Group'] != "Other"]

    if '烤三生產日期' in df.columns:
        df_25 = df[df['烤三生產日期'].dt.year == 2025].copy()
        if not df_25.empty:
            df_25['Time_Group'] = "2025 (Full Year)"
            df = pd.concat([df, df_25], ignore_index=True)

    # --- SIDEBAR UI ---
    st.sidebar.header("🔎 Dashboard Filters")
    all_periods = sorted(df['Time_Group'].unique())
    ui_selection = st.sidebar.multiselect("📅 Select Period(s):", options=["All"] + all_periods, default=["All"])
    
    if "All" in ui_selection or len(ui_selection) == 0:
        selected_periods = all_periods
    else:
        selected_periods = ui_selection
        
    df = df[df['Time_Group'].isin(selected_periods)]
    thickness_list = sorted(df['Actual_Thickness'].dropna().unique())

    # --- GLOBAL BOUNDS ---
    global_x_bounds = {}
    for feat in mech_features:
        if feat in df.columns:
            vd = df[[feat] + base_grades].dropna(subset=[feat]).copy()
            vd['T_Count'] = vd[base_grades].sum(axis=1)
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
                vd['T_Count'] = vd[base_grades].sum(axis=1)
                if not vd.empty:
                    fmin, fmax = global_x_bounds.get(feat, (0, 100))
                    cnts, _ = np.histogram(vd[feat], bins=np.linspace(fmin, fmax, 16), weights=vd['T_Count'])
                    max_y = max(max_y, cnts.max())
        return max_y * 1.35 if max_y > 0 else 50

    # --- TAB 1: YIELD SUMMARY ---
    with tab1:
        st.header("1. Quality Yield Summary")
        group_cols = ['Time_Group', 'Actual_Thickness', 'HR_Material']
        if set(group_cols).issubset(df.columns):
            sum_df = df.groupby(group_cols, dropna=False)[base_grades].sum().reset_index()
            sum_df['Total_Qty'] = sum_df[base_grades].sum(axis=1)
            for col in base_grades: 
                sum_df[f"% {col}"] = ((sum_df[col] / sum_df['Total_Qty'].replace(0, np.nan)) * 100).fillna(0).round(1)
            
            sum_df.rename(columns={'Time_Group': 'Period', 'Actual_Thickness': 'Thickness'}, inplace=True)
            for period in selected_periods:
                p_data = sum_df[sum_df['Period'] == period]
                if not p_data.empty:
                    st.markdown(f"### 📅 Period: **{period}**")
                    st.dataframe(p_data.drop(columns=['Period']), use_container_width=True, hide_index=True)
            
            towrite = io.BytesIO()
            sum_df.to_excel(towrite, index=False, engine='openpyxl')
            towrite.seek(0)
            st.download_button("📥 Download Summary (Excel)", data=towrite, file_name="Yield_Summary.xlsx")

    # --- TAB 2 & 3 FUNCTIONS ---
    def calc_stats(v_arr, w_arr, k_f):
        m_s = np.average(v_arr, weights=w_arr)
        s_s = np.sqrt(np.average((v_arr - m_s)**2, weights=w_arr))
        try:
            exp_v = np.repeat(v_arr, w_arr.astype(int))
            q1, q3 = np.percentile(exp_v, 25), np.percentile(exp_v, 75)
            iqr = q3 - q1
            mask = (v_arr >= q1 - k_f * iqr) & (v_arr <= q3 + k_f * iqr)
            vf, wf = v_arr[mask], w_arr[mask]
            if len(vf) > 0 and sum(wf) > 0:
                m_i = np.average(vf, weights=wf)
                s_i = np.sqrt(np.average((vf - m_i)**2, weights=wf))
            else: m_i, s_i = m_s, s_s
        except: m_i, s_i = m_s, s_s
        return (m_s, s_s), (m_i, s_i)

    def plot_dist(ax, data, feat, title, y_lim):
        c_map = {'A-B+': '#2ca02c', 'A-B': '#1f77b4', 'A-B-': '#ff7f0e', 'B+': '#9467bd', 'B': '#d62728'}
        f_min, f_max = global_x_bounds.get(feat, (0, 100))
        bins = np.linspace(f_min, f_max, 16)
        v_list, w_list, c_list = [], [], []
        
        for c_n in base_grades:
            if c_n in data.columns and feat in data.columns:
                td = data[[feat, c_n]].dropna()
                td = td[td[c_n] > 0]
                if not td.empty:
                    v, w = td[feat].values, td[c_n].values
                    v_list.append(v); w_list.append(w); c_list.append(c_map.get(c_n, '#7f7f7f'))
                    m = np.average(v, weights=w)
                    ax.axvline(m, color=c_map.get(c_n, '#7f'), ls='--', lw=1.5)
        
        if v_list:
            ax.hist(v_list, bins=bins, weights=w_list, color=c_list, stacked=True, edgecolor='white', alpha=0.8)
        ax.set_xlim(f_min, f_max); ax.set_ylim(0, y_lim)
        ax.set_title(title, fontsize=10, fontweight='bold')

    with tab3:
        st.markdown("##### ⚙️ Production Configuration")
        c1, c2, c3, c4 = st.columns(4)
        sig_rel = c1.number_input("Release Range (Sigma)", value=2.0, step=0.1)
        sig_mill = c2.number_input("Mill Range (Sigma)", value=1.0, step=0.1)
        iqr_k = c3.number_input("IQR Filter (k)", value=1.5, step=0.1)
        chart_m = c4.radio("I-MR Limits Based On:", ["Standard Method", "IQR Filtered Method"])

    specs = {"YS": (405, 500), "TS": (415, 550), "EL": (25, None), "YPE": (4, None)}
    g_cols = ['A-B+', 'A-B']

    for period in selected_periods:
        df_p = df[df['Time_Group'] == period]
        if df_p.empty: continue

        with tab2:
            st.markdown(f"## 📅 Period: **{period}**")
            ov_y = get_shared_y(df_p, mech_features)
            cols = st.columns(2)
            for i, f in enumerate([x for x in ['YS', 'TS', 'EL', 'YPE'] if x in df.columns]):
                with cols[i%2]:
                    fig, ax = plt.subplots(figsize=(8, 4))
                    plot_dist(ax, df_p, f, f"{f} (Overall)", ov_y)
                    st.pyplot(fig); plt.close(fig)
            
            for thick in thickness_list:
                df_t = df_p[df_p['Actual_Thickness'] == thick]
                if df_t.empty: continue
                st.markdown(f"**📏 Thickness: {thick}**")
                l_y = get_shared_y(df_t, mech_features)
                t_cols = st.columns(2)
                for i, f in enumerate([x for x in ['YS', 'TS', 'EL', 'YPE'] if x in df.columns]):
                    with t_cols[i%2]:
                        fig, ax = plt.subplots(figsize=(8, 4))
                        plot_dist(ax, df_t, f, f"{f} (Thick: {thick})", l_y)
                        st.pyplot(fig); plt.close(fig)
            st.markdown("---")

        with tab3:
            st.markdown(f"## 📅 Period: **{period}**")
            for thick in thickness_list:
                df_t = df_p[df_p['Actual_Thickness'] == thick]
                if df_t.empty: continue
                
                st.markdown(f"**📏 Thickness: {thick}**")
                p_data, p_dict = [], {}
                
                for f in mech_features:
                    if f in df.columns:
                        tc = df_t[[f] + g_cols].dropna()
                        if not tc.empty:
                            tc['G_Qty'] = tc[g_cols].sum(axis=1)
                            tc = tc[tc['G_Qty'] > 0]
                        if not tc.empty:
                            v, w = tc[f].values, tc['G_Qty'].values
                            (ms, ss), (mi, si) = calc_stats(v, w, iqr_k)
                            p_dict[f] = {'v': v, 'm': ms if chart_m == "Standard Method" else mi, 's': ss if chart_m == "Standard Method" else si}
                            for m_n, m_v, s_v in [("Standard", ms, ss), (f"IQR (k={iqr_k})", mi, si)]:
                                p_data.append({
                                    "Feature": f if m_n == "Standard" else "", "Method": m_n,
                                    "TARGET": int(round(m_v)), "TOLERANCE": int(round(s_v)),
                                    f"MILL {sig_mill}σ": f"{max(0, int(round(m_v - sig_mill*s_v)))}-{int(round(m_v + sig_mill*s_v))}",
                                    f"RELEASE {sig_rel}σ": f"{max(0, int(round(m_v - sig_rel*s_v)))}-{int(round(m_v + sig_rel*s_v))}"
                                })
                
                if p_data: st.dataframe(pd.DataFrame(p_data), use_container_width=True, hide_index=True)
                
                # I-MR Plots
                imr_c = st.columns(2)
                for idx, f in enumerate([x for x in ['YS', 'TS', 'EL', 'YPE'] if x in p_dict]):
                    with imr_c[idx%2]:
                        d = p_dict[f]
                        if len(d['v']) > 1:
                            fig, (a1, a2) = plt.subplots(2, 1, figsize=(8, 6), gridspec_kw={'height_ratios': [2, 1]})
                            ucl, lcl = d['m'] + sig_rel*d['s'], max(0, d['m'] - sig_rel*d['s'])
                            a1.plot(d['v'], marker='o', color='#1f77b4', lw=1)
                            a1.axhline(d['m'], color='green', ls='--'); a1.axhline(ucl, color='red', ls=':'); a1.axhline(lcl, color='red', ls=':')
                            a1.set_title(f"I-Chart: {f}", fontsize=10)
                            
                            mr = np.abs(np.diff(d['v']))
                            mrm, mru = np.mean(mr), 3.267 * np.mean(mr)
                            a2.plot(mr, marker='o', color='orange', lw=1)
                            a2.axhline(mrm, color='green', ls='--'); a2.axhline(mru, color='red', ls=':')
                            a2.set_title("Moving Range", fontsize=9)
                            fig.tight_layout(); st.pyplot(fig); plt.close(fig)
            st.markdown("---")

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
