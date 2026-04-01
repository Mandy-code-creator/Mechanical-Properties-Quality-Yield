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
    if 'Thickness' in df.columns:
        df.rename(columns={'Thickness': 'Actual_Thickness'}, inplace=True)
    else:
        for i, c in enumerate(df.columns):
            if '型式' in c and i > 0:
                df.rename(columns={df.columns[i-1]: 'Actual_Thickness'}, inplace=True)
                break

    if 'Actual_Thickness' in df.columns:
        df['Actual_Thickness'] = pd.to_numeric(df['Actual_Thickness'], errors='coerce').round(3)

    # --- 2. DATE PARSING ---
    if '烤三生產日期' in df.columns:
        d_str = df['烤三生產日期'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df['烤三生產日期'] = pd.to_datetime(d_str, format='%Y%m%d', errors='coerce').fillna(pd.to_datetime(d_str, errors='coerce'))

    # --- 3. GRADE COLUMNS MERGING ---
    base_grades = ['A-B+', 'A-B', 'A-B-', 'B+', 'B']
    for g in base_grades:
        match_cols = [c for c in df.columns if c == g or str(c).startswith(f"{g}.")]
        df[g] = df[match_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) if match_cols else 0
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
        df = df[df['Time_Group'] != "Other"]
        df_25 = df[df['烤三生產日期'].dt.year == 2025].copy()
        if not df_25.empty:
            df_25['Time_Group'] = "2025 (Full Year)"
            df = pd.concat([df, df_25], ignore_index=True)
    else:
        df['Time_Group'] = "Unknown"

    tab0, tab1, tab2, tab3 = st.tabs(["0. Raw Data Check", "1. Yield Summary", "2. Distribution", "3. Control Limits & I-MR"])

    with tab0:
        st.header("0. System Health & Raw Data Check")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Rows", len(df))
        c2.write("**Thickness:**")
        c2.write(df['Actual_Thickness'].value_counts().dropna().to_dict())
        c3.write("**Periods:**")
        c3.write(df['Time_Group'].value_counts().to_dict())
        st.dataframe(df.head(20), use_container_width=True)

    # FILTERING
    df = df[df['Actual_Thickness'].isin([0.5, 0.6, 0.75, 0.8])]
    st.sidebar.header("🔎 Filters")
    all_periods = sorted(df['Time_Group'].unique())
    ui_selection = st.sidebar.multiselect("📅 Select Period(s):", options=["All"] + all_periods, default=["All"])
    selected_periods = all_periods if ("All" in ui_selection or not ui_selection) else ui_selection
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
                fmin, fmax = q1 - (q99-q1)*0.2, q99 + (q99-q1)*0.2
                global_x_bounds[feat] = (fmin, fmax)

    def get_shared_y(data, features):
        max_y = 0
        for feat in features:
            if feat in data.columns:
                vd = data.dropna(subset=[feat]).copy()
                vd['T_Count'] = vd[base_grades].sum(axis=1)
                if not vd.empty:
                    fmin, fmax = global_x_bounds.get(feat, (vd[feat].min(), vd[feat].max()))
                    cnts, _ = np.histogram(vd[feat], bins=np.linspace(fmin, fmax, 16), weights=vd['T_Count'])
                    max_y = max(max_y, cnts.max())
        return max_y * 1.3 if max_y > 0 else 50

    # --- TAB 1: YIELD ---
    with tab1:
        st.header("1. Quality Yield Summary")
        group_cols = ['Time_Group', 'Actual_Thickness', 'HR_Material']
        sum_df = df.groupby(group_cols, dropna=False)[base_grades].sum().reset_index()
        sum_df['Total_Qty'] = sum_df[base_grades].sum(axis=1)
        for col in base_grades: sum_df[f"% {col}"] = ((sum_df[col] / sum_df['Total_Qty'].replace(0, np.nan)) * 100).fillna(0).round(1)
        sum_df.rename(columns={'Time_Group': 'Period', 'Actual_Thickness': 'Thickness'}, inplace=True)
        for period in selected_periods:
            p_data = sum_df[sum_df['Period'] == period]
            if not p_data.empty:
                st.markdown(f"### 📅 Period: **{period}**")
                st.dataframe(p_data.drop(columns=['Period']), use_container_width=True, hide_index=True)

    # --- TAB 2 & 3 CORE ---
    def calc_stats(v, w, k):
        m_s = np.average(v, weights=w)
        s_s = np.sqrt(np.average((v - m_s)**2, weights=w))
        try:
            exp_v = np.repeat(v, w.astype(int))
            q1, q3 = np.percentile(exp_v, 25), np.percentile(exp_v, 75)
            mask = (v >= q1 - k*(q3-q1)) & (v <= q3 + k*(q3-q1))
            m_i = np.average(v[mask], weights=w[mask])
            s_i = np.sqrt(np.average((v[mask] - m_i)**2, weights=w[mask]))
            return (m_s, s_s), (m_i, s_i)
        except: return (m_s, s_s), (m_s, s_s)

    def plot_dist(ax, data, feat, title, y_lim):
        c_map = {'A-B+': '#2ca02c', 'A-B': '#1f77b4', 'A-B-': '#ff7f0e', 'B+': '#9467bd', 'B': '#d62728'}
        f_min, f_max = global_x_bounds.get(feat, (data[feat].min(), data[feat].max()))
        v_list, w_list, colors, mean_info = [], [], [], []
        for g in base_grades:
            td = data[[feat, g]].dropna()
            td = td[td[g] > 0]
            if not td.empty:
                v, w = td[feat].values, td[g].values
                v_list.append(v); w_list.append(w); colors.append(c_map[g])
                m = np.average(v, weights=w)
                ax.axvline(m, color=c_map[g], ls='--', lw=1.2)
                mean_info.append({'v': m, 'c': c_map[g]})
        if v_list:
            ax.hist(v_list, bins=np.linspace(f_min, f_max, 16), weights=w_list, color=colors, stacked=True, edgecolor='white', alpha=0.7)
            # --- KHÔI PHỤC MEAN LABELS TẠI ĐÂY ---
            mean_info.sort(key=lambda x: x['v'])
            for i, info in enumerate(mean_info):
                h = y_lim * (0.9 - (i % 3) * 0.1)
                ax.text(info['v'], h, f"{info['v']:.1f}", color='white', fontweight='bold', fontsize=8, ha='center', bbox=dict(facecolor=info['c'], alpha=0.8, boxstyle='round,pad=0.2'))
        ax.set_xlim(f_min, f_max); ax.set_ylim(0, y_lim); ax.set_title(title, fontsize=9, fontweight='bold')

    with tab3:
        st.markdown("##### ⚙️ Production Settings")
        c1, c2, c3, c4 = st.columns(4)
        sig_rel = c1.number_input("Release Range (Sigma)", value=2.0)
        sig_mil = c2.number_input("Mill Range (Sigma)", value=1.0)
        iqr_k = c3.number_input("IQR Filter (k)", value=1.5)
        chart_m = c4.radio("I-MR Based On:", ["Standard", "IQR Filtered"])

    for period in selected_periods:
        df_p = df[df['Time_Group'] == period]
        if df_p.empty: continue
        with tab2:
            st.markdown(f"## 📅 Period: **{period}**")
            ov_y = get_shared_y(df_p, ['YS', 'TS', 'EL', 'YPE'])
            cols = st.columns(2)
            for idx, f in enumerate([x for x in ['YS', 'TS', 'EL', 'YPE'] if x in df_p.columns]):
                with cols[idx%2]:
                    fig, ax = plt.subplots(figsize=(7, 4))
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
                        fig, ax = plt.subplots(figsize=(7, 4))
                        plot_dist(ax, df_t, f, f"{f} (Thick: {thick})", ly)
                        st.pyplot(fig); plt.close(fig)
        with tab3:
            st.markdown(f"## 📅 Period: **{period}**")
            for thick in thickness_list:
                df_t = df_p[df_p['Actual_Thickness'] == thick]
                if df_t.empty: continue
                st.markdown(f"**📏 Thickness: {thick}**")
                p_data, p_dict = [], {}
                for f in [x for x in ['YS', 'TS', 'EL', 'YPE'] if x in df_t.columns]:
                    tc = df_t[[f, 'A-B+', 'A-B']].dropna()
                    tc['G_Qty'] = tc[['A-B+', 'A-B']].sum(axis=1)
                    tc = tc[tc['G_Qty'] > 0]
                    if not tc.empty:
                        (ms, ss), (mi, si) = calc_stats(tc[f].values, tc['G_Qty'].values, iqr_k)
                        curr_m, curr_s = (ms, ss) if chart_m == "Standard" else (mi, si)
                        p_dict[f] = {'v': tc[f].values, 'm': curr_m, 's': curr_s}
                        for n, mv, sv in [("Standard", ms, ss), ("IQR Filtered", mi, si)]:
                            p_data.append({"Feature": f if n=="Standard" else "", "Method": n, "TARGET": int(round(mv)), "TOL": int(round(sv)), 
                                           f"MILL {sig_mil}σ": f"{max(0,int(round(mv-sig_mil*sv)))}-{int(round(mv+sig_mil*sv))}",
                                           f"RELEASE {sig_rel}σ": f"{max(0,int(round(mv-sig_rel*sv)))}-{int(round(mv+sig_rel*sv))}"})
                if p_data: st.dataframe(pd.DataFrame(p_data), use_container_width=True, hide_index=True)
                imr_c = st.columns(2)
                for idx, f in enumerate(p_dict):
                    with imr_c[idx%2]:
                        d = p_dict[f]
                        if len(d['v']) > 1:
                            fig, (a1, a2) = plt.subplots(2, 1, figsize=(7, 5), gridspec_kw={'height_ratios': [2, 1]})
                            u, l = d['m']+sig_rel*d['s'], max(0, d['m']-sig_rel*d['s'])
                            a1.plot(d['v'], marker='o', ms=3, lw=1); a1.axhline(d['m'], color='g', ls='--'); a1.axhline(u, color='r', ls=':'); a1.axhline(l, color='r', ls=':')
                            a1.set_title(f"I-Chart: {f}", fontsize=9); mr = np.abs(np.diff(d['v']))
                            a2.plot(mr, marker='o', ms=3, lw=1, color='orange'); a2.axhline(np.mean(mr), color='g', ls='--'); a2.axhline(3.267*np.mean(mr), color='r', ls=':')
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
