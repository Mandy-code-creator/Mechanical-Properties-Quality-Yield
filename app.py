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

    # --- ĐÃ SỬA: THÊM TAB 3 VÀO DANH SÁCH KHAI BÁO ---
    tab0, tab1, tab2, tab3 = st.tabs(["0. Raw Check", "1. Yield Summary", "2. Distribution Analysis", "3.🔍 Root Cause & Diagnostic Analysis"])

    with tab0:
        st.dataframe(df.head(10), use_container_width=True)

    # --- GLOBAL FILTERING ---
    df = df[df['Actual_Thickness'].isin([0.5, 0.6, 0.75, 0.8])]
    st.sidebar.header("🔎 Filters")
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

        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df.to_excel(writer, index=False, sheet_name='Yield_Summary')
        st.download_button(
            label="📥 Download Yield Summary (Excel)",
            data=output.getvalue(),
            file_name="Yield_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- TAB 2: DISTRIBUTION ---
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

    with tab2:
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

# --- TAB 3: DIAGNOSTIC FLOW (4 STEPS) ---
    with tab3:
        st.header("🔍 4-Step Root Cause Diagnostic Flow")
        
        # --- STEP 1: HEATMAP ---
        defect_str = ", ".join(bad_grades)
        st.subheader(f"Step 1: Locate the Hotspot (% Defect Rate: {defect_str})")
        df_filtered['Defect_Rate'] = (df_filtered['Bad_Qty'] / df_filtered['Total_Qty'] * 100).fillna(0)
        df_filtered['Spec_Label'] = df_filtered['Actual_Thickness'].astype(str) + "mm (" + df_filtered['HR_Material'] + ")"
        
        heat_data = df_filtered.groupby(['Spec_Label', 'Time_Group'])['Defect_Rate'].mean().unstack()

        if not heat_data.empty:
            fig, ax = plt.subplots(figsize=(12, 6))
            sns.heatmap(heat_data, annot=True, fmt=".1f", cmap="YlOrRd", linewidths=.5, ax=ax)
            ax.set_title("HOTSPOT MAP: Redder = Higher Defect Rate", fontsize=12, fontweight='bold', color='#d62728', pad=15)
            ax.set_ylabel("Specification (Thickness & Material)")
            ax.set_xlabel("Production Period")
            st.pyplot(fig); plt.close(fig)

        # --- STEP 2: PARETO ---
        st.markdown("---")
        st.subheader("Step 2: Identify Main Defect Category (Pareto)")
        defect_sums = df_filtered[bad_grades].sum().sort_values(ascending=False)
        if defect_sums.sum() > 0:
            pareto_df = pd.DataFrame({'Count': defect_sums})
            pareto_df['Cum_Pct'] = pareto_df['Count'].cumsum() / pareto_df['Count'].sum() * 100
            fig, ax1 = plt.subplots(figsize=(8, 4))
            ax1.bar(pareto_df.index, pareto_df['Count'], color="#d62728", alpha=0.8)
            ax1.set_ylabel("Defective Coils")
            ax2 = ax1.twinx()
            ax2.plot(pareto_df.index, pareto_df['Cum_Pct'], color="#1f77b4", marker="D", ms=5)
            ax2.axhline(80, color="orange", linestyle="--")
            st.pyplot(fig); plt.close(fig)

        # --- STEP 3: OVERLAY ---
        st.markdown("---")
        st.subheader("Step 3: Property Shift Analysis (GOOD vs BAD)")
        active_mechs = [f for f in ['YS', 'TS', 'EL', 'YPE'] if f in df_filtered.columns]
        feat_diag = st.selectbox("Select property to diagnose shift:", active_mechs, key="diag_feat")
        
        fig, ax = plt.subplots(figsize=(10, 4))
        g_vals = df_filtered[df_filtered['Good_Qty'] > 0][feat_diag].dropna()
        b_vals = df_filtered[df_filtered['Bad_Qty'] > 0][feat_diag].dropna()
        
        if not g_vals.empty: sns.kdeplot(g_vals, ax=ax, label="GOOD", fill=True, color="green", alpha=0.3)
        if not b_vals.empty: sns.kdeplot(b_vals, ax=ax, label="BAD", fill=True, color="red", alpha=0.3)
        ax.set_title(f"Distribution Shift: {feat_diag} (Good vs Bad)")
        ax.legend(); st.pyplot(fig); plt.close(fig)

       # --- STEP 4: I-MR CHART WITH SPECS ---
        st.markdown("---")
        st.subheader(f"Step 4: Time-Series Stability Tracking (I-MR Chart for {feat_diag})")
        st.info("Filter down to a specific specification to view its timeline against Standard Specs.")
        
        # --- THÊM 2 DÒNG NÀY ĐỂ SỬA LỖI NAMERROR ---
        thickness_list = sorted(df_filtered['Actual_Thickness'].dropna().unique())
        material_list = sorted(df_filtered['HR_Material'].astype(str).unique())
        
        c1, c2, c3 = st.columns(3)
        sel_period = c1.selectbox("Filter by Period:", selected_periods)
        sel_thick = c2.selectbox("Filter by Thickness:", thickness_list)
        sel_mat = c3.selectbox("Filter by Material:", material_list)

        # ... (phần code vẽ I-MR Chart bên dưới giữ nguyên) ...
        
        c1, c2, c3 = st.columns(3)
        sel_period = c1.selectbox("Filter by Period:", selected_periods)
        sel_thick = c2.selectbox("Filter by Thickness:", thickness_list)
        sel_mat = c3.selectbox("Filter by Material:", material_list)

        imr_data = df_filtered[(df_filtered['Time_Group'] == sel_period) & 
                               (df_filtered['Actual_Thickness'] == sel_thick) & 
                               (df_filtered['HR_Material'] == sel_mat)]

        if not imr_data.empty and len(imr_data[feat_diag].dropna()) > 1:
            vals = imr_data[feat_diag].dropna().values
            mean_val = np.mean(vals)
            
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 7), gridspec_kw={'height_ratios': [2, 1]})
            
            # I-CHART
            ax1.plot(vals, marker='o', ms=4, lw=1.5, color='#1f77b4', label='Individual Value')
            ax1.axhline(mean_val, color='green', ls='--', label='Process Mean')
            
            # ADD GLOBAL SPECS
            if feat_diag in GLOBAL_SPECS:
                s = GLOBAL_SPECS[feat_diag]
                if s.get('min'): ax1.axhline(s['min'], color='red', ls='-', lw=2, label=f"Lower Spec ({s['min']})")
                if s.get('max'): ax1.axhline(s['max'], color='red', ls='-', lw=2, label=f"Upper Spec ({s['max']})")
                if s.get('target'): ax1.axhline(s['target'], color='black', ls=':', lw=1.5, label=f"Target ({s['target']})")
                if s.get('min') and s.get('max'):
                    ax1.axhspan(s['min'], s['max'], color='green', alpha=0.05)

            ax1.set_title(f"I-Chart: {feat_diag} | {sel_period} | {sel_thick}mm ({sel_mat})", fontweight='bold')
            ax1.legend(loc='upper right', fontsize=8)

            # MR-CHART
            mr = np.abs(np.diff(vals))
            mr_mean = np.mean(mr)
            ax2.plot(mr, marker='o', ms=4, lw=1.5, color='orange')
            ax2.axhline(mr_mean, color='green', ls='--', label='MR Mean')
            ax2.axhline(3.267 * mr_mean, color='red', ls=':', label='UCL (MR)')
            ax2.set_title("Moving Range (MR-Chart)", fontweight='bold')
            ax2.legend(loc='upper right', fontsize=8)

            fig.tight_layout()
            st.pyplot(fig); plt.close(fig)
        else:
            st.warning("Not enough data to generate I-MR chart for this specific combination.")
    # --- EXPORT SECTION ---
    # (Giữ nguyên logic Export ban đầu của bạn...)
    st.sidebar.header("📥 Export Options")
  
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
