import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import scipy.stats as stats
import math
import io
from fpdf import FPDF
import os
from matplotlib.patches import Patch

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="QC Mechanical Properties Optimizer", layout="wide")

st.title("📊 Mechanical Properties & Quality Yield Optimizer (Standard QC View)")
st.markdown("---")

# --- 1. FILE UPLOAD ---
uploaded_file = st.file_uploader("Upload your Excel data (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip() 

    # Khởi tạo biến toàn cục cho Export
    all_export_data = []
    overall_export_data = [] # <--- THÊM BIẾN LƯU DỮ LIỆU OVERALL
    
    # --- 2. DATA PREPROCESSING ---
    count_cols = ['A-B+數', 'A-B數', 'A-B-數', 'B+數', 'B數']
    count_cols = [col for col in count_cols if col in df.columns]
    
    for col in count_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['Total_Count'] = df[count_cols].sum(axis=1)

    mech_features = ['YS', 'TS', 'EL', 'YPE', 'HARDNESS']
    mech_features = [feat for feat in mech_features if feat in df.columns]
    for feat in mech_features:
        df[feat] = pd.to_numeric(df[feat], errors='coerce')

    thickness_list = sorted(df['厚度歸類'].dropna().unique(), key=lambda x: float(x))

    # --- TÍNH TOÁN KHUNG TRỤC X CỐ ĐỊNH ---
    global_x_bounds = {}
    for feat in mech_features:
        vd = df[[feat] + count_cols].dropna(subset=[feat]).copy()
        vd['T_Count'] = vd[count_cols].sum(axis=1)
        vd = vd[vd['T_Count'] > 0]
        if not vd.empty:
            q1 = np.percentile(vd[feat], 1)
            q99 = np.percentile(vd[feat], 99)
            iqr = q99 - q1
            fmin = max(vd[feat].min(), q1 - 0.5*iqr)
            fmax = min(vd[feat].max(), q99 + 0.5*iqr)
            if fmin >= fmax:
                fmin -= 5; fmax += 5
            buf = (fmax - fmin) * 0.05 if (fmax - fmin) > 0 else 5
            global_x_bounds[feat] = (fmin - buf, fmax + buf)

    # --- HÀM TÍNH TOÁN THANG ĐO Y ---
    def get_shared_y(data, features):
        max_y = 0
        for feat in features:
            if feat in data.columns:
                vd = data.dropna(subset=[feat]).copy()
                vd['T_Count'] = vd[count_cols].sum(axis=1)
                vd = vd[vd['T_Count'] > 0]
                if not vd.empty:
                    fmin, fmax = global_x_bounds.get(feat, (vd[feat].min(), vd[feat].max()))
                    bins = np.linspace(fmin, fmax, 16)
                    cnts, _ = np.histogram(vd[feat], bins=bins, weights=vd['T_Count'])
                    max_y = max(max_y, cnts.max())
        return max_y * 1.35 if max_y > 0 else 50 

    # --- 3. CREATE TABS ---
    tab1, tab2, tab3 = st.tabs([
        "1. Summary & Yields", 
        "2. Distribution Analysis",
        "3. Control Limits & Goals"
    ])

    # --- TAB 1: SUMMARY ---
    with tab1:
        st.header("1. Quality Summary by Thickness")
        summary_df = df.groupby('厚度歸類')[count_cols].sum().reset_index()
        summary_df['Total Data'] = summary_df[count_cols].sum(axis=1)
        for col in count_cols:
            summary_df[f"% {col.replace('數','')}"] = (summary_df[col] / summary_df['Total Data'] * 100).fillna(0).round(2)
        display_df = summary_df.copy()
        display_df.rename(columns={'厚度歸類': 'Thickness'}, inplace=True)
        for col in count_cols:
            display_df.rename(columns={col: col.replace('數','')}, inplace=True)
        display_df.insert(0, 'No.', range(1, len(display_df) + 1))
        
        int_cols = [c.replace('數','') for c in count_cols] + ['Total Data', 'No.']
        for c in int_cols:
            if c in display_df.columns:
                display_df[c] = display_df[c].astype(int)
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    # --- TAB 2: DISTRIBUTION ANALYSIS ---
    with tab2:
        st.header("2. Mechanical Properties Distribution Analysis")

        def plot_qc_dist(ax, data, feat, title, custom_y_limit, is_right=False):
            k_b = 15 
            color_map = {
                'A-B+數': '#2ca02c', 'A-B數': '#1f77b4', 'A-B-數': '#ff7f0e', 
                'B+數': '#9467bd', 'B數': '#d62728'
            }
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
                    
                    vals_list.append(vals)
                    wgts_list.append(wgts)
                    colors_list.append(color)
                    
                    m = np.average(vals, weights=wgts)
                    s = np.sqrt(np.average((vals - m)**2, weights=wgts))
                    
                    ax.axvline(m, color=color, ls='--', lw=1.5, zorder=3)
                    mean_inf.append({'val': m, 'color': color})

                    if len(vals) > 2 and s > 0:
                        x_r = np.linspace(f_min, f_max, 100)
                        bin_w = (f_max - f_min) / k_b
                        ax.plot(x_r, stats.norm.pdf(x_r, m, s) * wgts.sum() * bin_w, color=color, lw=2, zorder=4)

            if vals_list:
                ax.hist(vals_list, bins=bins_arr, weights=wgts_list, color=colors_list, 
                        stacked=True, edgecolor='white', alpha=0.8, zorder=2)

            ax.set_xlim(f_min, f_max)
            ax.set_ylim(0, custom_y_limit)

            if mean_inf:
                mean_inf.sort(key=lambda x: x['val'])
                levels = [0.90, 0.82, 0.74, 0.66, 0.58]
                for i, info in enumerate(mean_inf):
                    y_pos = custom_y_limit * levels[i % len(levels)]
                    ax.text(info['val'], y_pos, f"{int(round(info['val']))}", 
                            color=info['color'], fontsize=9, fontweight='bold',
                            ha='center', va='center', zorder=5,
                            bbox=dict(facecolor='white', alpha=0.9, edgecolor=info['color'], boxstyle='round,pad=0.3'))

            ax.set_title(f"{feat} (Thick: {title})", fontsize=12, fontweight='bold', pad=10)
            ax.set_ylabel("Count", fontsize=10)
            ax.set_xlabel("")
            
            if is_right:
                legend_elements = [Patch(facecolor=color_map[k], edgecolor='white', label=k.replace('數',''), alpha=0.5) for k in color_map]
                ax.legend(handles=legend_elements, title="Grade", title_fontsize='10', fontsize='9', 
                          bbox_to_anchor=(1.02, 1), loc='upper left', borderaxespad=0.)

        st.subheader("🌐 Factory Overall Distribution (Standard Combined View)")
        overall_y = get_shared_y(df, ['YS', 'TS', 'EL', 'YPE'])
        ov_cols = st.columns(2)
        for idx, feat in enumerate(['YS', 'TS', 'EL', 'YPE']):
            if feat in mech_features:
                with ov_cols[idx % 2]:
                    fig_ov, ax_ov = plt.subplots(figsize=(10, 5))
                    plot_qc_dist(ax_ov, df, feat, "Overall", custom_y_limit=overall_y, is_right=(idx % 2 != 0))
                    st.pyplot(fig_ov)
                    fig_ov.savefig(f"overall_{feat}.png", bbox_inches='tight')

        st.markdown("---")
        
        st.subheader("🔍 Detailed Distribution per Thickness Category")
        for thick in thickness_list:
            df_thick = df[df['厚度歸類'] == thick]
            st.markdown(f"### 📏 Category: **{thick}**")
            local_y = get_shared_y(df_thick, ['YS', 'TS', 'EL', 'YPE']) 
            cols_dist = st.columns(2)
            for idx, feat in enumerate(['YS', 'TS', 'EL', 'YPE']):
                if feat in mech_features:
                    with cols_dist[idx % 2]:
                        fig, ax = plt.subplots(figsize=(10, 5))
                        plot_qc_dist(ax, df_thick, feat, thick, custom_y_limit=local_y, is_right=(idx % 2 != 0))
                        st.pyplot(fig)
                        fig.savefig(f"dist_{feat}_{thick}.png", bbox_inches='tight')
            st.markdown("---")

# --- TAB 3: OPTIMIZATION & I-MR CHARTS (MULTI-METHOD & DYNAMIC PARAMS) ---
    with tab3:
        st.header("3. Production Control Limits & Goals (A-B & Above Focused)")
        
        # --- BẢNG ĐIỀU KHIỂN THÔNG SỐ (DYNAMIC CONTROLS) ---
        st.markdown("##### ⚙️ Parameter Configuration")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            sigma_release = st.number_input("Release Range (Sigma)", min_value=1.0, max_value=6.0, value=2.0, step=0.1)
        with col2:
            sigma_mill = st.number_input("Mill Range (Sigma)", min_value=0.5, max_value=4.0, value=1.0, step=0.1)
        with col3:
            iqr_k = st.number_input("IQR Filter Factor (k)", min_value=1.0, max_value=4.0, value=1.5, step=0.1)
        with col4:
            chart_method = st.radio("I-MR Chart Limits Based On:", ["Standard Method", "IQR Filtered Method"])

        spec_limits = {"YS": (405, 500), "TS": (415, 550), "EL": (25, None), "YPE": (4, None)}
        good_cols = [c for c in ['A-B+數', 'A-B數'] if c in df.columns]

        # --- HÀM TÍNH TOÁN DÙNG CHUNG (STANDARD & IQR) ---
        def calculate_stats(v_arr, w_arr, k_factor):
            # 1. Standard Method
            m_std = np.average(v_arr, weights=w_arr)
            s_std = np.sqrt(np.average((v_arr - m_std)**2, weights=w_arr))
            
            # 2. IQR Filtered Method
            try:
                expanded_v = np.repeat(v_arr, w_arr.astype(int))
                q1 = np.percentile(expanded_v, 25)
                q3 = np.percentile(expanded_v, 75)
                iqr = q3 - q1
                
                lower_iqr = q1 - k_factor * iqr
                upper_iqr = q3 + k_factor * iqr
                
                mask = (v_arr >= lower_iqr) & (v_arr <= upper_iqr)
                vf, wf = v_arr[mask], w_arr[mask]
                
                if len(vf) > 0 and sum(wf) > 0:
                    m_iqr = np.average(vf, weights=wf)
                    s_iqr = np.sqrt(np.average((vf - m_iqr)**2, weights=wf))
                else:
                    m_iqr, s_iqr = m_std, s_std
            except:
                m_iqr, s_iqr = m_std, s_std
                
            return (m_std, s_std), (m_iqr, s_iqr)

        # 1. OVERALL FACTORY
        st.markdown("---")
        st.subheader("🌐 Overall Factory Performance Goals")
        
        overall_export_data = [] 
        total_n_overall = df[count_cols].sum().sum()
        seg_dist_overall = "N/A" if total_n_overall == 0 else ", ".join([f"{k.replace('數','')}:{int(round(df[k].sum()/total_n_overall*100))}%" for k in count_cols])

        for feat in mech_features:
            if good_cols:
                df_ov = df[[feat] + good_cols].dropna(subset=[feat]).copy()
                df_ov['Good_Qty'] = df_ov[good_cols].sum(axis=1)
                df_ov = df_ov[df_ov['Good_Qty'] > 0]
                
                low, high = spec_limits.get(feat, (None, None))
                spec_str_ov = f"{int(low)}-{int(high)}" if low and high else (f">={int(low)}" if low else "N/A")

                if not df_ov.empty:
                    v, w = df_ov[feat].values, df_ov['Good_Qty'].values
                    (m_std, s_std), (m_iqr, s_iqr) = calculate_stats(v, w, iqr_k)
                    
                    methods_data = [
                        ("Standard", m_std, s_std),
                        (f"IQR (k={iqr_k})", m_iqr, s_iqr)
                    ]
                    
                    for method_name, m_val, s_val in methods_data:
                        mill_lower = max(0, int(round(m_val - sigma_mill * s_val)))
                        release_lower = max(0, int(round(m_val - sigma_release * s_val)))
                        
                        overall_export_data.append({
                            "Feature": feat, 
                            "Method": method_name,
                            "Current Limit (2025/12)": spec_str_ov, 
                            "Segment Distribution": seg_dist_overall,
                            "TARGET GOAL": int(round(m_val)),
                            "TOLERANCE": int(round(s_val)),
                            f"MILL RANGE {sigma_mill}σ": f"{mill_lower}-{int(round(m_val + sigma_mill*s_val))}",
                            f"RELEASE RANGE {sigma_release}σ": f"{release_lower}-{int(round(m_val + sigma_release*s_val))}"
                        })
        
        st.dataframe(pd.DataFrame(overall_export_data), use_container_width=True, hide_index=True)

        st.markdown("---")
        
        # 2. LOCAL THICKNESS
        st.subheader("🔍 Local Control Limits & I-MR Trending")
        plot_data_dict = {}

        for thick in thickness_list:
            st.markdown(f"#### 📏 Thickness Category: **{thick}**")
            df_t = df[df['厚度歸類'] == thick]
            plot_data_dict[thick] = {}
            thick_status = []
            
            for feat in mech_features:
                temp_calc = df_t[[feat] + good_cols].dropna(subset=[feat]).copy() if good_cols else pd.DataFrame()
                if not temp_calc.empty:
                    temp_calc['Good_Qty'] = temp_calc[good_cols].sum(axis=1)
                    temp_calc = temp_calc[temp_calc['Good_Qty'] > 0]
                
                low, high = spec_limits.get(feat, (None, None))
                spec_str = f"{int(low)}-{int(high)}" if low and high else (f">={int(low)}" if low else "N/A")

                if not temp_calc.empty:
                    v, w = temp_calc[feat].values, temp_calc['Good_Qty'].values
                    (m_std, s_std), (m_iqr, s_iqr) = calculate_stats(v, w, iqr_k)
                    
                    # Lưu trữ cả 2 bộ thông số để dùng cho I-MR chart tùy theo chọn lựa của user
                    plot_data_dict[thick][feat] = {
                        'values': v, 
                        'mean_std': m_std, 'std_std': s_std,
                        'mean_iqr': m_iqr, 'std_iqr': s_iqr
                    }
                    
                    total_n = df_t[count_cols].sum().sum()
                    seg_dist = "N/A" if total_n == 0 else ", ".join([f"{k.replace('數','')}:{int(round(df_t[k].sum()/total_n*100))}%" for k in count_cols])
                    
                    methods_data = [
                        ("Standard", m_std, s_std),
                        (f"IQR (k={iqr_k})", m_iqr, s_iqr)
                    ]
                    
                    for method_name, m_val, s_val in methods_data:
                        mill_lower = max(0, int(round(m_val - sigma_mill * s_val)))
                        release_lower = max(0, int(round(m_val - sigma_release * s_val)))
                        
                        row = {
                            "Feature": feat, 
                            "Method": method_name,
                            "Current Limit (2025/12)": spec_str,
                            "Segment Distribution": seg_dist,
                            "TARGET GOAL": int(round(m_val)),
                            "TOLERANCE": int(round(s_val)),
                            f"MILL RANGE {sigma_mill}σ": f"{mill_lower}-{int(round(m_val + sigma_mill*s_val))}",
                            f"RELEASE RANGE {sigma_release}σ": f"{release_lower}-{int(round(m_val + sigma_release*s_val))}"
                        }
                        thick_status.append(row)
                        
                        exp_row = row.copy()
                        exp_row['Thickness'] = thick
                        all_export_data.append(exp_row)

            st.dataframe(pd.DataFrame(thick_status), use_container_width=True, hide_index=True)
            
            # --- VẼ BIỂU ĐỒ I-MR TÙY BIẾN THEO METHOD CHỌN ---
            cols_imr = st.columns(2)
            top4 = [f for f in ['YS', 'TS', 'EL', 'YPE'] if f in plot_data_dict[thick]]
            for idx, f in enumerate(top4):
                with cols_imr[idx % 2]:
                    d = plot_data_dict[thick][f]
                    v = d['values']
                    
                    # Lựa chọn thông số Mean/Std dựa trên Radio Button
                    if chart_method == "Standard Method":
                        mv, sv = d['mean_std'], d['std_std']
                    else:
                        mv, sv = d['mean_iqr'], d['std_iqr']
                        
                    if len(v) > 1:
                        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 7), gridspec_kw={'height_ratios': [2, 1]})
                        
                        # I-MR dùng Sigma của Release Range làm giới hạn cảnh báo
                        ucl, lcl = mv + sigma_release*sv, max(0, mv - sigma_release*sv)
                        
                        # I-Chart
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
                        ax1.text(1.02, v_max, f"Max: {v_max:.1f}", color='gray', transform=trans1, va='center')
                        ax1.text(1.02, v_min, f"Min: {v_min:.1f}", color='gray', transform=trans1, va='center')

                        ax1.set_title(f"I-Chart: {f} ({chart_method})", fontsize=11, fontweight='bold')
                        ax1.set_ylabel("Value")
                        
                        # MR-Chart
                        mr = np.abs(np.diff(v))
                        mrm = np.mean(mr)
                        mru = 3.267 * mrm
                        
                        ax2.plot(mr, marker='o', color='orange', ms=4, lw=1, zorder=1)
                        mr_outs = np.where(mr > mru)[0]
                        if len(mr_outs) > 0:
                            ax2.scatter(mr_outs, mr[mr_outs], color='red', s=60, zorder=2)
                            
                        ax2.axhline(mrm, color='green', ls='--', lw=1.5)
                        ax2.axhline(mru, color='red', ls='--', lw=1.2)
                        
                        mr_max = np.max(mr) if len(mr) > 0 else 0
                        ax2.axhline(mr_max, color='gray', ls=':', lw=1, alpha=0.7)
                        
                        trans2 = ax2.get_yaxis_transform()
                        ax2.text(1.02, mrm, f"Mean: {mrm:.1f}", color='green', transform=trans2, va='center', fontweight='bold')
                        ax2.text(1.02, mru, f"UCL: {mru:.1f}", color='red', transform=trans2, va='center', fontweight='bold')
                        ax2.text(1.02, mr_max, f"Max: {mr_max:.1f}", color='gray', transform=trans2, va='center')

                        ax2.set_title("Moving Range", fontsize=10)
                        ax2.set_ylabel("Range")
                        
                        fig.tight_layout()
                        fig.subplots_adjust(right=0.85)
                        st.pyplot(fig)
                        fig.savefig(f"imr_{f}_{thick}.png", bbox_inches='tight')
            st.markdown("---")

    # --- EXPORT SECTION ---
    st.sidebar.header("📥 Export Options")
    if st.sidebar.button("Download Excel Report"):
        towrite = io.BytesIO()
        pd.DataFrame(all_export_data).to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.sidebar.download_button(label="Click to Download Excel", data=towrite, file_name="QC_Optimization_Report.xlsx")

    # --- PDF EXPORT SECTION ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("🖨️ PDF Reports")
    
    def clean(t): return str(t).replace('±', '+/-').replace('–', '-').encode('latin-1', 'ignore').decode('latin-1')

    # NÚT XUẤT BÁO CÁO TỔNG QUAN
    if st.sidebar.button("Generate OVERALL PDF (Executive)"):
        pdf_ov = FPDF(orientation='L')
        
        pdf_ov.add_page()
        pdf_ov.set_font('Arial', 'B', 16); pdf_ov.cell(0, 10, "QC MECHANICAL PROPERTIES - EXECUTIVE SUMMARY", ln=True, align="C"); pdf_ov.ln(5)
        pdf_ov.set_font('Arial', 'B', 12); pdf_ov.cell(0, 10, "1. Quality Summary", ln=True)
        pdf_ov.set_font('Arial', 'B', 8); cw = [10, 20] + [15]*len(count_cols) + [20] + [15]*len(count_cols)
        for i, col in enumerate(display_df.columns): pdf_ov.cell(cw[i] if i < len(cw) else 20, 8, clean(col), border=1, align='C')
        pdf_ov.ln(); pdf_ov.set_font('Arial', '', 8)
        for _, r in display_df.iterrows():
            for i, v in enumerate(r): pdf_ov.cell(cw[i] if i < len(cw) else 20, 8, clean(v), border=1, align='C')
            pdf_ov.ln()

        pdf_ov.add_page(); pdf_ov.set_font('Arial', 'B', 12); pdf_ov.cell(0, 10, "2. Factory Overall Distribution", ln=True); ys = pdf_ov.get_y()
        for idx, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
            path = f"overall_{f}.png"
            if os.path.exists(path): pdf_ov.image(path, x=(10 if idx%2==0 else 150), y=(ys if idx<2 else ys+75), w=135)

        pdf_ov.add_page(); pdf_ov.set_font('Arial', 'B', 12); pdf_ov.cell(0, 10, "3. Overall Factory Performance Goals", ln=True)
        
        # Cập nhật PDF có thêm cột "Method" và Dynamic Sigma titles
        heads = ["Feature", "Method", "Limit (25/12)", "Segment Dist", "TARGET", "TOL", f"MILL {sigma_mill}σ", f"RELEASE {sigma_release}σ"]
        # Phân bổ lại độ rộng cột sao cho tổng = ~270mm (A4 Landscape usable space)
        c_w3 = [16, 22, 24, 60, 15, 12, 28, 30] 
        
        pdf_ov.set_font('Arial', 'B', 8)
        for i, h in enumerate(heads): pdf_ov.cell(c_w3[i], 7, clean(h), border=1, align='C')
        pdf_ov.ln(); pdf_ov.set_font('Arial', '', 7)
        
        for row in overall_export_data:
            v_list = [row["Feature"], row["Method"], row["Current Limit (2025/12)"], row["Segment Distribution"], str(row["TARGET GOAL"]), str(row["TOLERANCE"]), row[f"MILL RANGE {sigma_mill}σ"], row[f"RELEASE RANGE {sigma_release}σ"]]
            for i, v in enumerate(v_list): pdf_ov.cell(c_w3[i], 7, clean(v), border=1, align='C')
            pdf_ov.ln()

        pdf_ov.output("QC_Overall_Report.pdf")
        with open("QC_Overall_Report.pdf", "rb") as f:
            st.sidebar.download_button("📥 Download OVERALL PDF", f.read(), "QC_Overall_Report.pdf", "application/pdf")

    # NÚT XUẤT BÁO CÁO ĐẦY ĐỦ
    if st.sidebar.button("Generate FULL PDF (Detailed)"):
        pdf = FPDF(orientation='L')
        
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16); pdf.cell(0, 10, "QC MECHANICAL PROPERTIES - FULL REPORT", ln=True, align="C"); pdf.ln(5)
        pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, "1. Quality Summary", ln=True)
        pdf.set_font('Arial', 'B', 8); cw = [10, 20] + [15]*len(count_cols) + [20] + [15]*len(count_cols)
        for i, col in enumerate(display_df.columns): pdf.cell(cw[i] if i < len(cw) else 20, 8, clean(col), border=1, align='C')
        pdf.ln(); pdf.set_font('Arial', '', 8)
        for _, r in display_df.iterrows():
            for i, v in enumerate(r): pdf.cell(cw[i] if i < len(cw) else 20, 8, clean(v), border=1, align='C')
            pdf.ln()

        pdf.add_page(); pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, "2. Factory Overall Distribution", ln=True); ys = pdf.get_y()
        for idx, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
            path = f"overall_{f}.png"
            if os.path.exists(path): pdf.image(path, x=(10 if idx%2==0 else 150), y=(ys if idx<2 else ys+75), w=135)

        for thick in thickness_list:
            pdf.add_page(); pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, f"3. Distribution - Thick: {thick}", ln=True); ys = pdf.get_y()
            for idx, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
                path = f"dist_{f}_{thick}.png"
                if os.path.exists(path): pdf.image(path, x=(10 if idx%2==0 else 150), y=(ys if idx<2 else ys+75), w=135)
            
            pdf.add_page(); pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, f"4. Control Limits - Thick: {thick}", ln=True)
            
            heads = ["Feature", "Method", "Limit (25/12)", "Segment Dist", "TARGET", "TOL", f"MILL {sigma_mill}σ", f"RELEASE {sigma_release}σ"]
            c_w3 = [16, 22, 24, 60, 15, 12, 28, 30] 
            pdf.set_font('Arial', 'B', 8)
            for i, h in enumerate(heads): pdf.cell(c_w3[i], 7, clean(h), border=1, align='C')
            pdf.ln(); pdf.set_font('Arial', '', 7)
            for row in all_export_data:
                if row['Thickness'] == thick:
                    v_list = [row["Feature"], row["Method"], row["Current Limit (2025/12)"], row["Segment Distribution"], str(row["TARGET GOAL"]), str(row["TOLERANCE"]), row[f"MILL RANGE {sigma_mill}σ"], row[f"RELEASE RANGE {sigma_release}σ"]]
                    for i, v in enumerate(v_list): pdf.cell(c_w3[i], 7, clean(v), border=1, align='C')
                    pdf.ln()
            
            pdf.add_page(); pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, f"5. I-MR Charts - Thick: {thick}", ln=True)
            y_imr = pdf.get_y() + 2 
            for idx, f in enumerate(['YS', 'TS', 'EL', 'YPE']):
                path = f"imr_{f}_{thick}.png"
                if os.path.exists(path): 
                    pdf.image(path, x=(10 if idx%2==0 else 150), y=(y_imr if idx<2 else y_imr+90), w=130)

        pdf.output("QC_Full_Report.pdf")
        with open("QC_Full_Report.pdf", "rb") as f:
            st.sidebar.download_button("📥 Download FULL PDF", f.read(), "QC_Full_Report.pdf", "application/pdf")
