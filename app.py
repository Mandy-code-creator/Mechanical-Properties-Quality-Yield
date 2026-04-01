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
    tab0, tab1, tab2, tab3 = st.tabs(["0. Raw Check", "1. Yield Summary", "2. Distribution Analysis", "3.🔍 Root Cause & Diagnostic Analysis", "4. I-MR Analysis"])

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

# --- TAB 3: DIAGNOSTIC FLOW (OPTIMIZED TIMELINE) ---
    with tab3:
        st.header("🔍 Quality Diagnostic & Global Stability Flow")
        st.info("Global analysis of production defects and continuous time-series stability.")
        
        # Define Global Specs for Highlighting
        GLOBAL_SPECS = {
            'YS': {'min': 400, 'max': 460, 'target': 430},
            'TS': {'min': 410, 'max': 470, 'target': 440},
            'EL': {'min': 25, 'max': None, 'target': None},
            'YPE': {'min': 4, 'max': None, 'target': None}
        }
        
        # --- STEP 1: HEATMAP (OPTIMIZED WEIGHTED AVERAGE) ---
        defect_str = ", ".join(bad_grades)
        st.subheader(f"Step 1: Defect Hotspot Map (% Rate: {defect_str})")
        
        # Nhãn quy cách
        df_filtered['Spec_Label'] = df_filtered['Actual_Thickness'].astype(str) + "mm (" + df_filtered['HR_Material'] + ")"
        
        # THUẬT TOÁN ĐÚNG: Tính Weighted Average cho từng nhóm Spec + Time
        heat_data = df_filtered.groupby(['Spec_Label', 'Time_Group']).apply(
            lambda x: (x['Bad_Qty'].sum() / x['Total_Qty'].sum() * 100) if x['Total_Qty'].sum() > 0 else 0
        ).unstack()

        if not heat_data.empty:
            fig, ax = plt.subplots(figsize=(12, 6))
            import seaborn as sns
            
            # Đặt vmax (ví dụ 30%) để làm nổi bật sự tương phản màu sắc
            # Ô nào > 30% sẽ đỏ đậm max level
            vmax_threshold = 30.0 if heat_data.max().max() > 30 else heat_data.max().max()
            
            sns.heatmap(heat_data, annot=True, fmt=".1f", cmap="YlOrRd", linewidths=.5, vmax=vmax_threshold, ax=ax)
            ax.set_title(f"HOTSPOT MAP: ACTUAL DEFECT RATE ({defect_str})", fontsize=12, fontweight='bold', color='#d62728', pad=15)
            ax.set_ylabel("Specification (Thickness & Material)")
            ax.set_xlabel("Production Period")
            st.pyplot(fig); plt.close(fig)
            
            # --- AUTO DETECT TOP 3 HOTSPOTS ---
            st.markdown("### 🚨 Top 3 Critical Hotspots (Action Required)")
            # Biến đổi bảng heatmap thành dạng cột để dễ sort
            stacked_data = heat_data.stack().reset_index()
            stacked_data.columns = ['Spec', 'Period', 'Defect_Rate']
            top_3 = stacked_data[stacked_data['Defect_Rate'] > 0].sort_values(by='Defect_Rate', ascending=False).head(3)
            
            if not top_3.empty:
                for idx, row in top_3.iterrows():
                    # Cảnh báo màu đỏ cho lỗi > 15%, màu cam cho lỗi > 5%
                    alert_color = "🔴" if row['Defect_Rate'] > 15 else "🟠"
                    st.error(f"{alert_color} **{row['Spec']}** during **{row['Period']}**: Defect Rate hits **{row['Defect_Rate']:.1f}%**")
            else:
                st.success("✅ Quy trình đang ổn định. Không phát hiện điểm nóng.")

        # --- STEP 2: PARETO ---
        st.markdown("---")
        st.subheader("Step 2: Defect Category Breakdown (Pareto)")
        defect_sums = df_filtered[bad_grades].sum().sort_values(ascending=False)
        if defect_sums.sum() > 0:
            pareto_df = pd.DataFrame({'Count': defect_sums})
            pareto_df['Cum_Pct'] = pareto_df['Count'].cumsum() / pareto_df['Count'].sum() * 100
            
            fig, ax1 = plt.subplots(figsize=(10, 4))
            ax1.bar(pareto_df.index, pareto_df['Count'], color="#d62728", alpha=0.8)
            ax2 = ax1.twinx()
            ax2.plot(pareto_df.index, pareto_df['Cum_Pct'], color="#1f77b4", marker="D", ms=5)
            ax2.axhline(80, color="orange", linestyle="--")
            plt.title("Pareto Principle: Focus on 80% of Quality Issues")
            st.pyplot(fig); plt.close(fig)

        # --- STEP 3: GLOBAL TIMELINE I-MR CHARTS ---
        st.markdown("---")
        st.subheader("Step 3: Continuous Time-Series Tracking (2024 - 2026)")
        st.warning("X-Axis displays actual Production Dates. Red dots indicate Out-of-Spec violations in specific time periods.")

        # Lọc bỏ dữ liệu trùng lặp "2025 (Full Year)" để dải thời gian vẽ lên biểu đồ không bị lặp lại
        df_timeline = df_filtered[df_filtered['Time_Group'] != "2025 (Full Year)"].copy()
        if df_timeline.empty:
            df_timeline = df_filtered.copy()

        for feat in ['YS', 'TS', 'EL', 'YPE']:
            if feat in df_timeline.columns:
                # Sắp xếp dữ liệu theo ngày tháng sản xuất để vẽ Time-series chuẩn xác
                df_ts = df_timeline.dropna(subset=[feat, '烤三生產日期']).sort_values(by='烤三生產日期').reset_index(drop=True)
                
                if len(df_ts) > 1:
                    st.markdown(f"#### 📈 Production Timeline: **{feat}**")
                    dates = df_ts['烤三生產日期']
                    vals = df_ts[feat].values
                    mean_val = np.mean(vals)
                    
                    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 8), gridspec_kw={'height_ratios': [2, 1]})
                    
                    # --- I-CHART (VẼ THEO NGÀY THÁNG) ---
                    ax1.plot(dates, vals, marker='o', ms=3, lw=1, color='#1f77b4', alpha=0.6, label=f'Individual {feat}', zorder=1)
                    ax1.axhline(mean_val, color='green', ls='--', label='Process Mean')
                    
                    # Bắt lỗi vượt Spec
                    out_dates, out_y = [], []
                    if feat in GLOBAL_SPECS:
                        s = GLOBAL_SPECS[feat]
                        min_s, max_s = s.get('min'), s.get('max')
                        
                        for i, v in enumerate(vals):
                            if (min_s is not None and v < min_s) or (max_s is not None and v > max_s):
                                out_dates.append(dates.iloc[i])
                                out_y.append(v)

                        if min_s: ax1.axhline(min_s, color='red', ls='-', lw=2, label=f"Lower Spec ({min_s})")
                        if max_s: ax1.axhline(max_s, color='red', ls='-', lw=2, label=f"Upper Spec ({max_s})")
                        if s.get('target'): ax1.axhline(s['target'], color='black', ls=':', label=f"Target ({s['target']})")
                        if min_s and max_s: ax1.axhspan(min_s, max_s, color='green', alpha=0.05)
                    
                    # Tô đỏ chấm vi phạm theo đúng ngày
                    if out_dates:
                        ax1.scatter(out_dates, out_y, color='red', s=50, zorder=5, label='Violation')

                    # VẼ VẠCH PHÂN CÁCH CÁC NĂM
                    for year in dates.dt.year.unique():
                        year_start = pd.Timestamp(year=year, month=1, day=1)
                        if year_start >= dates.min():
                            ax1.axvline(year_start, color='black', linestyle='-.', alpha=0.5)
                            ax2.axvline(year_start, color='black', linestyle='-.', alpha=0.5)
                            ax1.text(year_start + pd.Timedelta(days=10), ax1.get_ylim()[1]*0.98, f"Start of {year}", color='black', fontweight='bold', alpha=0.7, va='top', rotation=90)

                    ax1.set_title(f"I-Chart: {feat} (Chronological Timeline)", fontweight='bold')
                    ax1.legend(loc='upper right', fontsize=8)
                    ax1.grid(True, linestyle='--', alpha=0.3)

                    # --- MR-CHART (VẼ THEO NGÀY THÁNG) ---
                    mr = np.abs(np.diff(vals))
                    mr_mean = np.mean(mr)
                    ucl_mr = 3.267 * mr_mean
                    mr_dates = dates.iloc[1:] # MR chart bị lùi 1 nhịp so với I-Chart
                    
                    ax2.plot(mr_dates, mr, marker='o', ms=3, lw=1, color='orange', alpha=0.6, zorder=1)
                    ax2.axhline(mr_mean, color='green', ls='--', label='MR Mean')
                    ax2.axhline(ucl_mr, color='red', ls=':', label=f'UCL ({ucl_mr:.1f})')
                    
                    # Tô đỏ điểm MR vượt ngưỡng
                    v_out_dates, v_out_y = [], []
                    for i, m_val in enumerate(mr):
                        if m_val > ucl_mr:
                            v_out_dates.append(mr_dates.iloc[i])
                            v_out_y.append(m_val)
                    if v_out_dates:
                        ax2.scatter(v_out_dates, v_out_y, color='red', s=40, zorder=5)

                    ax2.set_title("Moving Range (Process Variation)", fontweight='bold')
                    ax2.legend(loc='upper right', fontsize=8)
                    ax2.grid(True, linestyle='--', alpha=0.3)

                    # Tự động xoay nghiêng định dạng ngày tháng ở trục X cho dễ đọc
                    fig.autofmt_xdate()
                    fig.tight_layout()
                    st.pyplot(fig); plt.close(fig)
                    st.markdown("---")
    # --- TASK 4: I-MR CHART (TIMELINE STABILITY) ---
    with tab5:
        st.header("📈 Task 4: I-MR Stability Tracking (Chronological)")
        st.info("Analysis based on production date from 2024 to 2026. Red dots = Out of Spec.")

        # Filters for Task 5
        imr_thicks = sorted(df_filtered['Actual_Thickness'].dropna().unique())
        imr_mats = sorted(df_filtered['HR_Material'].astype(str).unique())
        
        c1, c2 = st.columns(2)
        sel_t = c1.selectbox("Filter Thickness:", imr_thicks, key="t5_t")
        sel_m = c2.selectbox("Filter Material:", imr_mats, key="t5_m")

        imr_df = df_filtered[(df_filtered['Actual_Thickness'] == sel_t) & 
                            (df_filtered['HR_Material'] == sel_m)].sort_values(by='烤三生產日期').reset_index(drop=True)

        if not imr_df.empty:
            for feat in ['YS', 'TS', 'EL', 'YPE']:
                if feat in imr_df.columns:
                    valid_data = imr_df.dropna(subset=[feat, '烤三生產日期']).copy()
                    if len(valid_data) > 1:
                        st.markdown(f"### 🛡️ Chronological Stability: **{feat}**")
                        dates, vals = valid_data['烤三生產日期'], valid_data[feat].values
                        mean_v = np.mean(vals)
                        
                        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 8), gridspec_kw={'height_ratios': [2, 1]})
                        
                        # I-Chart
                        ax1.plot(dates, vals, marker='o', ms=4, lw=1, color='#1f77b4', alpha=0.6, label=feat)
                        ax1.axhline(mean_v, color='green', ls='--', label='Mean')
                        
                        # Add Global Specs & Highlights
                        if feat in GLOBAL_SPECS:
                            s = GLOBAL_SPECS[feat]
                            if s['min']: ax1.axhline(s['min'], color='red', lw=2, label=f"Min ({s['min']})")
                            if s['max']: ax1.axhline(s['max'], color='red', lw=2, label=f"Max ({s['max']})")
                            if s['target']: ax1.axhline(s['target'], color='black', ls=':', label='Target')
                            
                            # Violation Detection
                            v_x, v_y = [], []
                            for i, v in enumerate(vals):
                                if (s['min'] and v < s['min']) or (s['max'] and v > s['max']):
                                    v_x.append(dates.iloc[i]); v_y.append(v)
                            if v_x: ax1.scatter(v_x, v_y, color='red', s=60, zorder=5, label='Violation')
                        
                        # Year Separation Lines
                        for year in dates.dt.year.unique():
                            y_start = pd.Timestamp(year=year, month=1, day=1)
                            if dates.min() <= y_start <= dates.max():
                                ax1.axvline(y_start, color='black', ls='-.', alpha=0.3)
                                ax1.text(y_start, ax1.get_ylim()[1], f" {year}", fontsize=10, va='top')

                        ax1.set_title(f"Individual Chart (I) - {feat}", fontweight='bold')
                        ax1.legend(loc='upper right', fontsize=8)

                        # MR-Chart
                        mr = np.abs(np.diff(vals))
                        mr_mean = np.mean(mr)
                        ucl_mr = 3.267 * mr_mean
                        ax2.plot(dates.iloc[1:], mr, marker='o', ms=3, color='orange', alpha=0.6)
                        ax2.axhline(mr_mean, color='green', ls='--')
                        ax2.axhline(ucl_mr, color='red', ls=':', label=f'UCL ({ucl_mr:.1f})')
                        
                        # High variation points
                        hv_x, hv_y = [], []
                        for i, m_val in enumerate(mr):
                            if m_val > ucl_mr: hv_x.append(dates.iloc[i+1]); hv_y.append(m_val)
                        if hv_x: ax2.scatter(hv_x, hv_y, color='red', s=40, zorder=5)

                        ax2.set_title("Moving Range Chart (MR)", fontweight='bold')
                        fig.autofmt_xdate()
                        st.pyplot(fig); plt.close(fig)
                        
                        # Insight
                        if v_x:
                            st.error(f"⚠️ **Stability Insight:** Found {len(v_x)} out-of-spec points for {feat}. Check production logs around these dates.")
                        else:
                            st.success(f"✅ **Stability Insight:** {feat} is within specification limits for this period.")
        else:
            st.warning("No data found for the selected combination.")                
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
