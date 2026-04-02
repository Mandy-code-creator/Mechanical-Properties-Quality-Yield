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
# 👇👇👇 CHÈN ĐOẠN NÀY VÀO NGAY ĐÂY 👇👇👇
GLOBAL_SPECS = {
    'YS': {'min': 400, 'max': 460, 'target': 430},
    'TS': {'min': 410, 'max': 470, 'target': 440},
    'EL': {'min': 25, 'max': None, 'target': None},
    'YPE': {'min': 4, 'max': None, 'target': None}
}
# 👆👆👆 ============================== 👆👆👆
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
    tab0, tab1, tab2, tab3, tab4 = st.tabs(["0. Raw Check", "1. Yield Summary", "2. Distribution Analysis", "3.🔍 Root Cause & Diagnostic Analysis", "4. I-MR Analysis"])

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

   # --- TAB 1: YIELD SUMMARY ---
    with tab1:
        st.header("1. Quality Yield Summary & Worst Offenders")
        st.info("Overview of production yield. Chronologically sorted from 2024 onwards.")

        # --- EXECUTIVE SUMMARY: CHỈ TÍNH LỖI NẶNG (B+ TRỞ XUỐNG) ---
        st.subheader("📊 Executive Summary: Production Quality Timeline")
        
        # Tạo cột tính riêng tổng số lượng lỗi nặng
        severe_grades = ['B+', 'B']
        df_filtered['Severe_Bad_Qty'] = df_filtered[severe_grades].sum(axis=1)
        
        # NÂNG CẤP: TÍNH HÀNG ĐẠT BẰNG CÁCH LẤY TỔNG TRỪ ĐI LỖI NẶNG (Bao gồm cả A-B-)
        df_filtered['Acceptable_Qty'] = df_filtered['Total_Qty'] - df_filtered['Severe_Bad_Qty']
        
        # Gom nhóm theo Thời gian, Độ dày và Vật liệu
        period_summary = df_filtered.groupby(['Time_Group', 'Actual_Thickness', 'HR_Material'])[['Total_Qty', 'Acceptable_Qty', 'Severe_Bad_Qty']].sum().reset_index()
        
        # Lọc bỏ các dòng không có dữ liệu sản xuất
        period_summary = period_summary[period_summary['Total_Qty'] > 0]
        
        # Tính tỷ lệ % chuẩn xác (Đảm bảo Yield + Defect Rate = 100%)
        period_summary['Yield (%)'] = (period_summary['Acceptable_Qty'] / period_summary['Total_Qty'] * 100).fillna(0).round(2)
        period_summary['Defect_Rate (%)'] = (period_summary['Severe_Bad_Qty'] / period_summary['Total_Qty'] * 100).fillna(0).round(2)
        
        # --- BƯỚC QUAN TRỌNG: SẮP XẾP THEO THỨ TỰ THỜI GIAN (CHRONOLOGICAL) ---
        time_order_map = {
            "2024 (Full Year)": 1,
            "2025 H1 (Until 06/28)": 2,
            "2025 Q3 (06/29 - 09/30)": 3,
            "2025 Q4": 4,
            "2025 (Full Year)": 5,
            "Unknown": 99
        }
        # Tạo cột ảo để sort rồi xóa đi
        period_summary['Sort_Key'] = period_summary['Time_Group'].map(time_order_map).fillna(90)
        period_summary = period_summary.sort_values(by=['Sort_Key', 'Actual_Thickness'], ascending=[True, True])
        period_summary = period_summary.drop(columns=['Sort_Key'])
        
        # Đổi tên cột cho rõ nghĩa trên bảng hiển thị
        period_summary.rename(columns={'Severe_Bad_Qty': 'Bad_Qty (B+, B)'}, inplace=True)

        if not period_summary.empty:
            # Hiển thị bảng với format loại bỏ số 0 thập phân dư thừa
            st.dataframe(
                period_summary.style.background_gradient(subset=['Defect_Rate (%)'], cmap='Reds')
                                    .background_gradient(subset=['Yield (%)'], cmap='Greens')
                                    .format({
                                        'Actual_Thickness': '{:.2f}', 
                                        'Total_Qty': '{:.0f}',        
                                        'Acceptable_Qty': '{:.0f}',         
                                        'Bad_Qty (B+, B)': '{:.0f}',          
                                        'Yield (%)': '{:.2f}%', 
                                        'Defect_Rate (%)': '{:.2f}%'
                                    }),
                use_container_width=True, 
                hide_index=True
            )
        else:
            st.success("✅ No production data available for the selected filters.")

        st.markdown("---")

        # --- DETAILED YIELD BY PERIOD ---
        st.subheader("📑 Detailed Yield by Period (All Grades)")
        sum_df = df_filtered.groupby(['Time_Group', 'Actual_Thickness', 'HR_Material'], dropna=False)[base_grades].sum().reset_index()
        sum_df['Total_Qty'] = sum_df[base_grades].sum(axis=1)
        
        for col in base_grades: 
            sum_df[f"% {col}"] = ((sum_df[col] / sum_df['Total_Qty'].replace(0, np.nan)) * 100).fillna(0).round(1)
        
        # Sắp xếp chi tiết cũng theo thời gian luôn cho đồng bộ
        sum_df['Sort_Key'] = sum_df['Time_Group'].map(time_order_map).fillna(90)
        sum_df = sum_df.sort_values(by=['Sort_Key', 'Actual_Thickness']).drop(columns=['Sort_Key'])
        
        sum_df.rename(columns={'Time_Group': 'Period', 'Actual_Thickness': 'Thickness'}, inplace=True)
        
        # Lấy danh sách Period theo đúng thứ tự thời gian để render bảng
        ordered_periods = sorted(sum_df['Period'].unique(), key=lambda x: time_order_map.get(x, 99))
        
        for period in ordered_periods:
            p_data = sum_df[sum_df['Period'] == period]
            if not p_data.empty:
                st.markdown(f"#### 📅 Period: **{period}**")
                
                # Format dọn dẹp số 0 cho bảng chi tiết
                format_dict = {'Thickness': '{:.2f}', 'Total_Qty': '{:.0f}'}
                for col in base_grades:
                    format_dict[col] = '{:.0f}'
                    format_dict[f"% {col}"] = '{:.1f}%'
                
                st.dataframe(
                    p_data.drop(columns=['Period']).style.format(format_dict), 
                    use_container_width=True, 
                    hide_index=True
                )

        # --- XUẤT EXCEL CÓ MÀU (FORMATTED EXCEL) ---
        st.markdown("---")
        output = io.BytesIO()
        
        # Sử dụng engine xlsxwriter để có thể format màu sắc
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            period_summary.to_excel(writer, index=False, sheet_name='Executive_Summary')
            sum_df.to_excel(writer, index=False, sheet_name='Detailed_Yield')
            
            workbook  = writer.book
            worksheet = writer.sheets['Executive_Summary']

            # 1. Định nghĩa các format
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
            num_format = workbook.add_format({'align': 'center', 'border': 1})
            pct_format = workbook.add_format({'num_format': '0.00"%"', 'align': 'center', 'border': 1})

            # 2. Format Header và Border cho toàn bảng
            for col_num, value in enumerate(period_summary.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 15) # Độ rộng cột

            # 3. Apply Conditional Formatting cho cột Yield (Cột G - Index 6)
            # Màu xanh: Càng cao càng xanh
            worksheet.conditional_format(1, 6, len(period_summary), 6, {
                'type': '2_color_scale',
                'min_color': "#F7FCF5", # Xanh nhạt
                'max_color': "#41AB5D"  # Xanh đậm
            })

            # 4. Apply Conditional Formatting cho cột Defect Rate (Cột H - Index 7)
            # Màu đỏ: Càng cao càng đỏ
            worksheet.conditional_format(1, 7, len(period_summary), 7, {
                'type': '2_color_scale',
                'min_color': "#FFF5F0", # Đỏ nhạt
                'max_color': "#EF3B2C"  # Đỏ đậm
            })

            # Format định dạng số cho toàn bộ nội dung
            for row in range(1, len(period_summary) + 1):
                for col in range(len(period_summary.columns)):
                    if col >= 6: # Các cột %
                        worksheet.write(row, col, period_summary.iloc[row-1, col]/100 if isinstance(period_summary.iloc[row-1, col], (int, float)) else period_summary.iloc[row-1, col], pct_format)
                    else:
                        worksheet.write(row, col, period_summary.iloc[row-1, col], num_format)

        st.download_button(
            label="📥 Download Formatted Excel (With Colors)",
            data=output.getvalue(),
            file_name="Quality_Timeline_Report_Colored.xlsx",
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

# --- TAB 3: EXECUTIVE AUTO-INSIGHT & ROOT CAUSE ---
    with tab3:
        st.header("🧠 Executive Auto-Insight & Root Cause")
        st.info("Automated diagnostic engine: Quantifying impact and recommending actions based on severe defects (B+, B).")

        # --- CHUẨN BỊ DỮ LIỆU LỖI NẶNG ---
        severe_grades = ['B+', 'B']
        df_filtered['Severe_Bad_Qty'] = df_filtered[severe_grades].sum(axis=1)
        df_filtered['Spec_Label'] = df_filtered['Actual_Thickness'].astype(str) + "mm (" + df_filtered['HR_Material'] + ")"

        # --- BƯỚC 1: XÁC ĐỊNH TOP PROBLEM (AUTO HIGHLIGHT) ---
        heat_data = df_filtered.groupby(['Spec_Label', 'Time_Group']).apply(
            lambda x: (x['Severe_Bad_Qty'].sum() / x['Total_Qty'].sum() * 100) if x['Total_Qty'].sum() > 0 else 0
        )
        
        if not heat_data.empty and heat_data.max() > 0:
            heatmap_long = heat_data.reset_index()
            heatmap_long.columns = ['Spec', 'Period', 'Defect_Rate']
            top_issues = heatmap_long[heatmap_long['Defect_Rate'] > 0].sort_values('Defect_Rate', ascending=False).head(5)

            # --- BƯỚC 2: QUANTIFY IMPACT (ROOT CAUSE SCORE) ---
            # Tập trung phân tích vào TOP 3 SEGMENTS bị lỗi nặng nhất
            top_3_specs = top_issues.head(3)['Spec'].tolist()
            top_3_periods = top_issues.head(3)['Period'].tolist()
            
            # Lọc dữ liệu thuộc về Top 3 lỗi
            df_top3 = df_filtered[
                (df_filtered['Spec_Label'].isin(top_3_specs)) & 
                (df_filtered['Time_Group'].isin(top_3_periods))
            ]

            rc_results = {}
            for f in ['YS', 'TS', 'EL', 'YPE']:
                if f in df_top3.columns:
                    # Tính Mean của hàng Tốt trong toàn bộ data (để làm chuẩn - Benchmark)
                    good_mean_global = df_filtered[df_filtered['Good_Qty'] > 0][f].mean()
                    
                    # Tính Mean của hàng Lỗi Nặng chỉ trong Top 3 vùng Hotspot
                    bad_mean_top3 = df_top3[df_top3['Severe_Bad_Qty'] > 0][f].mean()
                    
                    if pd.notnull(good_mean_global) and pd.notnull(bad_mean_top3):
                        rc_results[f] = bad_mean_top3 - good_mean_global

            rc_s = pd.Series(rc_results).dropna().sort_values(key=abs, ascending=False)

            # ... (Phần code hiển thị Conclusion giữ nguyên hoặc cập nhật tiêu đề cho rõ nghĩa) ...

            with col2:
                st.warning("🧠 ROOT CAUSE DRIVER (Top 3 Hotspots)")
                st.info("Analysis: Mean difference of Bad Coils (in Top 3 problem areas) vs. Global Good Coils.")
                
                rc_df = rc_s.reset_index()
                rc_df.columns = ['Mechanical Feature', 'Impact Gap (Top 3 Bad vs Good)']
                
                def color_gap(val):
                    color = '#d62728' if abs(val) > 5 else ('#ff7f0e' if abs(val) > 0 else 'black')
                    return f'color: {color}; font-weight: bold'

                st.dataframe(
                    rc_df.style.map(color_gap, subset=['Impact Gap (Top 3 Bad vs Good)'])
                               .format({'Impact Gap (Top 3 Bad vs Good)': '{:+.2f}'}), 
                    use_container_width=True, hide_index=True
                )

            # --- BƯỚC 4: AUTO CONCLUSION ---
            if not top_issues.empty and not rc_s.empty:
                top_issue = top_issues.iloc[0]
                top_driver = rc_s.index[0]
                gap_val = rc_s.iloc[0]
                direction = "HIGHER ⬆️" if gap_val > 0 else "LOWER ⬇️"

                st.success(f"""
                ### 🎯 EXECUTIVE CONCLUSION & ACTION PLAN:
                
                * 🚨 **Biggest Hotspot:** Specification **{top_issue['Spec']}** during **{top_issue['Period']}** (Severe Defect Rate hits **{top_issue['Defect_Rate']:.1f}%**).
                * 🧠 **Main Root Cause Driver:** **{top_driver}** is the primary culprit causing these severe defects.
                * 📊 **Quantified Impact:** Defective coils have a {top_driver} that is on average **{abs(gap_val):.1f} {direction}** than good coils.
                * 🛠️ **Recommended Action:** Immediate parameter adjustment and SPC audit required for **{top_driver}** control limits on the {top_issue['Spec']} line.
                """)

            # --- BỐ CỤC HIỂN THỊ CHI TIẾT THEO CỘT ---
            st.markdown("---")
            col1, col2 = st.columns(2)
            
            with col1:
                st.error("🔥 TOP 5 PROBLEM SEGMENTS (Where to fix)")
                st.dataframe(
                    top_issues.style.background_gradient(subset=['Defect_Rate'], cmap='Reds')
                                    .format({'Defect_Rate': '{:.1f}%'}), 
                    use_container_width=True, hide_index=True
                )

            with col2:
                st.warning("🧠 ROOT CAUSE DRIVER (What to fix)")
                rc_df = rc_s.reset_index()
                rc_df.columns = ['Mechanical Feature', 'Impact Gap (Bad vs Good)']
                
                def color_gap(val):
                    color = '#d62728' if abs(val) > 5 else ('#ff7f0e' if abs(val) > 0 else 'black')
                    return f'color: {color}; font-weight: bold'

                st.dataframe(
                    rc_df.style.map(color_gap, subset=['Impact Gap (Bad vs Good)'])
                               .format({'Impact Gap (Bad vs Good)': '{:+.2f}'}), 
                    use_container_width=True, hide_index=True
                )

            # --- BƯỚC 3: DRILL DOWN ---
            if not rc_s.empty:
                st.markdown("---")
                top_driver = rc_s.index[0]
                st.info(f"📏 DRILL DOWN: {top_driver} Shift by Thickness (Isolating the issue)")
                
                drill_data = []
                for th in df_filtered['Actual_Thickness'].dropna().unique():
                    th_df = df_filtered[df_filtered['Actual_Thickness'] == th]
                    g_val = th_df[th_df['Good_Qty'] > 0][top_driver].mean()
                    b_val = th_df[th_df['Severe_Bad_Qty'] > 0][top_driver].mean()
                    
                    if pd.notnull(g_val) or pd.notnull(b_val):
                        drill_data.append({
                            'Thickness': th,
                            'GOOD Coils (Mean)': g_val,
                            'BAD Coils (Mean)': b_val,
                            'Impact Gap': (b_val - g_val) if (pd.notnull(g_val) and pd.notnull(b_val)) else None
                        })
                        
                drill_df = pd.DataFrame(drill_data).sort_values('Impact Gap', key=abs, ascending=False)
                st.dataframe(
                    drill_df.style.map(color_gap, subset=['Impact Gap'])
                                  .format({
                                      'Thickness': '{:.2f}mm', 
                                      'GOOD Coils (Mean)': '{:.1f}', 
                                      'BAD Coils (Mean)': '{:.1f}', 
                                      'Impact Gap': '{:+.1f}'
                                  }), 
                    use_container_width=True, hide_index=True
                )
                
            # Đẩy cái Heatmap hình ảnh xuống dưới cùng để làm "Bằng chứng" (Evidence)
            st.markdown("---")
            st.subheader("🗺️ Evidence: Visual Hotspot Map (Grades B+ and Below)")
            heat_pivot = heat_data.unstack()
            fig, ax = plt.subplots(figsize=(12, 5))
            vmax_threshold = 30.0 if heat_pivot.max().max() > 30 else heat_pivot.max().max()
            import seaborn as sns
            sns.heatmap(heat_pivot, annot=True, fmt=".1f", cmap="YlOrRd", linewidths=.5, vmax=vmax_threshold, ax=ax)
            ax.set_title("SEVERE DEFECT RATE (%)", fontweight='bold', color='#d62728')
            ax.set_ylabel(""); ax.set_xlabel("")
            fig.tight_layout()
            plt.savefig("export_heatmap.png", bbox_inches='tight', dpi=150)
            st.pyplot(fig); plt.close(fig)

        else:
            st.success("✅ Process is completely stable. No severe defect patterns detected to analyze.")

        # =====================================================================
        # TÍCH HỢP BIỂU ĐỒ I-MR CHO TOÀN BỘ DỮ LIỆU LỖI NẶNG (2024-2025) VÀO TAB 3
        # =====================================================================
        st.markdown("---")
        st.header("📈 Global I-MR Stability Tracking (Severe Defects: B+ and Below)")
        st.info("Chronological view of all severe defects across the entire dataset to identify global trends.")

        # Lọc lấy toàn bộ hàng bị lỗi B+ hoặc B
        df_severe_global = df_filtered[(df_filtered['B+'] > 0) | (df_filtered['B'] > 0)].sort_values(by='烤三生產日期').reset_index(drop=True)

        if not df_severe_global.empty:
            for feat in ['YS', 'TS', 'EL', 'YPE']:
                if feat in df_severe_global.columns:
                    valid_data = df_severe_global.dropna(subset=[feat, '烤三生產日期']).reset_index(drop=True)
                    if len(valid_data) > 1:
                        st.markdown(f"#### 🛡️ Global Stability: **{feat}** (Grades B+ and Below)")
                        
                        dates = valid_data['烤三生產日期']
                        vals = valid_data[feat].values
                        
                        # Dùng mảng số thứ tự (0, 1, 2...) làm trục X
                        x_seq = np.arange(len(vals)) 
                        mean_v = np.mean(vals)
                        
                        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 8), gridspec_kw={'height_ratios': [2, 1]})
                        
                        # --- I-Chart ---
                        ax1.plot(x_seq, vals, marker='o', ms=4, lw=1, color='#1f77b4', alpha=0.6, label=f"Value ({feat})")
                        
                        # Hiển thị số Mean
                        ax1.axhline(mean_v, color='green', ls='--', label=f'Mean: {mean_v:.1f}')
                        ax1.text(x_seq[-1], mean_v, f" Mean: {mean_v:.1f}", va='bottom', color='green', fontweight='bold')
                        
                        v_x, v_y = [], []
                        if feat in GLOBAL_SPECS:
                            s = GLOBAL_SPECS[feat]
                            
                            # Hiển thị số Min (LSL)
                            if s['min']: 
                                ax1.axhline(s['min'], color='red', lw=2, label=f"Min: {s['min']}")
                                ax1.text(x_seq[-1], s['min'], f" Min: {s['min']}", va='bottom', color='red', fontweight='bold')
                            
                            # Hiển thị số Max (USL)
                            if s['max']: 
                                ax1.axhline(s['max'], color='red', lw=2, label=f"Max: {s['max']}")
                                ax1.text(x_seq[-1], s['max'], f" Max: {s['max']}", va='bottom', color='red', fontweight='bold')
                            
                            for i, v in enumerate(vals):
                                if (s['min'] and v < s['min']) or (s['max'] and v > s['max']):
                                    v_x.append(i); v_y.append(v)
                            if v_x: ax1.scatter(v_x, v_y, color='red', s=60, zorder=5, label='Out of Spec')
                        
                        # Vẽ vạch phân tách năm mới
                        for i in range(1, len(dates)):
                            if dates.iloc[i].year != dates.iloc[i-1].year:
                                ax1.axvline(i, color='black', ls='-.', alpha=0.3)
                                ax1.text(i, ax1.get_ylim()[1], f" {dates.iloc[i].year}", fontsize=10, va='top')

                        ax1.set_title(f"Individual Chart (I) - {feat} (Severe Defects: B+, B)", fontweight='bold')
                        # Hiển thị Legend ở góc phải bên ngoài để không che biểu đồ
                        ax1.legend(loc='upper right', fontsize=9, bbox_to_anchor=(1.15, 1))
                        ax1.set_xticks([]) 

                        # --- MR-Chart ---
                        mr = np.abs(np.diff(vals))
                        mr_mean = np.mean(mr)
                        ucl_mr = 3.267 * mr_mean
                        
                        ax2.plot(x_seq[1:], mr, marker='o', ms=3, color='orange', alpha=0.6, label="Moving Range")
                        
                        # Hiển thị số MR Mean và UCL
                        ax2.axhline(mr_mean, color='green', ls='--', label=f'MR Mean: {mr_mean:.1f}')
                        ax2.text(x_seq[-1], mr_mean, f" Mean: {mr_mean:.1f}", va='bottom', color='green', fontweight='bold')
                        
                        ax2.axhline(ucl_mr, color='red', ls=':', label=f'UCL: {ucl_mr:.1f}')
                        ax2.text(x_seq[-1], ucl_mr, f" UCL: {ucl_mr:.1f}", va='bottom', color='red', fontweight='bold')
                        
                        hv_x, hv_y = [], []
                        for i, m_val in enumerate(mr):
                            if m_val > ucl_mr: hv_x.append(i+1); hv_y.append(m_val)
                        if hv_x: ax2.scatter(hv_x, hv_y, color='red', s=40, zorder=5)

                        ax2.set_title(f"Moving Range Chart (MR) - {feat} (Severe Defects: B+, B)", fontweight='bold')
                        ax2.legend(loc='upper right', fontsize=9, bbox_to_anchor=(1.15, 1))
                        
                        # Trang trí lại trục X bằng ngày tháng
                        step = max(1, len(x_seq) // 12) 
                        ax2.set_xticks(x_seq[::step])
                        ax2.set_xticklabels(dates.dt.strftime('%Y-%m-%d').iloc[::step], rotation=45, ha='right')

                        fig.tight_layout()
                        
                        # Lưu ảnh I-MR toàn cục (tên file riêng để không đụng hàng với tab 4)
                        plt.savefig(f"export_imr_global_{feat}.png", bbox_inches='tight', dpi=150)
                        st.pyplot(fig); plt.close(fig)
                        
                        if v_x:
                            st.error(f"⚠️ **Global Insight:** Found {len(v_x)} out-of-spec points for {feat} across the entire period.")
                        else:
                            st.success(f"✅ **Global Insight:** {feat} is globally within specification limits.")
        else:
            st.warning("No severe defects found across the entire dataset.")
   # --- TAB 4: I-MR CHART (TIMELINE STABILITY) ---
    with tab4:
        st.header("📈 Task 4: I-MR Stability Tracking (Chronological)")
        st.info("Analysis based on production sequence from 2024 to 2026. Red dots = Out of Spec.")

        GLOBAL_SPECS = {
            'YS': {'min': 400, 'max': 460, 'target': 430},
            'TS': {'min': 410, 'max': 470, 'target': 440},
            'EL': {'min': 25, 'max': None, 'target': None},
            'YPE': {'min': 4, 'max': None, 'target': None}
        }

        # Filters for Task 4
        imr_periods = ["All Periods"] + sorted(df_filtered['Time_Group'].dropna().unique().tolist())
        imr_thicks = sorted(df_filtered['Actual_Thickness'].dropna().unique())
        imr_mats = sorted(df_filtered['HR_Material'].astype(str).unique())
        
        c1, c2, c3 = st.columns(3)
        sel_p = c1.selectbox("Filter Period:", imr_periods, key="t4_p")
        sel_t = c2.selectbox("Filter Thickness:", imr_thicks, key="t4_t")
        sel_m = c3.selectbox("Filter Material:", imr_mats, key="t4_m")

        if sel_p == "All Periods":
            imr_df = df_filtered[(df_filtered['Actual_Thickness'] == sel_t) & 
                                (df_filtered['HR_Material'] == sel_m)]
        else:
            imr_df = df_filtered[(df_filtered['Time_Group'] == sel_p) & 
                                (df_filtered['Actual_Thickness'] == sel_t) & 
                                (df_filtered['HR_Material'] == sel_m)]

        imr_df = imr_df.sort_values(by='烤三生產日期').reset_index(drop=True)

        if not imr_df.empty:
            for feat in ['YS', 'TS', 'EL', 'YPE']:
                if feat in imr_df.columns:
                    valid_data = imr_df.dropna(subset=[feat, '烤三生產日期']).copy()
                    if len(valid_data) > 1:
                        st.markdown(f"### 🛡️ Stability: **{feat}**")
                        
                        # Khúc này đã được canh lề chuẩn 100%
                        valid_data = valid_data.reset_index(drop=True)
                        dates = valid_data['烤三生產日期']
                        vals = valid_data[feat].values
                        
                        x_seq = np.arange(len(vals)) 
                        mean_v = np.mean(vals)
                        
                        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 7), gridspec_kw={'height_ratios': [2, 1]})
                        
                        # --- I-Chart ---
                        ax1.plot(x_seq, vals, marker='o', ms=4, lw=1, color='#1f77b4', alpha=0.6, label=feat)
                        ax1.axhline(mean_v, color='green', ls='--', label='Mean')
                        
                        if feat in GLOBAL_SPECS:
                            s = GLOBAL_SPECS[feat]
                            if s['min']: ax1.axhline(s['min'], color='red', lw=2)
                            if s['max']: ax1.axhline(s['max'], color='red', lw=2)
                            
                            v_x, v_y = [], []
                            for i, v in enumerate(vals):
                                if (s['min'] and v < s['min']) or (s['max'] and v > s['max']):
                                    v_x.append(i); v_y.append(v)
                            if v_x: ax1.scatter(v_x, v_y, color='red', s=60, zorder=5)
                        
                        # Vạch chia năm
                        if sel_p == "All Periods":
                            for i in range(1, len(dates)):
                                if dates.iloc[i].year != dates.iloc[i-1].year:
                                    ax1.axvline(i, color='black', ls='-.', alpha=0.3)
                                    ax1.text(i, ax1.get_ylim()[1], f" {dates.iloc[i].year}", fontsize=10, va='top')

                        ax1.set_title(f"Individual Chart (I) - {feat}", fontweight='bold')
                        ax1.legend(loc='upper right', fontsize=8)
                        ax1.set_xticks([]) # Ẩn nhãn X trục trên cho thoáng

                        # --- MR-Chart ---
                        mr = np.abs(np.diff(vals))
                        mr_mean = np.mean(mr)
                        ucl_mr = 3.267 * mr_mean
                        
                        ax2.plot(x_seq[1:], mr, marker='o', ms=3, color='orange', alpha=0.6)
                        ax2.axhline(mr_mean, color='green', ls='--')
                        ax2.axhline(ucl_mr, color='red', ls=':')
                        
                        hv_x, hv_y = [], []
                        for i, m_val in enumerate(mr):
                            if m_val > ucl_mr: hv_x.append(i+1); hv_y.append(m_val)
                        if hv_x: ax2.scatter(hv_x, hv_y, color='red', s=40, zorder=5)

                        ax2.set_title("Moving Range Chart (MR)", fontweight='bold')
                        
                        # Trục X: tự động căn chỉnh hiển thị 10-15 nhãn ngày
                        step = max(1, len(x_seq) // 12) 
                        ax2.set_xticks(x_seq[::step])
                        ax2.set_xticklabels(dates.dt.strftime('%Y-%m-%d').iloc[::step], rotation=45, ha='right')

                        fig.tight_layout()
                        
                        safe_p_imr = "".join([c if c.isalnum() else "_" for c in sel_p])
                        plt.savefig(f"export_imr_{feat}.png", bbox_inches='tight', dpi=150)
                        st.pyplot(fig); plt.close(fig)
                        
                        if v_x:
                            st.error(f"⚠️ **Stability Insight:** Found {len(v_x)} out-of-spec points for {feat}.")
                        else:
                            st.success(f"✅ **Stability Insight:** {feat} is within specification limits.")
        else:
            st.warning("No data found for the selected combination.")                
    # --- EXPORT SECTION ---
    # (Giữ nguyên logic Export ban đầu của bạn...)
    st.sidebar.header("📥 Export Options")
  
# --- EXPORT PDF VISUAL REPORT ---
    st.sidebar.header("📥 Export PDF Report")
    st.sidebar.info("Navigate through the tabs to generate and update charts, then click below to compile them into a PDF.")

    if st.sidebar.button("🖨️ Generate PDF containing Charts"):
        pdf = FPDF(orientation='L', unit='mm', format='A4') # Khổ giấy A4 nằm ngang
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # --- PAGE 1: HEATMAP ---
        if os.path.exists("export_heatmap.png"):
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, "1. DEFECT HOTSPOT DIAGNOSTIC MAP", ln=True, align='C')
            pdf.image("export_heatmap.png", x=15, y=25, w=260) # Chèn hình Heatmap
            
        # --- PAGE 2: PARETO ---
        if os.path.exists("export_pareto.png"):
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, "2. PARETO ANALYSIS (MAIN DEFECTS)", ln=True, align='C')
            pdf.image("export_pareto.png", x=30, y=25, w=220) # Chèn hình Pareto
            
        # --- PAGE 3+: I-MR CHARTS ---
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, "3. I-MR PROCESS STABILITY TRACKING", ln=True, align='C')
        
        y_pos = 25
        chart_count = 0
        for feat in ['YS', 'TS', 'EL', 'YPE']:
            img_path = f"export_imr_{feat}.png"
            if os.path.exists(img_path):
                if chart_count == 2: # Nếu đã dán 2 hình thì qua trang mới
                    pdf.add_page()
                    y_pos = 20
                    chart_count = 0
                pdf.image(img_path, x=20, y=y_pos, w=250) # Chèn hình I-MR
                y_pos += 90 # Đẩy tọa độ Y xuống cho hình tiếp theo
                chart_count += 1
                
        # Xuất file PDF
        pdf.output("Quality_Visual_Report.pdf")
        
        # Hiển thị nút tải về
        with open("Quality_Visual_Report.pdf", "rb") as f:
            st.sidebar.download_button(
                label="✅ Click to Download your PDF Report", 
                data=f.read(), 
                file_name="Quality_Visual_Report.pdf", 
                mime="application/pdf"
            )
        st.sidebar.success("PDF Generated Successfully!")
