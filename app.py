import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import io
from matplotlib.patches import Patch
from fpdf import FPDF
from PIL import Image as PILImage
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
                df.rename(columns={df.columns[i - 1]: 'Actual_Thickness'}, inplace=True)
                break

    if 'Actual_Thickness' in df.columns:
        df['Actual_Thickness'] = pd.to_numeric(df['Actual_Thickness'], errors='coerce').round(3)

    if '烤三生產日期' in df.columns:
        d_str = df['烤三生產日期'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df['烤三生產日期'] = pd.to_datetime(d_str, format='%Y%m%d', errors='coerce').fillna(
            pd.to_datetime(d_str, errors='coerce')
        )

    # --- Coil ID ---
    COIL_ID_COL = '鋼捲號碼'
    if COIL_ID_COL in df.columns:
        df[COIL_ID_COL] = df[COIL_ID_COL].astype(str).str.strip()
    else:
        df[COIL_ID_COL] = df.index.astype(str)

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
        if f in df.columns:
            df[f] = pd.to_numeric(df[f], errors='coerce')

    if '熱軋材質' in df.columns:
        df['HR_Material'] = df['熱軋材質'].astype(str).str.strip().replace(['nan', ''], 'Unknown')
    else:
        df['HR_Material'] = 'Unknown'

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

    time_order_map = {
        "2024 (Full Year)": 1,
        "2025 H1 (Until 06/28)": 2,
        "2025 Q3 (06/29 - 09/30)": 3,
        "2025 Q4": 4,
        "2025 (Full Year)": 5,
        "2026 Q1": 6,
        "Unknown": 99
    }

    # --- TABS ---
    tab0, tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "0. Intro",             # <- Thêm tên tab thứ 0 của bạn vào đây (ví dụ: "0. Intro" hoặc "0. Home")
    "1. Overview", 
    "2. Capability", 
    "3. Root Cause", 
    "4. Group Detail", 
    "5. Tail Scrap", 
    "6. Control Limits"
    ])

    with tab0:
        st.dataframe(df.head(10), use_container_width=True)

    # --- GLOBAL FILTERING ---
    df = df[df['Actual_Thickness'].isin([0.5, 0.6, 0.75, 0.8])]
    st.sidebar.header("🔎 Filters")
    all_periods = sorted(df['Time_Group'].unique())
    ui_selection = st.sidebar.multiselect(
        "📅 Select Period(s):", options=["All"] + all_periods, default=["All"]
    )
    selected_periods = all_periods if ("All" in ui_selection or not ui_selection) else ui_selection
    df_filtered = df[df['Time_Group'].isin(selected_periods)].copy()
    thickness_list = sorted(df['Actual_Thickness'].dropna().unique())

    # ==========================================================
    # TAB 1: YIELD SUMMARY
    # ==========================================================
    with tab1:
        st.header("1. Quality Yield Summary & Worst Offenders")
        st.info("Overview of production yield. Chronologically sorted from 2024 onwards.")
        st.subheader("📊 Executive Summary: Production Quality Timeline")

        severe_grades = ['B+', 'B']
        df_filtered['Severe_Bad_Qty'] = df_filtered[severe_grades].sum(axis=1)
        df_filtered['Acceptable_Qty'] = df_filtered['Total_Qty'] - df_filtered['Severe_Bad_Qty']

        period_summary = df_filtered.groupby(['Time_Group', 'Actual_Thickness', 'HR_Material'])[
            ['Total_Qty', 'Acceptable_Qty', 'Severe_Bad_Qty']
        ].sum().reset_index()
        period_summary = period_summary[period_summary['Total_Qty'] > 0]
        period_summary['Yield (%)'] = (
            period_summary['Acceptable_Qty'] / period_summary['Total_Qty'] * 100
        ).fillna(0).round(2)
        period_summary['Defect_Rate (%)'] = (
            period_summary['Severe_Bad_Qty'] / period_summary['Total_Qty'] * 100
        ).fillna(0).round(2)
        period_summary['Sort_Key'] = period_summary['Time_Group'].map(time_order_map).fillna(90)
        period_summary = period_summary.sort_values(
            by=['Sort_Key', 'Actual_Thickness']
        ).drop(columns=['Sort_Key'])
        period_summary.rename(columns={'Severe_Bad_Qty': 'Bad_Qty (B+, B)'}, inplace=True)

        if not period_summary.empty:
            st.dataframe(
                period_summary.style
                    .background_gradient(subset=['Defect_Rate (%)'], cmap='Reds')
                    .background_gradient(subset=['Yield (%)'], cmap='Greens')
                    .format({
                        'Actual_Thickness': '{:.2f}', 'Total_Qty': '{:.0f}',
                        'Acceptable_Qty': '{:.0f}', 'Bad_Qty (B+, B)': '{:.0f}',
                        'Yield (%)': '{:.2f}%', 'Defect_Rate (%)': '{:.2f}%'
                    }),
                use_container_width=True, hide_index=True
            )

            fig_s, ax_s = plt.subplots(figsize=(13, 5))
            pivot_data = period_summary.pivot_table(
                index='Time_Group', columns='Actual_Thickness',
                values='Defect_Rate (%)', aggfunc='mean'
            )
            pivot_data = pivot_data.reindex(
                sorted(pivot_data.index, key=lambda x: time_order_map.get(x, 99))
            )
            pivot_data.plot(kind='bar', ax=ax_s, colormap='Reds', edgecolor='white')
            ax_s.set_title("Defect Rate (%) by Period & Thickness", fontweight='bold', fontsize=13)
            ax_s.set_xlabel("")
            ax_s.set_ylabel("Defect Rate (%)")
            ax_s.legend(title="Thickness (mm)", bbox_to_anchor=(1.02, 1), loc='upper left')
            ax_s.tick_params(axis='x', rotation=30)
            fig_s.tight_layout()
            plt.savefig("export_tab1_defect_rate_bar.png", bbox_inches='tight', dpi=150)
            st.pyplot(fig_s)
            plt.close(fig_s)

            fig_y, ax_y = plt.subplots(figsize=(13, 5))
            pivot_yield = period_summary.pivot_table(
                index='Time_Group', columns='Actual_Thickness',
                values='Yield (%)', aggfunc='mean'
            )
            pivot_yield = pivot_yield.reindex(
                sorted(pivot_yield.index, key=lambda x: time_order_map.get(x, 99))
            )
            pivot_yield.plot(kind='bar', ax=ax_y, colormap='Greens', edgecolor='white')
            ax_y.set_title("Yield (%) by Period & Thickness", fontweight='bold', fontsize=13)
            ax_y.set_xlabel("")
            ax_y.set_ylabel("Yield (%)")
            ax_y.set_ylim(0, 110)
            ax_y.legend(title="Thickness (mm)", bbox_to_anchor=(1.02, 1), loc='upper left')
            ax_y.tick_params(axis='x', rotation=30)
            fig_y.tight_layout()
            plt.savefig("export_tab1_yield_bar.png", bbox_inches='tight', dpi=150)
            st.pyplot(fig_y)
            plt.close(fig_y)

        st.markdown("---")
        st.subheader("📑 Detailed Yield by Period (All Grades)")
        sum_df = df_filtered.groupby(
            ['Time_Group', 'Actual_Thickness', 'HR_Material'], dropna=False
        )[base_grades].sum().reset_index()
        sum_df['Total_Qty'] = sum_df[base_grades].sum(axis=1)
        for col in base_grades:
            sum_df[f"% {col}"] = (
                (sum_df[col] / sum_df['Total_Qty'].replace(0, np.nan)) * 100
            ).fillna(0).round(1)
        sum_df['Sort_Key'] = sum_df['Time_Group'].map(time_order_map).fillna(90)
        sum_df = sum_df.sort_values(by=['Sort_Key', 'Actual_Thickness']).drop(columns=['Sort_Key'])
        sum_df.rename(columns={'Time_Group': 'Period', 'Actual_Thickness': 'Thickness'}, inplace=True)
        ordered_periods = sorted(sum_df['Period'].unique(), key=lambda x: time_order_map.get(x, 99))

        for period in ordered_periods:
            p_data = sum_df[sum_df['Period'] == period]
            if not p_data.empty:
                st.markdown(f"#### 📅 Period: **{period}**")
                format_dict = {'Thickness': '{:.2f}', 'Total_Qty': '{:.0f}'}
                for col in base_grades:
                    format_dict[col] = '{:.0f}'
                    format_dict[f"% {col}"] = '{:.1f}%'
                st.dataframe(
                    p_data.drop(columns=['Period']).style.format(format_dict),
                    use_container_width=True, hide_index=True
                )

        st.markdown("---")
        st.subheader("📊 Grade Distribution by Time Period (%)")
        grade_dist = df_filtered.groupby('Time_Group')[base_grades].sum()
        grade_dist['Total'] = grade_dist.sum(axis=1)
        for g in base_grades:
            grade_dist[f'pct_{g}'] = (
                grade_dist[g] / grade_dist['Total'].replace(0, np.nan) * 100
            ).fillna(0).round(1)
        grade_dist_display = grade_dist[[f'pct_{g}' for g in base_grades]].copy()
        grade_dist_display.columns = base_grades
        grade_dist_display.index.name = 'Time Period'
        grade_dist_display['_sort'] = grade_dist_display.index.map(
            lambda x: time_order_map.get(x, 99)
        )
        grade_dist_display = grade_dist_display.sort_values('_sort').drop(columns=['_sort'])
        grade_dist_pct = grade_dist_display.map(lambda x: f"{x:.1f}%")

        header_color = "#1a3a5c"
        alt_row_color = "#dce6f1"
        html = f"""
        <style>
        .grade-table {{ width:100%; border-collapse:collapse; font-family:sans-serif; font-size:14px; margin-bottom:24px; }}
        .grade-table th {{ background-color:{header_color}; color:white; padding:10px 16px; text-align:center; }}
        .grade-table td {{ padding:9px 16px; text-align:center; border-bottom:1px solid #ccc; }}
        .grade-table tr:nth-child(odd) td {{ background-color:{alt_row_color}; }}
        .grade-table tr:nth-child(even) td {{ background-color:#ffffff; }}
        .grade-table tr:hover td {{ background-color:#b8cce4; }}
        </style>
        <table class="grade-table">
            <thead><tr><th>Time Period</th>{''.join(f'<th>{g}</th>' for g in base_grades)}</tr></thead>
            <tbody>
        """
        for period, row in grade_dist_pct.iterrows():
            html += "<tr>"
            html += f"<td><b>{period}</b></td>"
            for g in base_grades:
                val = float(row[g].replace('%', ''))
                if g in ['B+', 'B'] and val > 1.0:
                    html += f'<td style="color:#c00000;font-weight:bold">{row[g]}</td>'
                elif g in ['A-B+', 'A-B'] and val > 30.0:
                    html += f'<td style="color:#2e7d32;font-weight:bold">{row[g]}</td>'
                else:
                    html += f'<td>{row[g]}</td>'
            html += "</tr>"
        html += "</tbody></table>"
        st.markdown(html, unsafe_allow_html=True)

        st.markdown("---")
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                period_summary.to_excel(writer, index=False, sheet_name='Executive_Summary')
                sum_df.to_excel(writer, index=False, sheet_name='Detailed_Yield')
                grade_dist_display.to_excel(writer, sheet_name='Grade_Distribution_%')
                workbook = writer.book
                worksheet = writer.sheets['Executive_Summary']
                header_fmt = workbook.add_format({
                    'bold': True, 'bg_color': '#D7E4BC', 'border': 1,
                    'align': 'center', 'valign': 'vcenter'
                })
                num_fmt = workbook.add_format({'align': 'center', 'border': 1})
                pct_fmt = workbook.add_format({'num_format': '0.00"%"', 'align': 'center', 'border': 1})
                for col_num, value in enumerate(period_summary.columns.values):
                    worksheet.write(0, col_num, value, header_fmt)
                    worksheet.set_column(col_num, col_num, 15)
                worksheet.conditional_format(1, 6, len(period_summary), 6,
                    {'type': '2_color_scale', 'min_color': "#F7FCF5", 'max_color': "#41AB5D"})
                worksheet.conditional_format(1, 7, len(period_summary), 7,
                    {'type': '2_color_scale', 'min_color': "#FFF5F0", 'max_color': "#EF3B2C"})
                for row in range(1, len(period_summary) + 1):
                    for col in range(len(period_summary.columns)):
                        val = period_summary.iloc[row - 1, col]
                        if col >= 6 and isinstance(val, (int, float)):
                            worksheet.write(row, col, val / 100, pct_fmt)
                        else:
                            worksheet.write(row, col, val, num_fmt)
            st.download_button(
                label="📥 Download Formatted Excel", data=output.getvalue(),
                file_name="Colored_Yield.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception:
            st.download_button(
                label="📥 Download Basic CSV",
                data=period_summary.to_csv(index=False).encode('utf-8'),
                file_name="Yield_Summary.csv", mime="text/csv"
            )

    # ==========================================================
    # ==========================================================
# ==========================================================
    # --- TAB 2: DISTRIBUTION + Cp / Cpk / Ca ---
    # ==========================================================
    with tab2:
        st.header("📊 Distribution & Process Capability (SPC)")
        st.info(
            "Visualizing mechanical property distribution and capability indices (Cp, Cpk, Ca). "
            "Data is grouped by time period and thickness to identify precise process shifts."
        )

        local_order_map = {
            "2024 (Full Year)": 1,
            "2025 H1 (Until 06/28)": 2,
            "2025 Q3 (06/29 - 09/30)": 3,
            "2025 Q4": 4,
            "2025 (Full Year)": 5,
            "2026 Q1": 6
        }
        
        ordered_periods = sorted(selected_periods, key=lambda x: local_order_map.get(x, 99))

        # ----------------------------------------------------------
        # CAPABILITY INDEX HELPERS
        # ----------------------------------------------------------
        def calc_capability(values, feat):
            """
            Return dict with mean, std, Cp, Cpk, Ca (None if spec missing).
            """
            vals = np.array(values, dtype=float)
            vals = vals[~np.isnan(vals)]
            if len(vals) < 2:
                return None
            mu  = np.mean(vals)
            std = np.std(vals, ddof=1)
            if std == 0:
                return None

            spec = GLOBAL_SPECS.get(feat, {})
            lsl  = spec.get('min')
            usl  = spec.get('max')
            tgt  = spec.get('target')

            result = {'mean': mu, 'std': std, 'n': len(vals),
                      'Cp': None, 'Cpk': None, 'Ca': None,
                      'LSL': lsl, 'USL': usl, 'Target': tgt}

            if lsl is not None and usl is not None:
                cp   = (usl - lsl) / (6 * std)
                cpu  = (usl - mu)  / (3 * std)
                cpl  = (mu  - lsl) / (3 * std)
                cpk  = min(cpu, cpl)
                result['Cp']  = round(cp,  3)
                result['Cpk'] = round(cpk, 3)
                if tgt is not None:
                    ca = (mu - tgt) / ((usl - lsl) / 2) * 100
                    result['Ca'] = round(ca, 2)
                else:
                    mid = (usl + lsl) / 2
                    ca  = (mu - mid) / ((usl - lsl) / 2) * 100
                    result['Ca'] = round(ca, 2)
            elif lsl is not None:
                cpl = (mu - lsl) / (3 * std)
                result['Cp']  = round(cpl, 3)
                result['Cpk'] = round(cpl, 3)
            elif usl is not None:
                cpu = (usl - mu) / (3 * std)
                result['Cp']  = round(cpu, 3)
                result['Cpk'] = round(cpu, 3)

            return result

        def cpk_color(cpk):
            """Traffic-light color for Cpk."""
            if cpk is None: return '#888888'
            if cpk >= 1.67: return '#2e7d32'   # excellent
            if cpk >= 1.33: return '#66bb6a'   # capable
            if cpk >= 1.00: return '#ffa726'   # marginal
            return '#d62728'                   # not capable

        def cpk_label(cpk):
            if cpk is None: return 'N/A'
            if cpk >= 1.67: return '✅ Excellent'
            if cpk >= 1.33: return '✅ Capable'
            if cpk >= 1.00: return '⚠️ Marginal'
            return '❌ Not Capable'

        def render_capability_badge(cap, feat):
            """Render a compact colored HTML badge row below each chart."""
            if cap is None:
                return
            cp_v   = f"{cap['Cp']:.3f}"   if cap['Cp']  is not None else 'N/A'
            cpk_v  = f"{cap['Cpk']:.3f}"  if cap['Cpk'] is not None else 'N/A'
            ca_v   = f"{cap['Ca']:.1f}%"  if cap['Ca']  is not None else 'N/A (no target)'
            mu_v   = f"{cap['mean']:.2f}"
            std_v  = f"{cap['std']:.3f}"
            n_v    = str(cap['n'])
            clr    = cpk_color(cap['Cpk'])
            lbl    = cpk_label(cap['Cpk'])
            spec   = GLOBAL_SPECS.get(feat, {})
            lsl_v  = str(spec.get('min', '—'))
            usl_v  = str(spec.get('max', '—'))

            html_badge = f"""
            <div style="background:#f8f9fa;border-left:5px solid {clr};
                        border-radius:6px;padding:8px 14px;margin:4px 0 10px 0;
                        font-family:monospace;font-size:13px;line-height:1.8;">
              <span style="font-size:14px;font-weight:bold;color:{clr};">{lbl}</span>
              &nbsp;&nbsp;|&nbsp;&nbsp;
              <b>LSL</b>: {lsl_v} &nbsp; <b>USL</b>: {usl_v}
              &nbsp;&nbsp;|&nbsp;&nbsp;
              <b>n</b>: {n_v} &nbsp;
              <b>Mean</b>: {mu_v} &nbsp;
              <b>Std</b>: {std_v}
              <br>
              <b style="color:{clr};">Cpk = {cpk_v}</b>
              &nbsp;&nbsp;
              <b>Cp = {cp_v}</b>
              &nbsp;&nbsp;
              <b>Ca = {ca_v}</b>
            </div>
            """
            st.markdown(html_badge, unsafe_allow_html=True)

        def build_capability_summary(df_src, feat, label):
            vals = df_src[feat].dropna().values if feat in df_src.columns else []
            cap = calc_capability(vals, feat)
            if cap is None: return None
            return {
                'Period / Segment': label,
                'Feature': feat,
                'n': cap['n'],
                'Mean': round(cap['mean'], 2),
                'Std': round(cap['std'], 3),
                'LSL': GLOBAL_SPECS.get(feat, {}).get('min'),
                'USL': GLOBAL_SPECS.get(feat, {}).get('max'),
                'Ca (%)': cap['Ca'],
                'Cp': cap['Cp'],
                'Cpk': cap['Cpk'],
                'Verdict': cpk_label(cap['Cpk']).replace('✅ ', '').replace('⚠️ ', '').replace('❌ ', '')
            }

        global_x_bounds = {}
        for feat in mech_features:
            if feat in df.columns:
                vd = df[[feat, 'Total_Qty']].dropna().copy()
                vd = vd[vd['Total_Qty'] > 0]
                if not vd.empty:
                    q1, q99 = np.percentile(vd[feat], 1), np.percentile(vd[feat], 99)
                    global_x_bounds[feat] = (q1 - (q99 - q1) * 0.25, q99 + (q99 - q1) * 0.25)

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
            c_map = {
                'A-B+': '#2ca02c', 'A-B': '#1f77b4',
                'A-B-': '#ff7f0e', 'B+': '#9467bd', 'B': '#d62728'
            }
            spec = GLOBAL_SPECS.get(feat, {})
            lsl, usl, tgt = spec.get('min'), spec.get('max'), spec.get('target')

            fmin, fmax = global_x_bounds.get(feat, (
                data[feat].min() if not data.empty else 0,
                data[feat].max() if not data.empty else 100
            ))
            v_l, w_l, clrs, m_info = [], [], [], []
            for g in base_grades:
                td = data[[feat, g]].dropna()
                td = td[td[g] > 0]
                if not td.empty:
                    v_l.append(td[feat].values)
                    w_l.append(td[g].values)
                    clrs.append(c_map[g])
                    m = np.average(td[feat].values, weights=td[g].values)
                    ax.axvline(m, color=c_map[g], ls='--', lw=1.2)
                    m_info.append({'v': m, 'c': c_map[g], 'label': g})

            if v_l:
                ax.hist(v_l, bins=np.linspace(fmin, fmax, 16), weights=w_l, color=clrs,
                        stacked=True, edgecolor='white', alpha=0.7)
                m_info.sort(key=lambda x: x['v'])
                x_range = fmax - fmin
                min_gap = x_range * 0.045
                positions = [info['v'] for info in m_info]
                for _ in range(50):
                    moved = False
                    for i in range(1, len(positions)):
                        if positions[i] - positions[i - 1] < min_gap:
                            mid = (positions[i] + positions[i - 1]) / 2
                            positions[i - 1] = mid - min_gap / 2
                            positions[i] = mid + min_gap / 2
                            moved = True
                    if not moved: break
                y_levels = [y_lim * (0.92 - (i % 4) * 0.13) for i in range(len(m_info))]
                for i, info in enumerate(m_info):
                    x_pos = positions[i]
                    y_pos = y_levels[i]
                    ax.annotate(
                        f"{info['v']:.1f}",
                        xy=(info['v'], y_pos * 0.6), xytext=(x_pos, y_pos),
                        color='white', fontweight='bold', fontsize=8,
                        ha='center', va='center',
                        bbox=dict(facecolor=info['c'], alpha=0.85, boxstyle='round,pad=0.25'),
                        arrowprops=dict(arrowstyle='-', color=info['c'], lw=1.0, alpha=0.6)
                        if abs(x_pos - info['v']) > min_gap * 0.3 else None
                    )

            y_top = y_lim * 0.98
            if lsl is not None:
                ax.axvline(lsl, color='red', lw=2, ls='-', zorder=3)
                ax.text(lsl, y_top, f' LSL\n {lsl}', color='red', fontsize=7.5, fontweight='bold', va='top', ha='left')
            if usl is not None:
                ax.axvline(usl, color='red', lw=2, ls='-', zorder=3)
                ax.text(usl, y_top, f' USL\n {usl}', color='red', fontsize=7.5, fontweight='bold', va='top', ha='right')
            if tgt is not None:
                ax.axvline(tgt, color='#1a7abf', lw=1.5, ls=':', zorder=3)
                ax.text(tgt, y_top * 0.75, f' TGT\n {tgt}', color='#1a7abf', fontsize=7, fontweight='bold', va='top', ha='left')

            ax.legend(handles=[Patch(facecolor=c_map[g], label=g) for g in base_grades if g in data.columns],
                      loc='upper right', fontsize=7)
            ax.set_xlim(fmin, fmax)
            ax.set_ylim(0, y_lim)
            ax.set_title(title, fontsize=10, fontweight='bold')

        # ----------------------------------------------------------
        # Build cross-period capability summary
        # ----------------------------------------------------------
        cap_summary_rows = []
        for _p in ordered_periods:
            _dfp = df_filtered[df_filtered['Time_Group'] == _p]
            for _f in ['YS', 'TS', 'EL', 'YPE']:
                row = build_capability_summary(_dfp, _f, _p)
                if row:
                    cap_summary_rows.append(row)

        if cap_summary_rows:
            st.subheader("📊 Process Capability Summary (All Selected Periods)")
            st.caption(
                "**Cp** = spread capability (tolerance / 6σ) | "
                "**Cpk** = centred capability (worst-side) | "
                "**Ca** = centering accuracy (0% = perfectly centred). "
                "Cpk ≥ 1.67 ✅ Excellent | ≥ 1.33 ✅ Capable | ≥ 1.00 ⚠️ Marginal | < 1.00 ❌ Not Capable"
            )
            cap_df = pd.DataFrame(cap_summary_rows)

            # ==========================================================
            # CROSS-PERIOD COMPARISON MATRIX
            # ==========================================================
            st.markdown("### 🔄 Cross-Period Comparison (Cpk, Cp, Ca Trend)")
            
            trend_data = []
            for p in ordered_periods:
                if p not in cap_df['Period / Segment'].values: continue
                row = {'Period': p}
                for feat in ['YS', 'TS', 'EL', 'YPE']:
                    df_fp = cap_df[(cap_df['Period / Segment'] == p) & (cap_df['Feature'] == feat)]
                    if not df_fp.empty:
                        row[f"{feat} Cpk"] = df_fp.iloc[0]['Cpk']
                        row[f"{feat} Cp"] = df_fp.iloc[0]['Cp']
                        row[f"{feat} Ca (%)"] = df_fp.iloc[0]['Ca (%)']
                if len(row) > 1:
                    trend_data.append(row)
                    
            if trend_data:
                trend_df = pd.DataFrame(trend_data).set_index('Period')
                
                def style_trend_table(df):
                    styles = pd.DataFrame('', index=df.index, columns=df.columns)
                    for r_idx in df.index:
                        for c in df.columns:
                            val = df.at[r_idx, c]
                            if pd.isna(val): continue
                            
                            if "Cpk" in c:
                                if val >= 1.67: styles.at[r_idx, c] = 'background-color: #2e7d32; color: white; font-weight: bold;'
                                elif val >= 1.33: styles.at[r_idx, c] = 'background-color: #66bb6a; color: white; font-weight: bold;'
                                elif val >= 1.00: styles.at[r_idx, c] = 'background-color: #ffa726; color: black; font-weight: bold;'
                                else: styles.at[r_idx, c] = 'background-color: #d62728; color: white; font-weight: bold;'
                            elif "Cp" in c:
                                if val >= 1.67: styles.at[r_idx, c] = 'color: #2e7d32; font-weight: bold;'
                                elif val >= 1.33: styles.at[r_idx, c] = 'color: #66bb6a; font-weight: bold;'
                                elif val >= 1.00: styles.at[r_idx, c] = 'color: #ffa726; font-weight: bold;'
                                else: styles.at[r_idx, c] = 'color: #d62728; font-weight: bold;'
                            elif "Ca" in c:
                                av = abs(val)
                                if av <= 12.5: styles.at[r_idx, c] = 'color: #2e7d32; font-weight: bold;'
                                elif av <= 25.0: styles.at[r_idx, c] = 'color: #ffa726; font-weight: bold;'
                                else: styles.at[r_idx, c] = 'color: #d62728; font-weight: bold;'
                    return styles
                
                format_dict = {c: "{:.1f}%" if "Ca" in c else "{:.3f}" for c in trend_df.columns}
                styled_trend = trend_df.style.apply(style_trend_table, axis=None).format(format_dict, na_rep="—")
                st.dataframe(styled_trend, use_container_width=True)

                # ==========================================================
                # AUTOMATED TREND CONCLUSION
                # ==========================================================
                st.markdown("#### 💡 Automated Trend Conclusion")
                periods_list = trend_df.index.tolist()
                
                if len(periods_list) >= 2:
                    first_p = periods_list[0]
                    last_p = periods_list[-1]
                    
                    for feat in ['YS', 'TS', 'EL', 'YPE']:
                        cpk_col = f"{feat} Cpk"
                        if cpk_col in trend_df.columns:
                            val_first = trend_df.at[first_p, cpk_col]
                            val_last = trend_df.at[last_p, cpk_col]
                            
                            if pd.isna(val_first) or pd.isna(val_last): continue
                            
                            diff = val_last - val_first
                            
                            if diff >= 0.05: trend_status = "📈 **Improving**"
                            elif diff <= -0.05: trend_status = "📉 **Declining**"
                            else: trend_status = "➡️ **Stable**"
                            
                            if val_last >= 1.33: risk_status = "✅ **Safe** (Capable)"
                            elif val_last >= 1.00: risk_status = "⚠️ **Warning** (Marginal)"
                            else: risk_status = "❌ **Danger** (Not Capable)"
                            
                            st.info(f"**{feat}:** Compared to [{first_p}], the capability (Cpk) in [{last_p}] shifted from **{val_first:.2f}** to **{val_last:.2f}** ➔ Trend: {trend_status}. Current Status: {risk_status}.")
                else:
                    st.caption("ℹ️ Please select at least 2 periods in the sidebar filter to enable automated trend comparison.")

            # ==========================================================
            # DETAILED CAPABILITY LOG (SHOWING MEAN & STD)
            # ==========================================================
            st.markdown("### 📋 Detailed Capability Log")
            st.caption("Review the **Mean** values here to understand the root cause of high Ca (Centering Accuracy) deviations.")
            
            def color_cpk_cell(val):
                if pd.isna(val): return ''
                c = cpk_color(val)
                return f'background-color:{c};color:white;font-weight:bold;text-align:center'

            def color_ca_cell(val):
                if pd.isna(val): return ''
                av = abs(val)
                if av <= 12.5: clr = '#2e7d32'
                elif av <= 25:  clr = '#ffa726'
                else:           clr = '#d62728'
                return f'color:{clr};font-weight:bold;text-align:center'

            fmt = {
                'Mean': '{:.2f}', 'Std': '{:.3f}',
                'Cp': '{:.3f}', 'Cpk': '{:.3f}',
                'Ca (%)': lambda v: f'{v:.1f}%' if pd.notnull(v) else '—',
                'LSL': lambda v: str(int(v)) if pd.notnull(v) else '—',
                'USL': lambda v: str(int(v)) if pd.notnull(v) else '—',
            }
            styled = (
                cap_df.style
                .map(color_cpk_cell, subset=['Cpk'])
                .map(color_ca_cell,  subset=['Ca (%)'])
                .background_gradient(subset=['Cp'], cmap='RdYlGn', vmin=0.67, vmax=2.0)
                .format(fmt, na_rep='—')
            )
            st.dataframe(styled, use_container_width=True, hide_index=True)

            cap_xlsx = io.BytesIO()
            try:
                with pd.ExcelWriter(cap_xlsx, engine='xlsxwriter') as _w:
                    cap_df.to_excel(_w, index=False, sheet_name='Capability_Log')
                    if trend_data:
                        trend_df.to_excel(_w, sheet_name='Capability_Trends')
                
                st.download_button(
                    label="📥 Download Capability Summary & Trends (Excel)",
                    data=cap_xlsx.getvalue(),
                    file_name="Capability_Summary_Trends.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception:
                pass

            st.markdown("---")

        # ----------------------------------------------------------
        # Per-period distribution charts WITH inline Cp/Cpk/Ca badge
        # ----------------------------------------------------------
        tab2_saved_files = []
        thickness_list = sorted(df_filtered['Actual_Thickness'].dropna().unique())
        
        for period in ordered_periods:
            df_p = df_filtered[df_filtered['Time_Group'] == period]
            if df_p.empty: continue
            
            st.markdown(f"## 📅 Period: **{period}**")
            ov_y = get_shared_y(df_p, ['YS', 'TS', 'EL', 'YPE'])
            safe_period = "".join([c if c.isalnum() else "_" for c in period])
            cols = st.columns(2)
            
            for idx, f in enumerate([x for x in ['YS', 'TS', 'EL', 'YPE'] if x in df_p.columns]):
                with cols[idx % 2]:
                    fig, ax = plt.subplots(figsize=(8, 4.5))
                    plot_dist(ax, df_p, f, f"{f} (Overall - {period})", ov_y)
                    fname = f"export_tab2_{safe_period}_overall_{f}.png"
                    plt.savefig(fname, bbox_inches='tight', dpi=150)
                    tab2_saved_files.append(fname)
                    st.pyplot(fig)
                    plt.close(fig)
                    
                    vals_all = df_p[f].dropna().values if f in df_p.columns else []
                    render_capability_badge(calc_capability(vals_all, f), f)

            for thick in thickness_list:
                df_t = df_p[df_p['Actual_Thickness'] == thick]
                if df_t.empty: continue
                
                st.markdown(f"**📏 Thickness: {thick}mm**")
                ly = get_shared_y(df_t, ['YS', 'TS', 'EL', 'YPE'])
                tcols = st.columns(2)
                
                for idx, f in enumerate([x for x in ['YS', 'TS', 'EL', 'YPE'] if x in df_t.columns]):
                    with tcols[idx % 2]:
                        fig, ax = plt.subplots(figsize=(8, 4.5))
                        plot_dist(ax, df_t, f, f"{f} (Thick:{thick} - {period})", ly)
                        fname = f"export_tab2_{safe_period}_t{str(thick).replace('.', 'p')}_{f}.png"
                        plt.savefig(fname, bbox_inches='tight', dpi=150)
                        tab2_saved_files.append(fname)
                        st.pyplot(fig)
                        plt.close(fig)
                        
                        vals_t = df_t[f].dropna().values if f in df_t.columns else []
                        render_capability_badge(calc_capability(vals_t, f), f)
    # ==========================================================
# TAB 3: ROOT CAUSE & DIAGNOSTIC
    # ==========================================================
    with tab3:
        st.header("🧠 Executive Auto-Insight & Root Cause")
        st.info("Automated diagnostic engine: Quantifying impact based on severe defects (B+, B).")

        st.success(
            "💡 **Q4 Improvement Verification:** '2025 Q4' has been automatically split into **October** "
            "and **Nov-Dec**. This isolates early-Q4 transition instability, providing evidence on whether "
            "defects dropped after the new corrective actions took effect."
        )

        df_tab3 = df_filtered.copy()

        def refine_q4(row):
            if row['Time_Group'] == '2025 Q4' and pd.notna(row['烤三生產日期']):
                if row['烤三生產日期'].month == 10:
                    return '2025 Q4 (Oct)'
                else:
                    return '2025 Q4 (Nov-Dec)'
            return row['Time_Group']
            
        if '烤三生產日期' in df_tab3.columns:
            df_tab3['Time_Group'] = df_tab3.apply(refine_q4, axis=1)

        severe_grades = ['B+', 'B']
        df_tab3['Severe_Bad_Qty'] = df_tab3[severe_grades].sum(axis=1)
        df_tab3['Spec_Label'] = (
            df_tab3['Actual_Thickness'].astype(str) + "mm (" + df_tab3['HR_Material'] + ")"
        )

        heat_data = df_tab3.groupby(['Spec_Label', 'Time_Group']).apply(
            lambda x: (x['Severe_Bad_Qty'].sum() / x['Total_Qty'].sum() * 100)
            if x['Total_Qty'].sum() > 0 else 0
        )

        df_unique_global = df_tab3.drop_duplicates(subset=['Row_ID'])

        if not heat_data.empty and heat_data.max() > 0:
            heatmap_long = heat_data.reset_index()
            heatmap_long.columns = ['Spec', 'Period', 'Defect_Rate']
            top_issues = heatmap_long[heatmap_long['Defect_Rate'] > 0].sort_values(
                'Defect_Rate', ascending=False
            ).head(5)
            
            top_3_subsets = []
            for _, row in top_issues.head(3).iterrows():
                subset = df_tab3[
                    (df_tab3['Spec_Label'] == row['Spec']) &
                    (df_tab3['Time_Group'] == row['Period'])
                ]
                top_3_subsets.append(subset)
            
            if top_3_subsets:
                df_top3 = pd.concat(top_3_subsets).drop_duplicates(subset=['Row_ID'])
            else:
                df_top3 = pd.DataFrame(columns=df_tab3.columns)

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
* 🚨 **Biggest Hotspot:** Specification **{top_issue['Spec']}** during **{top_issue['Period']}** (Severe Defect Rate: **{top_issue['Defect_Rate']:.1f}%**).
* 🧠 **Main Root Cause Driver:** **{top_driver}** is the primary culprit.
* 📊 **Quantified Impact:** Defective coils have a {top_driver} that is on average **{abs(gap_val):.1f} {direction}** than good coils.
                """)

            st.markdown("---")
            col1, col2 = st.columns(2)
            with col1:
                st.error("🔥 TOP 5 PROBLEM SEGMENTS")
                st.dataframe(
                    top_issues.style.background_gradient(subset=['Defect_Rate'], cmap='Reds')
                        .format({'Defect_Rate': '{:.1f}%'}),
                    use_container_width=True, hide_index=True
                )
            with col2:
                st.warning("🧠 ROOT CAUSE DRIVER (Top 3 Hotspots)")
                st.info("Mean difference: Bad Coils (Top 3) vs Global Good Coils (Deduplicated).")
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

            if not rc_s.empty:
                st.markdown("---")
                top_driver = rc_s.index[0]
                st.info(f"📏 DRILL DOWN: {top_driver} Shift by Thickness")
                drill_data = []
                for th in df_unique_global['Actual_Thickness'].dropna().unique():
                    th_df = df_unique_global[df_unique_global['Actual_Thickness'] == th]
                    g_val = th_df[th_df['Good_Qty'] > 0][top_driver].mean()
                    b_val = th_df[th_df['Severe_Bad_Qty'] > 0][top_driver].mean()
                    if pd.notnull(g_val) or pd.notnull(b_val):
                        drill_data.append({
                            'Thickness': th, 'GOOD Coils (Mean)': g_val,
                            'BAD Coils (Mean)': b_val,
                            'Impact Gap': (b_val - g_val)
                            if (pd.notnull(g_val) and pd.notnull(b_val)) else None
                        })
                drill_df = pd.DataFrame(drill_data).sort_values('Impact Gap', key=abs, ascending=False)
                st.dataframe(
                    drill_df.style.map(color_gap, subset=['Impact Gap'])
                        .format({
                            'Thickness': '{:.2f}mm', 'GOOD Coils (Mean)': '{:.1f}',
                            'BAD Coils (Mean)': '{:.1f}', 'Impact Gap': '{:+.1f}'
                        }),
                    use_container_width=True, hide_index=True
                )

            # ==========================================================
            # UPDATED: Visual Hotspot Map (Column order & Font sizes)
            # ==========================================================
            st.markdown("---")
            st.subheader("🗺️ Evidence: Visual Hotspot Map (Grades B+ and Below)")
            heat_pivot = heat_data.unstack()
            
            # --- Move 2025 Full Year next to 2024 Full Year by setting its weight to 1.5 ---
            local_order_map = {
                "2024 (Full Year)": 1,
                "2025 (Full Year)": 1.5,
                "2025 H1 (Until 06/28)": 2,
                "2025 Q3 (06/29 - 09/30)": 3,
                "2025 Q4 (Oct)": 4.1,
                "2025 Q4 (Nov-Dec)": 4.2,
                "2026 Q1": 6
            }
            
            if heat_pivot is not None and not heat_pivot.empty:
                ordered_cols = sorted(heat_pivot.columns, key=lambda x: local_order_map.get(x, 99))
                heat_pivot = heat_pivot.reindex(columns=ordered_cols)

                fig, ax = plt.subplots(figsize=(12, 5))
                vmax_threshold = 30.0 if heat_pivot.max().max() > 30 else heat_pivot.max().max()
                
                # Grid lines set to black
                sns.heatmap(heat_pivot, annot=True, fmt=".1f", cmap="YlOrRd",
                            linewidths=0.8, linecolor='black', vmax=vmax_threshold, ax=ax,
                            annot_kws={"size": 10}) # Default font size

                # Enlarge font size for ALL hotspots (value >= 3.0) across any period
                for text in ax.texts:
                    val_str = text.get_text()
                    if val_str:
                        try:
                            val = float(val_str)
                            if val >= 3.0:
                                text.set_fontsize(15)
                                text.set_fontweight('heavy')
                        except ValueError:
                            pass

                ax.set_title("SEVERE DEFECT RATE (%)", fontweight='bold', color='#d62728')
                ax.set_ylabel("")
                ax.set_xlabel("")
                fig.tight_layout()
                plt.savefig("export_heatmap.png", bbox_inches='tight', dpi=150)
                st.pyplot(fig)
                plt.close(fig)

        else:
            st.success("✅ Process is completely stable. No severe defect patterns detected.")

        st.markdown("---")
        st.header("📈 Global I-MR Stability Tracking (Severe Defects: B+ and Below)")
        df_severe_global = df_unique_global[
            (df_unique_global['B+'] > 0) | (df_unique_global['B'] > 0)
        ].sort_values(by='烤三生產日期').reset_index(drop=True)

        if not df_severe_global.empty:
            for feat in ['YS', 'TS', 'EL', 'YPE']:
                if feat in df_severe_global.columns:
                    valid_data = df_severe_global.dropna(
                        subset=[feat, '烤三生產日期']
                    ).reset_index(drop=True)
                    if len(valid_data) > 1:
                        st.markdown(f"#### 🛡️ Global Stability: **{feat}**")
                        dates = valid_data['烤三生產日期']
                        vals = valid_data[feat].values
                        x_seq = np.arange(len(vals))
                        mean_v = np.mean(vals)

                        fig, (ax1, ax2) = plt.subplots(
                            2, 1, figsize=(14, 8), gridspec_kw={'height_ratios': [2, 1]}
                        )
                        ax1.plot(x_seq, vals, marker='o', ms=5, lw=1.5,
                                 color='#004C99', alpha=0.9, label=f"Value ({feat})")
                        ax1.axhline(mean_v, color='black', ls='--', lw=1.5,
                                    label=f'Mean: {mean_v:.1f}')
                        ax1.text(x_seq[-1], mean_v, f" Mean: {mean_v:.1f}",
                                 va='bottom', color='black', fontweight='bold')
                        if feat in GLOBAL_SPECS:
                            s = GLOBAL_SPECS[feat]
                            if s['min']:
                                ax1.axhline(s['min'], color='red', lw=2)
                                ax1.text(x_seq[-1], s['min'], f" Min: {s['min']}",
                                         va='bottom', color='red', fontweight='bold')
                            if s['max']:
                                ax1.axhline(s['max'], color='red', lw=2)
                                ax1.text(x_seq[-1], s['max'], f" Max: {s['max']}",
                                         va='bottom', color='red', fontweight='bold')
                            v_x, v_y = [], []
                            for i, v in enumerate(vals):
                                if (s['min'] and v < s['min']) or (s['max'] and v > s['max']):
                                    v_x.append(i); v_y.append(v)
                            if v_x:
                                ax1.scatter(v_x, v_y, color='red', s=60, zorder=5)
                        for i in range(1, len(dates)):
                            if dates.iloc[i].year != dates.iloc[i - 1].year:
                                ax1.axvline(i, color='gray', ls=':', alpha=0.5)
                                ax1.text(i, ax1.get_ylim()[1],
                                         f" {dates.iloc[i].year}", fontsize=10, va='top')
                        ax1.set_title(f"Individual Chart (I) - {feat}", fontweight='bold')
                        ax1.legend(loc='upper right', fontsize=9, bbox_to_anchor=(1.15, 1))
                        ax1.set_xticks([])

                        mr = np.abs(np.diff(vals))
                        mr_mean = np.mean(mr)
                        ucl_mr = 3.267 * mr_mean
                        ax2.plot(x_seq[1:], mr, marker='o', ms=5, lw=1.5,
                                 color='#4B0082', alpha=0.9, label="Moving Range")
                        ax2.axhline(mr_mean, color='black', ls='--', lw=1.5,
                                    label=f'MR Mean: {mr_mean:.1f}')
                        ax2.axhline(ucl_mr, color='red', ls=':', lw=1.5,
                                    label=f'UCL: {ucl_mr:.1f}')
                        hv_x, hv_y = [], []
                        for i, m_val in enumerate(mr):
                            if m_val > ucl_mr:
                                hv_x.append(i + 1); hv_y.append(m_val)
                        if hv_x:
                            ax2.scatter(hv_x, hv_y, color='red', s=40, zorder=5)
                        ax2.set_title("Moving Range Chart (MR)", fontweight='bold')
                        ax2.legend(loc='upper right', fontsize=9, bbox_to_anchor=(1.15, 1))
                        step = max(1, len(x_seq) // 12)
                        ax2.set_xticks(x_seq[::step])
                        ax2.set_xticklabels(
                            dates.dt.strftime('%Y-%m-%d').iloc[::step], rotation=45, ha='right'
                        )
                        fig.tight_layout()
                        plt.savefig(f"export_imr_global_{feat}.png", bbox_inches='tight', dpi=150)
                        st.pyplot(fig)
                        plt.close(fig)
    # ==========================================================
    # TAB 4: I-MR CHART
    # ==========================================================
    with tab4:
        st.header("📈 Task 4: I-MR Stability Tracking (Chronological)")
        st.info("Analysis based on production sequence from 2024 to 2026. Red dots = Out of Spec.")

        @st.fragment
        def render_tab4():
            imr_periods = ["All Periods"] + sorted(
                df_filtered['Time_Group'].dropna().unique().tolist()
            )
            imr_thicks = sorted(df_filtered['Actual_Thickness'].dropna().unique())
            imr_mats = sorted(df_filtered['HR_Material'].astype(str).unique())
            c1, c2, c3 = st.columns(3)
            sel_p = c1.selectbox("Filter Period:", imr_periods, key="t4_p")
            sel_t = c2.selectbox("Filter Thickness:", imr_thicks, key="t4_t")
            sel_m = c3.selectbox("Filter Material:", imr_mats, key="t4_m")

            if sel_p == "All Periods":
                imr_df = df_filtered[
                    (df_filtered['Actual_Thickness'] == sel_t) &
                    (df_filtered['HR_Material'] == sel_m)
                ].drop_duplicates(subset=['Row_ID'])
            else:
                imr_df = df_filtered[
                    (df_filtered['Time_Group'] == sel_p) &
                    (df_filtered['Actual_Thickness'] == sel_t) &
                    (df_filtered['HR_Material'] == sel_m)
                ]
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

                            fig, (ax1, ax2) = plt.subplots(
                                2, 1, figsize=(12, 7), gridspec_kw={'height_ratios': [2, 1]}
                            )
                            ax1.plot(x_seq, vals, marker='o', ms=5, lw=1.5,
                                     color='#004C99', alpha=0.9, label=feat)
                            ax1.axhline(mean_v, color='black', ls='--', lw=1.5, label='Mean')
                            if feat in GLOBAL_SPECS:
                                s = GLOBAL_SPECS[feat]
                                if s['min']:
                                    ax1.axhline(s['min'], color='red', lw=2)
                                if s['max']:
                                    ax1.axhline(s['max'], color='red', lw=2)
                                v_x, v_y = [], []
                                for i, v in enumerate(vals):
                                    if (s['min'] and v < s['min']) or (s['max'] and v > s['max']):
                                        v_x.append(i); v_y.append(v)
                                if v_x:
                                    ax1.scatter(v_x, v_y, color='red', s=60, zorder=5)
                            if sel_p == "All Periods":
                                for i in range(1, len(dates)):
                                    if dates.iloc[i].year != dates.iloc[i - 1].year:
                                        ax1.axvline(i, color='gray', ls=':', alpha=0.5)
                                        ax1.text(i, ax1.get_ylim()[1],
                                                 f" {dates.iloc[i].year}", fontsize=10, va='top')
                            ax1.set_title(f"Individual Chart (I) - {feat}", fontweight='bold')
                            ax1.legend(loc='upper right', fontsize=8)
                            ax1.set_xticks([])

                            mr = np.abs(np.diff(vals))
                            mr_mean = np.mean(mr)
                            ucl_mr = 3.267 * mr_mean
                            ax2.plot(x_seq[1:], mr, marker='o', ms=5, lw=1.5,
                                     color='#4B0082', alpha=0.9)
                            ax2.axhline(mr_mean, color='black', ls='--', lw=1.5)
                            ax2.axhline(ucl_mr, color='red', ls=':', lw=1.5)
                            hv_x, hv_y = [], []
                            for i, m_val in enumerate(mr):
                                if m_val > ucl_mr:
                                    hv_x.append(i + 1); hv_y.append(m_val)
                            if hv_x:
                                ax2.scatter(hv_x, hv_y, color='red', s=40, zorder=5)
                            ax2.set_title("Moving Range Chart (MR)", fontweight='bold')
                            step = max(1, len(x_seq) // 12)
                            ax2.set_xticks(x_seq[::step])
                            ax2.set_xticklabels(
                                dates.dt.strftime('%Y-%m-%d').iloc[::step], rotation=45, ha='right'
                            )
                            fig.tight_layout()
                            plt.savefig(f"export_imr_{feat}.png", bbox_inches='tight', dpi=150)
                            st.pyplot(fig)
                            plt.close(fig)
            else:
                st.warning("No data found for the selected combination.")

        render_tab4()

    # ==========================================================
   # ==========================================================
# ==========================================================
    # --- TAB 5: TAIL SCRAP ANALYSIS (COIL-ID AWARE) ---
    # ==========================================================
    with tab5:
        st.header("5. Tail Scrap & Length Rejection Analysis")
        st.info(
            "Analysis of tail scrap rejection rate based on Measured Length "
            "and Tail Scrap Rejected.\n\n"
            "⚙️ **Coil-ID Aware Logic:** A coil running through the line multiple times "
            "is handled correctly. **Total Length** is based on the FIRST pass of ALL coils "
            "(including those with zero scrap). **Total Scrap** is the sum across all passes."
        )

        col_length = '實測長度'
        col_scrap  = '尾料剔退'

        if col_length not in df_filtered.columns or col_scrap not in df_filtered.columns:
            missing = [c for c in [col_length, col_scrap] if c not in df_filtered.columns]
            st.error(f"Missing columns: {missing}. Please check your data file.")
        else:
            # ----------------------------------------------------------------
            # 1. PRE-PROCESS & CLEANING
            # ----------------------------------------------------------------
            df_scrap_all = df_filtered.copy()
            
            # Fill NaN lengths and scraps with 0 immediately to ensure no data is lost
            df_scrap_all[col_length] = pd.to_numeric(df_scrap_all[col_length], errors='coerce').fillna(0)
            df_scrap_all[col_scrap]  = pd.to_numeric(df_scrap_all[col_scrap],  errors='coerce').fillna(0)
            
            # Secure Coil IDs: Prevent blank IDs from merging together
            df_scrap_all[COIL_ID_COL] = df_scrap_all[COIL_ID_COL].astype(str).str.strip().replace(['nan', 'None', '', 'NaN'], np.nan)
            missing_mask = df_scrap_all[COIL_ID_COL].isna()
            if missing_mask.any():
                df_scrap_all.loc[missing_mask, COIL_ID_COL] = [f"UNKNOWN_ID_{i}" for i in df_scrap_all[missing_mask].index]

            # --- NEW: Q4 Drilldown Logic for Consistent Timeline ---
            def refine_q4_scrap(row):
                if row['Time_Group'] == '2025 Q4' and pd.notna(row['烤三生產日期']):
                    if row['烤三生產日期'].month == 10:
                        return '2025 Q4 (Oct)'
                    else:
                        return '2025 Q4 (Nov-Dec)'
                return row['Time_Group']
                
            if '烤三生產日期' in df_scrap_all.columns:
                df_scrap_all['Time_Group'] = df_scrap_all.apply(refine_q4_scrap, axis=1)

            # Local Order Map updated to include the new Q4 periods
            local_order_map = {
                "2024 (Full Year)": 1,
                "2025 H1 (Until 06/28)": 2,
                "2025 Q3 (06/29 - 09/30)": 3,
                "2025 Q4 (Oct)": 4.1,
                "2025 Q4 (Nov-Dec)": 4.2,
                "2025 (Full Year)": 5,
                "2026 Q1": 6
            }

            # ----------------------------------------------------------------
            # 2. TOTAL SCRAP (Sum of ALL passes per coil per period)
            # ----------------------------------------------------------------
            scrap_sum_df = df_scrap_all.groupby(['Time_Group', COIL_ID_COL], as_index=False)[col_scrap].sum()
            scrap_sum_df.rename(columns={col_scrap: 'Total_Scrap'}, inplace=True)

            # ----------------------------------------------------------------
            # 3. ORIGINAL LENGTH (First pass only per coil per period)
            # ----------------------------------------------------------------
            # Sort by date so the earliest pass appears at the top
            df_sorted = df_scrap_all.sort_values(by=['Time_Group', COIL_ID_COL, '烤三生產日期'])
            
            # Drop duplicates to keep ONLY the first occurrence
            first_pass_df = df_sorted.drop_duplicates(subset=['Time_Group', COIL_ID_COL], keep='first')
            
            length_df = first_pass_df[['Time_Group', COIL_ID_COL, col_length, 'Actual_Thickness', 'HR_Material']].copy()
            length_df.rename(columns={col_length: 'Original_Length'}, inplace=True)

            # ----------------------------------------------------------------
            # 4. PASS COUNT (How many times a coil ran)
            # ----------------------------------------------------------------
            pass_count_df = df_scrap_all.groupby(['Time_Group', COIL_ID_COL], as_index=False).size()
            pass_count_df.rename(columns={'size': 'Pass_Count'}, inplace=True)

            # ----------------------------------------------------------------
            # 5. MASTER MERGE & RATE CALCULATION
            # ----------------------------------------------------------------
            df_coil = length_df.merge(scrap_sum_df, on=['Time_Group', COIL_ID_COL], how='left') \
                               .merge(pass_count_df, on=['Time_Group', COIL_ID_COL], how='left')
            
            # Calculate Rate: Total Scrap / Original Length
            df_coil['Scrap_Rate (%)'] = np.where(
                df_coil['Original_Length'] > 0,
                (df_coil['Total_Scrap'] / df_coil['Original_Length'] * 100),
                0
            ).round(2)

            # --- Display Multi-Pass Alert ---
            multi_pass = df_coil[df_coil['Pass_Count'] > 1]
            if not multi_pass.empty:
                st.info(
                    f"🔄 **{len(multi_pass)} coils** detected running through the line "
                    f"**more than once**. Scrap accumulated across all passes; "
                    f"length counted only once (original pass)."
                )
                with st.expander("📋 View Multi-Pass Coil Detail"):
                    st.dataframe(
                        multi_pass[[
                            COIL_ID_COL, 'Time_Group', 'Actual_Thickness', 'HR_Material',
                            'Original_Length', 'Total_Scrap', 'Pass_Count', 'Scrap_Rate (%)'
                        ]].sort_values('Pass_Count', ascending=False).style
                          .background_gradient(subset=['Scrap_Rate (%)'], cmap='Oranges')
                          .format({
                              'Original_Length': '{:,.1f}', 'Total_Scrap': '{:,.1f}',
                              'Scrap_Rate (%)': '{:.2f}%'
                          }),
                        use_container_width=True, hide_index=True
                    )

            # ================================================================
            # SECTION 1: SCRAP RATE BY TIME PERIOD
            # ================================================================
            st.markdown("---")
            st.subheader("📅 Section 1: Scrap Rate by Time Period")
            scrap_by_period = (
                df_coil.groupby('Time_Group').agg(
                    Total_Length=('Original_Length', 'sum'),
                    Total_Scrap=('Total_Scrap', 'sum'),
                    Coil_Count=(COIL_ID_COL, 'count'),
                ).reset_index()
            )
            scrap_by_period['Scrap_Rate (%)'] = np.where(
                scrap_by_period['Total_Length'] > 0,
                (scrap_by_period['Total_Scrap'] / scrap_by_period['Total_Length'] * 100),
                0
            ).round(2)
            scrap_by_period['Sort_Key'] = scrap_by_period['Time_Group'].map(local_order_map).fillna(99)
            scrap_by_period = scrap_by_period.sort_values('Sort_Key').drop(columns=['Sort_Key'])
            scrap_by_period.rename(columns={'Time_Group': 'Time Period'}, inplace=True)

            st.dataframe(
                scrap_by_period.style
                    .background_gradient(subset=['Scrap_Rate (%)'], cmap='Reds')
                    .format({
                        'Total_Length': '{:,.1f}', 'Total_Scrap': '{:,.1f}',
                        'Coil_Count': '{:.0f}', 'Scrap_Rate (%)': '{:.2f}%'
                    }),
                use_container_width=True, hide_index=True
            )

            fig1, ax1 = plt.subplots(figsize=(10, 4))
            colors = plt.cm.Reds(np.linspace(0.3, 0.85, len(scrap_by_period)))
            bars = ax1.bar(
                scrap_by_period['Time Period'], scrap_by_period['Scrap_Rate (%)'],
                color=colors, edgecolor='white'
            )
            for bar, val in zip(bars, scrap_by_period['Scrap_Rate (%)']):
                ax1.text(
                    bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.05,
                    f"{val:.2f}%", ha='center', va='bottom',
                    fontsize=10, fontweight='bold', color='#333'
                )
            ax1.set_title("Tail Scrap Rate (%) by Time Period", fontweight='bold', fontsize=13)
            ax1.set_xlabel("")
            ax1.set_ylabel("Scrap Rate (%)")
            ax1.tick_params(axis='x', rotation=20)
            ax1.set_ylim(0, scrap_by_period['Scrap_Rate (%)'].max() * 1.3 + 0.1)
            fig1.tight_layout()
            plt.savefig("export_tab5_scrap_by_period.png", bbox_inches='tight', dpi=150)
            st.pyplot(fig1)
            plt.close(fig1)

            # ================================================================
            # SECTION 2: BY PERIOD + THICKNESS + MATERIAL
            # ================================================================
            st.markdown("---")
            st.subheader("📏 Section 2: Scrap Rate by Period / Thickness / Material")
            scrap_detail = (
                df_coil.groupby(['Time_Group', 'Actual_Thickness', 'HR_Material']).agg(
                    Total_Length=('Original_Length', 'sum'),
                    Total_Scrap=('Total_Scrap', 'sum'),
                    Coil_Count=(COIL_ID_COL, 'count'),
                ).reset_index()
            )
            scrap_detail['Scrap_Rate (%)'] = np.where(
                scrap_detail['Total_Length'] > 0,
                (scrap_detail['Total_Scrap'] / scrap_detail['Total_Length'] * 100),
                0
            ).round(2)
            scrap_detail['Sort_Key'] = scrap_detail['Time_Group'].map(local_order_map).fillna(99)
            scrap_detail = scrap_detail.sort_values(
                by=['Sort_Key', 'Actual_Thickness', 'HR_Material']
            ).drop(columns=['Sort_Key'])
            scrap_detail = scrap_detail[scrap_detail['Total_Length'] > 0]

            st.dataframe(
                scrap_detail.style
                    .background_gradient(subset=['Scrap_Rate (%)'], cmap='Reds')
                    .format({
                        'Actual_Thickness': '{:.2f}mm', 'Total_Length': '{:,.1f}',
                        'Total_Scrap': '{:,.1f}', 'Coil_Count': '{:.0f}',
                        'Scrap_Rate (%)': '{:.2f}%'
                    }),
                use_container_width=True, hide_index=True
            )

            # Chart: by Thickness
            st.markdown("#### 📊 Scrap Rate by Thickness across Periods")
            pivot_scrap = scrap_detail.groupby(
                ['Time_Group', 'Actual_Thickness']
            )['Scrap_Rate (%)'].mean().unstack()
            pivot_scrap = pivot_scrap.reindex(
                sorted(pivot_scrap.index, key=lambda x: local_order_map.get(x, 99))
            )
            fig2, ax2 = plt.subplots(figsize=(12, 5))
            pivot_scrap.plot(kind='bar', ax=ax2, colormap='YlOrRd', edgecolor='white')
            for container in ax2.containers:
                labels = [
                    f"{v.get_height():.1f}%" if v.get_height() > 0 else ""
                    for v in container
                ]
                ax2.bar_label(container, labels=labels, label_type='edge',
                              fontsize=8, fontweight='bold', color='#333', padding=3)
            ax2.set_ylim(0, ax2.get_ylim()[1] * 1.18)
            ax2.set_title("Avg Scrap Rate (%) by Period & Thickness", fontweight='bold', fontsize=13)
            ax2.set_xlabel("")
            ax2.set_ylabel("Scrap Rate (%)")
            ax2.legend(title="Thickness (mm)", bbox_to_anchor=(1.02, 1), loc='upper left')
            ax2.tick_params(axis='x', rotation=25)
            fig2.tight_layout()
            plt.savefig("export_tab5_scrap_by_thickness.png", bbox_inches='tight', dpi=150)
            st.pyplot(fig2)
            plt.close(fig2)

            # Chart: by Material
            st.markdown("#### 🧱 Scrap Rate by Material across Periods")
            pivot_scrap_mat = scrap_detail.groupby(
                ['Time_Group', 'HR_Material']
            )['Scrap_Rate (%)'].mean().unstack()
            pivot_scrap_mat = pivot_scrap_mat.reindex(
                sorted(pivot_scrap_mat.index, key=lambda x: local_order_map.get(x, 99))
            )
            MATERIAL_COLORS = [
                '#E63946', '#2A9D8F', '#E9C46A', '#457B9D',
                '#F4A261', '#6A4C93', '#264653', '#A8DADC',
            ]
            mat_colors = MATERIAL_COLORS[:len(pivot_scrap_mat.columns)]
            fig3, ax3 = plt.subplots(figsize=(12, 5))
            pivot_scrap_mat.plot(
                kind='bar', ax=ax3, color=mat_colors, edgecolor='white', linewidth=0.8
            )
            for container in ax3.containers:
                labels = [
                    f"{v.get_height():.1f}%" if v.get_height() > 0 else ""
                    for v in container
                ]
                ax3.bar_label(container, labels=labels, label_type='edge',
                              fontsize=8, fontweight='bold', color='#333', padding=3)
            ax3.set_ylim(0, pivot_scrap_mat.max().max() * 1.25 + 1)
            ax3.set_title("Avg Scrap Rate (%) by Period & Material", fontweight='bold', fontsize=13)
            ax3.set_xlabel("")
            ax3.set_ylabel("Scrap Rate (%)")
            ax3.legend(title="Material", bbox_to_anchor=(1.02, 1), loc='upper left', framealpha=0.9)
            ax3.tick_params(axis='x', rotation=25)
            ax3.grid(axis='y', linestyle='--', alpha=0.4)
            ax3.set_axisbelow(True)
            fig3.tight_layout()
            plt.savefig("export_tab5_scrap_by_material.png", bbox_inches='tight', dpi=150)
            st.pyplot(fig3)
            plt.close(fig3)

            # ================================================================
            # SECTION 3: RAW vs CORRECTED COMPARISON
            # ================================================================
            st.markdown("---")
            st.subheader("🔍 Section 3: Impact of Coil-ID Fix (Raw vs Corrected)")
            st.caption(
                "A higher corrected rate = the raw method was undercounting "
                "(Denominator artificially inflated due to length duplication across passes)."
            )

            raw_by_period = (
                df_scrap_all.groupby('Time_Group')
                .apply(lambda x: x[col_scrap].sum() / x[col_length].sum() * 100
                       if x[col_length].sum() > 0 else 0)
                .reset_index()
            )
            raw_by_period.columns = ['Time_Group', 'Scrap_Rate_Raw (%)']

            corrected_by_period = scrap_by_period[['Time Period', 'Scrap_Rate (%)']].rename(
                columns={'Time Period': 'Time_Group', 'Scrap_Rate (%)': 'Scrap_Rate_Corrected (%)'}
            )
            compare_df = raw_by_period.merge(corrected_by_period, on='Time_Group', how='inner')
            compare_df['Difference (pp)'] = (
                compare_df['Scrap_Rate_Corrected (%)'] - compare_df['Scrap_Rate_Raw (%)']
            ).round(3)
            compare_df['Sort_Key'] = compare_df['Time_Group'].map(local_order_map).fillna(99)
            compare_df = compare_df.sort_values('Sort_Key').drop(columns=['Sort_Key'])

            def highlight_diff(val):
                if pd.isna(val): return ''
                if val > 0.5:  return 'color: #c00000; font-weight: bold'
                if val > 0.1:  return 'color: #e06000; font-weight: bold'
                return 'color: #2e7d32'

            st.dataframe(
                compare_df.style
                    .map(highlight_diff, subset=['Difference (pp)'])
                    .format({
                        'Scrap_Rate_Raw (%)':        '{:.3f}%',
                        'Scrap_Rate_Corrected (%)': '{:.3f}%',
                        'Difference (pp)':          '{:+.3f}'
                    }),
                use_container_width=True, hide_index=True
            )

            fig4, ax4 = plt.subplots(figsize=(11, 4))
            x_c = np.arange(len(compare_df))
            w   = 0.35
            ax4.bar(x_c - w / 2, compare_df['Scrap_Rate_Raw (%)'],
                    width=w, label='Raw (old method)', color='#9ecae1', edgecolor='white')
            ax4.bar(x_c + w / 2, compare_df['Scrap_Rate_Corrected (%)'],
                    width=w, label='Corrected (Coil-ID aware)', color='#d62728', edgecolor='white')
            ax4.set_xticks(x_c)
            ax4.set_xticklabels(compare_df['Time_Group'], rotation=20, ha='right')
            ax4.set_ylabel("Scrap Rate (%)")
            ax4.set_title("Raw vs Corrected Scrap Rate by Period", fontweight='bold', fontsize=13)
            ax4.legend()
            fig4.tight_layout()
            plt.savefig("export_tab5_comparison.png", bbox_inches='tight', dpi=150)
            st.pyplot(fig4)
            plt.close(fig4)

            # ================================================================
            # DOWNLOAD
            # ================================================================
            st.markdown("---")
            output_scrap = io.BytesIO()
            try:
                with pd.ExcelWriter(output_scrap, engine='xlsxwriter') as writer:
                    scrap_by_period.to_excel(writer, index=False, sheet_name='Scrap_By_Period')
                    scrap_detail.to_excel(
                        writer, index=False, sheet_name='Scrap_By_Period_Thick_Mat'
                    )
                    compare_df.to_excel(writer, index=False, sheet_name='Raw_vs_Corrected')
                    if not multi_pass.empty:
                        multi_pass.to_excel(writer, index=False, sheet_name='Multi_Pass_Coils')
                st.download_button(
                    label="📥 Download Scrap & Length Analysis Excel",
                    data=output_scrap.getvalue(),
                    file_name="Scrap_Length_Analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.warning(f"Could not generate Excel: {e}")
     # ==========================================================
# ==========================================================
    # --- TAB 6: DYNAMIC CONTROL LIMITS & I-MR CHARTS ---
    # ==========================================================
    with tab6:
        st.header("6. Dynamic Control Limits & I-MR Trending (Mill vs Release)")
        st.info(
            "Establish statistical control limits based on actual production data.\n\n"
            "🛡️ **Mill Range:** Tighter internal limits for early warning.\n"
            "🚛 **Release Range:** Wider limits determining product acceptance."
        )
        
        # --- DYNAMIC CONTROLS ---
        st.markdown("### ⚙️ Parameter Configuration")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            sigma_release = st.number_input("Release Range (Sigma)", min_value=1.0, max_value=6.0, value=3.0, step=0.1)
        with col2:
            sigma_mill = st.number_input("Mill Range (Sigma)", min_value=0.5, max_value=4.0, value=2.0, step=0.1)
        with col3:
            iqr_k = st.number_input("IQR Filter Factor (k)", min_value=1.0, max_value=4.0, value=1.5, step=0.1)
        with col4:
            chart_method = st.radio("I-MR Chart Limits Based On:", ["Standard Method", "IQR Filtered Method"])

        # Spec limits mapping
        spec_limits = {"YS": (GLOBAL_SPECS.get('YS', {}).get('min'), GLOBAL_SPECS.get('YS', {}).get('max')), 
                       "TS": (GLOBAL_SPECS.get('TS', {}).get('min'), GLOBAL_SPECS.get('TS', {}).get('max')), 
                       "EL": (GLOBAL_SPECS.get('EL', {}).get('min'), GLOBAL_SPECS.get('EL', {}).get('max')), 
                       "YPE": (GLOBAL_SPECS.get('YPE', {}).get('min'), GLOBAL_SPECS.get('YPE', {}).get('max'))}

        # --- CALCULATION LOGIC (STANDARD & IQR) ---
        def calculate_stats(v_arr, w_arr, k_factor):
            # 1. Standard Method
            m_std = np.average(v_arr, weights=w_arr)
            s_std = np.sqrt(np.average((v_arr - m_std)**2, weights=w_arr))
            
            # 2. IQR Filtered Method
            try:
                # Expand values based on weights for accurate percentiles
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

        all_export_data = []
        
        # ================================================================
        # 1. OVERALL FACTORY
        # ================================================================
        st.markdown("---")
        st.subheader("🌐 Overall Factory Performance Goals")
        
        overall_export_data = [] 
        
        for feat in ['YS', 'TS', 'EL', 'YPE']:
            if feat in df_filtered.columns:
                df_ov = df_filtered[[feat, 'Total_Qty']].dropna(subset=[feat]).copy()
                df_ov = df_ov[df_ov['Total_Qty'] > 0]
                
                low, high = spec_limits.get(feat, (None, None))
                spec_str_ov = f"{int(low)}-{int(high)}" if pd.notnull(low) and pd.notnull(high) else (f">={int(low)}" if pd.notnull(low) else "N/A")

                if not df_ov.empty:
                    v, w = df_ov[feat].values, df_ov['Total_Qty'].values
                    # Safe check for integer weights
                    w = np.where(pd.isna(w) | (w <= 0), 1, w).astype(int)
                    
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
                            "Current Limit": spec_str_ov, 
                            "TARGET GOAL": int(round(m_val)),
                            "TOLERANCE": round(s_val, 2),
                            f"MILL RANGE {sigma_mill}σ": f"{mill_lower} - {int(round(m_val + sigma_mill*s_val))}",
                            f"RELEASE RANGE {sigma_release}σ": f"{release_lower} - {int(round(m_val + sigma_release*s_val))}"
                        })
        
        st.dataframe(pd.DataFrame(overall_export_data), use_container_width=True, hide_index=True)

        # ================================================================
        # 2. LOCAL THICKNESS & I-MR CHARTS
        # ================================================================
        st.markdown("---")
        st.subheader("🔍 Local Control Limits & I-MR Trending")
        plot_data_dict = {}

        thickness_list = sorted(df_filtered['Actual_Thickness'].dropna().unique())

        for thick in thickness_list:
            st.markdown(f"#### 📏 Thickness Category: **{thick}mm**")
            df_t = df_filtered[df_filtered['Actual_Thickness'] == thick].sort_values(by='烤三生產日期')
            plot_data_dict[thick] = {}
            thick_status = []
            
            for feat in ['YS', 'TS', 'EL', 'YPE']:
                if feat in df_t.columns:
                    temp_calc = df_t[[feat, 'Total_Qty', '烤三生產日期']].dropna(subset=[feat]).copy()
                    temp_calc = temp_calc[temp_calc['Total_Qty'] > 0]
                
                    low, high = spec_limits.get(feat, (None, None))
                    spec_str = f"{int(low)}-{int(high)}" if pd.notnull(low) and pd.notnull(high) else (f">={int(low)}" if pd.notnull(low) else "N/A")

                    if not temp_calc.empty:
                        v, w = temp_calc[feat].values, temp_calc['Total_Qty'].values
                        w = np.where(pd.isna(w) | (w <= 0), 1, w).astype(int)
                        
                        (m_std, s_std), (m_iqr, s_iqr) = calculate_stats(v, w, iqr_k)
                        
                        plot_data_dict[thick][feat] = {
                            'values': v, 
                            'mean_std': m_std, 'std_std': s_std,
                            'mean_iqr': m_iqr, 'std_iqr': s_iqr
                        }
                        
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
                                "Current Limit": spec_str,
                                "TARGET GOAL": int(round(m_val)),
                                "TOLERANCE": round(s_val, 2),
                                f"MILL RANGE {sigma_mill}σ": f"{mill_lower} - {int(round(m_val + sigma_mill*s_val))}",
                                f"RELEASE RANGE {sigma_release}σ": f"{release_lower} - {int(round(m_val + sigma_release*s_val))}"
                            }
                            thick_status.append(row)
                            
                            exp_row = row.copy()
                            exp_row['Thickness'] = thick
                            all_export_data.append(exp_row)

            st.dataframe(pd.DataFrame(thick_status), use_container_width=True, hide_index=True)
            
            # --- I-MR CHARTS PLOTTING ---
            cols_imr = st.columns(2)
            top4 = [f for f in ['YS', 'TS', 'EL', 'YPE'] if f in plot_data_dict[thick]]
            
            for idx, f in enumerate(top4):
                with cols_imr[idx % 2]:
                    d = plot_data_dict[thick][f]
                    v = d['values']
                    
                    if chart_method == "Standard Method":
                        mv, sv = d['mean_std'], d['std_std']
                    else:
                        mv, sv = d['mean_iqr'], d['std_iqr']
                        
                    if len(v) > 1:
                        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 7), gridspec_kw={'height_ratios': [2, 1]})
                        
                        # Limits based on Release Range Sigma
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
                        plt.close(fig)
            st.markdown("---")

        # ================================================================
        # EXPORT SECTION
        # ================================================================
        output_opt = io.BytesIO()
        try:
            with pd.ExcelWriter(output_opt, engine='xlsxwriter') as writer:
                pd.DataFrame(all_export_data).to_excel(writer, index=False, sheet_name='Control_Optimization')
            st.download_button(
                label="📥 Download QC Optimization Report (Excel)",
                data=output_opt.getvalue(),
                file_name="QC_Optimization_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.warning(f"Could not generate Excel: {e}")        
    # ==========================================================
    # EXPORT PDF
    # ==========================================================
    st.sidebar.header("📥 Export PDF Report")
    st.sidebar.info(
        "💡 Tip: Browse through all tabs first so all charts are generated before exporting."
    )

    if st.sidebar.button("🖨️ Generate PDF Report"):
        try:
            def get_image_height_mm(img_path, width_mm):
                with PILImage.open(img_path) as img:
                    w_px, h_px = img.size
                return width_mm * h_px / w_px

            def add_section_title(pdf, title):
                pdf.set_font('Arial', 'B', 14)
                pdf.set_fill_color(230, 230, 230)
                pdf.cell(0, 10, title, ln=True, align='C', fill=True)
                pdf.ln(3)

            def add_images_to_pdf(pdf, img_files, title, page_width_mm=262, margin_x=15):
                if not img_files:
                    return
                pdf.add_page()
                add_section_title(pdf, title)
                bottom_margin = 10
                page_height = 210
                y_cursor = pdf.get_y()
                for img_path in img_files:
                    if not os.path.exists(img_path):
                        continue
                    img_h = get_image_height_mm(img_path, page_width_mm)
                    if y_cursor + img_h > page_height - bottom_margin:
                        pdf.add_page()
                        y_cursor = 15
                    pdf.image(img_path, x=margin_x, y=y_cursor, w=page_width_mm)
                    y_cursor += img_h + 4
                    pdf.set_y(y_cursor)

            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=False)

            if os.path.exists("export_heatmap.png"):
                pdf.add_page()
                add_section_title(pdf, "1. DEFECT HOTSPOT DIAGNOSTIC MAP (SEVERE DEFECTS)")
                img_w = 255
                img_h = get_image_height_mm("export_heatmap.png", img_w)
                pdf.image("export_heatmap.png", x=15, y=pdf.get_y(), w=img_w)

            tab1_files = sorted([
                f for f in os.listdir('.') if f.startswith("export_tab1_") and f.endswith(".png")
            ])
            if tab1_files:
                add_images_to_pdf(pdf, tab1_files, "2. YIELD SUMMARY CHARTS (TAB 1)")

            tab2_files = sorted([
                f for f in os.listdir('.') if f.startswith("export_tab2_") and f.endswith(".png")
            ])
            if tab2_files:
                add_images_to_pdf(pdf, tab2_files, "3. DISTRIBUTION ANALYSIS (TAB 2)")

            global_imr_files = [
                f"export_imr_global_{feat}.png" for feat in ['YS', 'TS', 'EL', 'YPE']
                if os.path.exists(f"export_imr_global_{feat}.png")
            ]
            if global_imr_files:
                add_images_to_pdf(
                    pdf, global_imr_files,
                    "4. GLOBAL PROCESS STABILITY - ALL SEVERE DEFECTS (TAB 3)"
                )

            filtered_imr_files = [
                f"export_imr_{feat}.png" for feat in ['YS', 'TS', 'EL', 'YPE']
                if os.path.exists(f"export_imr_{feat}.png")
            ]
            if filtered_imr_files:
                add_images_to_pdf(
                    pdf, filtered_imr_files,
                    "5. SPECIFIC I-MR TRACKING - FILTERED SEGMENT (TAB 4)"
                )

            tab5_files = sorted([
                f for f in os.listdir('.') if f.startswith("export_tab5_") and f.endswith(".png")
            ])
            if tab5_files:
                add_images_to_pdf(pdf, tab5_files, "6. SCRAP ANALYSIS - COIL-ID AWARE (TAB 5)")

            pdf.output("Quality_Visual_Report.pdf")
            with open("Quality_Visual_Report.pdf", "rb") as f:
                st.sidebar.download_button(
                    label="✅ Click to Download PDF Report",
                    data=f.read(),
                    file_name="Quality_Visual_Report.pdf",
                    mime="application/pdf"
                )
            st.sidebar.success("🎉 PDF Generated Successfully!")

        except Exception as e:
            st.sidebar.error(f"⚠️ Error generating PDF: {e}")
