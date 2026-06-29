import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Weekly Performance 綜合數據轉換器", layout="wide")

st.title("📊 Weekly Performance 綜合數據自動轉換工具")
st.write("上傳原始 `OverseaDataMarket` 檔案，自動生成 New Central 所需的「MM概況/城市/星級/各國籍/Eztravel」多維度分析週對週報表。")

# 1. 日期區間設定
st.subheader("🗓️ 步驟一：設定比對區間（格式：YYYY-MM-DD）")
col1, col2 = st.columns(2)
with col1:
    lw_start = st.text_input("上週 開始日期", "2026-06-12")
    lw_end = st.text_input("上週 結束日期", "2026-06-18")
with col2:
    tw_start = st.text_input("當週 開始日期", "2026-06-19")
    tw_end = st.text_input("當週 結束日期", "2026-06-25")

# 2. 檔案上傳
st.subheader("📂 步驟二：上傳原始 Excel 檔案")
uploaded_file = st.file_uploader("請上傳 weekly performance 原始檔 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        st.info("⚡ 正在深度解析多維度數據，請稍候...")
        
        # 讀取檔案
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = 'TW' if 'TW' in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        
        # 基本欄位檢查
        required_cols = ['Book Time', 'RN', 'ordamount_afterdiscount', 'MM', 'City', 'Star', 'Ctrip/Trip site', 'Secondary Booking Channel']
        missing_cols = [c for c in required_cols if c not in df.columns]
        
        if missing_cols:
            st.error(f"錯誤：檔案缺少必要欄位：{missing_cols}")
        else:
            # 時間清洗
            df['Book Time'] = pd.to_datetime(df['Book Time'])
            df['Book Date'] = df['Book Time'].dt.date
            
            d_lw_start = pd.to_datetime(lw_start).date()
            d_lw_end = pd.to_datetime(lw_end).date()
            d_tw_start = pd.to_datetime(tw_start).date()
            d_tw_end = pd.to_datetime(tw_end).date()
            
            # 標籤化
            lw_label = f"{lw_start.split('-')[1]}/{lw_start.split('-')[2]}-{lw_end.split('-')[1]}/{lw_end.split('-')[2]}"
            tw_label = f"{tw_start.split('-')[1]}/{tw_start.split('-')[2]}-{tw_end.split('-')[1]}/{tw_end.split('-')[2]}"
            
            # 定義通用的週對週 WoW 計算函式
            def generate_wow_report(df_data, group_col, standard_list=None):
                lw_sub = df_data[(df_data['Book Date'] >= d_lw_start) & (df_data['Book Date'] <= d_lw_end)]
                tw_sub = df_data[(df_data['Book Date'] >= d_tw_start) & (df_data['Book Date'] <= d_tw_end)]
                
                lw_g = lw_sub.groupby(group_col).agg(lw_RN=('RN', 'sum'), lw_REV=('ordamount_afterdiscount', 'sum')).reset_index()
                tw_g = tw_sub.groupby(group_col).agg(tw_RN=('RN', 'sum'), tw_REV=('ordamount_afterdiscount', 'sum')).reset_index()
                
                if standard_list:
                    base = pd.DataFrame({group_col: standard_list})
                else:
                    all_items = sorted(list(set(lw_g[group_col].dropna().unique()) | set(tw_g[group_col].dropna().unique())))
                    base = pd.DataFrame({group_col: all_items})
                    
                res = pd.merge(base, lw_g, on=group_col, how='left')
                res = pd.merge(res, tw_g, on=group_col, how='left').fillna(0)
                
                res['lw_ADR'] = res.apply(lambda r: r['lw_REV'] / r['lw_RN'] if r['lw_RN'] > 0 else 0.0, axis=1)
                res['tw_ADR'] = res.apply(lambda r: r['tw_REV'] / r['tw_RN'] if r['tw_RN'] > 0 else 0.0, axis=1)
                
                res['WoW RN 成長率'] = res.apply(lambda r: f"{((r['tw_RN'] - r['lw_RN']) / r['lw_RN'] * 100):.1f}%" if r['lw_RN'] > 0 else ("100.0%" if r['tw_RN'] > 0 else "0.0%"), axis=1)
                res['WoW REV 成長率'] = res.apply(lambda r: f"{((r['tw_REV'] - r['lw_REV']) / r['lw_REV'] * 100):.1f}%" if r['lw_REV'] > 0 else ("100.0%" if r['tw_REV'] > 0 else "0.0%"), axis=1)
                
                res = res.rename(columns={
                    'lw_RN': f'{lw_label} RN', 'lw_REV': f'{lw_label} REV', 'lw_ADR': f'{lw_label} ADR',
                    'tw_RN': f'{tw_label} RN', 'tw_REV': f'{tw_label} REV', 'tw_ADR': f'{tw_label} ADR'
                })
                return res

            # --- 1. MM 概況清洗邏輯 ---
            standard_mms = ["Abby Cheng （鄭任妤）", "Hino Chi （籍喆）", "Sandy Su （蘇筱鈞）", "Wayne Wang （王俊明）", "Carol Lin （林芝榕）", "Flora Tsao （曹芳綺）", "Justin Lee （李岳勳）"]
            def clean_mm(raw):
                if pd.isna(raw): return "Others"
                for m in standard_mms:
                    if m.split(' （')[1].replace('）','').strip() in str(raw): return m
                return "Others"
            df['Cleaned_MM'] = df['MM'].apply(clean_mm)
            df_mm_report = generate_wow_report(df, 'Cleaned_MM', standard_mms).rename(columns={'Cleaned_MM': '行政區 (MM)'})
            
            # --- 2. 城市概況 ---
            df_city_report = generate_wow_report(df, 'City').rename(columns={'City': '城市'})
            
            # --- 3. 星級概況 ---
            df['Star_Label'] = df['Star'].apply(lambda x: f"{int(x)}星級" if pd.notna(x) and str(x).replace('.0','').isdigit() else "其他")
            df_star_report = generate_wow_report(df, 'Star_Label').rename(columns={'Star_Label': '酒店星級'})
            
            # --- 4. 各國籍/站點概況 ---
            df_site_report = generate_wow_report(df, 'Ctrip/Trip site').rename(columns={'Ctrip/Trip site': '旅客來源站點'})
            
            # --- 5. Eztravel (ez share) 專案表現 ---
            df_ez = df[df['Secondary Booking Channel'] == 'Eztravel']
            df_ez_report = generate_wow_report(df_ez, 'Cleaned_MM', standard_mms).rename(columns={'Cleaned_MM': 'Eztravel 專案 (MM)'})
            
            st.success("🎉 全維度報表解析成功！")
            
            # 網頁分頁預覽效果
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["👤 MM 概況", "🏙️ 城市概況", "⭐ 星級概況", "✈️ 旅客站點", "🤝 Eztravel (ez share)"])
            with tab1: st.dataframe(df_mm_report, use_container_width=True)
            with tab2: st.dataframe(df_city_report, use_container_width=True)
            with tab3: st.dataframe(df_star_report, use_container_width=True)
            with tab4: st.dataframe(df_site_report, use_container_width=True)
            with tab5: st.dataframe(df_ez_report, use_container_width=True)
            
            # 匯出多個分頁的 Excel 檔案
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_mm_report.to_excel(writer, sheet_name="MM概況", index=False)
                df_city_report.to_excel(writer, sheet_name="城市概況", index=False)
                df_star_report.to_excel(writer, sheet_name="星級概況", index=False)
                df_site_report.to_excel(writer, sheet_name="各國籍站點概況", index=False)
                df_ez_report.to_excel(writer, sheet_name="Eztravel專案", index=False)
            
            st.markdown("---")
            st.download_button(
                label="📥 下載 New Central 完整多維度大報表 Excel",
                data=output.getvalue(),
                file_name=f"New_Central_全維度綜合分析_{tw_end.split('-')[1]}{tw_end.split('-')[2]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"運算時發生未知錯誤: {e}")
