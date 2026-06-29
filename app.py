import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Weekly Performance 數據自動轉換器", layout="wide")

st.title("📊 Weekly Performance 數據轉換工具")
st.write("請上傳每週從系統匯出的原始 `OverseaDataMarket` 檔案，系統將自動為你產出 New Central 的 MM 概況報表。")

# 1. 日期區間設定
st.subheader("🗓️ 步驟一：設定比對區間（請輸入西元年月日，格式如：2026-06-12）")
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
        st.info("正在讀取並解析檔案中，請稍候...")
        
        # 讀取 Excel 檔案（預設讀取第一個分頁，或嘗試尋找 'TW' 分頁）
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = 'TW' if 'TW' in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        
        # 檢查關鍵欄位是否存在
        required_cols = ['Book Time', 'RN', 'ordamount_afterdiscount', 'MM']
        missing_cols = [c for c in required_cols if c not in df.columns]
        
        if missing_cols:
            st.error(f"錯誤：上傳的檔案缺少必要欄位：{missing_cols}，請確認是否為正確的系統導出檔。")
        else:
            # 轉換時間欄位
            df['Book Time'] = pd.to_datetime(df['Book Time'])
            df['Book Date'] = df['Book Time'].dt.date
            
            # 定義日期邊界
            d_lw_start = pd.to_datetime(lw_start).date()
            d_lw_end = pd.to_datetime(lw_end).date()
            d_tw_start = pd.to_datetime(tw_start).date()
            d_tw_end = pd.to_datetime(tw_end).date()
            
            # 篩選上週與當週數據
            lw_df = df[(df['Book Date'] >= d_lw_start) & (df['Book Date'] <= d_lw_end)]
            tw_df = df[(df['Book Date'] >= d_tw_start) & (df['Book Date'] <= d_tw_end)]
            
            # 建立固定的 MM 清單並清理格式
            standard_mms = [
                "Abby Cheng （鄭任妤）", "Hino Chi （籍喆）", "Sandy Su （蘇筱鈞）", 
                "Wayne Wang （王俊明）", "Carol Lin （林芝榕）", "Flora Tsao （曹芳綺）", "Justin Lee （李岳勳）"
            ]
            
            # 定義名稱對應函式
            def find_standard_mm(raw_name):
                if pd.isna(raw_name):
                    return "Others"
                for s_mm in standard_mms:
                    # 提取中文名字進行模糊比對
                    chi_name = s_mm.split('（')[1].replace('）', '').strip()
                    if chi_name in str(raw_name):
                        return s_mm
                return "Others"
            
            df['Cleaned_MM'] = df['MM'].apply(find_standard_mm)
            
            # 重新按清洗後的 MM 篩選
            lw_df = df[(df['Book Date'] >= d_lw_start) & (df['Book Date'] <= d_lw_end)]
            tw_df = df[(df['Book Date'] >= d_tw_start) & (df['Book Date'] <= d_tw_end)]
            
            # 聚合上週數據
            lw_summary = lw_df.groupby('Cleaned_MM').agg(
                lw_RN=('RN', 'sum'),
                lw_REV=('ordamount_afterdiscount', 'sum')
            ).reset_index()
            lw_summary['lw_ADR'] = lw_summary.apply(lambda r: r['lw_REV'] / r['lw_RN'] if r['lw_RN'] > 0 else 0.0, axis=1)
            
            # 聚合當週數據
            tw_summary = tw_df.groupby('Cleaned_MM').agg(
                tw_RN=('RN', 'sum'),
                tw_REV=('ordamount_afterdiscount', 'sum')
            ).reset_index()
            tw_summary['tw_ADR'] = tw_summary.apply(lambda r: r['tw_REV'] / r['tw_RN'] if r['tw_RN'] > 0 else 0.0, axis=1)
            
            # 合併基礎架構
            base_df = pd.DataFrame({'行政區': standard_mms})
            
            # 合併上週
            output_df = pd.merge(base_df, lw_summary, left_on='行政區', right_on='Cleaned_MM', how='left').drop(columns=['Cleaned_MM'])
            # 合併當週
            output_df = pd.merge(output_df, tw_summary, left_on='行政區', right_on='Cleaned_MM', how='left').drop(columns=['Cleaned_MM'])
            
            # 填補空值
            output_df = output_df.fillna(0)
            
            # 計算 WoW 成長率
            output_df['WoW RN 成長率'] = output_df.apply(
                lambda r: f"{((r['tw_RN'] - r['lw_RN']) / r['lw_RN'] * 100):.1f}%" if r['lw_RN'] > 0 else ("100.0%" if r['tw_RN'] > 0 else "0.0%"), axis=1
            )
            output_df['WoW REV 成長率'] = output_df.apply(
                lambda r: f"{((r['tw_REV'] - r['lw_REV']) / r['lw_REV'] * 100):.1f}%" if r['lw_REV'] > 0 else ("100.0%" if r['tw_REV'] > 0 else "0.0%"), axis=1
            )
            
            # 重新命名表頭為 New Central 的精緻格式
            lw_label = f"{lw_start.split('-')[1]}/{lw_start.split('-')[2]}-{lw_end.split('-')[1]}/{lw_end.split('-')[2]}"
            tw_label = f"{tw_start.split('-')[1]}/{tw_start.split('-')[2]}-{tw_end.split('-')[1]}/{tw_end.split('-')[2]}"
            
            output_df = output_df.rename(columns={
                'lw_RN': f'{lw_label} RN',
                'lw_REV': f'{lw_label} REV',
                'lw_ADR': f'{lw_label} ADR',
                'tw_RN': f'{tw_label} RN',
                'tw_REV': f'{tw_label} REV',
                'tw_ADR': f'{tw_label} ADR'
            })
            
            # 欄位格式化呈現（僅影響網頁預覽，匯出時保留數值）
            preview_df = output_df.copy()
            for col in preview_df.columns:
                if 'REV' in col or 'ADR' in col:
                    preview_df[col] = preview_df[col].map(lambda x: f"${x:,.2f}")
                elif 'RN' in col and '成長率' not in col:
                    preview_df[col] = preview_df[col].map(lambda x: f"{int(x):,}")
            
            st.success("🎉 數據自動轉換完成！")
            
            # 網頁預覽
            st.subheader("📋 New Central 格式 - MM營運概況預覽")
            st.dataframe(preview_df, use_container_width=True)
            
            # 匯出成 Excel 檔案
            output = io.BytesIO()
            sheet_title = f"2026 {tw_end.split('-')[1]}{tw_end.split('-')[2]}"
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                output_df.to_excel(writer, sheet_name=sheet_title, index=False)
            
            st.download_button(
                label="📥 下載 New Central 整理分頁 Excel",
                data=output.getvalue(),
                file_name=f"New_Central_分頁_{tw_end.split('-')[1]}{tw_end.split('-')[2]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"處理檔案時發生錯誤，請檢查欄位格式或日期輸入。錯誤訊息: {e}")
