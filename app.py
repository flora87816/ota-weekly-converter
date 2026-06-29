import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="New Central 數據校正版", layout="wide")

st.title("📊 New Central 標準數據轉換工具 (Raw Data 精準對齊版)")

# 1. 日期區間
col1, col2 = st.columns(2)
with col1:
    lw_start = st.text_input("上週 開始日期", "2026-06-12")
    lw_end = st.text_input("上週 結束日期", "2026-06-18")
with col2:
    tw_start = st.text_input("當週 開始日期", "2026-06-19")
    tw_end = st.text_input("當週 結束日期", "2026-06-25")

uploaded_file = st.file_uploader("請上傳原始 Excel 檔案", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df['Book Time'] = pd.to_datetime(df['Book Time'])
        df['Book Date'] = df['Book Time'].dt.date
        
        # 篩選資料
        d_lw_start, d_lw_end = pd.to_datetime(lw_start).date(), pd.to_datetime(lw_end).date()
        d_tw_start, d_tw_end = pd.to_datetime(tw_start).date(), pd.to_datetime(tw_end).date()
        
        lw_sub = df[(df['Book Date'] >= d_lw_start) & (df['Book Date'] <= d_lw_end)]
        tw_sub = df[(df['Book Date'] >= d_tw_start) & (df['Book Date'] <= d_tw_end)]
        
        # 核心計算邏輯：直接加總，ADR = GMV / RN
        def get_vals(sub_df):
            rn = sub_df['RN'].sum()
            rev = sub_df['gmv'].sum()
            adr = rev / rn if rn > 0 else 0
            return rn, rev, adr

        # 整理 MM 數據
        standard_mms = ["Abby Cheng （鄭任妤）", "Hino Chi （籍喆）", "Sandy Su （蘇筱鈞）", "Wayne Wang （王俊明）", "Carol Lin （林芝榕）", "Flora Tsao （曹芳綺）", "Justin Lee （李岳勳）"]
        
        def map_mm(raw):
            raw = str(raw)
            for m in standard_mms:
                if any(name in raw for name in [m.split('（')[1].replace('）',''), m.split(' （')[0].split(' ')[0]]):
                    return m
            return "Others"
        
        df['Cleaned_MM'] = df['MM'].apply(map_mm)
        
        # 建立報表結構
        output_rows = []
        for mm in standard_mms + ["Others"]:
            l_rn, l_rev, l_adr = get_vals(lw_sub[lw_sub['Cleaned_MM'] == mm])
            t_rn, t_rev, t_adr = get_vals(tw_sub[tw_sub['Cleaned_MM'] == mm])
            
            # WoW 計算
            wow_rn = (t_rn - l_rn) / l_rn if l_rn != 0 else 0
            wow_rev = (t_rev - l_rev) / l_rev if l_rev != 0 else 0
            wow_adr = (t_adr - l_adr) / l_adr if l_adr != 0 else 0
            
            output_rows.append([mm, l_rn, l_rev, l_adr, t_rn, t_rev, t_adr, 
                               f"{wow_rn:.2%}", f"{wow_rev:.2%}", f"{wow_adr:.2%}"])
            
        final_df = pd.DataFrame(output_rows, columns=["項目", "LW RN", "LW REV", "LW ADR", "TW RN", "TW REV", "TW ADR", "WoW RN%", "WoW REV%", "WoW ADR%"])
        
        st.dataframe(final_df)
        
        output = io.BytesIO()
        final_df.to_excel(output, index=False)
        st.download_button("📥 下載對齊數據", data=output.getvalue(), file_name="Aligned_Data.xlsx")
        
    except Exception as e:
        st.error(f"錯誤: {e}")
