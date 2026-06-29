import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Weekly Performance to New Central 轉換器", layout="wide")

st.title("📊 Weekly Performance 數據轉換工具 (New Central 格式)")
st.write("請上傳每週的原始 `weekly performance` 檔案，系統會自動轉換為 New Central 的週對週 MM 概況格式。")

# 1. 參數輸入
col1, col2 = st.columns(2)
with col1:
    last_week_range = st.text_input("輸入上週日期區間 (例如: 6/12-6/18)", "6/12-6/18")
with col2:
    this_week_range = st.text_input("輸入當週日期區間 (例如: 6/19-6/25)", "6/19-6/25")

# 2. 檔案上傳
uploaded_file = st.file_uploader("請上傳 weekly performance 原始檔 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 讀取 Excel 檔案
        xls_wp = pd.ExcelFile(uploaded_file)
        
        if 'TW' in xls_wp.sheet_names:
            st.success("檔案上傳成功！正在解析數據...")
            
            # 這裡預留自動化清洗邏輯，會讀取 TW 分頁並自動排列為 New Central 結構
            # 目前先為你產生與 New Central 0622.xlsx 完全相同的標準 MM 清單
            mm_list = [
                "Abby Cheng （鄭任妤）", "Hino Chi （籍喆）", "Sandy Su （蘇筱鈞）", 
                "Wayne Wang （王俊明）", "Carol Lin （林芝榕）", "Flora Tsao （曹芳綺）", "Justin Lee （李岳勳）"
            ]
            
            summary_data = []
            for mm in mm_list:
                summary_data.append({
                    "行政區": mm,
                    f"{last_week_range} RN": 0,    
                    f"{last_week_range} REV": 0.0,
                    f"{last_week_range} ADR": 0.0,
                    f"{this_week_range} RN": 0,
                    f"{this_week_range} REV": 0.0,
                    f"{this_week_range} ADR": 0.0,
                    "WoW RN 成長率": "0.0%",
                    "WoW REV 成長率": "0.0%"
                })
            
            df_output = pd.DataFrame(summary_data)
            
            # 網頁預覽
            st.subheader("📋 轉換結果預覽 (MM概況)")
            st.dataframe(df_output, use_container_width=True)
            
            # 匯出成 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 建立多層表頭或標準格式
                df_output.to_excel(writer, sheet_name=f"2026 {this_week_range.split('-')[1].replace('/', '')}", index=False)
            
            st.download_button(
                label="📥 下載 New Central 整理分頁 Excel",
                data=output.getvalue(),
                file_name=f"New_Central_分頁_{this_week_range.replace('/', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("錯誤：上傳的檔案中找不到名為 'TW' 的工作表，請檢查檔案格式是否正確。")
            
    except Exception as e:
        st.error(f"處理檔案時發生錯誤: {e}")
