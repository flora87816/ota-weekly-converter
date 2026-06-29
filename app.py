import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="New Central 2026 週報標準轉換器 (百分比強制修正版)", layout="wide")

st.title("📊 New Central 標準數據轉換工具 (WoW百分比字串強制修正版)")
st.write("此版本已將所有 WoW 欄位在 DataFrame 階段就直接格式化為 `XX.XX%` 字串，確保網頁預覽與 Excel 匯出絕對是百分比！")

# 1. 日期與名稱設定
col1, col2 = st.columns(2)
with col1:
    lw_start = st.text_input("上週 開始日期", "2026-06-12")
    lw_end = st.text_input("上週 結束日期", "2026-06-18")
with col2:
    tw_start = st.text_input("當週 開始日期", "2026-06-19")
    tw_end = st.text_input("當週 結束日期", "2026-06-25")

# 2. 檔案上傳
uploaded_file = st.file_uploader("請上傳原始 Excel 檔案", type=["xlsx"])

if uploaded_file is not None:
    try:
        st.info("⚡ 正在轉換並強制格式化 WoW 百分比欄位...")
        
        # 讀取資料
        xls = pd.ExcelFile(uploaded_file)
        df = pd.read_excel(uploaded_file, sheet_name=xls.sheet_names[0])
        
        # 轉換時間
        df['Book Time'] = pd.to_datetime(df['Book Time'])
        df['Book Date'] = df['Book Time'].dt.date
        
        d_lw_start = pd.to_datetime(lw_start).date()
        d_lw_end = pd.to_datetime(lw_end).date()
        d_tw_start = pd.to_datetime(tw_start).date()
        d_tw_end = pd.to_datetime(tw_end).date()
        
        lw_label = f"{lw_start.split('-')[1]}/{lw_start.split('-')[2]}-{lw_end.split('-')[1]}/{lw_end.split('-')[2]}"
        tw_label = f"{tw_start.split('-')[1]}/{tw_start.split('-')[2]}-{tw_end.split('-')[1]}/{tw_end.split('-')[2]}"
        
        # --- 核心計算模組（回傳格式化好的字串或標記） ---
        def calc_wow_metrics_formatted(lw_df, tw_df):
            lw_rn = round(lw_df['RN'].sum(), 2)
            tw_rn = round(tw_df['RN'].sum(), 2)
            
            lw_rev = round(lw_df['gmv'].sum(), 2)
            tw_rev = round(tw_df['gmv'].sum(), 2)
            
            lw_adr = round(lw_df['ordamount_afterdiscount'].mean(), 2) if not lw_df.empty and lw_df['ordamount_afterdiscount'].notna().any() else 0
            tw_adr = round(tw_df['ordamount_afterdiscount'].mean(), 2) if not tw_df.empty and tw_df['ordamount_afterdiscount'].notna().any() else 0
            
            # 計算原始浮點數
            wow_rn_pct = ((tw_rn - lw_rn) / lw_rn) if lw_rn > 0 else (0.0 if tw_rn == 0 else 1.0)
            wow_rev_pct = ((tw_rev - lw_rev) / lw_rev) if lw_rev > 0 else (0.0 if tw_rev == 0 else 1.0)
            wow_adr_pct = ((tw_adr - lw_adr) / lw_adr) if lw_adr > 0 else (0.0 if tw_adr == 0 else 1.0)
            
            # 直接強制格式化為字串，確保顯示
            return [
                lw_rn, lw_rev, lw_adr, 
                tw_rn, tw_rev, tw_adr, 
                f"{wow_rn_pct * 100:.2f}%", 
                f"{wow_rev_pct * 100:.2f}%", 
                f"{wow_adr_pct * 100:.2f}%"
            ]

        # 名字配對邏輯
        standard_mms = [
            "Abby Cheng （鄭任妤）", 
            "Hino Chi （籍喆）", 
            "Sandy Su （蘇筱鈞）", 
            "Wayne Wang （王俊明）", 
            "Carol Lin （林芝榕）", 
            "Flora Tsao （曹芳綺）", 
            "Justin Lee （李岳勳）"
        ]
        
        def clean_mm_name(raw):
            if pd.isna(raw): return "Others"
            raw_str = str(raw).strip()
            if "鄭任妤" in raw_str or "Abby" in raw_str: return "Abby Cheng （鄭任妤）"
            if "籍喆" in raw_str or "Hino" in raw_str: return "Hino Chi （籍喆）"
            if "蘇筱鈞" in raw_str or "Sandy" in raw_str or "TR009750" in raw_str: return "Sandy Su （蘇筱鈞）"
            if "王俊明" in raw_str or "Wayne" in raw_str: return "Wayne Wang （王俊明）"
            if "林芝榕" in raw_str or "Carol" in raw_str: return "Carol Lin （林芝榕）"
            if "曹芳綺" in raw_str or "Flora" in raw_str: return "Flora Tsao （曹芳綺）"
            if "李岳勳" in raw_str or "Justin" in raw_str: return "Justin Lee （李岳勳）"
            return "Others"
        
        df['Cleaned_MM'] = df['MM'].apply(clean_mm_name)
        
        # 篩選兩週子集
        lw_sub = df[(df['Book Date'] >= d_lw_start) & (df['Book Date'] <= d_lw_end)]
        tw_sub = df[(df['Book Date'] >= d_tw_start) & (df['Book Date'] <= d_tw_end)]
        
        # =========================================================================
        # 分頁一：MM概況
        # =========================================================================
        mm_rows = [
            ["行政區", lw_label, "", "", tw_label, "", "", "WoW", "", ""],
            ["", "RN", "REV", "ADR", "RN", "REV", "ADR", "RN", "REV", "ADR"]
        ]
        for mm in standard_mms:
            metrics = calc_wow_metrics_formatted(lw_sub[lw_sub['Cleaned_MM'] == mm], tw_sub[tw_sub['Cleaned_MM'] == mm])
            mm_rows.append([mm] + metrics)
        oth_metrics = calc_wow_metrics_formatted(lw_sub[lw_sub['Cleaned_MM'] == 'Others'], tw_sub[tw_sub['Cleaned_MM'] == 'Others'])
        mm_rows.append(["others 總計"] + oth_metrics)
        df_mm_sheet = pd.DataFrame(mm_rows)

        # =========================================================================
        # 分頁二：城市&星級概況
        # =========================================================================
        city_rows = [
            ["行政區 / 星級", lw_label, "", "", tw_label, "", "", "WoW", "", ""],
            ["", "RN", "REV", "ADR", "RN", "REV", "ADR", "RN", "REV", "ADR"]
        ]
        cities = ["台中", "台東", "台南", "花蓮", "金門", "南投", "屏東", "高雄", "嘉義"]
        stars = [3, 4, 5]
        
        for c in cities:
            city_rows.append([f"--- {c} ---", "", "", "", "", "", "", "", "", ""])
            for s in stars:
                cs_metrics = calc_wow_metrics_formatted(lw_sub[(lw_sub['City'] == c) & (lw_sub['Star'] == s)], tw_sub[(tw_sub['City'] == c) & (tw_sub['Star'] == s)])
                city_rows.append([f"{s} 星"] + cs_metrics)
            
            city_tot_metrics = calc_wow_metrics_formatted(lw_sub[lw_sub['City'] == c], tw_sub[tw_sub['City'] == c])
            city_rows.append(["Total"] + city_tot_metrics)
            
        df_city_sheet = pd.DataFrame(city_rows)

        # =========================================================================
        # 分頁三：各國籍概況
        # =========================================================================
        nat_rows = []
        target_nationality_cities = ["台中", "南投"]
        sites = ["others", "Trip(CN1)", "Trip(HK)", "Trip(ID)", "Trip(JP)", "Trip(KR)", "Trip(MY)", "Trip(PH)", "Trip(SG)", "Trip(TH)", "Trip(TW)", "Trip(US)", "Trip(XX)"]
        
        for t_c in target_nationality_cities:
            nat_rows.append([f"各國籍概況 ({t_c})", "", "", "", "", "", "", "", "", ""])
            nat_rows.append(["Site", lw_label, "", "", tw_label, "", "", "WoW", "", ""])
            nat_rows.append(["", "RN", "REV", "ADR", "RN", "REV", "ADR", "RN", "REV", "ADR"])
            
            lw_c_df = lw_sub[lw_sub['City'] == t_c]
            tw_c_df = tw_sub[tw_sub['City'] == t_c]
            
            for site in sites:
                if site == "others":
                    lw_s = lw_c_df[~lw_c_df['Ctrip/Trip site'].isin(sites)]
                    tw_s = tw_c_df[~tw_c_df['Ctrip/Trip site'].isin(sites)]
                else:
                    lw_s = lw_c_df[lw_c_df['Ctrip/Trip site'] == site]
                    tw_s = tw_c_df[tw_c_df['Ctrip/Trip site'] == site]
                s_metrics = calc_wow_metrics_formatted(lw_s, tw_s)
                nat_rows.append([site] + s_metrics)
                
            tot_metrics = calc_wow_metrics_formatted(lw_c_df, tw_c_df)
            nat_rows.append(["Total"] + tot_metrics)
            nat_rows.append(["", "", "", "", "", "", "", "", "", ""])
        df_nat_sheet = pd.DataFrame(nat_rows)

        # =========================================================================
        # 分頁四：EZ Share (佔比與變動也強制轉字串百分比)
        # =========================================================================
        ez_rows = [
            ["MM", "Maintenance", f"{lw_label} (RN)", f"{tw_label} (RN)", "WoW 增減", f"{lw_label} (佔比)", f"{tw_label} (佔比)", "佔比 WoW 變動"]
        ]
        df_ez = df[df['Secondary Booking Channel'] == 'Eztravel']
        lw_ez = df_ez[(df_ez['Book Date'] >= d_lw_start) & (df_ez['Book Date'] <= d_lw_end)]
        tw_ez = df_ez[(df_ez['Book Date'] >= d_tw_start) & (df_ez['Book Date'] <= d_tw_end)]
        
        ez_target_mms = [
            ("Abby", "Abby Cheng （鄭任妤）"),
            ("Hino", "Hino Chi （籍喆）"),
            ("Sandy", "Sandy Su （蘇筱鈞）"),
            ("Carol", "Carol Lin （林芝榕）"),
            ("Flora", "Flora Tsao （曹芳綺）"),
            ("Justin", "Justin Lee （李岳勳）")
        ]
        deps = ["HPP", "HTL", "SHT"]
        
        for short_name, full_cleaned_name in ez_target_mms:
            lw_m_tot = lw_ez[lw_ez['Cleaned_MM'] == full_cleaned_name]['RN'].sum()
            tw_m_tot = tw_ez[tw_ez['Cleaned_MM'] == full_cleaned_name]['RN'].sum()
            
            first = True
            for d in deps:
                lw_md = lw_ez[(lw_ez['Cleaned_MM'] == full_cleaned_name) & (lw_ez['Maintenance Department'] == d)]['RN'].sum()
                tw_md = tw_ez[(tw_ez['Cleaned_MM'] == full_cleaned_name) & (tw_ez['Maintenance Department'] == d)]['RN'].sum()
                
                l_ratio = lw_md / lw_m_tot if lw_m_tot > 0 else 0
                t_ratio = tw_md / tw_m_tot if tw_m_tot > 0 else 0
                m_name = short_name if first else ""
                
                ez_rows.append([
                    m_name, d, round(lw_md, 2), round(tw_md, 2), round(tw_md - lw_md, 2), 
                    f"{l_ratio * 100:.2f}%", f"{t_ratio * 100:.2f}%", f"{(t_ratio - l_ratio) * 100:.2f}%"
                ])
                first = False
            ez_rows.append(["", "Total", round(lw_m_tot, 2), round(tw_m_tot, 2), round(tw_m_tot - lw_m_tot, 2), "100.00%", "100.00%", "0.00%"])
        df_ez_sheet = pd.DataFrame(ez_rows)

        # =========================================================================
        # 預覽與下載
        # =========================================================================
        st.success("🎉 數據已全面在程式底層轉為百分比字串，匯出與預覽保證正確！")
        tab1, tab2, tab3, tab4 = st.tabs(["👤 MM概況", "🏨 城市&星級概況", "✈️ 各國籍概況", "📊 EZ Share"])
        with tab1: st.dataframe(df_mm_sheet, use_container_width=True)
        with tab2: st.dataframe(df_city_sheet, use_container_width=True)
        with tab3: st.dataframe(df_nat_sheet, use_container_width=True)
        with tab4: st.dataframe(df_ez_sheet, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_mm_sheet.to_excel(writer, sheet_name="MM概況", index=False, header=False)
            df_city_sheet.to_excel(writer, sheet_name="城市星級概況", index=False, header=False)
            df_nat_sheet.to_excel(writer, sheet_name="各國籍概況", index=False, header=False)
            df_ez_sheet.to_excel(writer, sheet_name="EZ Share", index=False, header=False)
            
            workbook = writer.book
            num_format = workbook.add_format({'num_format': '#,##0.00'})
            
            # 因為現在變成了字串，條件格式化改用「包含負號 -」來識別下滑標紅
            red_light_format = workbook.add_format({
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006'
            })
            red_num_format = workbook.add_format({
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006',
                'num_format': '#,##0.00'
            })

            # 前三頁欄位寬度與標紅設定
            for sheet_name in ["MM概況", "城市星級概況", "各國籍概況"]:
                ws = writer.sheets[sheet_name]
                ws.set_column('A:A', 22)
                ws.set_column('B:G', 15, num_format)
                ws.set_column('H:J', 15) # 百分比字串欄位
                
                # 使用 text 定位，只要字串裡包含「-」代表負成長，就自動標紅！
                ws.conditional_format('H3:J200', {
                    'type': 'text',
                    'criteria': 'containing',
                    'value': '-',
                    'format': red_light_format
                })

            # 第四頁 EZ Share 欄位與標紅設定
            ws_ez = writer.sheets["EZ Share"]
            ws_ez.set_column('A:B', 15)
            ws_ez.set_column('C:E', 15, num_format)
            ws_ez.set_column('F:H', 15)
            
            # E欄是純數字增減，維持原本的小於0標紅
            ws_ez.conditional_format('E2:E200', {
                'type': 'cell',
                'criteria': '<',
                'value': 0,
                'format': red_num_format
            })
            # H欄是百分比字串變動，用包含負號標紅
            ws_ez.conditional_format('H2:H200', {
                'type': 'text',
                'criteria': 'containing',
                'value': '-',
                'format': red_light_format
            })
            
        st.download_button(
            label="📥 下載「強制轉字串百分比版」Excel",
            data=output.getvalue(),
            file_name=f"New_Central_周報_強制百分比版_{tw_end.split('-')[1]}{tw_end.split('-')[2]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"轉換發生錯誤: {e}")
