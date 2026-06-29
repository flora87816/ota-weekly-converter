import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="New Central 2026 0629 標準轉換器 (欄位校正版)", layout="wide")

st.title("📊 New Central 標準數據轉換工具 (REV: gmv / ADR: after_discount 平均)")
st.write("此版本已將 **REV 修正為 gmv 總和**，**ADR 修正為 ordamount_afterdiscount 平均值**，且所有 WoW 欄位皆以**百分比 (%) 四捨五入**呈現。")

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
        st.info("⚡ 正在嚴格依據全新欄位定義 (REV=gmv, ADR=after_discount mean) 計算中...")
        
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
        
        final_rows = []
        
        # --- 核心計算模組：處理全新的定義與百分比四捨五入 ---
        def calc_wow_metrics(lw_df, tw_df):
            # 1. RN 總和
            lw_rn = round(lw_df['RN'].sum(), 2)
            tw_rn = round(tw_df['RN'].sum(), 2)
            
            # 2. REV = gmv 欄位的總和
            lw_rev = round(lw_df['gmv'].sum(), 2)
            tw_rev = round(tw_df['gmv'].sum(), 2)
            
            # 3. ADR = ordamount_afterdiscount 欄位的平均值 (Mean)
            # 若無數據則給 0
            lw_adr = round(lw_df['ordamount_afterdiscount'].mean(), 2) if not lw_df.empty and lw_df['ordamount_afterdiscount'].notna().any() else 0
            tw_adr = round(tw_df['ordamount_afterdiscount'].mean(), 2) if not tw_df.empty and tw_df['ordamount_afterdiscount'].notna().any() else 0
            
            # 4. WoW 百分比變動計算 (嚴格四捨五入至小數點後兩位)
            wow_rn_pct = f"{round(((tw_rn - lw_rn) / lw_rn) * 100, 2):.2f}%" if lw_rn > 0 else "0.00%"
            wow_rev_pct = f"{round(((tw_rev - lw_rev) / lw_rev) * 100, 2):.2f}%" if lw_rev > 0 else "0.00%"
            wow_adr_pct = f"{round(((tw_adr - lw_adr) / lw_adr) * 100, 2):.2f}%" if lw_adr > 0 else "0.00%"
            
            return lw_rn, lw_rev, lw_adr, tw_rn, tw_rev, tw_adr, wow_rn_pct, wow_rev_pct, wow_adr_pct

        # =========================================================================
        # 區塊一：MM概況
        # =========================================================================
        final_rows.append(["", "MM概況", "", "", "", "", "", "", "", "", ""])
        final_rows.append(["", "行政區", lw_label, "", "", tw_label, "", "", "WoW", "", ""])
        final_rows.append(["", "", "RN", "REV", "ADR", "RN", "REV", "ADR", "RN", "REV", "ADR"])
        
        standard_mms = ["Abby Cheng （鄭任妤）", "Hino Chi （籍喆）", "Sandy Su （蘇筱鈞）", "Wayne Wang （王俊明）", "Carol Lin （林芝榕）", "Flora Tsao （曹芳綺）", "Justin Lee （李岳勳）"]
        
        def clean_mm_name(raw):
            if pd.isna(raw): return "Others"
            for m in standard_mms:
                chi = m.split('（')[1].replace('）','').strip()
                if chi in str(raw): return m
            return "Others"
        
        df['Cleaned_MM'] = df['MM'].apply(clean_mm_name)
        
        lw_sub = df[(df['Book Date'] >= d_lw_start) & (df['Book Date'] <= d_lw_end)]
        tw_sub = df[(df['Book Date'] >= d_tw_start) & (df['Book Date'] <= d_tw_end)]
        
        for mm in standard_mms:
            lw_m = lw_sub[lw_sub['Cleaned_MM'] == mm]
            tw_m = tw_sub[tw_sub['Cleaned_MM'] == mm]
            
            metrics = calc_wow_metrics(lw_m, tw_m)
            final_rows.append(["", mm] + list(metrics))
            
        # Others 總計
        lw_oth = lw_sub[lw_sub['Cleaned_MM'] == 'Others']
        tw_oth = tw_sub[tw_sub['Cleaned_MM'] == 'Others']
        oth_metrics = calc_wow_metrics(lw_oth, tw_oth)
        final_rows.append(["", "others 總計"] + list(oth_metrics))
        
        final_rows.append([""] * 11)
        
        # =========================================================================
        # 區塊二：城市&星級概況
        # =========================================================================
        final_rows.append(["", "城市&星級概況", "", "", "", "", "", "", "", "", ""])
        final_rows.append(["", "行政區", lw_label, "", "", tw_label, "", "", "WoW", "", ""])
        final_rows.append(["", "", "RN", "REV", "ADR", "RN", "REV", "ADR", "RN", "REV", "ADR"])
        
        cities = ["台中", "台東", "台南", "花蓮", "金門", "南投", "屏東", "高雄", "嘉義"]
        stars = [3, 4, 5]
        
        for c in cities:
            final_rows.append(["", c, "", "", "", "", "", "", "", "", ""])
            for s in stars:
                lw_cs = lw_sub[(lw_sub['City'] == c) & (lw_sub['Star'] == s)]
                tw_cs = tw_sub[(tw_sub['City'] == c) & (tw_sub['Star'] == s)]
                
                cs_metrics = calc_wow_metrics(lw_cs, tw_cs)
                final_rows.append(["", s] + list(cs_metrics))
                
        final_rows.append([""] * 11)
        
        # =========================================================================
        # 區塊三：各國籍概況 (台中 / 南投)
        # =========================================================================
        target_nationality_cities = ["台中", "南投"]
        sites = ["others", "Trip(CN1)", "Trip(HK)", "Trip(ID)", "Trip(JP)", "Trip(KR)", "Trip(MY)", "Trip(PH)", "Trip(SG)", "Trip(TH)", "Trip(TW)", "Trip(US)", "Trip(XX)"]
        
        for t_c in target_nationality_cities:
            final_rows.append(["", f"各國籍概況\n({t_c})", "", "", "", "", "", "", "", "", ""])
            final_rows.append(["", "Site", lw_label, "", "", tw_label, "", "", "WoW", "", ""])
            final_rows.append(["", "", "RN", "REV", "ADR", "RN", "REV", "ADR", "RN", "REV", "ADR"])
            
            lw_c_df = lw_sub[lw_sub['City'] == t_c]
            tw_c_df = tw_sub[tw_sub['City'] == t_c]
            
            for site in sites:
                if site == "others":
                    lw_s = lw_c_df[~lw_c_df['Ctrip/Trip site'].isin(sites)]
                    tw_s = tw_c_df[~tw_c_df['Ctrip/Trip site'].isin(sites)]
                else:
                    lw_s = lw_c_df[lw_c_df['Ctrip/Trip site'] == site]
                    tw_s = tw_c_df[tw_c_df['Ctrip/Trip site'] == site]
                    
                s_metrics = calc_wow_metrics(lw_s, tw_s)
                final_rows.append(["", site] + list(s_metrics))
                
            tot_metrics = calc_wow_metrics(lw_c_df, tw_c_df)
            final_rows.append(["", "Total"] + list(tot_metrics))
            final_rows.append([""] * 11)

        # =========================================================================
        # 區塊四：EZ Share
        # =========================================================================
        final_rows.append(["", "EZ Share", "", "", "", "", "", "", "", "", ""])
        final_rows.append(["", "MM", "Maintenance", f"{lw_label}\n(RN)", f"{tw_label}\n(RN)", "WoW", f"{lw_label}\n(佔比)", f"{tw_label}\n(佔比)", "WoW", "", ""])
        
        df_ez = df[df['Secondary Booking Channel'] == 'Eztravel']
        lw_ez = df_ez[(df_ez['Book Date'] >= d_lw_start) & (df_ez['Book Date'] <= d_lw_end)]
        tw_ez = df_ez[(df_ez['Book Date'] >= d_tw_start) & (df_ez['Book Date'] <= d_tw_end)]
        
        ez_mms = ["Abby", "Hino", "Sandy", "Carol", "Flora", "Justin"]
        deps = ["HPP", "HTL", "SHT"]
        
        for emm in ez_mms:
            lw_m_tot = lw_ez[lw_ez['Cleaned_MM'].str.contains(emm, na=False)]['RN'].sum()
            tw_m_tot = tw_ez[tw_ez['Cleaned_MM'].str.contains(emm, na=False)]['RN'].sum()
            
            first = True
            for d in deps:
                lw_md = lw_ez[(lw_ez['Cleaned_MM'].str.contains(emm, na=False)) & (lw_ez['Maintenance Department'] == d)]['RN'].sum()
                tw_md = tw_ez[(tw_ez['Cleaned_MM'].str.contains(emm, na=False)) & (tw_ez['Maintenance Department'] == d)]['RN'].sum()
                
                l_ratio = lw_md / lw_m_tot if lw_m_tot > 0 else 0
                t_ratio = tw_md / tw_m_tot if tw_m_tot > 0 else 0
                
                m_name = emm if first else ""
                
                # EZ Share 數量 WoW 維持量差，佔比變動採用百分比點數差額 (四捨五入)
                final_rows.append([
                    "", m_name, d, round(lw_md, 2), round(tw_md, 2), round(tw_md - lw_md, 2), 
                    f"{round(l_ratio * 100, 2):.2f}%", f"{round(t_ratio * 100, 2):.2f}%", f"{round((t_ratio - l_ratio) * 100, 2):.2f}%", "", ""
                ])
                first = False
                
            final_rows.append(["", "", "Total", round(lw_m_tot, 2), round(tw_m_tot, 2), round(tw_m_tot - lw_m_tot, 2), "100.00%", "100.00%", "0.00%", "", ""])

        # 輸出處理
        out_df = pd.DataFrame(final_rows)
        st.success("🎉 欄位校正與全四捨五入百分比報表轉換完成！")
        st.dataframe(out_df.fillna(""), use_container_width=True)
        
        output = io.BytesIO()
        sheet_title = f"2026 {tw_end.split('-')[1]}{tw_end.split('-')[2]}"
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            out_df.to_excel(writer, sheet_name=sheet_title, index=False, header=False)
            
        st.download_button(
            label="📥 下載最終校正版 Excel 報表",
            data=output.getvalue(),
            file_name=f"New_Central_周報_Final_{tw_end.split('-')[1]}{tw_end.split('-')[2]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"轉換發生錯誤: {e}")
