import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="New Central 2026 0629 標準轉換器 (WoW%)", layout="wide")

st.title("📊 New Central 標準數據轉換工具 (WoW % 修正版)")
st.write("此版本已將 WoW 的 RN 與 REV 計算邏輯修正為 **百分比 (%) 成長率**。")

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
        st.info("⚡ 正在嚴格依照 0629 百分比規則重組數據中...")
        
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
        
        # 空的容器，用來裝垂直排版的列數據
        final_rows = []
        
        # --- 通用處理計算 WoW % 的模組 ---
        def calc_wow_metrics(lw_rn, lw_rev, tw_rn, tw_rev):
            lw_adr = lw_rev / lw_rn if lw_rn > 0 else 0
            tw_adr = tw_rev / tw_rn if tw_rn > 0 else 0
            
            # WoW 百分比計算 (如果上週為 0，則呈現 0%)
            wow_rn_pct = f"{(tw_rn - lw_rn) / lw_rn:.2%}" if lw_rn > 0 else "0.00%"
            wow_rev_pct = f"{(tw_rev - lw_rev) / lw_rev:.2%}" if lw_rev > 0 else "0.00%"
            # ADR 的 WoW 通常看增減金額值
            wow_adr_diff = tw_adr - lw_adr
            
            return lw_rn, lw_rev, lw_adr, tw_rn, tw_rev, tw_adr, wow_rn_pct, wow_rev_pct, wow_adr_diff

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
            
            metrics = calc_wow_metrics(lw_m['RN'].sum(), lw_m['ordamount_afterdiscount'].sum(), tw_m['RN'].sum(), tw_m['ordamount_afterdiscount'].sum())
            final_rows.append(["", mm] + list(metrics))
            
        # Others 總計
        lw_oth = lw_sub[lw_sub['Cleaned_MM'] == 'Others']
        tw_oth = tw_sub[tw_sub['Cleaned_MM'] == 'Others']
        oth_metrics = calc_wow_metrics(lw_oth['RN'].sum(), lw_oth['ordamount_afterdiscount'].sum(), tw_oth['RN'].sum(), tw_oth['ordamount_afterdiscount'].sum())
        final_rows.append(["", "others 總計"] + list(oth_metrics))
        
        final_rows.append([""] * 11) # 空行
        
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
                
                cs_metrics = calc_wow_metrics(lw_cs['RN'].sum(), lw_cs['ordamount_afterdiscount'].sum(), tw_cs['RN'].sum(), tw_cs['ordamount_afterdiscount'].sum())
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
                    
                s_metrics = calc_wow_metrics(lw_s['RN'].sum(), lw_s['ordamount_afterdiscount'].sum(), tw_s['RN'].sum(), tw_s['ordamount_afterdiscount'].sum())
                final_rows.append(["", site] + list(s_metrics))
                
            # Total
            tot_metrics = calc_wow_metrics(lw_c_df['RN'].sum(), lw_c_df['ordamount_afterdiscount'].sum(), tw_c_df['RN'].sum(), tw_c_df['ordamount_afterdiscount'].sum())
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
                
                # EZ Share WoW 規則：量差用絕對值，佔比差用百分比點數(相減)
                final_rows.append([
                    "", m_name, d, lw_md, tw_md, tw_md - lw_md, 
                    f"{l_ratio:.2%}", f"{t_ratio:.2%}", f"{(t_ratio - l_ratio):.2%}", "", ""
                ])
                first = False
                
            final_rows.append(["", "", "Total", lw_m_tot, tw_m_tot, tw_m_tot - lw_m_tot, "100.00%", "100.00%", "0.00%", "", ""])

        # 輸出處理
        out_df = pd.DataFrame(final_rows)
        st.success("🎉 百分比格式報表轉換完成！")
        st.dataframe(out_df.fillna(""), use_container_width=True)
        
        output = io.BytesIO()
        sheet_title = f"2026 {tw_end.split('-')[1]}{tw_end.split('-')[2]}"
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            out_df.to_excel(writer, sheet_name=sheet_title, index=False, header=False)
            
        st.download_button(
            label="📥 下載修正版 Excel 報表",
            data=output.getvalue(),
            file_name=f"New_Central_周報_Pct_{tw_end.split('-')[1]}{tw_end.split('-')[2]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"轉換發生錯誤: {e}")
