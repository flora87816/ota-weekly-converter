import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="New Central 2026 0629 標準轉換器", layout="wide")

st.title("📊 New Central 標準數據轉換工具 (2026 0629 規則)")
st.write("請上傳原始 `OverseaDataMarket` 檔案。系統將嚴格遵循範例中的垂直排版架構，為您產出單一分頁的完整周報。")

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
        st.info("⚡ 正在嚴格依照 0629 規則重組數據中...")
        
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
        
        # 切分核心週數據
        lw_df = df[(df['Book Date'] >= d_lw_start) & (df['Book Date'] <= d_lw_end)]
        tw_df = df[(df['Book Date'] >= d_tw_start) & (df['Book Date'] <= d_tw_end)]
        
        # 預備一個空的 DataFrame 用來做垂直對齊
        final_rows = []
        
        # --- 通用處理計算 WoW 價差與量差的模組 ---
        def calc_wow_diff(lw_g, tw_g, base_df, join_on):
            m = pd.merge(base_df, lw_g, on=join_on, how='left')
            m = pd.merge(m, tw_g, on=join_on, how='left').fillna(0)
            
            # 計算 ADR
            m['lw_ADR'] = m.apply(lambda r: r['lw_REV'] / r['lw_RN'] if r['lw_RN'] > 0 else 0.0, axis=1)
            m['tw_ADR'] = m.apply(lambda r: r['tw_REV'] / r['tw_RN'] if r['tw_RN'] > 0 else 0.0, axis=1)
            
            # 0629 的 WoW 規則是相減 (當週 - 上週)
            m['wow_RN'] = m['tw_RN'] - m['lw_RN']
            m['wow_REV'] = m['tw_REV'] - m['lw_REV']
            m['wow_ADR'] = m['tw_ADR'] - m['lw_ADR']
            return m

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
        
        lw_mm = df[(df['Book Date'] >= d_lw_start) & (df['Book Date'] <= d_lw_end)].groupby('Cleaned_MM').agg(lw_RN=('RN', 'sum'), lw_REV=('ordamount_afterdiscount', 'sum')).reset_index()
        tw_mm = df[(df['Book Date'] >= d_tw_start) & (df['Book Date'] <= d_tw_end)].groupby('Cleaned_MM').agg(tw_RN=('RN', 'sum'), tw_REV=('ordamount_afterdiscount', 'sum')).reset_index()
        
        base_mm = pd.DataFrame({'Cleaned_MM': standard_mms})
        mm_res = calc_wow_diff(lw_mm, tw_mm, base_mm, 'Cleaned_MM')
        
        for _, r in mm_res.iterrows():
            final_rows.append(["", r['Cleaned_MM'], r['lw_RN'], r['lw_REV'], r['lw_ADR'], r['tw_RN'], r['tw_REV'], r['tw_ADR'], r['wow_RN'], r['wow_REV'], r['wow_ADR']])
            
        # Others 總計
        lw_oth_df = df[(df['Book Date'] >= d_lw_start) & (df['Book Date'] <= d_lw_end) & (df['Cleaned_MM'] == 'Others')]
        tw_oth_df = df[(df['Book Date'] >= d_tw_start) & (df['Book Date'] <= d_tw_end) & (df['Cleaned_MM'] == 'Others')]
        l_o_rn, l_o_rev = lw_oth_df['RN'].sum(), lw_oth_df['ordamount_afterdiscount'].sum()
        t_o_rn, t_o_rev = tw_oth_df['RN'].sum(), tw_oth_df['ordamount_afterdiscount'].sum()
        l_o_adr = l_o_rev / l_o_rn if l_o_rn > 0 else 0
        t_o_adr = t_o_rev / t_o_rn if t_o_rn > 0 else 0
        final_rows.append(["", "others 總計", l_o_rn, l_o_rev, l_o_adr, t_o_rn, t_o_rev, t_o_adr, t_o_rn - l_o_rn, t_o_rev - l_o_rev, t_o_adr - l_o_adr])
        
        final_rows.append([""] * 11) # 空行
        
        # =========================================================================
        # 區塊二：城市&星級概況
        # =========================================================================
        final_rows.append(["", "城市&星級概況", "", "", "", "", "", "", "", "", ""])
        final_rows.append(["", "行政區", lw_label, "", "", tw_label, "", "", "WoW", "", ""])
        final_rows.append(["", "", "RN", "REV", "ADR", "RN", "REV", "ADR", "RN", "REV", "ADR"])
        
        cities = ["台中", "台東", "台南", "花蓮", "金門", "南投", "屏東", "高雄", "嘉義"]
        stars = [3, 4, 5]
        
        lw_sub = df[(df['Book Date'] >= d_lw_start) & (df['Book Date'] <= d_lw_end)]
        tw_sub = df[(df['Book Date'] >= d_tw_start) & (df['Book Date'] <= d_tw_end)]
        
        for c in cities:
            final_rows.append(["", c, "", "", "", "", "", "", "", "", ""])
            for s in stars:
                lw_cs = lw_sub[(lw_sub['City'] == c) & (lw_sub['Star'] == s)]
                tw_cs = tw_sub[(tw_sub['City'] == c) & (tw_sub['Star'] == s)]
                
                l_rn, l_rev = lw_cs['RN'].sum(), lw_cs['ordamount_afterdiscount'].sum()
                t_rn, t_rev = tw_cs['RN'].sum(), tw_cs['ordamount_afterdiscount'].sum()
                l_adr = l_rev / l_rn if l_rn > 0 else 0
                t_adr = t_rev / t_rn if t_rn > 0 else 0
                
                final_rows.append(["", s, l_rn, l_rev, l_adr, t_rn, t_rev, t_adr, t_rn - l_rn, t_rev - l_rev, t_adr - l_adr])
                
        final_rows.append([""] * 11)
        
        # =========================================================================
        # 區塊三：各國籍概況 (台中 / 南投)
        # =========================================================================
        target_nationality_cities = ["台中", "南投"]
        sites = ["others", "Trip(CN1)", "Trip(HK)", "Trip(ID)", "Trip(JP)", "Trip(KR)", "Trip(MY)", "Trip(PH)", "Trip(SG)", "Trip(TH)", "Trip(TW)", "Trip(US)", "Trip(XX)"]
        
        for t_c in target_nationality_cities:
            final_rows.append(["", f"各國籍概況\n({t_c})", "", "", "", "", "", "", "", "", ""])
            final_rows.append(["", "Site", lw_label, "", "", 6/19-6/25, "", "", "WoW", "", ""])
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
                    
                l_rn, l_rev = lw_s['RN'].sum(), lw_s['ordamount_afterdiscount'].sum()
                t_rn, t_rev = tw_s['RN'].sum(), tw_s['ordamount_afterdiscount'].sum()
                l_adr = l_rev / l_rn if l_rn > 0 else 0
                t_adr = t_rev / t_rn if t_rn > 0 else 0
                
                final_rows.append(["", site, l_rn, l_rev, l_adr, t_rn, t_rev, t_adr, t_rn - l_rn, t_rev - l_rev, t_adr - l_adr])
                
            # Total
            l_t_rn, l_t_rev = lw_c_df['RN'].sum(), lw_c_df['ordamount_afterdiscount'].sum()
            t_t_rn, t_t_rev = tw_c_df['RN'].sum(), tw_c_df['ordamount_afterdiscount'].sum()
            final_rows.append(["", "Total", l_t_rn, l_t_rev, l_t_rev/l_t_rn if l_t_rn > 0 else 0, t_t_rn, t_t_rev, t_t_rev/t_t_rn if t_t_rn > 0 else 0, t_t_rn - l_t_rn, t_t_rev - l_t_rev, (t_t_rev/t_t_rn if t_t_rn > 0 else 0) - (l_t_rev/l_t_rn if l_t_rn > 0 else 0)])
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
            # 建立該 MM 總計基礎
            lw_m_tot = lw_ez[lw_ez['Cleaned_MM'].str.contains(emm, na=False)]['RN'].sum()
            tw_m_tot = tw_ez[tw_ez['Cleaned_MM'].str.contains(emm, na=False)]['RN'].sum()
            
            first = True
            for d in deps:
                lw_md = lw_ez[(lw_ez['Cleaned_MM'].str.contains(emm, na=False)) & (lw_ez['Maintenance Department'] == d)]['RN'].sum()
                tw_md = tw_ez[(tw_ez['Cleaned_MM'].str.contains(emm, na=False)) & (tw_ez['Maintenance Department'] == d)]['RN'].sum()
                
                l_ratio = lw_md / lw_m_tot if lw_m_tot > 0 else 0
                t_ratio = tw_md / tw_m_tot if tw_m_tot > 0 else 0
                
                m_name = emm if first else ""
                final_rows.append(["", m_name, d, lw_md, tw_md, tw_md - lw_md, f"{l_ratio:.2%}", f"{t_ratio:.2%}", f"{(t_ratio - l_ratio):.2%}", "", ""])
                first = False
                
            # Total 行
            lw_m_tot_val = lw_ez[lw_ez['Cleaned_MM'].str.contains(emm, na=False)]['RN'].sum()
            tw_m_tot_val = tw_ez[tw_ez['Cleaned_MM'].str.contains(emm, na=False)]['RN'].sum()
            final_rows.append(["", "", "Total", lw_m_tot_val, tw_m_tot_val, tw_m_tot_val - lw_m_tot_val, "100.00%", "100.00%", "0.00%", "", ""])

        # 輸出轉換成 DataFrame
        out_df = pd.DataFrame(final_rows)
        
        st.success("🎉 符合 0629 標準格式數據建構完畢！")
        st.dataframe(out_df.fillna(""), use_container_width=True)
        
        # 匯出單一分頁 Excel 檔案
        output = io.BytesIO()
        sheet_title = f"2026 {tw_end.split('-')[1]}{tw_end.split('-')[2]}"
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            out_df.to_excel(writer, sheet_name=sheet_title, index=False, header=False)
            
        st.download_button(
            label="📥 下載 New Central 標準垂直週報 Excel",
            data=output.getvalue(),
            file_name=f"New_Central_周報_{tw_end.split('-')[1]}{tw_end.split('-')[2]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"轉換發生錯誤，請確認上傳檔案。錯誤訊息: {e}")
