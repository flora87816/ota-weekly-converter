import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="New Central 自動化分析器", layout="wide")

st.title("📊 New Central 自動化分析系統")
st.subheader("Raw Data 原始數據精準對齊版")

# 匯率設定 (由使用者在側邊欄調整)
st.sidebar.header("💱 匯率設置")
rate_twd = st.sidebar.number_input("TWD 兌 RMB 匯率", value=0.222, format="%.4f")
fx_rates = {'TWD': rate_twd, 'RMB': 1.0, 'CNY': 1.0}

def get_rmb_gmv(row):
    return float(row['gmv']) * fx_rates.get(str(row['Currency']).upper().strip(), 1.0)

uploaded_file = st.file_uploader("請上傳 OverseaDataMarket 原始檔案", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Book Time'] = pd.to_datetime(df['Book Time'])
    df['gmv_rmb'] = df.apply(get_rmb_gmv, axis=1)
    
    # 篩選兩週數據 (以 Book Time 為錨點)
    lw = df[(df['Book Time'] >= "2026-06-12") & (df['Book Time'] < "2026-06-19")]
    tw = df[(df['Book Time'] >= "2026-06-19") & (df['Book Time'] < "2026-06-26")]

    def get_calc_row(l_df, t_df):
        l_rn, t_rn = l_df['RN'].sum(), t_df['RN'].sum()
        l_rev, t_rev = l_df['gmv_rmb'].sum(), t_df['gmv_rmb'].sum()
        l_adr = l_rev / l_rn if l_rn > 0 else 0
        t_adr = t_rev / t_rn if t_rn > 0 else 0
        w_rn = (t_rn - l_rn) / l_rn if l_rn != 0 else 0
        w_rev = (t_rev - l_rev) / l_rev if l_rev != 0 else 0
        w_adr = (t_adr - l_adr) / l_adr if l_adr != 0 else 0
        return [l_rn, l_rev, l_adr, t_rn, t_rev, t_adr, f"{w_rn:.2%}", f"{w_rev:.2%}", f"{w_adr:.2%}"]

    # 建立表格
    results = []
    cities = ["台中", "台南", "高雄", "花蓮", "南投"]
    for c in cities:
        results.append([f"--- {c} ---"] + [""]*9)
        for s in [3, 4, 5]:
            results.append([f"{s} 星"] + get_calc_row(lw[(lw['City']==c)&(lw['Star']==s)], tw[(tw['City']==c)&(tw['Star']==s)]))
        results.append(["Total"] + get_calc_row(lw[lw['City']==c], tw[tw['City']==c]))

    out_df = pd.DataFrame(results, columns=["項目", "LW RN", "LW REV", "LW ADR", "TW RN", "TW REV", "TW ADR", "WoW RN%", "WoW REV%", "WoW ADR%"])
    st.dataframe(out_df)

    # 下載檔案，包含條件格式化標紅
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        out_df.to_excel(writer, index=False)
        ws = writer.sheets['Sheet1']
        red_format = writer.book.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        # 自動標記 WoW 中含有負號 '-' 的儲存格為紅色
        ws.conditional_format('H2:J100', {'type': 'text', 'criteria': 'containing', 'value': '-', 'format': red_format})
        
    st.download_button("📥 下載 New Central 報表", data=buf.getvalue(), file_name="New_Central_Report.xlsx")
