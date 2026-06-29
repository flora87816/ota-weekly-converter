import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="New Central 2026 最終對齊版", layout="wide")

st.title("📊 New Central 標準數據轉換工具 (最終對齊版)")

# 1. 匯率設定
st.sidebar.header("💱 匯率設定")
rate_twd = st.sidebar.number_input("TWD 新台幣兌 RMB 匯率", value=0.222, format="%.4f")
fx_rates = {'TWD': rate_twd, 'RMB': 1.0, 'CNY': 1.0}

def get_rmb(row):
    return float(row['gmv']) * fx_rates.get(str(row['Currency']).upper().strip(), 1.0)

# 2. 上傳與處理
uploaded_file = st.file_uploader("請上傳原始資料", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Book Time'] = pd.to_datetime(df['Book Time'])
    df['gmv_rmb'] = df.apply(get_rmb, axis=1)
    
    lw_sub = df[(df['Book Time'] >= pd.Timestamp("2026-06-12")) & (df['Book Time'] < pd.Timestamp("2026-06-19"))]
    tw_sub = df[(df['Book Time'] >= pd.Timestamp("2026-06-19")) & (df['Book Time'] < pd.Timestamp("2026-06-26"))]

    def calc_row(l_df, t_df):
        l_rn, t_rn = l_df['RN'].sum(), t_df['RN'].sum()
        l_rev, t_rev = l_df['gmv_rmb'].sum(), t_df['gmv_rmb'].sum()
        l_adr = l_rev / l_rn if l_rn > 0 else 0
        t_adr = t_rev / t_rn if t_rn > 0 else 0
        w_rn = (t_rn - l_rn) / l_rn if l_rn != 0 else 0
        w_rev = (t_rev - l_rev) / l_rev if l_rev != 0 else 0
        w_adr = (t_adr - l_adr) / l_adr if l_adr != 0 else 0
        return [l_rn, l_rev, l_adr, t_rn, t_rev, t_adr, f"{w_rn:.2%}", f"{w_rev:.2%}", f"{w_adr:.2%}"]

    # 產生城市表格
    rows = []
    cities = ["台中", "台南", "高雄", "花蓮", "南投"]
    for c in cities:
        rows.append([f"--- {c} ---"] + [""]*9)
        for s in [3, 4, 5]:
            rows.append([f"{s} 星"] + calc_row(lw_sub[(lw_sub['City']==c)&(lw_sub['Star']==s)], tw_sub[(tw_sub['City']==c)&(tw_sub['Star']==s)]))
        rows.append(["Total"] + calc_row(lw_sub[lw_sub['City']==c], tw_sub[tw_sub['City']==c]))
    
    out_df = pd.DataFrame(rows, columns=["項目", "LW RN", "LW REV", "LW ADR", "TW RN", "TW REV", "TW ADR", "WoW RN%", "WoW REV%", "WoW ADR%"])
    st.dataframe(out_df)

    # 匯出與美化
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        out_df.to_excel(writer, index=False)
        ws = writer.sheets['Sheet1']
        red_fmt = writer.book.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        ws.conditional_format('H2:J100', {'type': 'text', 'criteria': 'containing', 'value': '-', 'format': red_fmt})
    st.download_button("📥 下載對齊數據", data=output.getvalue(), file_name="Aligned_Final.xlsx")
