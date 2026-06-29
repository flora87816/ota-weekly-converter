import streamlit as st
import pandas as pd
import datetime

st.set_page_config(layout="wide")
st.title("📊 週報數據自動化生成系統 (Weekly Progress)")
st.write("只需上傳原始 `OverseaDataMarket` CSV 檔，即可自動計算 MM 績效與 WoW 變幅。")

# 1. 檔案上傳介面
uploaded_file = st.file_uploader("請上傳原始資料 (CSV 格式)", type=["csv"])

if uploaded_file is not None:
    try:
        # 讀取資料
        df = pd.read_csv(uploaded_file)
        
        # 檢查必備欄位
        required_cols = ['Book Time', 'MM', 'RN', 'gmv']
        if not all(col in df.columns for col in required_cols):
            st.error(f"上傳的檔案缺少必要欄位！必須包含: {required_cols}")
        else:
            # 轉換時間格式
            df['Book Time'] = pd.to_datetime(df['Book Time'])
            
            # 讓用戶確認或選擇日期範圍
            min_date = df['Book Time'].min().date()
            max_date = df['Book Time'].max().date()
            st.info(f"偵測到資料時間範圍：{min_date} 至 {max_date}")
            
            # 自動推算最近的兩週（以最後一天往前推 14 天）
            end_date = df['Book Time'].max()
            start_date = end_date - datetime.timedelta(days=14)
            
            # 切分兩週的區間 (後7天為本週，前7天為上週)
            w2_start = end_date - datetime.timedelta(days=7)
            
            st.subheader("📅 自動劃分的週區間")
            col1, col2 = st.columns(2)
            with col1:
                st.metric("上週區間 (W1)", f"{start_date.strftime('%m/%d')}-{w2_start.strftime('%m/%d')}")
            with col2:
                st.metric("本週區間 (W2)", f"{(w2_start + datetime.timedelta(days=1)).strftime('%m/%d')}-{end_date.strftime('%m/%d')}")

            # 標記資料屬於哪一週
            def assign_week(x):
                if start_date <= x < (w2_start + datetime.timedelta(days=1)):
                    return 'W1'
                elif (w2_start + datetime.timedelta(days=1)) <= x <= end_date:
                    return 'W2'
                return 'Other'

            df['Week'] = df['Book Time'].apply(assign_week)
            
            # 過濾出這兩週的有效資料
            df_filtered = df[df['Week'].isin(['W1', 'W2'])].copy()
            
            # 2. 依據 MM 聚合計算
            # 分別計算各周的 RN 和 GMV(REV)
            pivot = df_filtered.groupby(['MM', 'Week']).agg({
                'RN': 'sum',
                'gmv': 'sum'
            }).unstack(fill_value=0)
            
            # 重建扁平化欄位結構
            report = pd.DataFrame(index=pivot.index)
            
            # W1 數據
            report['W1_RN'] = pivot[('RN', 'W1')]
            report['W1_REV'] = pivot[('gmv', 'W1')]
            report['W1_ADR'] = report['W1_REV'] / report['W1_RN'].replace(0, 1) # 避免除以0
            
            # W2 數據
            report['W2_RN'] = pivot[('RN', 'W2')]
            report['W2_REV'] = pivot[('gmv', 'W2')]
            report['W2_ADR'] = report['W2_REV'] / report['W2_RN'].replace(0, 1)
            
            # 3. 計算 WoW (百分比)
            report['RN_WoW'] = (report['W2_RN'] - report['W1_RN']) / report['W1_RN'].replace(0, 1)
            report['REV_WoW'] = (report['W2_REV'] - report['W1_REV']) / report['W1_REV'].replace(0, 1)
            report['ADR_WoW'] = (report['W2_ADR'] - report['W1_ADR']) / report['W1_ADR'].replace(0, 1)
            
            # 4. 計算 Total (總計行)
            total_w1_rn = report['W1_RN'].sum()
            total_w1_rev = report['W1_REV'].sum()
            total_w1_adr = total_w1_rev / max(total_w1_rn, 1)
            
            total_w2_rn = report['W2_RN'].sum()
            total_w2_rev = report['W2_REV'].sum()
            total_w2_adr = total_w2_rev / max(total_w2_rn, 1)
            
            total_row = pd.DataFrame([{
                'W1_RN': total_w1_rn, 'W1_REV': total_w1_rev, 'W1_ADR': total_w1_adr,
                'W2_RN': total_w2_rn, 'W2_REV': total_w2_rev, 'W2_ADR': total_w2_adr,
                'RN_WoW': (total_w2_rn - total_w1_rn) / max(total_w1_rn, 1),
                'REV_WoW': (total_w2_rev - total_w1_rev) / max(total_w1_rev, 1),
                'ADR_WoW': (total_w2_adr - total_w1_adr) / max(total_w1_adr, 1)
            }], index=['Total'])
            
            report = pd.concat([report, total_row])
            
            # 5. 格式化欄位名稱與顯示格式
            # 重新調整欄位順序以符合 New Central 樣式
            display_cols = [
                'W1_RN', 'W2_RN', 'RN_WoW',
                'W1_REV', 'W2_REV', 'REV_WoW',
                'W1_ADR', 'W2_ADR', 'ADR_WoW'
            ]
            report = report[display_cols]
            
            # 欄位重新命名顯示
            w1_label = f"{start_date.strftime('%m/%d')}-{w2_start.strftime('%m/%d')}"
            w2_label = f"{(w2_start + datetime.timedelta(days=1)).strftime('%m/%d')}-{end_date.strftime('%m/%d')}"
            
            report.columns = [
                f'RN ({w1_label})', f'RN ({w2_label})', 'RN WoW',
                f'REV ({w1_label})', f'REV ({w2_label})', 'REV WoW',
                f'ADR ({w1_label})', f'ADR ({w2_label})', 'ADR WoW'
            ]
            
            # 顯示結果表格並套用格式 (REV, ADR 四捨五入，WoW 轉百分比)
            st.subheader("📋 統計結果 (MM概況)")
            
            formatted_report = report.copy()
            # 格式化百分比與金額
            for col in formatted_report.columns:
                if 'WoW' in col:
                    formatted_report[col] = formatted_report[col].apply(lambda x: f"{x:.2%}")
                elif 'REV' in col or 'ADR' in col:
                    formatted_report[col] = formatted_report[col].apply(lambda x: f"${x:,.2f}")
                else:
                    formatted_report[col] = formatted_report[col].apply(lambda x: f"{int(x)}")

            st.dataframe(formatted_report, use_container_width=True)
            
            # 提供 Excel 下載功能
            @st.cache_data
            def convert_df(df_to_download):
                return df_to_download.to_csv().encode('utf-8-sig')

            csv_data = convert_df(report)
            st.download_button(
                label="📥 下載整理後的報表 (CSV)",
                data=csv_data,
                file_name=f'MM_Weekly_Report_{end_date.strftime("%Y%m%d")}.csv',
                mime='text/csv',
            )
            
    except Exception as e:
        st.error(f"處裡檔案時發生錯誤: {e}")
