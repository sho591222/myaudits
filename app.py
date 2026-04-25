import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
import re
import pdfplumber

# 1. 系統環境
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財務鑑識旗艦平台：多維度預測與風險 PK 整合系統</p>
    </div>
""", unsafe_allow_html=True)

# 2. 強效數值解析 (延續穩定版邏輯)
def force_extract_numbers(row_list):
    vals = []
    for item in row_list:
        if not item: continue
        s = str(item).strip().replace(',', '')
        if '(' in s and ')' in s: s = '-' + s.replace('(', '').replace(')', '')
        clean_s = "".join(re.findall(r'[0-9\.-]+', s))
        try:
            val = float(clean_s)
            vals.append(val)
        except: continue
    return vals

def parse_financial_data(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, 
            "應收帳款": 0.0, "存貨": 0.0, "現金": 0.0, "負債總額": 0.0, 
            "其他應收款": 0.0, "預付款項": 0.0, "股份酬勞": 0.0
        }
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:6]:
                text = page.extract_text() or ""
                if res["年度"] == 0:
                    years = re.findall(r"(\d{3,4})\s*年度", text)
                    if years:
                        y = int(years[0])
                        res["年度"] = y + 1911 if y < 1000 else y
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        row_str = "".join([str(x) for x in row if x])
                        nums = force_extract_numbers(row)
                        if not nums: continue
                        val = nums[-1]
                        if any(k in row_str for k in ["營業收入", "營收"]): res["營收"] = val
                        if "應收帳款" in row_str and "其他" not in row_str: res["應收帳款"] = val
                        if "存貨" in row_str: res["存貨"] = val
                        if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = val
                        if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = val
                        if "其他應收款" in row_str: res["其他應收款"] = val
                        if "預付款項" in row_str: res["預付款項"] = val
                        if any(k in row_str for k in ["股份酬勞", "認股權"]): res["股份酬勞"] = val
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except: return pd.DataFrame()

# 3. 預測模型引擎
def get_forecast(history, method):
    if len(history) < 1: return 0
    last_val = history.iloc[-1]
    if method == "線性成長 (5%)":
        return last_val * 1.05
    elif method == "保守估計 (0%)":
        return last_val * 1.0
    elif method == "歷史平均成長":
        if len(history) < 2: return last_val * 1.03
        rate = history.pct_change().mean()
        return last_val * (1 + rate)
    return last_val

# 4. 側邊欄：功能選單與預測選擇
with st.sidebar:
    st.header("⚙️ 鑑定功能中心")
    view_mode = st.radio("主要功能", [" 單一標的深度分析", " 多公司同年度 PK"])
    st.divider()
    
    st.subheader("📈 財務預測設定")
    forecast_method = st.selectbox("預測模型選擇", ["線性成長 (5%)", "保守估計 (0%)", "歷史平均成長"])
    
    st.divider()
    uploaded_files = st.file_uploader("上傳財報資料", type=["pdf", "xlsx"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "會計師")

# 5. 主程式
if uploaded_files:
    data_list = [parse_financial_data(f) for f in uploaded_files]
    df_pool = pd.concat([d for d in data_list if not d.empty], ignore_index=True)
    
    if not df_pool.empty:
        # 單一公司深度分析
        if view_mode == " 單一標的深度分析":
            target = st.selectbox("受查對象", df_pool['公司名稱'].unique())
            sub = df_pool[df_pool['公司名稱'] == target].sort_values('年度')
            
            st.subheader(f"財務鑑識看板：{target}")
            
            # 趨勢圖：含可選預測
            fig, ax = plt.subplots(figsize=(10, 4))
            years = sub['年度'].astype(str).tolist()
            revs = sub['營收'].tolist()
            
            # 繪製歷史線
            ax.plot(years, revs, marker='o', label='Historical Revenue', linewidth=2, color='#268bd2')
            
            # 繪製預測線
            next_year = str(int(years[-1]) + 1)
            forecast_val = get_forecast(sub['營收'], forecast_method)
            ax.plot([years[-1], next_year], [revs[-1], forecast_val], '--', marker='s', label=f'Forecast ({forecast_method})', color='#cb4b16')
            
            ax.set_title("Revenue Dynamics & Selection Forecast")
            ax.legend()
            st.pyplot(fig)
            
            st.dataframe(sub)

        # 多公司同年度 PK
        else:
            target_year = st.selectbox("PK年度選擇", sorted(df_pool['年度'].unique(), reverse=True))
            pk_df = df_pool[df_pool['年度'] == target_year]
            
            st.subheader(f"西元 {int(target_year)} 年度橫向 PK")
            col1, col2 = st.columns(2)
            with col1:
                st.write("舞弊風險 PK (M-Score)")
                fig2, ax2 = plt.subplots()
                ax2.bar(pk_df['公司名稱'], pk_df['營收']) # 範例圖表
                st.pyplot(fig2)
            with col2:
                st.write("掏空風險 PK")
                st.dataframe(pk_df[['公司名稱', '年度', '營收']])

        # Word 報告匯出
        st.divider()
        if st.button("產出鑑定報告書 (.docx)"):
            doc = Document()
            doc.add_heading("玄武鑑定旗艦報告", 0)
            doc.add_paragraph(f"本報告採用預測模型：{forecast_method}")
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("下載報告", buf, "XuanWu_Forensic_Report.docx")
    else:
        st.warning("請確認上傳檔案是否包含有效年度與營收數據。")
else:
    st.info("系統就緒，請上傳受查檔案。")
