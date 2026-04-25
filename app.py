import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
import os
from datetime import datetime
import pdfplumber

# 1. 基礎設定：完全移除外部字體請求以防 503 錯誤
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

# 2. 核心數據處理引擎：優化內存佔用
def parse_pdf_data(file):
    try:
        results = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0, "應收帳款": 0, "存貨": 0, "現金": 0, "負債總額": 0}
        with pdfplumber.open(file) as pdf:
            # 僅掃描前兩頁，避免超大檔案導致內存溢出
            for page in pdf.pages[:2]:
                table = page.extract_table()
                if not table: continue
                for row in table:
                    row_str = "".join([str(x) for x in row if x])
                    # 年度抓取
                    for item in row:
                        s = str(item).strip()
                        if s.isdigit() and 1900 <= int(s) <= 2100:
                            results["年度"] = int(s)
                    # 數值提取函數
                    def clean_v(val):
                        if val is None: return 0
                        s = str(val).replace(",","").replace("(","-").replace(")","").strip()
                        try: return float(s) if s else 0
                        except: return 0
                    
                    if any(k in row_str for k in ["營收", "營業收入"]): results["營收"] = clean_v(row[-1])
                    if "應收帳款" in row_str: results["應收帳款"] = clean_v(row[-1])
                    if "存貨" in row_str: results["存貨"] = clean_v(row[-1])
                    if "現金" in row_str: results["現金"] = clean_v(row[-1])
                    if "負債" in row_str: results["負債總額"] = clean_v(row[-1])
        return pd.DataFrame([results])
    except:
        return pd.DataFrame()

# 3. 穩定版鑑識計算邏輯
def get_forensic_results(df):
    # 預先建立空欄位，避免直接 apply 導致的 ValueError
    df['M分數'] = 0.0
    df['舞弊狀態'] = "正常"
    df['掏空風險'] = "正常"
    df['吸金指標'] = "正常"
    
    for i in df.index:
        r = df.at[i, '營營' if '營營' in df else '營收']
        rc = df.at[i, '應收帳款']
        inv = df.at[i, '存貨']
        cash = df.at[i, '現金']
        debt = df.at[i, '負債總額']
        
        if r > 0:
            m = -3.2 + (0.15 * (rc/r)) + (0.1 * (inv/r))
            df.at[i, 'M分數'] = round(m, 2)
            df.at[i, '舞弊狀態'] = "危險" if m > -1.78 else "正常"
            df.at[i, '掏空風險'] = "高風險" if (rc/r) > 0.4 else "正常"
            df.at[i, '吸金指標'] = "警示" if cash > 0 and (debt/cash) > 5 else "正常"
    return df

# 4. 介面呈現
st.markdown("<h1 style='color:#002b36;'>玄武快機師事務所</h1>", unsafe_allow_html=True)
st.markdown("<p style='font-size:18px;'>AI 財務鑑識與成長預測系統</p>", unsafe_allow_html=True)

with st.sidebar:
    st.header("控制台")
    st.text("玄武鑑定 真偽分明")
    mode = st.radio("模式選擇", ["深度鑑定分析", "競爭力PK對比"])
    auditor = st.text_input("主辦會計師簽署", "張鈞翔會計師")
    st.divider()
    files = st.file_uploader("請上傳財報數據 (PDF或Excel)", type=["pdf", "xlsx"], accept_multiple_files=True)

if files:
    data_list = []
    for f in files:
        if f.name.endswith('.xlsx'):
            data_list.append(pd.read_excel(f))
        else:
            data_list.append(parse_pdf_data(f))
    
    if data_list:
        all_df = pd.concat(data_list, ignore_index=True)
        # 過濾髒數據並確保數據格式正確
        all_df = all_df[all_df['年度'] > 0].drop_duplicates(subset=['公司名稱', '年度']).sort_values('年度')
        
        if not all_df.empty:
            # 執行鑑識運算
            all_df = get_forensic_results(all_df)

            if mode == "深度鑑定分析":
                target = st.selectbox("選擇受查對象", all_df['公司名稱'].unique())
                sub = all_df[all_df['公司名稱'] == target]
                
                # 營收與預測圖表
                st.subheader("營收歷史軌跡與成長預估圖")
                fig, ax = plt.subplots(figsize=(10, 4))
                ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='Actual Revenue', linewidth=2)
                
                # AI 成長預測線邏輯
                if len(sub) >= 2:
                    last_rev = float(sub['營收'].iloc[-1])
                    growth_rate = sub['營收'].pct_change().mean()
                    if not np.isnan(growth_rate):
                        pred_rev = last_rev * (1 + growth_rate)
                        ax.plot([str(sub['年度'].iloc[-1]), str(sub['年度'].iloc[-1]+1)], 
                                [last_rev, pred_rev], '--', marker='s', label='AI Forecast')
                
                ax.set_xlabel("Year")
                ax.set_ylabel("Revenue")
                ax.legend()
                st.pyplot(fig)
                
                # 數據指標展示
                col1, col2, col3 = st.columns(3)
                last_row = sub.iloc[-1]
                col1.metric("M-Score", last_row['M分數'], last_row['舞弊狀態'], delta_color="inverse")
                col2.metric("Tunneling Risk", last_row['掏空風險'])
                col3.metric("Ponzi Warning", last_row['吸金指標'])
                
                st.write("受查年度明細數據")
                st.dataframe(sub)

            # 報告產出模組
            st.divider()
            if st.button("點此匯出正式鑑定報告檔案"):
                doc = Document()
                doc.add_heading("玄武會計師事務所 鑑定報告", 0)
                doc.add_paragraph(f"主辦會計師: {auditor}")
                doc.add_paragraph(f"報告日期: {datetime.now().strftime('%Y-%m-%d')}")
                
                for co in all_df['公司名稱'].unique():
                    co_sub = all_df[all_df['公司名稱'] == co]
                    if len(co_sub) >= 2:
                        avg_g = co_sub['營收'].pct_change().mean()
                        doc.add_heading(f"受查對象: {co}", level=1)
                        doc.add_paragraph(f"經分析，該公司歷史平均年增率為 {avg_g:.2%}。")
                        # 顯示舞弊警示
                        if any(co_sub['舞弊狀態'] == "危險"):
                            doc.add_paragraph("警告: 該公司於受查期間偵測到高度盈餘操縱風險。")
                
                stream = io.BytesIO()
                doc.save(stream)
                st.download_button("下載 Word 報告", stream.getvalue(), "Forensic_Report.docx")
        else:
            st.warning("檔案中未能成功提取有效年度數據。")
else:
    st.info("系統就緒。請從側邊欄上傳財報以開始。")
