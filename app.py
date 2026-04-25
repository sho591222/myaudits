import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
import os
from datetime import datetime
import pdfplumber

# 1. 系統環境優化
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

# 2. 核心數據處理引擎
def parse_pdf_data(file):
    try:
        results = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0, "應收帳款": 0, "存貨": 0, "現金": 0, "負債總額": 0}
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if not table: continue
                for row in table:
                    row_str = "".join([str(x) for x in row if x])
                    # 年度抓取
                    for item in row:
                        s = str(item).strip()
                        if s.isdigit() and 1900 <= int(s) <= 2100: results["年度"] = int(s)
                    # 數值抓取邏輯
                    def clean_n(val):
                        if not val: return 0
                        s = str(val).replace(",","").replace("(","-").replace(")","").strip()
                        return pd.to_numeric(s, errors='coerce') if s else 0
                    if any(k in row_str for k in ["營收", "營業收入"]): results["營收"] = clean_n(row[-1])
                    if "應收帳款" in row_str: results["應收帳款"] = clean_n(row[-1])
                    if "存貨" in row_str: results["存貨"] = clean_n(row[-1])
                    if "現金" in row_str: results["現金"] = clean_n(row[-1])
                    if "負債" in row_str: results["負債總額"] = clean_n(row[-1])
        return pd.DataFrame([results])
    except:
        return pd.DataFrame()

# 3. 穩定版鑑識模型
def run_forensic(row):
    # 預設回傳 4 個穩定值
    out = pd.Series([0.0, "數據不足", "正常", "正常"], index=['M分數', '舞弊狀態', '掏空風險', '吸金指標'])
    try:
        r = float(row.get('營收', 0))
        if r <= 0: return out
        m = -3.2 + (0.15 * (float(row.get('應收帳款', 0))/r)) + (0.1 * (float(row.get('存貨', 0))/r))
        out['M分數'] = round(m, 2)
        out['舞弊狀態'] = "危險" if m > -1.78 else "正常"
        out['掏空風險'] = "高風險" if (float(row.get('應收帳款', 0))/r) > 0.4 else "正常"
        out['吸金指標'] = "警示" if float(row.get('現金', 0)) > 0 and (float(row.get('負債總額', 0))/float(row.get('現金', 0))) > 5 else "正常"
    except: pass
    return out

# 4. 網頁介面
st.markdown("<h1 style='color:#002b36;'>玄武快機師事務所</h1>", unsafe_allow_html=True)
st.markdown("<p>AI 財務鑑識與成長預測系統</p>", unsafe_allow_html=True)

with st.sidebar:
    st.header("控制台")
    st.text("玄武鑑定 真偽分明")
    mode = st.radio("模式", ["單一公司", "多公司分析"])
    auditor = st.text_input("主辦會計師", "張鈞翔會計師")
    files = st.file_uploader("上傳財報", type=["pdf", "xlsx"], accept_multiple_files=True)

if files:
    dfs = []
    for f in files:
        dfs.append(pd.read_excel(f) if f.name.endswith('.xlsx') else parse_pdf_data(f))
    
    if dfs:
        df = pd.concat(dfs, ignore_index=True)
        df = df[df['年度'] > 0].drop_duplicates(subset=['公司名稱', '年度']).sort_values('年度')
        
        if not df.empty:
            df[['M分數', '舞弊狀態', '掏空風險', '吸金指標']] = df.apply(run_forensic, axis=1)

            if mode == "深度鑑定":
                target = st.selectbox("選擇對象", df['公司名稱'].unique())
                sub = df[df['公司名稱'] == target]
                
                # 圖表：營收與預測
                fig, ax = plt.subplots(figsize=(10, 4))
                ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='歷史營收')
                # 簡易線性預測
                if len(sub) >= 2:
                    growth = sub['營營' if '營營' in sub else '營收'].pct_change().mean()
                    next_rev = sub['營收'].iloc[-1] * (1 + (growth if not np.isnan(growth) else 0))
                    ax.plot([str(sub['年度'].iloc[-1]), str(sub['年度'].iloc[-1]+1)], 
                            [sub['營收'].iloc[-1], next_rev], '--', marker='s', label='AI預測')
                ax.legend()
                st.pyplot(fig)
                st.dataframe(sub)

            if st.button("下載鑑定報告"):
                doc = Document()
                doc.add_heading("玄武會計師事務所 鑑定報告", 0)
                doc.add_paragraph(f"主辦會計師：{auditor}")
                for co in df['公司名稱'].unique():
                    doc.add_paragraph(f"受查對象：{co}，鑑定結論見下方表格。")
                buf = io.BytesIO()
                doc.save(buf)
                st.download_button("點此下載", buf.getvalue(), "鑑定報告.docx")
