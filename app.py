import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
from datetime import datetime
import pdfplumber

st.set_page_config(layout="wide", page_title="玄武鑑識中心")

def parse_pdf_data(file):
    try:
        # 初始化數據，年度設為 None 方便後續判斷
        res = {"公司名稱": file.name.replace(".pdf", ""), "年度": None, "營收": 0, "應收帳款": 0, "存貨": 0, "現金": 0, "負債總額": 0}
        
        with pdfplumber.open(file) as pdf:
            # 掃描前 3 頁，增加抓到表頭的機會
            for page in pdf.pages[:3]:
                text = page.extract_text()
                # 1. 優先從純文本中抓取年度 (支援 民國年 與 西元年)
                if not res["年度"]:
                    import re
                    # 匹配 "112年" 或 "2023年" 或 "2023"
                    years = re.findall(r"(\d{3,4})\s*年度", text)
                    if years:
                        y = int(years[0])
                        res["年度"] = y + 1911 if y < 1000 else y

                # 2. 處理表格
                table = page.extract_table()
                if not table: continue
                for row in table:
                    # 移除 None 並轉為字串
                    clean_row = [str(x) if x else "" for x in row]
                    row_str = "".join(clean_row)
                    
                    # 如果文本沒抓到年度，從表格第一列再試一次
                    if not res["年度"]:
                        for item in clean_row:
                            if item.isdigit() and 100 <= int(item) <= 2100:
                                y = int(item)
                                res["年度"] = y + 1911 if y < 1000 else y

                    # 數值提取 (取該列最後一個非空值，通常是本期數)
                    def get_val(r):
                        for v in reversed(r):
                            s = v.replace(",","").replace("(","-").replace(")","").strip()
                            try: return float(s)
                            except: continue
                        return 0

                    if any(k in row_str for k in ["營收", "營業收入", "Revenue"]): res["營收"] = get_val(clean_row)
                    if "應收帳款" in row_str: res["應收帳款"] = get_val(clean_row)
                    if "存貨" in row_str: res["存貨"] = get_val(clean_row)
                    if "現金" in row_str: res["現金"] = get_val(clean_row)
                    if "負債" in row_str: res["負債總額"] = get_val(clean_row)
        
        return pd.DataFrame([res]) if res["年度"] else pd.DataFrame()
    except:
        return pd.DataFrame()

# 鑑識邏輯優化 (確保不會因為除以零崩潰)
def forensic_logic(df):
    cols = ['M分數', '舞弊狀態', '掏空風險', '吸金指標']
    for c in cols: df[c] = 0.0 if c == 'M分數' else "正常"
    
    for i in df.index:
        rev = float(df.at[i, '營收'])
        if rev > 0:
            m = -3.2 + (0.15 * (df.at[i, '應收帳款']/rev)) + (0.1 * (df.at[i, '存貨']/rev))
            df.at[i, 'M分數'] = round(m, 2)
            df.at[i, '舞弊狀態'] = "危險" if m > -1.78 else "正常"
            df.at[i, '掏空風險'] = "高風險" if (df.at[i, '應收帳款']/rev) > 0.4 else "正常"
            cash = float(df.at[i, '現金'])
            df.at[i, '吸金指標'] = "警示" if cash > 0 and (df.at[i, '負債總額']/cash) > 5 else "正常"
    return df

# --- 主介面 ---
st.markdown("<h1 style='color:#002b36;'>玄武快機師事務所</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.header("控制台")
    files = st.file_uploader("請上傳財報 (PDF/Excel)", type=["pdf", "xlsx"], accept_multiple_files=True)
    auditor = st.text_input("主辦會計師", "張鈞翔會計師")

if files:
    all_data = []
    for f in files:
        data = pd.read_excel(f) if f.name.endswith('.xlsx') else parse_pdf_data(f)
        if not data.empty: all_data.append(data)
    
    if all_data:
        main_df = pd.concat(all_data, ignore_index=True)
        main_df = main_df.drop_duplicates(subset=['公司名稱', '年度']).sort_values('年度')
        main_df = forensic_logic(main_df)
        
        # 顯示結果
        target = st.selectbox("受查對象", main_df['公司名稱'].unique())
        sub = main_df[main_df['公司名稱'] == target]
        
        # 繪圖
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='Actual Revenue')
        if len(sub) >= 2:
            growth = sub['營收'].pct_change().mean()
            pred = sub['營收'].iloc[-1] * (1 + growth)
            ax.plot([str(sub['年度'].iloc[-1]), str(int(sub['年度'].iloc[-1])+1)], 
                    [sub['營收'].iloc[-1], pred], '--', marker='s', label='AI Forecast')
        ax.legend()
        st.pyplot(fig)
        
        st.dataframe(sub)
    else:
        st.error("無法從檔案中提取年度數據。請檢查 PDF 是否為掃描圖檔(無法選取文字)或是格式過於特殊。")
