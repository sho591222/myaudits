import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
import re
import pdfplumber

# 1. 系統環境與穩定性設定
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

# 移除所有表情符號，確保後端編碼不崩潰
st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財務鑑識旗艦平台：數值精準解析與風險預警</p>
    </div>
""", unsafe_allow_html=True)

# 2. 強效數值清理函數 (解決圖表點在 0 的關鍵)
def strong_clean_val(v):
    if v is None: return 0.0
    s = str(v).strip()
    # 處理會計格式：(1,234.56) 轉為 -1234.56
    if '(' in s and ')' in s:
        s = '-' + s.replace('(', '').replace(')', '')
    # 移除千分位逗號、全形空格、百分比符號與所有中文
    s = re.sub(r'[^\d\.\-]', '', s)
    try:
        return float(s) if s else 0.0
    except:
        return 0.0

# 3. 數據解析引擎 (PDF & Excel)
def parse_financial_data(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), 
            "年度": None, "營收": 0.0, "應收帳款": 0.0, 
            "存貨": 0.0, "現金": 0.0, "負債總額": 0.0, 
            "其他應收款": 0.0, "預付款項": 0.0
        }
        
        if file.name.endswith('.xlsx'):
            df_xlsx = pd.read_excel(file)
            df_xlsx.columns = [str(c).strip() for c in df_xlsx.columns]
            rename_map = {"年份": "年度", "Year": "年度", "營業收入": "營收"}
            df_xlsx = df_xlsx.rename(columns=rename_map)
            # 數值列全部強制清理
            for col in res.keys():
                if col in df_xlsx.columns and col != "公司名稱":
                    df_xlsx[col] = df_xlsx[col].apply(strong_clean_val)
            return df_xlsx

        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:5]:
                text = page.extract_text() or ""
                if not res["年度"]:
                    years = re.findall(r"(\d{3,4})\s*年度", text)
                    if years:
                        y = int(years[0])
                        res["年度"] = y + 1911 if y < 1000 else y

                table = page.extract_table()
                if not table: continue
                for row in table:
                    clean_row = [str(x) if x else "" for x in row]
                    row_str = "".join(clean_row)
                    
                    # 關鍵科目抓取
                    def get_last_val(r):
                        for v in reversed(r):
                            val = strong_clean_val(v)
                            if val != 0: return val
                        return 0.0

                    if any(k in row_str for k in ["營收", "營業收入"]): res["營收"] = get_last_val(clean_row)
                    if "應收帳款" in row_str: res["應收帳款"] = get_last_val(clean_row)
                    if "存貨" in row_str: res["存貨"] = get_last_val(clean_row)
                    if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = get_last_val(clean_row)
                    if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = get_last_val(clean_row)
                    if "其他應收款" in row_str: res["其他應收款"] = get_last_val(clean_row)
                    if "預付款項" in row_str: res["預付款項"] = get_last_val(clean_row)

        return pd.DataFrame([res]) if res["年度"] else pd.DataFrame()
    except:
        return pd.DataFrame()

# 4. 鑑識邏輯
def run_forensics(df):
    df['M分數'] = 0.0
    df['舞弊風險'] = "正常"
    df['掏空指數'] = 0.0
    for i in df.index:
        r = df.at[i, '營收']
        if r > 0:
            m = -3.2 + (0.15 * (df.at[i, '應收帳款']/r)) + (0.1 * (df.at[i, '存貨']/r))
            df.at[i, 'M分數'] = round(m, 2)
            df.at[i, '舞弊風險'] = "危險" if m > -1.78 else "正常"
            df.at[i, '掏空指數'] = round((df.at[i, '其他應收款'] + df.at[i, '預付款項']) / r, 3)
    return df

# 5. UI 與繪圖 (優化軸距，防止點擠在 0)
with st.sidebar:
    st.header("功能中心")
    mode = st.radio("視角選擇", ["單一深度分析", "多公司PK"])
    files = st.file_uploader("批次上傳數據 (PDF/Excel)", type=["pdf", "xlsx"], accept_multiple_files=True)

if files:
    main_df = pd.concat([parse_financial_data(f) for f in files if not parse_financial_data(f).empty], ignore_index=True)
    if not main_df.empty and '年度' in main_df.columns:
        main_df = main_df[main_df['年度'] > 0].sort_values('年度')
        main_df = run_forensics(main_df)
        
        target = st.selectbox("選擇公司", main_df['公司名稱'].unique())
        sub = main_df[main_df['公司名稱'] == target]
        
        # 繪圖區
        c1, c2 = st.columns(2)
        with c1:
            st.write("Revenue Dynamics (營收趨勢)")
            fig, ax = plt.subplots(figsize=(5, 3))
            ax.plot(sub['年度'].astype(str), sub['營營' if '營營' in sub else '營收'], marker='o', linewidth=2)
            # 自動調整 Y 軸刻度，避免擠在 0
            if sub['營收'].max() > 0:
                ax.set_ylim(sub['營收'].min() * 0.9, sub['營收'].max() * 1.1)
            st.pyplot(fig)
            
        with c2:
            st.write("Asset Tunelling Observe (掏空觀察)")
            fig2, ax2 = plt.subplots(figsize=(5, 3))
            ax2.bar(sub['年度'].astype(str), sub['其他應收款'], label='Other RCV', color='red', alpha=0.6)
            ax2.bar(sub['年度'].astype(str), sub['預付款項'], bottom=sub['其他應收款'], label='Prepay', color='orange', alpha=0.6)
            ax2.legend(prop={'size': 7})
            st.pyplot(fig2)
        
        st.write("鑑定數據清單")
        st.dataframe(sub)
    else:
        st.error("未能解析數值，請確認檔案中的數字格式是否清晰。")
