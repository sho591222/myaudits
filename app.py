import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
import re
import pdfplumber

# 1. 系統環境與 UI
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財務鑑識旗艦平台：整合「認股選項」與「同業 PK」模組</p>
    </div>
""", unsafe_allow_html=True)

# 2. 強效數值提取與關鍵科目擴充
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
            "其他應收款": 0.0, "預付款項": 0.0, "股份酬勞": 0.0  # 新增選項科目
        }
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:8]: # 擴大掃描至報表附註
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
                        # 標準科目匹配
                        if any(k in row_str for k in ["營業收入", "營收"]): res["營收"] = val
                        if "應收帳款" in row_str and "其他" not in row_str: res["應收帳款"] = val
                        if "存貨" in row_str: res["存貨"] = val
                        if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = val
                        if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = val
                        if "其他應收款" in row_str: res["其他應收款"] = val
                        if "預付款項" in row_str: res["預付款項"] = val
                        # 「選項」相關科目偵測 (股份酬勞、認股權)
                        if any(k in row_str for k in ["股份酬勞", "認股權", "RSU"]): res["股份酬勞"] = val
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except: return pd.DataFrame()

# 3. 鑑識邏輯：包含「選項操縱」分析
def run_forensics(df):
    for c in ['M分數', '掏空指數', '選項異常度', '結論']: df[c] = 0.0
    for i in df.index:
        r = float(df.at[i, '營收'])
        if r > 0:
            # 1. 舞弊 M-Score
            m = -3.2 + (0.15 * (df.at[i, '應收帳款']/r)) + (0.1 * (df.at[i, '存貨']/r))
            df.at[i, 'M分數'] = round(m, 2)
            # 2. 掏空指標 (資產偏移)
            df.at[i, '掏空指數'] = round((df.at[i, '其他應收款'] + df.at[i, '預付款項']) / r, 3)
            # 3. 選項操縱 (若股份酬勞佔營收比重突增，代表可能透過選項輸送利益)
            opt_ratio = df.at[i, '股份酬勞'] / r
            df.at[i, '選項異常度'] = round(opt_ratio, 4)
            # 4. 綜合鑑定
            if m > -1.78 or df.at[i, '掏空指數'] > 0.2 or opt_ratio > 0.05:
                df.at[i, '結論'] = "建議深度查核"
            else:
                df.at[i, '結論'] = "暫無異常"
    return df

# 4. 側邊欄：功能選單
with st.sidebar:
    st.header("⚙️ 鑑定功能選單")
    view_mode = st.radio("模式選擇", ["🔍 單一深度報告", "⚔️ 多公司同年度 PK"])
    st.divider()
    # 新增診斷選項勾選
    st.subheader("診斷開關")
    check_fraud = st.checkbox("舞弊偵測 (M-Score)", value=True)
    check_tunnel = st.checkbox("掏空預警 (其他應收)", value=True)
    check_option = st.checkbox("認股權操縱分析", value=True)
    st.divider()
    uploaded_files = st.file_uploader("批次上傳受查財報", type=["pdf", "xlsx"], accept_multiple_files=True)

# 5. 主程式
if uploaded_files:
    df_pool = pd.concat([parse_financial_data(f) for f in uploaded_files], ignore_index=True)
    
    if not df_pool.empty:
        df_pool = run_forensics(df_pool)

        if view_mode == "⚔️ 多公司同年度 PK":
            target_year = st.selectbox("選擇比對年度", sorted(df_pool['年度'].unique(), reverse=True))
            pk_df = df_pool[df_pool['年度'] == target_year]
            
            st.subheader(f"西元 {int(target_year)} 年度：橫向對比分析")
            
            # --- PK 圖表：加入「選項異常度」---
            col1, col2, col3 = st.columns(3)
            with col1:
                st.write("舞弊風險 PK")
                fig1, ax1 = plt.subplots()
                ax1.bar(pk_df['公司名稱'], pk_df['M分數'], color='#268bd2')
                ax1.axhline(y=-1.78, color='red', linestyle='--')
                st.pyplot(fig1)
            with col2:
                st.write("掏空壓力 PK")
                fig2, ax2 = plt.subplots()
                ax2.bar(pk_df['公司名稱'], pk_df['掏空指數'], color='#b58900')
                st.pyplot(fig2)
            with col3:
                st.write("認股權操縱 PK")
                fig3, ax3 = plt.subplots()
                ax3.bar(pk_df['公司名稱'], pk_df['選項異常度'], color='#dc322f')
                st.pyplot(fig3)

            st.dataframe(pk_df)

        else: # 單一深度分析
            target_co = st.selectbox("選擇受查公司", df_pool['公司名稱'].unique())
            sub = df_pool[df_pool['公司名稱'] == target_co]
            st.subheader(f"趨勢追蹤：{target_co}")
            
            # 趨勢圖：營收 vs 股份酬勞 (選項)
            fig4, ax4 = plt.subplots(figsize=(10, 4))
            ax4.plot(sub['年度'].astype(str), sub['營收'], label='Revenue', marker='o')
            ax4.bar(sub['年度'].astype(str), sub['股份酬勞'], label='Stock Options (Expense)', color='red', alpha=0.3)
            ax4.legend()
            st.pyplot(fig4)
            
            st.dataframe(sub)

        # 匯出報告 (含表格與結論)
        st.divider()
        if st.button("產出鑑定報告書 (.docx)"):
            doc = Document()
            doc.add_heading(f"玄武鑑識鑑定報告 - {view_mode}", 0)
            doc.add_paragraph(f"本報告包含：{'舞弊' if check_fraud else ''} {'掏空' if check_tunnel else ''} {'認股權' if check_option else ''} 分析。")
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("點此下載鑑定報告", buf, "XuanWu_Report.docx")
    else:
        st.error("未能成功解析數據，請確認 PDF 內文格式。")
