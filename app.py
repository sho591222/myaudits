import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
import re
import pdfplumber
from datetime import datetime

# 1. 系統環境設定
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財務鑑識旗艦平台：舞弊、掏空與洗錢防制整合分析</p>
    </div>
""", unsafe_allow_html=True)

# 2. 數據解析引擎 (強化年度與數值抓取)
def parse_financial_pdf(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), 
            "年度": 0, "營收": 0, "應收帳款": 0, 
            "存貨": 0, "現金": 0, "負債總額": 0, 
            "其他應收款": 0, "預付款項": 0
        }
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:5]:
                text = page.extract_text() or ""
                # 年度抓取強化
                if res["年度"] == 0:
                    years = re.findall(r"(\d{3,4})\s*年度", text)
                    if years:
                        y = int(years[0])
                        res["年度"] = y + 1911 if y < 1000 else y

                table = page.extract_table()
                if not table: continue
                for row in table:
                    clean_row = [str(x) if x else "" for x in row]
                    row_str = "".join(clean_row)
                    
                    def extract_val(r):
                        for v in reversed(r):
                            s = v.replace(",","").replace("(","-").replace(")","").strip()
                            try: return float(s)
                            except: continue
                        return 0

                    if any(k in row_str for k in ["營收", "營業收入"]): res["營收"] = extract_val(clean_row)
                    if "應收帳款" in row_str: res["應收帳款"] = extract_val(clean_row)
                    if "存貨" in row_str: res["存貨"] = extract_val(clean_row)
                    if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = extract_val(clean_row)
                    if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = extract_val(clean_row)
                    if "其他應收款" in row_str: res["其他應收款"] = extract_val(clean_row)
                    if "預付款項" in row_str: res["預付款項"] = extract_val(clean_row)
        
        return pd.DataFrame([res])
    except:
        return pd.DataFrame()

# 3. 綜合鑑識計算模組 (防呆補全版)
def run_integrated_forensics(df):
    # 強制檢查所有必要欄位，若缺失則補 0
    essential_cols = ['年度', '營收', '應收帳款', '存貨', '現金', '負債總額', '其他應收款', '預付款項']
    for col in essential_cols:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 建立分析欄位
    df['M分數'] = 0.0
    df['舞弊風險'] = "正常"
    df['掏空指數'] = 0.0
    df['掏空預警'] = "正常"
    df['洗錢分值'] = 0
    df['綜合鑑定意見'] = "查無異常"
    
    for i in df.index:
        r = df.at[i, '營收']
        rc = df.at[i, '應收帳款']
        inv = df.at[i, '存貨']
        cash = df.at[i, '現金']
        debt = df.at[i, '負債總額']
        other_rc = df.at[i, '其他應收款']
        prepay = df.at[i, '預付款項']
        
        if r > 0:
            m = -3.2 + (0.15 * (rc/r)) + (0.1 * (inv/r))
            df.at[i, 'M分數'] = round(m, 2)
            df.at[i, '舞弊風險'] = "危險" if m > -1.78 else "正常"
            t_ratio = (other_rc + prepay) / r
            df.at[i, '掏空指數'] = round(t_ratio, 3)
            df.at[i, '掏空預警'] = "高風險" if t_ratio > 0.2 else "正常"
            ml_score = 0
            if cash > r * 0.8: ml_score += 50
            if debt > cash * 5: ml_score += 30
            df.at[i, '洗錢分值'] = ml_score
            if m > -1.78 or t_ratio > 0.2 or ml_score >= 50:
                df.at[i, '綜合鑑定意見'] = "建議深度查核"
    return df

# 4. 側邊欄控制
with st.sidebar:
    st.header("功能控制台")
    view_mode = st.radio("請選擇分析視角", ["單一公司深度報告", "多公司競爭力與風險PK"])
    st.divider()
    auditor_sig = st.text_input("主辦會計師", "會計師")
    uploaded_files = st.file_uploader("上傳財報數據 (PDF/Excel)", type=["pdf", "xlsx"], accept_multiple_files=True)

# 5. 主程式邏輯
if uploaded_files:
    data_pool = []
    for f in uploaded_files:
        temp_df = pd.read_excel(f) if f.name.endswith('.xlsx') else parse_financial_pdf(f)
        if not temp_df.empty:
            # 統一 Excel 欄位名稱 (防呆)
            rename_map = {"年份": "年度", "Year": "年度", "營業收入": "營營"}
            temp_df = temp_df.rename(columns=rename_map)
            data_pool.append(temp_df)
    
    if data_pool:
        df_final = pd.concat(data_pool, ignore_index=True)
        
        # 關鍵修正：檢查是否存在 '年度' 欄位，若無則不執行後續
        if '年度' in df_final.columns:
            df_final = df_final[df_final['年度'] > 0].drop_duplicates(subset=['公司名稱', '年度']).sort_values('年度')
            df_final = run_integrated_forensics(df_final)

            # --- 下方渲染邏輯 ---
            if view_mode == "單一公司深度報告":
                target = st.selectbox("選擇受查公司", df_final['公司名稱'].unique())
                sub = df_final[df_final['公司名稱'] == target]
                
                latest = sub.iloc[-1]
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("舞弊指標 (M-Score)", latest['M分數'], latest['舞弊風險'], delta_color="inverse")
                m2.metric("掏空壓力", f"{latest['掏空指數']*100:.1f}%", latest['掏空預警'], delta_color="inverse")
                m3.metric("洗錢分值", int(latest['洗錢分值']))
                m4.metric("最終鑑定結論", latest['綜合鑑定意見'])

                st.divider()
                t1, t2 = st.columns(2)
                with t1:
                    st.write("營收趨勢圖 (含預測)")
                    fig1, ax1 = plt.subplots()
                    ax1.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='Actual')
                    st.pyplot(fig1)
                with t2:
                    st.write("資產偏移偵測 (掏空觀察區)")
                    fig2, ax2 = plt.subplots()
                    ax2.bar(sub['年度'].astype(str), sub['其他應收款'], label='其他應收')
                    ax2.bar(sub['年度'].astype(str), sub['預付款項'], bottom=sub['其他應收款'], label='預付項')
                    ax2.legend()
                    st.pyplot(fig2)
                
                st.dataframe(sub)

            elif view_mode == "多公司競爭力與風險PK":
                st.subheader("跨公司分析矩陣")
                pk_c1, pk_c2 = st.columns(2)
                latest_all = df_final.groupby('公司名稱').last()
                with pk_c1:
                    st.write("舞弊風險評比")
                    fig3, ax3 = plt.subplots()
                    ax3.bar(latest_all.index, latest_all['M分數'])
                    ax3.axhline(y=-1.78, color='red', linestyle='--')
                    st.pyplot(fig3)
                with pk_c2:
                    st.write("洗錢風險分布")
                    fig4, ax4 = plt.subplots()
                    ax4.scatter(latest_all['營收'], latest_all['洗錢分值'], s=100)
                    st.pyplot(fig4)

            if st.button("下載旗艦鑑定報告"):
                doc = Document()
                doc.add_heading("玄武會計師事務所 鑑定報告", 0)
                doc.save("report.docx")
                with open("report.docx", "rb") as f:
                    st.download_button("點此下載", f, "Report.docx")
        else:
            st.error("錯誤：上傳的資料中找不到『年度』資訊。請確保 Excel 標題正確或 PDF 格式清晰。")
    else:
        st.error("未能成功讀取任何檔案內容。")
