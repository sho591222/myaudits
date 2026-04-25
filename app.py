import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
import re
import pdfplumber
from datetime import datetime

# 1. 系統環境設定與基礎 UI
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財務鑑識旗艦平台：舞弊、掏空與洗錢防制整合分析</p>
    </div>
""", unsafe_allow_html=True)

# 2. 核心數據解析引擎 (支援多科目抓取)
def parse_financial_pdf(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), 
            "年度": None, "營收": 0, "應收帳款": 0, 
            "存貨": 0, "現金": 0, "負債總額": 0, 
            "其他應收款": 0, "預付款項": 0
        }
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:5]:
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

                if not res["年度"]:
                    text = page.extract_text()
                    years = re.findall(r"(\d{3,4})\s*年度", text)
                    if years:
                        y = int(years[0])
                        res["年度"] = y + 1911 if y < 1000 else y
        
        return pd.DataFrame([res]) if res["年度"] else pd.DataFrame()
    except:
        return pd.DataFrame()

# 3. 綜合鑑識計算模組 (舞弊+掏空+洗錢)
def run_integrated_forensics(df):
    # 確保資料型態為數值
    numeric_cols = ['營收', '應收帳款', '存貨', '現金', '負債總額', '其他應收款', '預付款項']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 建立分析結果欄位
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
            # A. 舞弊分析
            m = -3.2 + (0.15 * (rc/r)) + (0.1 * (inv/r))
            df.at[i, 'M分數'] = round(m, 2)
            df.at[i, '舞弊風險'] = "危險" if m > -1.78 else "正常"
            
            # B. 掏空分析 (搬錢比例)
            t_ratio = (other_rc + prepay) / r
            df.at[i, '掏空指數'] = round(t_ratio, 3)
            df.at[i, '掏空預警'] = "高風險" if t_ratio > 0.2 else "正常"
            
            # C. 洗錢分析 (資金異常堆積)
            ml_score = 0
            if cash > r * 0.8: ml_score += 50
            if debt > cash * 5: ml_score += 30
            df.at[i, '洗錢分值'] = ml_score
            
            # D. 結論
            if m > -1.78 or t_ratio > 0.2 or ml_score >= 50:
                df.at[i, '綜合鑑定意見'] = "建議深度查核"
                
    return df

# 4. 側邊欄控制
with st.sidebar:
    st.header("功能控制台")
    st.text("品牌口號: 玄武鑑定 真偽分明")
    view_mode = st.radio("請選擇分析視角", ["單一公司深度報告", "多公司競爭力與風險PK"])
    st.divider()
    auditor_sig = st.text_input("主辦會計師", "張鈞翔會計師")
    uploaded_files = st.file_uploader("上傳財報數據 (PDF/Excel)", type=["pdf", "xlsx"], accept_multiple_files=True)

# 5. 主程式邏輯
if uploaded_files:
    data_pool = []
    for f in uploaded_files:
        if f.name.endswith('.xlsx'):
            data_pool.append(pd.read_excel(f))
        else:
            data_pool.append(parse_financial_pdf(f))
    
    if data_pool:
        df_final = pd.concat(data_pool, ignore_index=True)
        df_final = df_final[df_final['年度'] > 0].drop_duplicates(subset=['公司名稱', '年度']).sort_values('年度')
        df_final = run_integrated_forensics(df_final)

        # 模式一：單一公司深度報告
        if view_mode == "單一公司深度報告":
            target = st.selectbox("選擇受查公司", df_final['公司名稱'].unique())
            sub = df_final[df_final['公司名稱'] == target]
            
            st.subheader(f"財務鑑定看板：{target}")
            
            # 指標卡片區
            latest = sub.iloc[-1]
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("舞弊指標 (M-Score)", latest['M分數'], latest['舞弊風險'], delta_color="inverse")
            m2.metric("掏空係數", f"{latest['掏空指數']*100:.1f}%", latest['掏空預警'], delta_color="inverse")
            m3.metric("洗錢風險分值", int(latest['洗錢分值']))
            m4.metric("最終鑑定結論", latest['綜合鑑定意見'])

            # 綜合分析圖表區
            st.divider()
            t1, t2 = st.columns(2)
            with t1:
                st.write("營收成長與預測趨勢")
                fig1, ax1 = plt.subplots()
                ax1.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='歷史營收', color='#268bd2')
                if len(sub) >= 2:
                    g = sub['營收'].pct_change().mean()
                    p = sub['營收'].iloc[-1] * (1 + g)
                    ax1.plot([str(sub['年度'].iloc[-1]), str(int(sub['年度'].iloc[-1])+1)], 
                             [sub['營收'].iloc[-1], p], '--', marker='s', label='AI預測', color='#cb4b16')
                ax1.legend()
                st.pyplot(fig1)
            
            with t2:
                st.write("資產偏移偵測 (掏空/洗錢觀察)")
                fig2, ax2 = plt.subplots()
                ax2.stackplot(sub['年度'].astype(str), sub['其他應收款'], sub['預付款項'], labels=['其他應收', '預付項'], alpha=0.5)
                ax2.set_title("非本業資產堆積觀察")
                ax2.legend(loc='upper left')
                st.pyplot(fig2)

            st.write("詳細鑑定數據表")
            st.dataframe(sub)

        # 模式二：多公司風險 PK
        else:
            st.subheader("跨公司犯罪風險與競爭力 PK")
            
            pk_col1, pk_col2 = st.columns(2)
            with pk_col1:
                st.write("各公司舞弊風險對比 (M-Score)")
                latest_all = df_final.groupby('公司名稱').last()
                fig3, ax3 = plt.subplots()
                ax3.bar(latest_all.index, latest_all['M分數'], color='#2aa198')
                ax3.axhline(y=-1.78, color='red', linestyle='--', label='舞弊警戒線')
                ax3.legend()
                st.pyplot(fig3)

            with pk_col2:
                st.write("跨公司洗錢風險矩陣")
                fig4, ax4 = plt.subplots()
                ax4.scatter(latest_all['營收'], latest_all['洗錢分值'], s=latest_all['現金']/1000, alpha=0.5)
                for i, txt in enumerate(latest_all.index):
                    ax4.annotate(txt, (latest_all['營收'].iloc[i], latest_all['洗錢分值'].iloc[i]))
                ax4.set_xlabel("營收規模")
                ax4.set_ylabel("洗錢可疑度")
                st.pyplot(fig4)

        # 報告匯出模組
        if st.button("產出整合式鑑定報告書"):
            doc = Document()
            doc.add_heading("玄武會計師事務所 旗艦鑑定報告", 0)
            doc.add_paragraph(f"主辦會計師：{auditor_sig}")
            for co in df_final['公司名稱'].unique():
                cur = df_final[df_final['公司名稱'] == co].iloc[-1]
                doc.add_heading(f"受查公司：{co}", level=1)
                doc.add_paragraph(f"鑑定結論：{cur['綜合鑑定意見']}")
                doc.add_paragraph(f"1. 舞弊偵測：M-Score 為 {cur['M分數']} ({cur['舞弊風險']})。")
                doc.add_paragraph(f"2. 掏空預警：掏空壓力指數為 {cur['掏空指數']*100:.1f}%。")
                doc.add_paragraph(f"3. 洗錢防制：洗錢可疑分值評定為 {int(cur['洗錢分值'])}。")
            
            buf = io.BytesIO()
            doc.save(buf)
            st.download_button("下載 Word 旗艦鑑定書", buf.getvalue(), "XuanWu_Flagship_Report.docx")
    else:
        st.error("未能成功提取數據。請確認 PDF 內容正確。")
else:
    st.info("系統準備就緒。請上傳受查數據以啟動 AI 鑑識。")
