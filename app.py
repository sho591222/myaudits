import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import pdfplumber

# --- 1. 系統環境設定 ---
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

if 'current_fig' not in st.session_state:
    st.session_state['current_fig'] = None

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財報鑑識旗艦系統：自動化深沈敘述、圖表嵌入與多維度比對</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. 暴力解析引擎：解決數據為 0 的問題 ---
def clean_num(text):
    if not text: return 0.0
    # 移除千分位、空格，處理括號負數
    s = str(text).strip().replace(',', '').replace('$', '')
    if '(' in s and ')' in s:
        s = '-' + s.replace('(', '').replace(')', '')
    match = re.search(r'[-+]?\d*\.\d+|\d+', s)
    return float(match.group()) if match else 0.0

def parse_financial_data(file):
    res = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, 
           "應收": 0.0, "存貨": 0.0, "現金": 0.0, "其他應收": 0.0, "預付": 0.0, "淨利": 0.0}
    try:
        with pdfplumber.open(file) as pdf:
            full_text = ""
            for page in pdf.pages[:12]:
                text = page.extract_text() or ""
                full_text += text
                
                # 抓取年度
                if res["年度"] == 0:
                    y_match = re.search(r"(\d{3,4})\s*年度", text)
                    if y_match:
                        y = int(y_match.group(1))
                        res["年度"] = y + 1911 if y < 1000 else y

            # 關鍵字暴力對位
            def find_val(keywords, content):
                for k in keywords:
                    # 搜尋關鍵字後 30 個字元內的數字
                    m = re.search(rf"{k}.{{0,30}}?([\d,]{{2,}}|\([\d,]{{2,}}\))", content)
                    if m: return clean_num(m.group(1))
                return 0.0

            res["營收"] = find_val(["營業收入", "營收合計"], full_text)
            res["應收"] = find_val(["應收帳款淨額", "應收帳款"], full_text)
            res["存貨"] = find_val(["存貨", "存貨淨額"], full_text)
            res["現金"] = find_val(["現金及約當現金", "現金及流動資產"], full_text)
            res["其他應收"] = find_val(["其他應收款", "其他應收"], full_text)
            res["預付"] = find_val(["預付款項", "預付費用"], full_text)
            res["淨利"] = find_val(["本期淨利", "本期損益"], full_text)
            
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except:
        return pd.DataFrame()

# --- 3. 鑑定邏輯 ---
def forensic_analysis(df):
    df['M分數'] = df.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
    df['掏空指數'] = df.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
    df['風險判定'] = df['M分數'].apply(lambda x: "⚠️ 注意" if x > -1.78 else "✅ 正常")
    return df

# --- 4. 介面控制 ---
with st.sidebar:
    st.header("⚙️ 鑑定功能中心")
    mode = st.radio("功能模式", ["🔍 單一深度比較 (多年趨勢)", "⚔️ 多公司橫向 PK (同年度比較)"])
    st.divider()
    uploaded_files = st.file_uploader("上傳受查 PDF", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "會計師")

# --- 5. 主程式 ---
if uploaded_files:
    data_list = [parse_financial_data(f) for f in uploaded_files]
    df_pool = pd.concat([d for d in data_list if not d.empty], ignore_index=True)
    
    if not df_pool.empty:
        df_pool = forensic_analysis(df_pool)

        if mode == "🔍 單一深度比較 (多年趨勢)":
            target = st.selectbox("選擇受查公司", df_pool['公司名稱'].unique())
            sub = df_pool[df_pool['公司名稱'] == target].sort_values('年度')
            
            st.subheader(f"📊 {target}：歷年趨勢鑑識看板")
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='營業收入', linewidth=2)
            ax.plot(sub['年度'].astype(str), sub['應收'], marker='x', label='應收帳款', linestyle='--')
            ax.bar(sub['年度'].astype(str), sub['其他應收'], alpha=0.3, label='其他應收 (掏空區)', color='red')
            ax.legend()
            st.pyplot(fig)
            st.session_state['current_fig'] = fig
            st.dataframe(sub)

        else: # PK 模式
            target_year = st.selectbox("選擇 PK 年度", sorted(df_pool['年度'].unique(), reverse=True))
            pk_df = df_pool[df_pool['年度'] == target_year]
            
            st.subheader(f"⚔️ {target_year} 年度：多公司橫向比對")
            fig, ax = plt.subplots(figsize=(10, 4))
            x = np.arange(len(pk_df))
            ax.bar(x - 0.2, pk_df['M分數'], 0.4, label='M-Score (舞弊)', color='skyblue')
            ax.bar(x + 0.2, pk_df['掏空指數']*10, 0.4, label='掏空指標 x10', color='salmon')
            ax.set_xticks(x)
            ax.set_xticklabels(pk_df['公司名稱'])
            ax.axhline(y=-1.78, color='red', linestyle='--', label='風險閾值')
            ax.legend()
            st.pyplot(fig)
            st.session_state['current_fig'] = fig
            st.dataframe(pk_df)

        # --- 6. 生成 Word 報告 ---
        st.divider()
        if st.button("🚀 生成「深沈敘述版」詳細鑑定報告"):
            doc = Document()
            title = doc.add_heading('財務鑑識鑑定意見書', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 第一部分
            doc.add_heading('壹、 鑑定背景與數據概況', level=1)
            doc.add_paragraph(f"本報告由 {auditor} 會計師執行。針對受查對象進行多維度財務比對分析。")
            
            # 第二部分：深沈敘述邏輯
            doc.add_heading('貳、 盈餘品質與資產風險鑑定', level=1)
            row = sub.iloc[-1] if mode == "🔍 單一深度比較 (多年趨勢)" else pk_df.iloc[0]
            
            narrative = (f"【鑑定結果】受查單位之 M-Score 為 {row['M分數']}，判定為「{row['風險判定']}」。"
                         f"其掏空指數為 {row['掏空指數']}。若數據偏離產業常態，建議查核人員應針對其「其他應收款」進行深度函證，"
                         "以排除非法資金挪用之可能性。")
            doc.add_paragraph(narrative)
            
            # 第三部分：嵌入圖表
            doc.add_heading('參、 鑑定圖表分析', level=1)
            if st.session_state['current_fig']:
                img_buf = io.BytesIO()
                st.session_state['current_fig'].savefig(img_buf, format='png', dpi=200)
                img_buf.seek(0)
                doc.add_picture(img_buf, width=Inches(5.5))
            
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 下載詳盡報告 (.docx)", buf, "Forensic_Report.docx")
    else:
        st.warning("⚠️ 掃描不到數據，請確認 PDF 文字可選取。")
