import streamlit as st
import pandas as pd
import re
import pdfplumber
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
import io

# --- 系統配置 ---
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>【最終解析版】暴力掃描技術：解決數據為 0 與圖表空白問題</p>
    </div>
""", unsafe_allow_html=True)

# --- 核心解析：暴力掃描函數 ---
def force_extract_numbers(text, keywords):
    """
    暴力掃描技術：在整頁文字中尋找關鍵字，並提取其後方最接近的數字
    """
    for key in keywords:
        # 正則表達式：尋找關鍵字後方 20 個字元內的數字 (處理千分位與括號)
        pattern = rf"{key}.{{0,20}}?([\d,]{{2,}}|\([\d,]{{2,}}\))"
        match = re.search(pattern, text)
        if match:
            val_str = match.group(1).replace(',', '').replace('(', '-').replace(')', '')
            try:
                return float(val_str)
            except:
                continue
    return 0.0

def parse_pdf_ultimate(file):
    res = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, "應收": 0.0, 
           "存貨": 0.0, "現金": 0.0, "其他應收": 0.0, "預付": 0.0, "淨利": 0.0}
    
    try:
        with pdfplumber.open(file) as pdf:
            # 掃描前 15 頁，確保抓到資產負債表與損益表
            all_text = ""
            for page in pdf.pages[:15]:
                all_text += (page.extract_text() or "")
            
            # 1. 抓年度 (民國或西元)
            y_match = re.search(r"(\d{3,4})\s*年度", all_text)
            if y_match:
                y = int(y_match.group(1))
                res["年度"] = y + 1911 if y < 1000 else y
            
            # 2. 暴力抓取各項數值
            res["營收"] = force_extract_numbers(all_text, ["營業收入", "營收合計"])
            res["應收"] = force_extract_numbers(all_text, ["應收帳款淨額", "應收帳款", "應收帳款－淨額"])
            res["存貨"] = force_extract_numbers(all_text, ["存貨", "存貨淨額"])
            res["現金"] = force_extract_numbers(all_text, ["現金及約當現金", "現金及流動資產"])
            res["其他應收"] = force_extract_numbers(all_text, ["其他應收款", "其他應收"])
            res["預付"] = force_extract_numbers(all_text, ["預付款項", "預付費用"])
            res["淨利"] = force_extract_numbers(all_text, ["本期淨利", "本期損益", "淨利（損）"])
            
    except Exception as e:
        st.error(f"解析失敗: {e}")
        
    return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()

# --- 側邊欄 ---
with st.sidebar:
    st.header("⚙️ 鑑識中心")
    uploaded_files = st.file_uploader("上傳財報 PDF", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽證會計師", "會計師")

# --- 主程式 ---
if uploaded_files:
    dfs = []
    for f in uploaded_files:
        df_single = parse_pdf_ultimate(f)
        if not df_single.empty:
            dfs.append(df_single)
    
    if dfs:
        df_pool = pd.concat(dfs, ignore_index=True)
        # 計算鑑定指標
        df_pool['M分數'] = df_pool.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
        df_pool['掏空指數'] = df_pool.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
        
        target = st.selectbox("選擇公司", df_pool['公司名稱'].unique())
        sub = df_pool[df_pool['公司名稱'] == target].sort_values('年度')
        
        # 繪圖
        st.subheader(f"歷年趨勢鑑識看板：{target}")
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='營業收入')
        ax.plot(sub['年度'].astype(str), sub['應收'], marker='x', label='應收帳款')
        ax.bar(sub['年度'].astype(str), sub['其他應收'], alpha=0.3, label='其他應收 (掏空指標)', color='red')
        ax.legend()
        st.pyplot(fig)
        
        st.dataframe(sub)
        
        # Word 報告產出
        if st.button("🚀 生成深沈敘述鑑定報告"):
            doc = Document()
            doc.add_heading(f"玄武鑑識鑑定報告 - {target}", 0)
            doc.add_heading("一、 盈餘操縱與舞弊分析", level=1)
            latest = sub.iloc[-1]
            doc.add_paragraph(f"本報告由 {auditor} 簽署。經鑑定 M 分數為 {latest['M分數']}，顯示該公司...")
            
            # 插入圖表
            img_buf = io.BytesIO()
            fig.savefig(img_buf, format='png', dpi=200)
            img_buf.seek(0)
            doc.add_picture(img_buf, width=Inches(5.5))
            
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 下載報告", buf, f"Report_{target}.docx")
    else:
        st.warning("⚠️ 暴力掃描仍無法讀取數據。請檢查 PDF 是否為加密文件或純圖片檔。")
