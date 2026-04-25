import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import matplotlib.font_manager as fm
import os
import requests
from datetime import datetime

# --- 1. 環境設定：中文字體 ---
@st.cache_resource
def load_chinese_font():
    font_url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
    font_path = "NotoSansCJKtc-Regular.otf"
    if not os.path.exists(font_path):
        try:
            response = requests.get(font_url)
            with open(font_path, "wb") as f:
                f.write(response.content)
        except: return None
    return font_path

font_p = load_chinese_font()
def apply_font_logic(font_path):
    if font_path:
        custom_font = fm.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = custom_font.get_name()
        fm.fontManager.addfont(font_path)
        plt.rcParams['axes.unicode_minus'] = False
        return custom_font
    return None
font_prop = apply_font_logic(font_p)

# --- 2. 核心鑑識邏輯引擎 ---
def forensic_analysis_logic(row):
    # 讀取 Excel 欄位數據
    r = row.get('營收', 0)
    rc = row.get('應收帳款', 0)
    inv = row.get('存貨', 0)
    c = row.get('現金', 0)
    a = row.get('總資產', 0)
    ni = row.get('淨利', 0)
    ocf = row.get('營業現金流', 0)
    
    # 指標計算
    m_score = -3.2 + (0.1 * (rc/r if r!=0 else 0)) + (0.2 * (inv/r if r!=0 else 0))
    cash_ratio = ocf / ni if ni > 0 else 0
    
    tags = []
    if m_score > -1.78: tags.append("財報舞弊高風險")
    if (rc / r) > 0.45 if r!=0 else False: tags.append("資產掏空警訊")
    if cash_ratio < 0.15 and ni > 0: tags.append("龐氏吸金預警")
    if (rc + inv) / a > 0.4 if a!=0 else False: tags.append("異常洗錢風險")
    
    return pd.Series([m_score, " | ".join(tags) if tags else "未見明顯異常"])

# --- 3. 頁面配置 ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計鑑定系統 V2")

with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.header("數據與報告設定")
    auditor_name = st.text_input("主辦會計師", "張鈞翔會計師")
    firm_name = st.text_input("事務所名稱", "玄武聯合會計師事務所")
    
    st.divider()
    uploaded_file = st.file_uploader("上傳數據底稿 (Excel 或 CSV)", type=["xlsx", "csv"])

# --- 4. 數據處理與分析 ---
if uploaded_file:
    if uploaded_file.name.endswith('.xlsx'):
        df = pd.read_excel(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file)
    
    # 執行鑑定運算
    df[['M分數', '鑑定結論']] = df.apply(forensic_analysis_logic, axis=1)

    st.header("數據鑑定分析看板")
    
    # 圖表呈現
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    
    # 圖 1：規模與債權對比
    df.plot(kind='bar', x=df.columns[0], y=['營收', '應收帳款'], ax=ax1, color=['#1f77b4', '#ff7f0e'])
    ax1.set_title("營收與債權品質對比圖", fontproperties=font_prop)
    
    # 圖 2：M-Score 風險監測
    ax2.plot(df[df.columns[0]], df['M分數'], color='red', marker='D', linewidth=2)
    ax2.axhline(y=-1.78, color='black', linestyle='--')
    ax2.fill_between(df[df.columns[0]], -1.78, df['M分數'].max()+0.5, where=(df['M分數'] > -1.78), color='red', alpha=0.2)
    ax2.set_title("舞弊偵測模型警戒監控 (高於虛線為風險)", fontproperties=font_prop)
    
    st.pyplot(fig)

    # 數據表展示
    st.subheader("鑑定底稿清單")
    st.dataframe(df)

    # 法律聲明
    st.divider()
    st.caption("法律聲明：本鑑定報告係由自動化模型產出，偵測結果屬風險預警性質。最終法律結論應以會計師簽署之正式紙本報告為準。")

    # --- 報告生成按鈕區 ---
    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        # 下載 Excel 鑑定底稿
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='鑑定底稿')
        st.download_button("下載 Excel 鑑定底稿", output_excel.getvalue(), "鑑定底稿.xlsx")

    with col_dl2:
        # 下載 Word 鑑定意見書
        if st.button("準備 Word 報告資料"):
            doc = Document()
            if os.path.exists("logo.png"):
                doc.add_picture("logo.png", width=Inches(2.0))
            
            doc.add_heading("鑑識會計鑑定意見書", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"事務所：{firm_name}\n會計師：{auditor_name}\n報告日期：{datetime.now().strftime('%Y/%m/%d')}")
            
            doc.add_heading("一、 鑑定結論摘要", level=1)
            for _, row in df.iterrows():
                p = doc.add_paragraph()
                p.add_run(f"標的名稱：{row[df.columns[0]]}\n").bold = True
                p.add_run(f"鑑定結論：{row['鑑定結論']}\n")
                p.add_run(f"舞弊分數 (M-Score)：{round(row['M分數'], 2)}")

            doc.add_page_break()
            doc.add_heading("二、 法律聲明與簽章", level=1)
            doc.add_paragraph("本報告偵測之異常態樣屬量化推論，旨在識別查核重點。鑑定結論之最終效力應由主辦會計師輔以實質查核程序後定論。")
            
            doc.add_paragraph("\n\n會計師簽署：____________________")
            
            buf_word = io.BytesIO()
            doc.save(buf_word)
            st.download_button("點此下載 Word 鑑定意見書", buf_word.getvalue(), "鑑定意見書.docx")

else:
    st.info("請上傳包含「營收、應收帳款、存貨、總資產、淨利、營業現金流」欄位的 Excel 檔案開始分析。")
