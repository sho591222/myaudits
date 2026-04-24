import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
import io
import matplotlib.font_manager as fm
import os
import requests
from datetime import datetime

# --- 1. 解決圖表中文亂碼 (下載思源黑體) ---
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

# --- 2. 頁面設定 ---
st.set_page_config(layout="wide", page_title="專業鑑識會計鑑定系統")
st.title("專業鑑識會計：一站式深度鑑定報告系統")

# --- 3. 側邊欄：模式切換與基本資訊 ---
with st.sidebar:
    st.header("系統模式切換")
    user_mode = st.radio("請選擇操作身分", ["一般公司模式", "會計師專業模式"])
    
    st.divider()
    st.header("基本資訊填寫")
    co_name = st.text_input("受調查公司/對象名稱", "XX股份有限公司")
    
    if user_mode == "會計師專業模式":
        auditor_name = st.text_input("主辦會計師簽署", "陳會計師 (CPA)")
        firm_name = st.text_input("所屬會計師事務所", "誠信聯合會計師事務所")
        audit_advice = st.text_area("查核建議 (手動填寫)", "請輸入針對此案的初步查核建議...")
    
    st.divider()
    files = st.file_uploader("批次上傳年度財報 PDF", type=["pdf"], accept_multiple_files=True)

# --- 4. 數據模擬與鑑定邏輯 (包含四大報表數據) ---
def advanced_forensic_engine(i):
    # 模擬數據：營收、應收、現金、負債
    rev = 5000 + (i * 300)
    rec = 200 + (i * 1800)
    cash = 2000 - (i * 400)
    liab = 3000 + (i * 600)
    
    # 舞弊/掏空風險計算
    m_score = -2.5 + (i * 0.5)
    z_score = 4.0 - (i * 1.2)
    
    analysis = ""
    if rec > rev * 0.4:
        analysis = "【警訊】應收帳款成長異常，可能存在虛增營收或資金掏空之情事。"
    elif cash < 500:
        analysis = "【警訊】現金流極度匱乏，營運持續性存在重大疑慮。"
    else:
        analysis = "財務數據尚在監控範圍，需持續追蹤關係人交易項目。"
        
    return rev, rec, cash, liab, m_score, z_score, analysis

# --- 5. 主內容顯示區 ---
if files:
    all_data = []
    sorted_files = sorted(files, key=lambda x: x.name)
    
    for i, f in enumerate(sorted_files):
        rev, rec, cash, liab, m, z, msg = advanced_forensic_engine(i)
        all_data.append({
            "年度": f.name.replace(".pdf", ""),
            "營業收入": rev, "應收帳款": rec, "現金預算": cash, "流動負債": liab,
            "M分數": m, "Z分數": z, "詳細敘述": msg
        })
    
    df = pd.DataFrame(all_data)

    # 圖表呈現 (各個項目的幅度圖表)
    st.subheader("一、 財務趨勢與幅度分析圖表")
    col1, col2 = st.columns(2)
    
    with col1:
        fig1, ax1 = plt.subplots()
        ax1.plot(df["年度"], df["營業收入"], label="營業收入", marker="o", linewidth=2)
        ax1.plot(df["年度"], df["應收帳款"], label="應收帳款", marker="s", color="orange")
        ax1.set_title("營收與應收帳款異常對比", fontproperties=font_prop)
        ax1.legend(prop=font_prop)
        st.pyplot(fig1)
        
    with col2:
        fig2, ax2 = plt.subplots()
        ax2.bar(df["年度"], df["現金預算"], label="現金水位", color="skyblue")
        ax2.set_title("現金流與償債能力分析", fontproperties=font_prop)
        ax2.legend(prop=font_prop)
        st.pyplot(fig2)

    # 詳細敘述版與專家分析
    st.subheader("二、 財報異常分析與詳細敘述")
    for idx, row in df.iterrows():
        with st.expander(f"年度 {row['年度']} - 鑑定分析報告詳細敘述"):
            st.write(f"**分析結果：** {row['詳細敘述']}")
            if user_mode == "會計師專業模式":
                st.info(f"**專業預警：** M分數為 {round(row['M分數'], 2)} (警戒線 -1.78)，顯示舞弊風險隨年份遞增。")

    # --- 6. 產生 DOC 文件 (一鍵產生分析報告) ---
    if st.sidebar.button("產生完整鑑定 DOC 文件"):
        doc = Document()
        doc.add_heading(f"{co_name} 鑑定分析報告", 0)
        
        # 基本資訊
        if user_mode == "會計師專業模式":
            doc.add_paragraph(f"會計師事務所：{firm_name}")
            doc.add_paragraph(f"主辦會計師：{auditor_name}")
            doc.add_heading("查核建議", level=1)
            doc.add_paragraph(audit_advice)
        
        doc.add_heading("各年度詳細分析敘述", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"年度：{r['年度']}", level=2)
            doc.add_paragraph(f"分析結論：{r['詳細敘述']}")
            doc.add_paragraph(f"關鍵數據：營收 {r['營業收入']} / 應收 {r['應收帳款']} / M分數 {round(r['M分數'], 2)}")

        doc_buf = io.BytesIO()
        doc.save(doc_buf)
        doc_buf.seek(0)
        st.sidebar.download_button("下載產出的鑑定 DOC", doc_buf, f"{co_name}_鑑定分析報告.docx")

else:
    st.info("請完成模式選擇並上傳財報 PDF 以啟動鑑定分析。")
