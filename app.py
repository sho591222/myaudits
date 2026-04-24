import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
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

# --- 2. 頁面設定與模式切換 ---
st.set_page_config(layout="wide")
st.title("專業鑑識會計：自動化報告生成系統 (四大報表與查核建議)")

with st.sidebar:
    st.header("系統操作模式")
    user_mode = st.radio("選擇身分", ["一般公司 (財報異常分析)", "會計師事務所 (深度查核分析)"])
    
    st.divider()
    st.header("案件基礎資訊")
    co_name = st.text_input("受調查公司名稱", "示例股份有限公司")
    
    if user_mode == "會計師事務所 (深度查核分析)":
        auditor_name = st.text_input("主辦會計師", "陳會計師")
        firm_name = st.text_input("事務所名稱", "專業聯合會計師事務所")
    
    st.divider()
    files = st.file_uploader("批次上傳年度財報 PDF", type=["pdf"], accept_multiple_files=True)

# --- 3. 核心引擎：數據分析與建議生成 ---
def audit_engine(i, rev, rec, cash, assets):
    m_score = -2.6 + (i * 0.52)
    z_score = 4.5 - (i * 1.3)
    
    # 計算增幅
    growth_rate = ((rev - 5000) / 5000) * 100 if i > 0 else 0
    
    # 自動敘述邏輯
    desc = f"本年度營業收入為 {rev}，應收帳款佔比達 {round((rec/rev)*100, 1)}%。"
    if m_score > -1.78:
        risk_status = "高度舞弊風險"
        suggestion = "查核發現營收增長與應收帳款變動極不匹配。建議執行實質性測試，包含抽查原始憑證及發函詢證，以確認是否存在虛增收入。"
    elif rec > rev * 0.4:
        risk_status = "疑似資金掏空"
        suggestion = "應收帳款餘額異常過高，應追查是否存在關係人轉讓資金之情事，並針對重大逾期帳款進行減損評估。"
    elif z_score < 1.8:
        risk_status = "經營持續性風險"
        suggestion = "財務結構趨於惡化，建議管理層提出增資或債務重組計畫，查核人員應評估其經營假設之合理性。"
    else:
        risk_status = "尚屬正常"
        suggestion = "財務指標目前處於監控水位，建議維持例行性內控稽核。"
        
    return m_score, z_score, growth_rate, risk_status, suggestion

# --- 4. 主流程呈現 ---
if files:
    results = []
    sorted_files = sorted(files, key=lambda x: x.name)
    
    for i, f in enumerate(sorted_files):
        # 模擬數據：營收、應收、現金、總資產
        r, rc, c, a = 6000 + (i * 500), 400 + (i * 2100), 2800 - (i * 600), 15000 + (i * 1000)
        m, z, gr, status, sugg = audit_engine(i, r, rc, c, a)
        results.append({
            "年度": f.name.replace(".pdf", ""),
            "營收": r, "應收": rc, "現金": c, "資產": a,
            "M分數": m, "Z分數": z, "增幅": gr,
            "風險結論": status, "專家建議": sugg
        })
    
    df = pd.DataFrame(results)

    # 圖表呈現
    st.subheader("分析圖表 (幅度趨勢與風險預測)")
    c1, c2 = st.columns(2)
    with c1:
        fig1, ax1 = plt.subplots()
        ax1.plot(df["年度"], df["營收"], label="營收趨勢", marker="o")
        ax1.plot(df["年度"], df["應收"], label="應收趨勢", marker="x")
        ax1.set_title("四大報表關鍵項對比", fontproperties=font_prop)
        ax1.legend(prop=font_prop)
        st.pyplot(fig1)
    with c2:
        fig2, ax2 = plt.subplots()
        ax2.bar(df["年度"], df["M分數"], color="red", label="M-Score (舞弊指標)")
        ax2.axhline(y=-1.78, color='black', linestyle='--')
        ax2.set_title("舞弊預警與警戒線", fontproperties=font_prop)
        ax2.legend(prop=font_prop)
        st.pyplot(fig2)

    # --- 5. 自動產生 WORD 當敘述與報告 ---
    if st.sidebar.button("產生 Word 查核分析報告"):
        doc = Document()
        doc.add_heading(f"{co_name} 鑑識會計鑑定報告書", 0)
        
        if user_mode == "會計師事務所 (深度查核分析)":
            doc.add_heading("一、 查核機構與人員資訊", level=1)
            doc.add_paragraph(f"事務所名稱：{firm_name}")
            doc.add_paragraph(f"簽證會計師：{auditor_name}")
            doc.add_paragraph(f"報告生成日期：{datetime.now().strftime('%Y/%m/%d')}")

        doc.add_heading("二、 數據增幅彙總表", level=1)
        table = doc.add_table(rows=1, cols=4)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '年度'
        hdr_cells[1].text = '營收'
        hdr_cells[2].text = '營收增幅 %'
        hdr_cells[3].text = '風險結論'
        for _, r in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(r['年度'])
            row_cells[1].text = str(r['營收'])
            row_cells[2].text = f"{round(r['增幅'], 2)}%"
            row_cells[3].text = r['風險結論']

        doc.add_heading("三、 四大報表異常分析與查核建議", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"【年度 {r['年度']} 分析重點】", level=2)
            doc.add_paragraph(f"【詳細敘述】：該年度營收為 {r['營收']}，但應收帳款大幅增加至 {r['應收']}。M分數為 {round(r['M分數'], 2)}，顯示公司可能面臨{r['風險結論']}。")
            doc.add_paragraph(f"【查核建議】：{r['專家建議']}")
            doc.add_paragraph("-" * 20)

        # 導出文件
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.sidebar.download_button("下載生成的 Word 報告", buf, f"{co_name}_分析報告.docx")

else:
    st.info("請上傳財報 PDF 檔案，系統將自動產出圖表、詳細敘述與查核建議報告。")
