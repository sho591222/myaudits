import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import matplotlib.font_manager as fm
import os
import requests
from datetime import datetime

# --- 1. 解決圖表中文亂碼 ---
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

# --- 2. 頁面配置與側邊欄 Logo ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計系統")

with st.sidebar:
    # 網頁 Logo 顯示
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    else:
        st.info("💡 請將 logo.png 上傳至 GitHub 儲存庫以顯示 Logo")
    
    st.header("系統模式")
    user_mode = st.radio("切換身分", ["一般公司模式", "會計師專業模式"])
    
    st.divider()
    co_name = st.text_input("受調查公司名稱", "示例企業股份有限公司")
    
    if user_mode == "會計師專業模式":
        auditor_name = st.text_input("主辦會計師", "陳會計師 (CPA)")
        firm_name = st.text_input("事務所全銜", "誠信聯合會計師事務所")
    
    st.divider()
    files = st.file_uploader("批次上傳年度財報 PDF", type=["pdf"], accept_multiple_files=True)

# --- 3. 自動化鑑定引擎與敘述生成 ---
def forensic_engine(i, r, rc, c, a):
    m_score = -2.7 + (i * 0.58)
    z_score = 4.4 - (i * 1.35)
    
    # 四大報表詳細自動敘述
    narrative = f"經查核，該年度營業收入錄得 {r} 萬元，應收帳款餘額達 {rc} 萬元。資產負債表中現金水位為 {c} 萬元。"
    
    if m_score > -1.78:
        status = "高度財報不實風險"
        sugg = "【查核建議】：M分數已跨越舞弊警戒線，顯示營收與應收帳款之關聯性存在重大異常。建議針對前十大客戶執行實地盤點，並對收入確認時點進行穿透式審查。"
    elif rc > r * 0.45:
        status = "資金掏空預警"
        sugg = "【查核建議】：應收帳款佔營收比例過高，資產流動性有枯竭風險。應詳查關係人往來明細，確認是否存在無實質交易基礎之資金撥貸。"
    elif z_score < 1.8:
        status = "經營假設疑慮 (Going Concern)"
        sugg = "【查核建議】：Z分數跌入破產區間，償債能力嚴重不足。會計師應評估管理層之改善計畫，並考慮在查核報告中加入強調事項段。"
    else:
        status = "營運狀態穩定"
        sugg = "【查核建議】：目前各項核心指標尚屬穩健。建議維持現行內部控制制度，並定期追蹤大額應收帳款回收狀況。"
        
    return m_score, z_score, narrative, status, sugg

# --- 4. 主介面顯示 ---
if files:
    results = []
    for i, f in enumerate(sorted(files, key=lambda x: x.name)):
        # 模擬從四大報表擷取數據
        r, rc, c, a = 5500 + (i * 600), 350 + (i * 2400), 3200 - (i * 700), 18000 + (i * 1500)
        m, z, nar, stat, sugg = forensic_engine(i, r, rc, c, a)
        results.append({
            "年度": f.name.replace(".pdf", ""),
            "營收": r, "應收": rc, "現金": c, "資產": a,
            "M分數": m, "Z分數": z, "詳細敘述": nar, "結論": stat, "查核建議": sugg
        })
    
    df = pd.DataFrame(results)

    # 圖表區
    st.subheader("一、 財報數據趨勢與風險幅度分析")
    col1, col2 = st.columns(2)
    with col1:
        fig1, ax1 = plt.subplots()
        ax1.plot(df["年度"], df["營收"], label="營收幅度", marker="o")
        ax1.plot(df["年度"], df["應收"], label="應收幅度", marker="x", color="red")
        ax1.set_title("收入實質性鑑定趨勢", fontproperties=font_prop)
        ax1.legend(prop=font_prop)
        st.pyplot(fig1)
    with col2:
        fig2, ax2 = plt.subplots()
        ax2.bar(df["年度"], df["M分數"], color="orange", label="M-Score (舞弊預警)")
        ax2.axhline(y=-1.78, color='black', linestyle='--', label="警戒線")
        ax2.set_title("舞弊動態風險模型", fontproperties=font_prop)
        ax2.legend(prop=font_prop)
        st.pyplot(fig2)

    # --- 5. WORD 報告生成功能 (含 Logo) ---
    if st.sidebar.button("產生完整查核鑑定報告 (Word)"):
        doc = Document()
        
        # Word 報告頂部插入 Logo (置中)
        if os.path.exists("logo.png"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture("logo.png", width=Inches(2.5))
        
        title = doc.add_heading(f"{co_name} 鑑識會計鑑定報告書", 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if user_mode == "會計師專業模式":
            doc.add_paragraph(f"事務所全銜：{firm_name}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
            doc.add_paragraph(f"主辦會計師：{auditor_name}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
            doc.add_paragraph(f"報告生成日：{datetime.now().strftime('%Y/%m/%d')}").alignment = WD_ALIGN_PARAGRAPH.RIGHT

        doc.add_heading("一、 四大報表詳細敘述分析", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"【年度 {r['年度']} 鑑定細節】", level=2)
            doc.add_paragraph(f"● 數據分析：{r['詳細敘述']}")
            doc.add_paragraph(f"● 鑑定結論：該年度判定為「{r['結論']}」。")
            doc.add_paragraph(f"● 自動查核建議：{r['查核建議']}")
            doc.add_paragraph("-" * 30)

        # 數據表插入
        doc.add_heading("二、 關鍵數據對照表", level=1)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text, hdr_cells[1].text = '年度', '營收'
        hdr_cells[2].text, hdr_cells[3].text = 'M分數', '鑑定結論'
        for _, r in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text, row_cells[1].text = str(r['年度']), str(r['營收'])
            row_cells[2].text, row_cells[3].text = str(round(r['M分數'], 2)), r['結論']

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.sidebar.download_button("📩 下載 Word 分析報告", buf, f"{co_name}_鑑識報告.docx")
else:
    st.info("請上傳財報 PDF 檔案開始自動化鑑定分析。")
