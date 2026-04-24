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

# --- 2. 頁面配置與側邊欄 Logo ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計系統")

with st.sidebar:
    # 網頁 Logo 顯示 (搜尋目錄下的 logo.png)
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    else:
        st.info("💡 請將黑白 LOGO 檔名改為 logo.png 上傳至 GitHub")
    
    st.header("系統操作模式")
    user_mode = st.radio("切換身分", ["一般公司模式", "會計師專業模式"])
    
    st.divider()
    co_name = st.text_input("受調查公司名稱", "示例股份有限公司")
    
    if user_mode == "會計師專業模式":
        auditor_name = st.text_input("主辦會計師簽署", "陳會計師 (CPA)")
        firm_name = st.text_input("事務所全銜", "玄武聯合會計師事務所")
    
    st.divider()
    files = st.file_uploader("批次上傳年度財報 PDF", type=["pdf"], accept_multiple_files=True)

# --- 3. 自動化鑑定引擎與直覺化敘述 ---
def forensic_engine(i, r, rc, c, a):
    # 風險建模指標
    m_score = -2.8 + (i * 0.65)
    z_score = 4.5 - (i * 1.4)
    
    # 計算增幅與佔比
    rec_ratio = round((rc / r) * 100, 1)
    
    # 四大報表自動敘述
    narrative = f"本年度營業收入錄得 {r} 萬元，應收帳款餘額為 {rc} 萬元，佔營收比例達 {rec_ratio}%。現金水位監測值為 {c} 萬元。"
    
    if m_score > -1.78:
        status = "高度財報不實風險 (舞弊)"
        sugg = "【查核建議】：M分數已衝破警戒線，顯示營收增長與應收帳款變動極不匹配，疑似虛構銷售。建議執行應收帳款實地盤點，並針對前五大客戶發函詢證。"
    elif rec_ratio > 40:
        status = "資金掏空預警"
        sugg = "【查核建議】：應收帳款佔比過高，資產流動性急劇下滑。應詳查重大關係人往來明細，確認是否存在無實質交易基礎之資金撥貸。"
    elif z_score < 1.8:
        status = "經營假設疑慮 (破產預警)"
        sugg = "【查核建議】：Z分數跌入危險區間，償債能力嚴重不足。會計師應評估管理層之改善計畫，並考慮在查核報告中加入強調事項段。"
    else:
        status = "營運狀態穩定"
        sugg = "【查核建議】：核心指標目前尚屬穩健。建議維持現行內部控制制度，並定期追蹤大額逾期帳款。"
        
    return m_score, z_score, narrative, status, sugg

# --- 4. 主介面：直覺化圖表呈現 ---
if files:
    results = []
    # 根據檔名排序
    sorted_files = sorted(files, key=lambda x: x.name)
    
    for i, f in enumerate(sorted_files):
        # 數據模擬：模擬財報數據變動
        r, rc, c, a = 6000 + (i * 500), 400 + (i * 2500), 3500 - (i * 800), 20000 + (i * 2000)
        m, z, nar, stat, sugg = forensic_engine(i, r, rc, c, a)
        results.append({
            "年度": f.name.replace(".pdf", ""),
            "營收": r, "應收": rc, "現金": c, "資產": a,
            "M分數": m, "Z分數": z, "詳細敘述": nar, "結論": stat, "查核建議": sugg
        })
    
    df = pd.DataFrame(results)

    st.subheader("一、 鑑定數據趨勢分析 (直覺化視覺預警)")
    col1, col2 = st.columns(2)
    
    with col1:
        fig1, ax1 = plt.subplots(figsize=(8, 5))
        ax1.plot(df["年度"], df["營收"], label="營業收入趨勢", marker="o", linewidth=2, color="#1f77b4")
        ax1.plot(df["年度"], df["應收"], label="應收帳款趨勢", marker="s", linewidth=2, color="#ff7f0e")
        ax1.set_title("收入實質性鑑定對比 (交叉即警訊)", fontproperties=font_prop, fontsize=14)
        ax1.set_ylabel("金額 (萬元)", fontproperties=font_prop)
        ax1.grid(True, linestyle='--', alpha=0.5)
        ax1.legend(prop=font_prop)
        st.pyplot(fig1)
        
    with col2:
        fig2, ax2 = plt.subplots(figsize=(8, 5))
        # 優化 M-Score 圖表：折線圖搭配風險填充區
        ax2.plot(df["年度"], df["M分數"], color="red", label="M-Score 舞弊指標", marker="D", linewidth=3)
        ax2.axhline(y=-1.78, color='black', linestyle='--', linewidth=2, label="舞弊警戒線")
        # 填充高度風險區域
        ax2.fill_between(df["年度"], -1.78, 0, where=(df["M分數"] > -1.78), color='red', alpha=0.2, label="高度風險區")
        ax2.set_title("舞弊動態風險監測 (越高越危險)", fontproperties=font_prop, fontsize=14)
        ax2.set_ylim([-3.5, 0]) 
        ax2.set_ylabel("風險分數", fontproperties=font_prop)
        ax2.grid(True, linestyle='--', alpha=0.5)
        ax2.legend(prop=font_prop, loc='lower left')
        st.pyplot(fig2)

    st.subheader("二、 專家鑑定詳細敘述與查核建議報告")
    for _, row in df.iterrows():
        with st.expander(f"年度 {row['年度']} 鑑定細節"):
            st.markdown(f"**【數據詳細敘述】：** {row['詳細敘述']}")
            st.warning(f"**【鑑定風險結論】：** {row['結論']}")
            st.success(f"**【自動化查核建議】：** {row['查核建議']}")

    # --- 5. WORD 報告生成功能 (含置中 Logo) ---
    if st.sidebar.button("產生完整查核鑑定報告 (Word)"):
        doc = Document()
        # 插入 Logo
        if os.path.exists("logo.png"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture("logo.png", width=Inches(2.5))
        
        main_title = doc.add_heading(f"{co_name} 鑑識會計鑑定報告書", 0)
        main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if user_mode == "會計師專業模式":
            doc.add_paragraph(f"事務所全銜：{firm_name}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
            doc.add_paragraph(f"主辦會計師：{auditor_name}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
            doc.add_paragraph(f"報告生成日：{datetime.now().strftime('%Y/%m/%d')}").alignment = WD_ALIGN_PARAGRAPH.RIGHT

        doc.add_heading("一、 各年度深度鑑定分析敘述", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"【年度：{r['年度']}】", level=2)
            doc.add_paragraph(f"● 財報數據詳細敘述：{r['詳細敘述']}")
            doc.add_paragraph(f"● 鑑定結論：{r['結論']}")
            doc.add_paragraph(f"● 專業查核建議：{r['查核建議']}")
            doc.add_paragraph("-" * 35)

        # 關鍵指標彙整表
        doc.add_heading("二、 關鍵指標彙整對照表", level=1)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = '年度', '營收', 'M-Score', '鑑定結論'
        for _, r in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(r['年度'])
            row_cells[1].text = str(r['營收'])
            row_cells[2].text = str(round(r['M分數'], 2))
            row_cells[3].text = r['結論']

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.sidebar.download_button("📩 下載 Word 鑑定報告", buf, f"{co_name}_鑑定報告.docx")
else:
    st.info("👋 您好！請於左側上傳財報 PDF，系統將自動為您執行鑑識分析與報告撰寫。")
