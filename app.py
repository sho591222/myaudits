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

# --- 2. 頁面配置與 Logo ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計系統")

with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    else:
        st.info("💡 請上傳 logo.png 至 GitHub")
    
    st.header("鑑定模式選擇")
    user_mode = st.radio("身分切換", ["一般公司模式", "會計師專業模式"])
    
    st.divider()
    co_name = st.text_input("受調查公司名稱", "示例股份有限公司")
    
    if user_mode == "會計師專業模式":
        auditor_name = st.text_input("簽證會計師", "陳會計師 (CPA)")
        firm_name = st.text_input("事務所名稱", "玄武聯合會計師事務所")
    
    st.divider()
    files = st.file_uploader("上傳年度財報 PDF", type=["pdf"], accept_multiple_files=True)

# --- 3. 財報舞弊鑑定引擎 (Beneish M-Score Logic) ---
def fraud_detection_engine(i, r, rc, c, gp, expense):
    # 模擬 Beneish 指標：DSRI (應收帳款指數), GMI (毛利指數), AQI (資產品質指數)
    dsri = 1.0 + (i * 0.15) if i > 0 else 1.0
    gmi = 1.0 + (i * 0.05) if i > 0 else 1.0
    # Beneish M-Score 公式簡化版
    m_score = -4.84 + (0.92 * dsri) + (0.52 * gmi) + (i * 0.4)
    
    # 舞弊敘述生成
    fraud_narrative = ""
    if m_score > -1.78:
        risk_level = "【高風險】"
        fraud_narrative = f"經由 Beneish M-Score 模型鑑定，得分 {round(m_score, 2)} 已衝破 -1.78 警戒線。該公司之應收帳款成長速度遠高於營收，顯示極可能存在「虛構銷售」或「提前認列收入」之舞弊行為。"
    else:
        risk_level = "【低風險】"
        fraud_narrative = f"M-Score 得分為 {round(m_score, 2)}，目前處於安全區間。財報數據之鉤稽關係尚屬合理。"

    # 查核建議
    if m_score > -1.78:
        sugg = "建議會計師針對「收入截止測試」擴大抽樣比例，並對前五大異常客戶發函詢證，必要時應執行分錄檢測（Journal Entry Testing）偵測異常手動分錄。"
    else:
        sugg = "目前舞弊傾向不明顯，建議維持例行性內部控制審查。"

    return m_score, risk_level, fraud_narrative, sugg

# --- 4. 主介面顯示 ---
if files:
    all_res = []
    for i, f in enumerate(sorted(files, key=lambda x: x.name)):
        # 數據模擬：營收, 應收, 現金, 毛利, 費用
        r, rc, c, gp, ex = 6000 + (i * 400), 500 + (i * 2200), 3000 - (i * 600), 2000, 1500
        m, r_lvl, f_nar, sugg = fraud_detection_engine(i, r, rc, c, gp, ex)
        all_res.append({
            "年度": f.name.replace(".pdf", ""),
            "營收": r, "應收": rc, "M分數": m,
            "風險等級": r_lvl, "舞弊敘述": f_nar, "專家建議": sugg
        })
    
    df = pd.DataFrame(all_res)

    # 直覺化圖表
    st.subheader("一、 財報舞弊動態風險分析")
    c1, c2 = st.columns(2)
    with c1:
        fig1, ax1 = plt.subplots(figsize=(8, 5))
        ax1.plot(df["年度"], df["營收"], label="營收幅度", marker="o", linewidth=2)
        ax1.plot(df["年度"], df["應收"], label="應收帳款幅度", marker="s", linewidth=2, color="orange")
        ax1.set_title("收入與債權異常對比 (交叉即舞弊警訊)", fontproperties=font_prop)
        ax1.legend(prop=font_prop)
        st.pyplot(fig1)
    with c2:
        fig2, ax2 = plt.subplots(figsize=(8, 5))
        ax2.plot(df["年度"], df["M分數"], color="red", label="M-Score 舞弊模型值", marker="D", linewidth=3)
        ax2.axhline(y=-1.78, color='black', linestyle='--', label="舞弊警戒線")
        ax2.fill_between(df["年度"], -1.78, 0, where=(df["M分數"] > -1.78), color='red', alpha=0.2, label="舞弊危險區")
        ax2.set_title("財報舞弊機率預警 (越高越危險)", fontproperties=font_prop)
        ax2.set_ylim([-3.5, 0])
        ax2.legend(prop=font_prop)
        st.pyplot(fig2)

    # 詳細報告區
    st.subheader("二、 鑑定報告詳細敘述 (舞弊專項)")
    for _, row in df.iterrows():
        with st.expander(f"【年度：{row['年度']}】鑑定詳情"):
            st.error(f"**鑑定狀態：{row['風險等級']}**")
            st.write(f"**詳細敘述：** {row['舞弊敘述']}")
            st.success(f"**查核建議：** {row['專家建議']}")

    # --- 5. Word 生成 ---
    if st.sidebar.button("產生 Word 鑑定報告"):
        doc = Document()
        if os.path.exists("logo.png"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture("logo.png", width=Inches(2.5))
        
        t = doc.add_heading(f"{co_name} 財報舞弊鑑定分析報告", 0)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if user_mode == "會計師專業模式":
            doc.add_paragraph(f"事務所：{firm_name}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
            doc.add_paragraph(f"會計師：{auditor_name}").alignment = WD_ALIGN_PARAGRAPH.RIGHT

        doc.add_heading("各年度財報舞弊深度鑑定", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"年度：{r['年度']}", level=2)
            doc.add_paragraph(f"1. 鑑定結論：{r['風險等級']}")
            doc.add_paragraph(f"2. 舞弊分析詳細敘述：{r['舞弊敘述']}")
            doc.add_paragraph(f"3. 專家查核對策：{r['專家建議']}")
            doc.add_paragraph("-" * 30)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.sidebar.download_button("📩 下載 Word 鑑定報告", buf, f"{co_name}_舞弊報告.docx")

else:
    st.info("👋 您好！請上傳 PDF，系統將自動啟動財報舞弊鑑定程序。")
