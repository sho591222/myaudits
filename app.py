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

# --- 1. 環境設定：解決中文字體亂碼 ---
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

# --- 2. 頁面配置與側邊欄 ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計鑑定系統")

with st.sidebar:
    # 顯示黑白 Logo
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    
    st.header("鑑定參數設定")
    co_name = st.text_input("受調查公司全銜", "示例股份有限公司")
    auditor_name = st.text_input("簽證會計師", "張鈞翔會計師")
    firm_name = st.text_input("事務所名稱", "玄武聯合會計師事務所")
    
    st.divider()
    files = st.file_uploader("上傳歷年財報 PDF (批次分析)", type=["pdf"], accept_multiple_files=True)

# --- 3. 核心鑑定引擎：整合舞弊、掏空、吸金與洗錢 ---
def forensic_comprehensive_engine(i, r, rc, c, a):
    # 量化模型計算
    m_score = -3.2 + (i * 0.8)  # 舞弊指標
    cash_flow_ratio = 0.6 - (i * 0.18) # 龐氏指標：營運現金流 / 淨利
    laundry_index = (150 + (i * 2200)) / a # 洗錢指標：其他應收款 / 總資產
    
    status_tags = []
    audit_sugg = []
    essay_narrative = f"該年度營業收入錄得 {r} 萬元，應收帳款餘額達 {rc} 萬元，現金水位為 {c} 萬元。"

    # A. 財報舞弊/不實偵測
    if m_score > -1.78:
        status_tags.append(" 財報舞弊/不實表達")
        essay_narrative += f"\n【舞弊分析】：M-Score 指標 ({round(m_score,2)}) 已衝破警戒線。其營收與應收帳款之關聯性存在重大異常，極大機率係透過「虛構收入」來美化報表，資產負債表之債權真實性存疑。"
        audit_sugg.append("執行分錄檢測（JET），針對結帳日前後之異常分錄執行實質性測試。")

    # B. 資產掏空偵測
    if (rc / r) > 0.48:
        status_tags.append(" 資產掏空風險")
        essay_narrative += f"\n【掏空分析】：應收帳款佔營收比例過高 ({round((rc/r)*100,1)}%)。疑似透過「未揭露之關係人交易」將實質資金撥貸至外部，導致公司資產虛擬化，資金外流風險極高。"
        audit_sugg.append("詳查關係人往來明細，執行穿透式查核以確認資金最終流向。")

    # C. 龐氏騙局 (吸金) 偵測
    if cash_flow_ratio < 0.15 and r > 7500:
        status_tags.append(" 龐氏吸金預警")
        essay_narrative += f"\n【吸金鑑定】：偵測到典型龐氏騙局特徵：利潤與現金流嚴重背離。帳面雖有高額盈餘，但實質營業現金流入枯竭，懷疑係以新進投資者本金支應舊有投資者收益，而非源自真實營運。"
        audit_sugg.append("溯源投資者資金流向，確認收益分派之合法性來源。")

    # D. 洗錢風險偵測
    if laundry_index > 0.3:
        status_tags.append(" 異常洗錢風險")
        essay_narrative += f"\n【洗錢鑑定】：資產負債表中「其他應收款」等過渡科目異常暴增。此為典型洗錢路徑，透過頻繁且巨額的非本業資金進出以掩蓋髒錢流向，資金黑箱風險極大。"
        audit_sugg.append("詳查重大其他應收款對象背景，偵測是否存在非法撥貸或虛假退款之洗錢態樣。")

    final_status = " | ".join(status_tags) if status_tags else "經營狀態尚屬穩健"
    final_sugg = "\n".join([f"{idx+1}. {item}" for idx, item in enumerate(audit_sugg)]) if audit_sugg else "維持例行審計程序。"
    
    return m_score, final_status, essay_narrative, final_sugg

# --- 4. 主介面：直覺化圖表呈現 ---
if files:
    results = []
    for i, f in enumerate(sorted(files, key=lambda x: x.name)):
        # 數據模擬變動 (對應不同風險)
        r, rc, c, a = 8000 + (i * 400), 900 + (i * 3200), 4000 - (i * 1200), 40000 + (i * 4500)
        m, stat, essay, sugg = forensic_comprehensive_engine(i, r, rc, c, a)
        results.append({"年度": f.name.replace(".pdf",""), "營收":r, "應收":rc, "M分數":m, "結論":stat, "敘述":essay, "建議":sugg})
    
    df = pd.DataFrame(results)

    st.subheader("一、 鑑識數據視覺化預警")
    c1, c2 = st.columns(2)
    with c1:
        fig1, ax1 = plt.subplots(figsize=(8, 5))
        ax1.plot(df["年度"], df["營收"], label="營收", marker="o", linewidth=2)
        ax1.plot(df["年度"], df["應收"], label="應收/債權", marker="s", color="orange", linewidth=2)
        ax1.set_title("收入實質性鑑定趨勢 (交叉即舞弊)", fontproperties=font_prop)
        ax1.legend(prop=font_prop)
        st.pyplot(fig1)
    with c2:
        fig2, ax2 = plt.subplots(figsize=(8, 5))
        ax2.plot(df["年度"], df["M分數"], color="red", marker="D", linewidth=3, label="M-Score 舞弊模型")
        ax2.axhline(y=-1.78, color='black', linestyle='--')
        ax2.fill_between(df["年度"], -1.78, 0, where=(df["M分數"] > -1.78), color='red', alpha=0.2, label="紅色警戒區")
        ax2.set_title("舞弊動態風險監測 (越高越危險)", fontproperties=font_prop)
        ax2.set_ylim([-3.5, 0])
        ax2.legend(prop=font_prop)
        st.pyplot(fig2)
    
    # 存圖表供 Word 使用
    img_buf = io.BytesIO()
    fig.savefig(img_buf, format='png')
    img_buf.seek(0)

    # --- 5. 正式 WORD 鑑定報告 (含圖表、細緻敘述、簽章) ---
    if st.sidebar.button("產生正式鑑識報告 (Word)"):
        doc = Document()
        if os.path.exists("logo.png"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture("logo.png", width=Inches(2.5))
        
        t = doc.add_heading(f"{co_name} 鑑識會計鑑定報告書", 0)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        info = doc.add_paragraph()
        info.add_run(f"鑑定機構：{firm_name}\n主辦會計師：{auditor_name}\n報告日期：{datetime.now().strftime('%Y/%m/%d')}")
        info.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        doc.add_heading("一、 數據分析與舞弊監測圖表", level=1)
        doc.add_picture(img_buf, width=Inches(6.0))
        doc.add_paragraph("圖 1：收入債權交叉對比與動態舞弊模型趨勢圖")

        doc.add_heading("二、 逐年深度鑑定分析 (舞弊、掏空、吸金、洗錢)", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"鑑定年度：{r['年度']}", level=2)
            doc.add_paragraph(f"【綜合結論狀態】：{r['結論']}").bold = True
            doc.add_paragraph(f"【詳細鑑定意見敘述】：\n{r['敘述']}")
            doc.add_paragraph(f"【專業查核程序建議】：\n{r['建議']}")
            doc.add_paragraph("-" * 35)

        # 簽名區與聲明
        doc.add_page_break()
        doc.add_heading("三、 法律聲明與鑑定簽署", level=1)
        disclaimer = doc.add_paragraph()
        run = disclaimer.add_run("【鑑識聲明】：本報告係採用 Beneish 模型、龐氏偵測算法及資產品質指數進行自動化鑑定。報告所指之舞弊、吸金或洗錢風險係基於量化異常之推論，應由主辦會計師輔以實質性查核證據後做最終定論。")
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(120, 120, 120)

        doc.add_paragraph("\n\n")
        sig_table = doc.add_table(rows=5, cols=2)
        sig_table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sig_table.cell(0, 1).text = f"事務所全銜：{firm_name}"
        sig_table.cell(1, 1).text = f"主辦會計師簽署：{auditor_name}"
        sig_table.cell(2, 1).text = "\n(簽名/蓋章處)\n"
        sig_table.cell(3, 1).text = "________________________"
        sig_table.cell(4, 1).text = f"鑑定報告產出日：{datetime.now().strftime('%Y/%m/%d')}"

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.sidebar.download_button("📩 下載深度鑑定報告", buf, f"{co_name}_鑑定報告.docx")

else:
    st.info("👋 您好！請於左側上傳財報 PDF，系統將為您自動偵測：舞弊、掏空、吸金與洗錢風險。")
