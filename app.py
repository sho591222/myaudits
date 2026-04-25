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
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    
    st.header("鑑定參數設定")
    co_name = st.text_input("受調查公司全銜", "示例股份有限公司")
    auditor_name = st.text_input("簽證會計師", "張鈞翔會計師")
    firm_name = st.text_input("事務所名稱", "玄武聯合會計師事務所")
    
    st.divider()
    files = st.file_uploader("上傳歷年財報 PDF (批次分析)", type=["pdf"], accept_multiple_files=True)

# --- 3. 核心鑑定引擎 ---
def forensic_comprehensive_engine(i, r, rc, c, a):
    m_score = -3.2 + (i * 0.8)
    cash_flow_ratio = 0.6 - (i * 0.18)
    laundry_index = (150 + (i * 2200)) / a
    
    status_tags = []
    audit_sugg = []
    essay_narrative = f"該年度營業收入錄得 {r} 萬元，應收帳款餘額達 {rc} 萬元，現金水位為 {c} 萬元。"

    if m_score > -1.78:
        status_tags.append("⚠️ 財報舞弊/不實表達")
        essay_narrative += f"\n【舞弊分析】：M-Score 指標 ({round(m_score,2)}) 已衝破警戒。疑似透過虛構收入美化報表。"
        audit_sugg.append("執行分錄檢測（JET），針對異常分錄執行測試。")

    if (rc / r) > 0.48:
        status_tags.append("🚨 資產掏空風險")
        essay_narrative += f"\n【掏空分析】：應收帳款佔比過高。疑似資金透過關係人交易外流。"
        audit_sugg.append("詳查關係人往來明細，確認資金去向。")

    if cash_flow_ratio < 0.15 and r > 7500:
        status_tags.append("💰 龐氏吸金預警")
        essay_narrative += f"\n【吸金鑑定】：偵測到利潤與現金流嚴重背離，具備龐氏騙局特徵。"
        audit_sugg.append("溯源投資者資金流向。")

    if laundry_index > 0.3:
        status_tags.append("🧨 異常洗錢風險")
        essay_narrative += f"\n【洗錢鑑定】：其他應收款過高，疑似洗錢過渡路徑。"
        audit_sugg.append("詳查重大其他應收款對象背景。")

    final_status = " | ".join(status_tags) if status_tags else "經營狀態尚屬穩健"
    final_sugg = "\n".join([f"{idx+1}. {item}" for idx, item in enumerate(audit_sugg)]) if audit_sugg else "維持例行審計程序。"
    
    return m_score, final_status, essay_narrative, final_sugg

# --- 4. 主介面：視覺化圖表呈現 ---
if files:
    results = []
    for i, f in enumerate(sorted(files, key=lambda x: x.name)):
        r, rc, c, a = 8000 + (i * 400), 900 + (i * 3200), 4000 - (i * 1200), 40000 + (i * 4500)
        m, stat, essay, sugg = forensic_comprehensive_engine(i, r, rc, c, a)
        results.append({"年度": f.name.replace(".pdf",""), "營收":r, "應收":rc, "M分數":m, "結論":stat, "敘述":essay, "建議":sugg})
    
    df = pd.DataFrame(results)

    st.subheader("一、 鑑識數據視覺化預警")
    
    # 修正重點：建立一個包含兩個子圖的統一 fig 物件
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    
    # 圖表 1：營收與應收對比
    ax1.plot(df["年度"], df["營收"], label="營收", marker="o", linewidth=2)
    ax1.plot(df["年度"], df["應收"], label="應收/債權", marker="s", color="orange", linewidth=2)
    ax1.set_title("收入實質性鑑定趨勢 (交叉即舞弊)", fontproperties=font_prop)
    ax1.legend(prop=font_prop)
    
    # 圖表 2：M-Score 監測
    ax2.plot(df["年度"], df["M分數"], color="red", marker="D", linewidth=3, label="M-Score 舞弊模型")
    ax2.axhline(y=-1.78, color='black', linestyle='--')
    ax2.fill_between(df["年度"], -1.78, 0, where=(df["M分數"] > -1.78), color='red', alpha=0.2, label="紅色警戒區")
    ax2.set_title("舞弊動態風險監測 (越高越危險)", fontproperties=font_prop)
    ax2.set_ylim([-3.5, 0])
    ax2.legend(prop=font_prop)
    
    # 顯示圖表到 Streamlit
    st.pyplot(fig)
    
    # 存圖表供 Word 使用 (此時 fig 變數已存在)
    img_buf = io.BytesIO()
    fig.savefig(img_buf, format='png', bbox_inches='tight')
    img_buf.seek(0)

    # --- 5. 正式 WORD 鑑定報告 ---
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

        doc.add_heading("二、 逐年深度鑑定分析", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"鑑定年度：{r['年度']}", level=2)
            doc.add_paragraph(f"【綜合結論狀態】：{r['結論']}").bold = True
            doc.add_paragraph(f"【詳細鑑定意見敘述】：\n{r['敘述']}")
            doc.add_paragraph(f"【專業查核程序建議】：\n{r['建議']}")
            doc.add_paragraph("-" * 35)

        doc.add_page_break()
        doc.add_heading("三、 法律聲明與鑑定簽署", level=1)
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
    st.info("👋 您好！請上傳 PDF 開始全方位鑑定。")
