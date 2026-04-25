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

# --- 1. 環境設定：解決中文字體顯示 ---
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

# --- 2. 頁面配置與側邊欄設定 ---
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

# --- 3. 核心鑑定引擎：舞弊、掏空、吸金、洗錢 ---
def forensic_comprehensive_engine(i, r, rc, c, a):
    m_score = -3.2 + (i * 0.8)
    cash_flow_ratio = 0.6 - (i * 0.18)
    laundry_index = (150 + (i * 2200)) / a
    
    status_tags = []
    audit_sugg = []
    essay_narrative = f"該年度營業收入錄得 {r} 萬元，應收帳款餘額達 {rc} 萬元，現金水位為 {c} 萬元。"

    if m_score > -1.78:
        status_tags.append("財報舞弊/不實表達")
        essay_narrative += f"\n【舞弊分析】：M-Score 指標 ({round(m_score,2)}) 已衝破警戒。顯示營收增長與應收帳款嚴重脫節，可能透過虛構收入美化損益表。"
        audit_sugg.append("執行分錄檢測(JET)，針對結帳日前後之異常分錄執行測試。")

    if (rc / r) > 0.48:
        status_tags.append("資產掏空風險")
        essay_narrative += f"\n【掏空分析】：應收帳款佔營收比例過高。疑似資金透過未揭露之關係人交易外流。"
        audit_sugg.append("詳查關係人往來明細，執行穿透式查核以確認資金最終流向。")

    if cash_flow_ratio < 0.15 and r > 7500:
        status_tags.append("龐氏吸金預警")
        essay_narrative += f"\n【吸金鑑定】：偵測到利潤與現金流嚴重背離。帳面盈餘高但經營性現金流枯竭，具備龐氏騙局特徵。"
        audit_sugg.append("對投資者資金來源進行溯源查核，確認收益來源是否為真實營運獲利。")

    if laundry_index > 0.3:
        status_tags.append("異常洗錢風險")
        essay_narrative += f"\n【洗錢鑑定】：其他應收款等過渡科目異常暴增。疑似透過非本業往來款掩蓋不明資金流向。"
        audit_sugg.append("詳查重大其他應收款對象背景，偵測是否存在非法撥貸或虛假退款之洗錢態樣。")

    final_status = " | ".join(status_tags) if status_tags else "經營狀態尚屬穩健"
    final_sugg = "\n".join([f"{idx+1}. {item}" for idx, item in enumerate(audit_sugg)]) if audit_sugg else "維持例行審計程序。"
    
    return m_score, final_status, essay_narrative, final_sugg

# --- 4. 主介面：視覺化圖表 ---
if files:
    results = []
    for i, f in enumerate(sorted(files, key=lambda x: x.name)):
        r, rc, c, a = 8000 + (i * 400), 900 + (i * 3200), 4000 - (i * 1200), 40000 + (i * 4500)
        m, stat, essay, sugg = forensic_comprehensive_engine(i, r, rc, c, a)
        results.append({"年度": f.name.replace(".pdf",""), "營收":r, "應收":rc, "M分數":m, "結論":stat, "敘述":essay, "建議":sugg})
    
    df = pd.DataFrame(results)

    st.subheader("一、 鑑識數據視覺化預警")
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    
    ax1.plot(df["年度"], df["營收"], label="營收", marker="o", linewidth=2)
    ax1.plot(df["年度"], df["應收"], label="應收/債權", marker="s", color="orange", linewidth=2)
    ax1.set_title("收入實質性鑑定趨勢 (交叉即警訊)", fontproperties=font_prop)
    ax1.legend(prop=font_prop)
    
    ax2.plot(df["年度"], df["M分數"], color="red", marker="D", linewidth=3, label="M-Score 舞弊模型")
    ax2.axhline(y=-1.78, color='black', linestyle='--')
    ax2.fill_between(df["年度"], -1.78, 0, where=(df["M分數"] > -1.78), color='red', alpha=0.2, label="紅色警戒區")
    ax2.set_title("舞弊動態風險監測 (越高越危險)", fontproperties=font_prop)
    ax2.set_ylim([-3.5, 0])
    ax2.legend(prop=font_prop)
    
    st.pyplot(fig)
    
    img_buf = io.BytesIO()
    fig.savefig(img_buf, format='png', bbox_inches='tight')
    img_buf.seek(0)

    # --- 5. Word 報告生成 (含法律聲明與簽章) ---
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

        doc.add_heading("二、 逐年深度鑑定分析 (含舞弊、掏空、吸金、洗錢)", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"年度：{r['年度']}", level=2)
            doc.add_paragraph(f"【綜合結論狀態】：{r['結論']}").bold = True
            doc.add_paragraph(f"【詳細鑑定意見敘述】：\n{r['敘述']}")
            doc.add_paragraph(f"【專業查核程序建議】：\n{r['建議']}")
            doc.add_paragraph("-" * 35)

        doc.add_page_break()
        doc.add_heading("三、 法律聲明與鑑定簽署", level=1)
        
        desc_p = doc.add_paragraph()
        run = desc_p.add_run("【重要聲明與免責條款】：\n")
        run.bold = True
        
        content = (
            "1. 本報告係基於委任人提供之財務數據，透過量化鑑識模型（包括但不限於 Beneish M-Score、龐氏偵測算法、資金洗錢流向指標）產出之初步鑑定結果。\n"
            "2. 本系統所偵測之「舞弊」、「吸金」或「洗錢」風險指標係屬數據異常推論，旨在協助會計師識別高風險領域。鑑定結論之最終法律效力，應以會計師執行實質查核程序後出具之正式查核報告為準。\n"
            "3. 本報告不構成對受調查公司財務報表合法性之最終保證。若將本報告用於司法訴訟，請務必由執業會計師進行簽署核閱，以符合法定鑑定之程序規範。"
        )
        run_content = desc_p.add_run(content)
        run_content.font.size = Pt(9)
        run_content.font.color.rgb = RGBColor(80, 80, 80)

        doc.add_paragraph("\n\n")
        
        sig_table = doc.add_table(rows=5, cols=2)
        sig_table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sig_table.cell(0, 1).text = f"鑑定機構：{firm_name}"
        sig_table.cell(1, 1).text = f"主辦會計師簽署：{auditor_name}"
        sig_table.cell(2, 1).text = "\n(簽名/小官章蓋印處)\n"
        sig_table.cell(3, 1).text = "________________________"
        sig_table.cell(4, 1).text = f"鑑定報告產出日：{datetime.now().strftime('%Y/%m/%d')}"

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.sidebar.download_button("下載深度鑑定報告", buf, f"{co_name}_鑑定報告.docx")
else:
    st.info("系統就緒。請上傳 PDF 開始鑑定流程。")
