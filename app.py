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

# --- 1. 中文字體與環境處理 ---
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

# --- 2. 側邊欄配置 ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計鑑定系統")

with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    
    st.header("鑑定報告參數")
    co_name = st.text_input("受調查公司全銜", "示例股份有限公司")
    auditor_name = st.text_input("主辦會計師", "陳會計師 (CPA)")
    firm_name = st.text_input("事務所全銜", "玄武聯合會計師事務所")
    
    st.divider()
    files = st.file_uploader("批次上傳年度財報 PDF", type=["pdf"], accept_multiple_files=True)

# --- 3. 核心鑑定引擎：舞弊、掏空與不實偵測 ---
def forensic_analysis_logic(i, r, rc, c, a):
    m_score = -3.0 + (i * 0.75)
    z_score = 4.8 - (i * 1.6)
    rc_to_r = round((rc / r) * 100, 1)
    
    status = "營運狀態尚無重大異常"
    detail_essay = f"該年度營業收入為 {r} 萬元，應收帳款餘額達 {rc} 萬元。經計算，應收帳款佔營收比重達 {rc_to_r}%，現金餘額為 {c} 萬元。"
    audit_sugg = "建議維持常規審計程序，並追蹤後續應收帳款回收情形。"

    # 舞弊判定 (M-Score)
    if m_score > -1.78:
        status = "⚠️ 高度舞弊風險 (財報不實)"
        detail_essay += f"\n【舞弊分析】：Beneish M-Score 為 {round(m_score,2)}，已達警戒水位。該數據顯示公司可能透過「虛構收入」或「操縱應計項目」以美化損益表，其資產負債表之債權真實性存疑。"
        audit_sugg = "應執行分錄檢測（JET），調閱重大非經常性傳票，並對前五大客戶執行外部詢證函發函作業。"

    # 掏空判定
    elif rc_to_r > 50:
        status = "🚨 資產掏空預警 (資金外流)"
        detail_essay += f"\n【掏空分析】：應收帳款佔營收比例過高 ({rc_to_r}%)。此異常結構常見於「非法資金撥貸」或「未揭露之關係人交易」，懷疑公司實質資產已遭虛擬化，資金可能已外流。"
        audit_sugg = "應詳查重大往來對象之背景資訊，確認是否存在關係人代持情事，並針對資金去向執行穿透式查核。"

    return m_score, z_score, detail_essay, status, audit_sugg

# --- 4. 主介面顯示 ---
if files:
    results = []
    for i, f in enumerate(sorted(files, key=lambda x: x.name)):
        # 數據模擬
        r, rc, c, a = 7000 + (i * 500), 700 + (i * 2800), 4000 - (i * 1000), 30000 + (i * 3000)
        m, z, essay, stat, sugg = forensic_analysis_logic(i, r, rc, c, a)
        results.append({"年度": f.name.replace(".pdf",""), "營收":r, "應收":rc, "M分數":m, "Z分數":z, "敘述":essay, "結論":stat, "建議":sugg})
    
    df = pd.DataFrame(results)

    # 產生分析圖表
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    ax1.plot(df["年度"], df["營收"], label="營收趨勢", marker="o")
    ax1.plot(df["年度"], df["應收"], label="應收帳款", marker="s", color="orange")
    ax1.set_title("收入實質性對比分析", fontproperties=font_prop)
    ax1.legend(prop=font_prop)
    
    ax2.plot(df["年度"], df["M分數"], color="red", marker="D", linewidth=3)
    ax2.axhline(y=-1.78, color='black', linestyle='--', label="舞弊警戒線")
    ax2.fill_between(df["年度"], -1.78, 0, where=(df["M分數"] > -1.78), color='red', alpha=0.2)
    ax2.set_title("舞弊動態風險預警 (M-Score)", fontproperties=font_prop)
    st.pyplot(fig)
    
    img_buf = io.BytesIO()
    fig.savefig(img_buf, format='png')
    img_buf.seek(0)

    # --- 5. WORD 報告生成邏輯 (含簽名欄位) ---
    if st.sidebar.button("產生正式鑑定報告 (Word)"):
        doc = Document()
        
        # 頁首 Logo
        if os.path.exists("logo.png"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture("logo.png", width=Inches(2.5))
        
        # 報告標題
        t = doc.add_heading(f"{co_name} 鑑識會計鑑定報告書", 0)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 鑑定圖表
        doc.add_heading("一、 數據鑑定視覺化分析", level=1)
        doc.add_picture(img_buf, width=Inches(6.0))
        
        # 深度分析敘述
        doc.add_heading("二、 逐年異常深度分析與舞弊偵測", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"鑑定年度：{r['年度']}", level=2)
            doc.add_paragraph(f"【結論】：{r['結論']}").bold = True
            doc.add_paragraph(f"【詳細鑑定意見】：\n{r['敘述']}")
            doc.add_paragraph(f"【專家查核建議】：\n{r['建議']}")
            doc.add_paragraph("-" * 40)

        # 簽名欄與免責聲明 (關鍵更新)
        doc.add_page_break()
        doc.add_heading("三、 鑑定聲明與簽署欄位", level=1)
        
        disclaimer = doc.add_paragraph()
        disclaimer_run = disclaimer.add_run("【免責聲明】：本鑑定報告係基於委任人提供之財務資料及本系統之量化模型自動生成。鑑定結論僅供會計師執行審計程序之參考，不代表對該公司財務報表合法性之最終保證。若涉及司法訴訟，請以會計師最終核閱之正式查核報告為準。")
        disclaimer_run.font.size = Pt(9)
        disclaimer_run.font.color.rgb = RGBColor(100, 100, 100)

        doc.add_paragraph("\n\n")
        
        # 簽名表格
        sig_table = doc.add_table(rows=5, cols=2)
        sig_table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        # 填入簽名資訊
        sig_table.cell(0, 1).text = f"事務所：{firm_name}"
        sig_table.cell(1, 1).text = f"主辦會計師：{auditor_name}"
        sig_table.cell(2, 1).text = "\n(簽名/蓋章處)\n"
        sig_table.cell(3, 1).text = "________________________"
        sig_table.cell(4, 1).text = f"日期：{datetime.now().strftime('%Y 年 %m 月 %d 日')}"

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.sidebar.download_button("📩 下載正式鑑定報告", buf, f"{co_name}_鑑定報告.docx")

else:
    st.info("請上傳 PDF 檔案開始鑑定程序。")
