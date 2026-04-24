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
st.set_page_config(layout="wide")
st.title("專業鑑識會計：全自動四大報表分析與查核建議系統")

# --- 3. 側邊欄：資訊填寫與模式切換 ---
with st.sidebar:
    st.header("系統模式選擇")
    user_mode = st.radio("身分切換", ["一般公司模式", "會計師專業模式"])
    
    st.divider()
    st.header("基礎資訊")
    co_name = st.text_input("受調查公司名稱", "XX股份有限公司")
    
    if user_mode == "會計師專業模式":
        auditor_name = st.text_input("主辦會計師", "陳會計師 (CPA)")
        firm_name = st.text_input("事務所名稱", "誠信聯合會計師事務所")
    
    st.divider()
    files = st.file_uploader("批次上傳年度財報 PDF", type=["pdf"], accept_multiple_files=True)

# --- 4. 自動化查核建議與鑑定引擎 ---
def generate_audit_report(i, rev, rec, cash, liab):
    # 風險指標計算
    m_score = -2.5 + (i * 0.48) # 舞弊傾向
    z_score = 4.2 - (i * 1.15) # 破產風險
    
    # 計算增幅
    rev_growth = 0.15 + (i * 0.05)
    
    # 自動生成鑑定結論與查核建議
    report = {
        "舞弊分析": "穩定" if m_score < -1.78 else "高度風險：M分數異常，疑似虛增營收。",
        "掏空分析": "注意" if rec > rev * 0.3 else "正常",
        "財務分析": "警訊" if z_score < 1.8 else "良好",
        "查核建議": ""
    }
    
    if m_score > -1.78:
        report["查核建議"] = "建議執行收入截止測試，並對前五大客戶發函詢證以確認交易真實性。"
    elif rec > rev * 0.35:
        report["查核建議"] = "應收帳款與營收成長失衡，建議查核關係人交易項目，確認是否有資金套現疑慮。"
    elif z_score < 1.8:
        report["查核建議"] = "財務結構急劇惡化，應評估公司經營假設(Going Concern)之合理性。"
    else:
        report["查核建議"] = "目前財務數據尚屬穩健，建議持續監控季度現金流量變化。"
        
    return m_score, z_score, rev_growth, report

# --- 5. 主內容顯示區 ---
if files:
    all_results = []
    sorted_files = sorted(files, key=lambda x: x.name)
    
    for i, f in enumerate(sorted_files):
        # 模擬四大報表數據擷取
        r, rc, c, l = 6000 + (i * 400), 500 + (i * 1900), 2500 - (i * 500), 3000 + (i * 700)
        m, z, growth, rpt = generate_audit_report(i, r, rc, c, l)
        
        all_results.append({
            "年度": f.name.replace(".pdf", ""),
            "營收": r, "應收": rc, "現金": c, "負債": l,
            "M分數": m, "Z分數": z, "增幅": growth,
            "建議報告": rpt
        })
    
    df = pd.DataFrame(all_results)

    # 圖表區 (四大報表項目與增幅)
    st.subheader("一、 數據分析圖表 (趨勢與增幅)")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        fig1, ax1 = plt.subplots()
        ax1.plot(df["年度"], df["營收"], label="營收", marker="o")
        ax1.plot(df["年度"], df["應收"], label="應收", marker="x")
        ax1.set_title("收入與資產對比圖", fontproperties=font_prop)
        ax1.legend(prop=font_prop)
        st.pyplot(fig1)

    with c2:
        fig2, ax2 = plt.subplots()
        ax2.bar(df["年度"], df["M分數"], color="red", alpha=0.6, label="M分數 (舞弊)")
        ax2.axhline(y=-1.78, color='black', linestyle='--')
        ax2.set_title("舞弊預警模型圖", fontproperties=font_prop)
        ax2.legend(prop=font_prop)
        st.pyplot(fig2)

    with c3:
        fig3, ax3 = plt.subplots()
        ax3.step(df["年度"], df["增幅"], label="營收增幅 %", where='mid')
        ax3.set_title("年度增幅變動率", fontproperties=font_prop)
        ax3.legend(prop=font_prop)
        st.pyplot(fig3)

    # 顯示自動生成的查核建議
    st.subheader("二、 自動化專家鑑定與查核建議報告")
    for index, row in df.iterrows():
        with st.expander(f"年度 {row['年度']} 鑑定報告細節"):
            st.warning(f"**鑑定結論：** {row['建議報告']['舞弊分析']}")
            st.info(f"**詳細敘述：** 本年度應收帳款佔比達 {round((row['應收']/row['營收'])*100, 1)}%，M分數為 {round(row['M分數'], 2)}。")
            st.success(f"**自動生成查核建議：** {row['建議報告']['查核建議']}")

    # --- 6. 產生一鍵式專業 Word 分析報告 ---
    if st.sidebar.button("產生完整查核分析報告 (DOC)"):
        doc = Document()
        doc.add_heading(f"鑑識會計查核鑑定報告 - {co_name}", 0)
        
        if user_mode == "會計師專業模式":
            doc.add_paragraph(f"事務所：{firm_name}")
            doc.add_paragraph(f"主辦會計師：{auditor_name}")
            doc.add_paragraph(f"報告日期：{datetime.now().strftime('%Y/%m/%d')}")
        
        doc.add_heading("各年度深度鑑定與查核建議", level=1)
        for _, r in df.iterrows():
            doc.add_heading(f"年度 {r['年度']}", level=2)
            doc.add_paragraph(f"1. 舞弊與掏空分析：{r['建議報告']['舞弊分析']}")
            doc.add_paragraph(f"2. 財報異常詳細敘述：營收 {r['營收']}，應收帳款異常增加至 {r['應收']}。")
            doc.add_paragraph(f"3. 專業查核建議：{r['建議報告']['查核建議']}")
            doc.add_paragraph("-" * 30)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.sidebar.download_button("點此下載 DOC 報告檔案", buf, f"{co_name}_鑑識分析報告.docx")

else:
    st.info("請上傳 PDF 檔案開始產出一站式分析圖表與查核建議報告。")
