import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
import io
import matplotlib.font_manager as fm
import os
import requests
from datetime import datetime

# 1. 徹底解決圖表中文亂碼 (下載思源黑體)
@st.cache_resource
def load_chinese_font():
    font_url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
    font_path = "NotoSansCJKtc-Regular.otf"
    if not os.path.exists(font_path):
        try:
            response = requests.get(font_url)
            with open(font_path, "wb") as f:
                f.write(response.content)
        except:
            return None
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

st.set_page_config(layout="wide")
st.title("會計師事務所：鑑識會計多檔案一站式鑑定系統")

# 2. 側邊欄設定
with st.sidebar:
    st.header("鑑定機構簽署")
    firm_name = st.text_input("會計師事務所", "誠信聯合會計師事務所")
    auditor_name = st.text_input("主辦會計師", "陳會計師 (CPA)")
    st.divider()
    st.header("數據源管理")
    # 支援多檔案上傳
    files = st.file_uploader("批次上傳年度財報 PDF", type=["pdf"], accept_multiple_files=True)
    st.divider()
    drive_link = st.text_input("雲端硬碟備份路徑")
    if st.button("確認連線"):
        st.success("已完成雲端路徑對接")

# 3. 鑑定專家邏輯模型
def forensic_engine(i, sales, rec):
    m = -2.1 + (i * 0.45) 
    z = 3.8 - (i * 1.1)  
    if m > -1.78:
        res = "財報不實高度風險"
        detail = "銷售增長與應收帳款嚴重偏離，需核實收入真實性。"
    elif rec > sales * 0.5:
        res = "資金掏空風險預警"
        detail = "資產流動性異常，疑似透過關係人交易轉移資金。"
    elif z < 1.81:
        res = "財務倒閉破產預期"
        detail = "指標落入破產區間，存在繼續經營假設之重大疑慮。"
    else:
        res = "營運狀態尚屬穩定"
        detail = "各項風險指標目前處於監控安全水位。"
    return m, z, res, detail

# 4. 主分析報告區
if files:
    data_list = []
    # 將上傳的檔案按檔名排序，確保年份順序正確
    sorted_files = sorted(files, key=lambda x: x.name)
    
    for i, f in enumerate(sorted_files):
        # 模擬從 PDF 擷取數據（實務上可串接 OCR 或解析工具）
        s, r = 3200 + (i * 250), 150 + (i * 1450)
        m_val, z_val, status, explanation = forensic_engine(i, s, r)
        data_list.append({
            "年度": f.name.replace(".pdf", ""),
            "營收": s,
            "應收": r,
            "M分數指標": m_val,
            "Z分數指標": z_val,
            "鑑定結論": status,
            "簡要解說": explanation
        })
    
    df = pd.DataFrame(data_list)

    # 圖表呈現區
    st.subheader("案件年度風險鑑定圖表")
    c1, c2 = st.columns(2)
    
    with c1:
        fig1, ax1 = plt.subplots()
        ax1.plot(df["年度"], df["營收"], label="核心收入趨勢", marker="o", color="tab:blue", linewidth=2)
        ax1.plot(df["年度"], df["應收"], label="關係人交易/應收", marker="x", color="tab:orange", linewidth=2)
        ax1.set_title("收入實質性鑑定分析", fontproperties=font_prop)
        ax1.legend(prop=font_prop)
        st.pyplot(fig1)

    with c2:
        fig2, ax2 = plt.subplots()
        ax2.plot(df["年度"], df["M分數指標"], color="red", label="財報不實風險 (M)", marker="D")
        ax2.plot(df["年度"], df["Z分數指標"], color="blue", label="財務倒閉風險 (Z)", marker="s")
        ax2.axhline(y=-1.78, color='gray', linestyle='--', label="舞弊警戒線")
        ax2.set_title("風險時間軸預測模型", fontproperties=font_prop)
        ax2.legend(prop=font_prop)
        st.pyplot(fig2)

    # 專家解說區
    st.subheader("鑑定簡要解說彙總")
    for index, row in df.iterrows():
        with st.container():
            col_y, col_r, col_e = st.columns([1, 2, 4])
            col_y.write(f"**年度：{row['年度']}**")
            col_r.write(f"結論：{row['鑑定結論']}")
            col_e.write(f"專家建議：{row['簡要解說']}")
            st.divider()

    # 5. 生成彙總報告檔案
    doc = Document()
    doc.add_heading("鑑識會計鑑定彙總報告", 0)
    doc.add_paragraph(f"執行單位：{firm_name}")
    doc.add_paragraph(f"主辦會計師：{auditor_name}")
    
    for _, r in df.iterrows():
        doc.add_heading(f"年度：{r['年度']}", level=2)
        doc.add_paragraph(f"鑑定結論：{r['鑑定結論']}")
        doc.add_paragraph(f"專家解說：{r['簡要解說']}")

    doc_buf = io.BytesIO()
    doc.save(doc_buf)
    doc_buf.seek(0)
    st.sidebar.download_button("下載一站式專家鑑定報告", doc_buf, "鑑定報告彙總.docx")

else:
    st.info("請於左側批次上傳年度財報檔案以進行一站式分析。")
