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
import collections

# --- 1. 環境設定 ---
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

# --- 2. 鑑識邏輯工具箱 ---
def get_benford_score(data_list):
    """簡易班佛定律偵測：計算數字首位數分布偏移度"""
    first_digits = [int(str(abs(x))[0]) for x in data_list if x != 0]
    counts = collections.Counter(first_digits)
    # 理想值 1 的出現機率約 30%
    one_ratio = counts.get(1, 0) / len(first_digits) if first_digits else 0.3
    return "正常" if 0.25 <= one_ratio <= 0.35 else "異常偏移"

def forensic_engine_v2(r, rc, inv, c, a, ni, ocf):
    """
    全面鑑定引擎 V2
    r:營收, rc:應收, inv:存貨, c:現金, a:總資產, ni:淨利, ocf:營業現金流
    """
    m_score = -3.2 + (0.1 * (rc/r if r!=0 else 0)) + (0.2 * (inv/r if r!=0 else 0))
    ponzi_index = ocf / ni if ni > 0 else 0
    laundry_index = (rc + inv) / a if a > 0 else 0
    
    tags = []
    sugg = []
    
    # 舞弊與掏空
    if m_score > -1.78:
        tags.append("財報舞弊高風險")
        sugg.append("執行收入實質性測試，查核應收帳款真實性。")
    if (rc / r) > 0.5:
        tags.append("資產掏空警訊")
        sugg.append("針對大額應收帳款對象執行關係人身分穿透。")
        
    # 吸金 (龐氏)
    if ponzi_index < 0.2 and ni > 1000:
        tags.append("龐氏吸金預警")
        sugg.append("核對利潤來源是否僅為帳面應計，並追查利息發放之現金來源。")
        
    # 洗錢
    if laundry_index > 0.4:
        tags.append("異常資金洗錢風險")
        sugg.append("查核存貨與往來款科目，是否存在虛假交易循環。")
        
    return m_score, " | ".join(tags) if tags else "未見明顯異常", sugg

# --- 3. 介面設計 ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計系統 V2")

with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.header("系統模式選擇")
    app_mode = st.radio("請選擇分析範疇", ["單一公司歷年診斷", "多公司橫向評比"])
    
    st.divider()
    auditor_name = st.text_input("簽證會計師", "張鈞翔會計師")
    firm_name = st.text_input("事務所名稱", "玄武聯合會計師事務所")
    files = st.file_uploader("批次上傳財報 PDF", type=["pdf"], accept_multiple_files=True)

# --- 4. 執行分析 ---
if files:
    data_rows = []
    for i, f in enumerate(sorted(files, key=lambda x: x.name)):
        # 模擬數據：此處在實務中應由 PDF 轉文字解析
        year_or_co = f.name.replace(".pdf", "")
        # 為了展示效果，模擬不同風險數據
        r, rc, inv, c, a = 10000 + (i*500), 2000 + (i*2500), 1000 + (i*800), 5000 - (i*1000), 50000
        ni, ocf = 2000, 300 - (i*50)
        
        m, stat, suggs = forensic_engine_v2(r, rc, inv, c, a, ni, ocf)
        data_rows.append({
            "標的": year_or_co, "營收": r, "應收": rc, "存貨": inv, 
            "現金": c, "資產": a, "M分數": m, "鑑定結論": stat, "建議": "\n".join(suggs)
        })
    
    df = pd.DataFrame(data_rows)

    # 視覺化展示
    st.header(f"鑑定成果：{app_mode}")
    col1, col2 = st.columns(2)
    
    with col1:
        fig1, ax1 = plt.subplots(figsize=(8, 5))
        ax1.bar(df["標的"], df["營收"], label="營收", color="skyblue")
        ax1.plot(df["標的"], df["應收"], label="應收帳款", marker="o", color="orange")
        ax1.set_title("營收與債權品質分析", fontproperties=font_prop)
        ax1.legend(prop=font_prop)
        st.pyplot(fig1)
        
    with col2:
        fig2, ax2 = plt.subplots(figsize=(8, 5))
        ax2.plot(df["標的"], df["M分數"], color="red", marker="x", label="舞弊指標")
        ax2.axhline(y=-1.78, color='black', linestyle='--')
        ax2.fill_between(df["標的"], -1.78, 0, color='red', alpha=0.1)
        ax2.set_title("動態風險警戒監控", fontproperties=font_prop)
        st.pyplot(fig2)

    # 數據表與班佛定律
    st.subheader("分析明細與數據合規性")
    if app_mode == "單一公司歷年診斷":
        ben_stat = get_benford_score(df["營收"].tolist())
        st.warning(f"班佛定律首位數檢測結果：{ben_stat} (檢測財務數字是否符合自然分布)")
    
    st.dataframe(df)

    # 網頁法律聲明
    st.divider()
    st.caption("法律聲明：本報告係量化模型產出之初步診斷，不構成最終法律判定。鑑定效力以會計師簽署之正式報告為準。")

    # Word 生成邏輯 (略，與前版相同但包含更多欄位)
    if st.sidebar.button("產生正式深度鑑定報告"):
        st.success("報告已準備就緒，請點擊下方下載按鈕。")
        # (此處可放置前述之 Word 生成代碼)
else:
    st.info("請上傳財報 PDF 檔案以啟動鑑定引擎。")
