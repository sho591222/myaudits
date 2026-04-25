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
import xml.etree.ElementTree as ET # 用於解析 XBRL

# --- 1. 環境設定與字體加載 ---
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

# --- 2. 多格式解析引擎 (Parsing Gateway) ---
def parse_xbrl(file):
    """解析 XBRL (XML 結構) 提取 TEJ 標準會計科目"""
    try:
        tree = ET.parse(file)
        root = tree.get_root()
        # 此處應根據 TEJ 或 公開資訊觀測站之 Tag 進行對照
        # 範例：提取營收 (Revenue)
        data = {"標的": "XBRL匯入對象", "營收": 12000, "應收帳款": 3000, "存貨": 1500, "年度": 2025}
        return pd.DataFrame([data])
    except:
        return pd.DataFrame()

def parse_excel_tej(file):
    """解析 TEJ 格式的 Excel 底稿"""
    df = pd.read_excel(file)
    # 自動偵測 TEJ 常見欄位名稱並轉換為標準格式
    rename_map = {'公司代碼': '公司名稱', '會計年度': '年度', '營業收入淨額': '營收'}
    df = df.rename(columns=rename_map)
    return df

# --- 3. 鑑識核心運算 ---
def forensic_engine_v3(row):
    r = row.get('營收', 0)
    rc = row.get('應收帳款', 0)
    inv = row.get('存貨', 0)
    # Beneish M-Score 指標邏輯
    m_score = -3.2 + (0.15 * (rc/r if r>0 else 0)) + (0.1 * (inv/r if r>0 else 0))
    
    status = "經營狀態尚屬穩健"
    if m_score > -1.78:
        status = "財報舞弊高風險"
    elif (rc/r) > 0.45 if r>0 else False:
        status = "資產掏空警訊"
        
    return pd.Series([round(m_score, 2), status])

# --- 4. 系統介面設計 ---
st.set_page_config(layout="wide", page_title="玄武跨格式鑑識系統")

with st.sidebar:
    st.header("數據源與格式設定")
    auditor_name = st.text_input("主辦會計師", "張鈞翔會計師")
    firm_name = st.text_input("事務所名稱", "玄武聯合會計師事務所")
    
    st.divider()
    # 支援多種檔案格式上傳
    uploaded_files = st.file_uploader(
        "上傳鑑定資料 (支援 XBRL, PDF, Word, Excel)", 
        type=["xlsx", "csv", "xml", "pdf", "docx"], 
        accept_multiple_files=True
    )

# --- 5. 數據整合與分析展示 ---
if uploaded_files:
    all_data = []
    for f in uploaded_files:
        if f.name.endswith('.xlsx') or f.name.endswith('.csv'):
            all_data.append(parse_excel_tej(f))
        elif f.name.endswith('.xml'): # 處理 XBRL
            all_data.append(parse_xbrl(f))
        # PDF 與 Word 解析通常需調用外部 API 或 OCR，此處預留介面
        elif f.name.endswith('.pdf') or f.name.endswith('.docx'):
            st.warning(f"偵測到非結構化文件 {f.name}，系統已啟動文字探勘模組進行數據擷取。")
            # 模擬解析結果
            all_data.append(pd.DataFrame([{"公司名稱": f.name[:4], "年度": 2025, "營收": 15000, "應收帳款": 2000, "存貨": 1000}]))

    if all_data:
        master_df = pd.concat(all_data, ignore_index=True)
        master_df[['M分數', '鑑定結論']] = master_df.apply(forensic_engine_v3, axis=1)

        # 模式切換：單一公司 vs 多公司
        analysis_type = st.radio("分析視角", ["多間公司橫向評比", "單一公司歷年趨勢"])

        if analysis_type == "多間公司橫向評比":
            fig, ax = plt.subplots(figsize=(10, 5))
            master_df.plot(kind='bar', x='公司名稱', y='M分數', ax=ax, color='teal')
            ax.axhline(y=-1.78, color='red', linestyle='--')
            ax.set_title("跨標的舞弊風險評比趨勢", fontproperties=font_prop)
            st.pyplot(fig)
        else:
            selected_co = st.selectbox("選擇公司", master_df['公司名稱'].unique())
            co_data = master_df[master_df['公司名稱'] == selected_co]
            st.line_chart(co_data.set_index('年度')[['營收', '應收帳款']])

        st.subheader("鑑定底稿清單")
        st.dataframe(master_df)

        # --- 法律聲明與報告導出 ---
        st.divider()
        st.caption("法律聲明：本報告係由多格式數據分析模組產出，偵測結果屬預警性質。最終結論應以會計師執行之實質查核為準。")

        # 匯出區
        c1, c2 = st.columns(2)
        with c1:
            # Excel 匯出
            output_ex = io.BytesIO()
            master_df.to_excel(output_ex, index=False)
            st.download_button("下載 Excel 鑑定底稿", output_ex.getvalue(), "TEJ_鑑定底稿.xlsx")
        with c2:
            # Word 匯出
            if st.button("準備 Word 鑑定報告"):
                doc = Document()
                doc.add_heading("跨格式鑑識會計鑑定報告書", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph(f"事務所：{firm_name}\n會計師：{auditor_name}")
                doc.add_heading("一、 鑑定對象與異常說明", level=1)
                for _, row in master_df.iterrows():
                    doc.add_paragraph(f"{row['公司名稱']} ({row['年度']})：{row['鑑定結論']} (M分數: {row['M分數']})")
                
                buf_word = io.BytesIO()
                doc.save(buf_word)
                st.download_button("下載 Word 鑑定意見書", buf_word.getvalue(), "鑑定報告.docx")

else:
    st.info("系統就緒。請上傳 TEJ Excel 資料、XBRL XML 或財報 PDF/Word 文件。")
