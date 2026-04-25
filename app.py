import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import matplotlib.font_manager as fm
import os
import requests
from datetime import datetime
import pdfplumber  # 導入 PDF 解析工具

# --- 1. 環境設定：中文字體 ---
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

# --- 2. 數據解析與鑑定引擎 ---
def parse_pdf_to_df(file):
    """從 PDF 提取表格數據並轉換為 DataFrame"""
    try:
        all_text = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_text.extend(table)
        # 轉換為 DF 並進行基礎清洗
        pdf_df = pd.DataFrame(all_text[1:], columns=all_text[0])
        # 這裡假設 PDF 欄位名稱正確，若不正確需在此處進行對照清洗
        return pdf_df
    except Exception as e:
        st.error(f"PDF 解析失敗 ({file.name}): {e}")
        return pd.DataFrame()

def forensic_engine(row):
    r = pd.to_numeric(row.get('營收', 0), errors='coerce') or 0
    rc = pd.to_numeric(row.get('應收帳款', 0), errors='coerce') or 0
    inv = pd.to_numeric(row.get('存貨', 0), errors='coerce') or 0
    m_score = -3.2 + (0.15 * (rc/r if r>0 else 0)) + (0.1 * (inv/r if r>0 else 0))
    status = "風險預警" if m_score > -1.78 else "經營穩健"
    return pd.Series([round(m_score, 2), status])

def get_forecast(df, years=3):
    if len(df) < 2: return pd.DataFrame()
    last_year = df['年度'].max()
    growth = df['營收'].pct_change().mean()
    curr_rev = df['營收'].iloc[-1]
    f_list = []
    for i in range(1, years + 1):
        curr_rev *= (1 + growth)
        f_list.append({'年度': f"{int(last_year)+i}(預測)", '營收': round(curr_rev, 2), '分析類型': '未來推估'})
    return pd.DataFrame(f_list)

# --- 3. 側邊欄與功能切換 ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計旗艦平台")

with st.sidebar:
    st.header("功能導航中心")
    analysis_mode = st.radio(
        "選擇分析模式",
        [
            "單一公司：歷年趨勢與財務預測",
            "單一公司：特定單年深度診斷",
            "對比模式：單一公司 vs 多家同業",
            "群體模式：多公司多年風險掃描"
        ]
    )
    st.divider()
    auditor = st.text_input("主辦會計師", "張鈞翔會計師")
    # 支援 PDF 與 Excel 多選上傳
    files = st.file_uploader("批次上傳數據 (支援 Excel 與 PDF)", type=["xlsx", "pdf"], accept_multiple_files=True)

# --- 4. 數據整合流程 ---
if files:
    all_data = []
    for f in files:
        if f.name.endswith('.xlsx'):
            tmp = pd.read_excel(f)
        elif f.name.endswith('.pdf'):
            tmp = parse_pdf_to_df(f)
        
        if not tmp.empty:
            if '公司名稱' not in tmp.columns:
                tmp['公司名稱'] = f.name.replace(".xlsx", "").replace(".pdf", "")
            all_data.append(tmp)
    
    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        # 數值轉換確保計算正確
        for col in ['營收', '應收帳款', '存貨', '年度']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        df[['M分數', '結論']] = df.apply(forensic_engine, axis=1)

        # --- 5. 介面呈現 (模式分界) ---
        if analysis_mode == "單一公司：歷年趨勢與財務預測":
            target = st.selectbox("選擇公司", df['公司名稱'].unique())
            sub = df[df['公司名稱'] == target].sort_values('年度')
            st.header(f"{target} 深度鑑定與成長預測")
            
            f_df = get_forecast(sub)
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='歷史實際營收')
            if not f_df.empty:
                ax.plot(f_df['年度'], f_df['營收'], '--', marker='s', color='gray', label='模型預測趨勢')
            ax.set_title("營收軌跡與 3 年未來預測", fontproperties=font_prop)
            ax.legend(prop=font_prop)
            st.pyplot(fig)
            st.dataframe(sub)

        elif analysis_mode == "對比模式：單一公司 vs 多家同業":
            year = st.selectbox("比較基準年度", sorted(df['年度'].unique(), reverse=True))
            target = st.selectbox("主選公司 (標色)", df['公司名稱'].unique())
            year_df = df[df['年度'] == year]
            
            st.header(f"{year} 年度同業風險評比")
            fig2, ax2 = plt.subplots(figsize=(10, 5))
            colors = ['red' if c == target else 'skyblue' for c in year_df['公司名稱']]
            ax2.bar(year_df['公司名稱'], year_df['M分數'], color=colors)
            ax2.axhline(y=-1.78, color='black', linestyle='--')
            ax2.set_title(f"風險對照圖 (紅色為主選標的)", fontproperties=font_prop)
            st.pyplot(fig2)

        # 這裡可繼續加入其餘模式邏輯 (如單年診斷、群體掃描)...

        # --- 6. 報告導出 ---
        st.divider()
        col_ex, col_wd = st.columns(2)
        with col_ex:
            out_ex = io.BytesIO()
            df.to_excel(out_ex, index=False)
            st.download_button("下載整合鑑定底稿 (Excel)", out_ex.getvalue(), "鑑定底稿.xlsx")
        with col_wd:
            if st.button("生成綜合報告 (Word)"):
                doc = Document()
                doc.add_heading("鑑識會計鑑定報告書", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph(f"主辦會計師：{auditor}\n日期：{datetime.now().strftime('%Y/%m/%d')}")
                # 報告內容邏輯...
                buf_word = io.BytesIO()
                doc.save(buf_word)
                st.download_button("下載 Word 報告", buf_word.getvalue(), "報告書.docx")
else:
    st.info("系統就緒。請上傳數據檔案並選擇分析模式。")
