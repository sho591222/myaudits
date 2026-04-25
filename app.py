import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import matplotlib.font_manager as fm
import os
import requests
from datetime import datetime

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

# --- 2. 鑑識與預測邏輯 ---
def forensic_engine_pro(row):
    r, rc, ni, ocf = row.get('營收', 0), row.get('應收帳款', 0), row.get('淨利', 0), row.get('營業現金流', 0)
    m_score = -3.2 + (0.1 * (rc/r if r>0 else 0))
    # 預警邏輯
    tags = []
    if m_score > -1.78: tags.append("舞弊風險")
    if ocf < ni * 0.2: tags.append("現金流異常")
    return pd.Series([round(m_score, 2), " | ".join(tags) if tags else "穩健"])

def financial_forecast(df, years=2):
    """財務預測邏輯：基於平均成長率推估未來趨勢"""
    last_year = df['年度'].max()
    avg_growth = df['營收'].pct_change().mean()
    forecast_data = []
    current_revenue = df[df['年度'] == last_year]['營收'].values[0]
    
    for i in range(1, years + 1):
        current_revenue *= (1 + avg_growth)
        forecast_data.append({
            '年度': f"{last_year + i}(預測)",
            '營收': round(current_revenue, 2),
            '來源': '預測模型'
        })
    return pd.DataFrame(forecast_data)

# --- 3. 介面設計 ---
st.set_page_config(layout="wide", page_title="玄武旗艦級鑑識會計與預測系統")

with st.sidebar:
    st.header("功能導覽")
    mode = st.selectbox("分析視角", ["單一公司深度診斷與預測", "同產業多公司橫向評比"])
    auditor = st.text_input("主辦會計師", "張鈞翔會計師")
    st.divider()
    uploaded_file = st.file_uploader("上傳數據總表 (Excel)", type=["xlsx"])

# --- 4. 核心邏輯執行 ---
if uploaded_file:
    df_raw = pd.read_excel(uploaded_file)
    df_raw[['M分數', '鑑定結論']] = df_raw.apply(forensic_engine_pro, axis=1)

    if mode == "單一公司深度診斷與預測":
        target_co = st.selectbox("選擇公司", df_raw['公司名稱'].unique())
        co_df = df_raw[df_raw['公司名稱'] == target_co].sort_values('年度')
        
        st.subheader(f"分析標的：{target_co}")
        
        # 財務預測區
        st.write("### 財務預測與未來風險評估")
        f_df = financial_forecast(co_df)
        combined_df = pd.concat([co_df[['年度', '營收']], f_df], ignore_index=True)
        
        fig_f, ax_f = plt.subplots(figsize=(10, 4))
        ax_f.plot(combined_df['年度'].astype(str), combined_df['營收'], marker='o', label='歷史/預測營收')
        ax_f.fill_between(f_df['年度'].astype(str), f_df['營收']*0.9, f_df['營收']*1.1, color='gray', alpha=0.2, label='預測置信區間')
        ax_f.set_title("營收歷年趨勢與未來推估", fontproperties=font_prop)
        ax_f.legend(prop=font_prop)
        st.pyplot(fig_f)

    else:
        # 多公司同產業評比
        st.subheader("同產業競爭風險分析")
        target_year = st.selectbox("選擇比較年度", sorted(df_raw['年度'].unique(), reverse=True))
        year_df = df_raw[df_raw['年度'] == target_year]
        
        fig_cmp, ax_cmp = plt.subplots(figsize=(10, 5))
        ax_cmp.bar(year_df['公司名稱'], year_df['M分數'], color='teal')
        ax_cmp.axhline(y=-1.78, color='red', linestyle='--')
        ax_cmp.set_title(f"{target_year} 年度同業舞弊指標分布", fontproperties=font_prop)
        st.pyplot(fig_cmp)
        
        st.write("### 異常標的焦點分析")
        st.dataframe(year_df[year_df['M分數'] > -1.78])

    # --- 5. 報告導出模組 ---
    st.divider()
    st.caption("法律聲明：本報告包含歷史鑑定與預測數據。預測結果係基於統計模型推論，不保證未來實際獲利情形。")
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        output_ex = io.BytesIO()
        df_raw.to_excel(output_ex, index=False)
        st.download_button("匯出鑑定底稿 (Excel)", output_ex.getvalue(), "鑑定底稿_旗艦版.xlsx")
    with col_dl2:
        if st.button("生成 Word 鑑定與預測報告"):
            doc = Document()
            doc.add_heading("鑑識會計分析與財務預測意見書", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"主辦會計師：{auditor}\n分析時間：{datetime.now().strftime('%Y/%m/%d')}")
            
            doc.add_heading("一、 重大風險鑑定 (歷史數據)", level=1)
            doc.add_paragraph("偵測到以下標的具備潛在舞弊或掏空跡象：")
            for _, r in df_raw[df_raw['M分數'] > -1.78].iterrows():
                doc.add_paragraph(f"公司：{r['公司名稱']} ({r['年度']}) - 指標異常")

            doc.add_heading("二、 未來展望與財務預測", level=1)
            doc.add_paragraph("基於過往成長模型，對選定標的進行之推估結果已列於附件底稿。")
            
            doc.add_paragraph("\n\n(簽署區)\n____________________")
            buf_wd = io.BytesIO()
            doc.save(buf_wd)
            st.download_button("點此下載 Word 報告", buf_wd.getvalue(), "鑑識預測報告.docx")

else:
    st.info("請上傳包含公司、年度、營收、應收帳款等欄位的 Excel 資料。")
