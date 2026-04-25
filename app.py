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

# --- 1. 環境與字體設定 ---
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

# --- 2. 核心邏輯模組：鑑定與預測 ---
def forensic_logic(row):
    r, rc, inv, ni, ocf = row.get('營收', 0), row.get('應收帳款', 0), row.get('存貨', 0), row.get('淨利', 0), row.get('營業現金流', 0)
    m_score = -3.2 + (0.15 * (rc/r if r>0 else 0)) + (0.1 * (inv/r if r>0 else 0))
    tags = []
    if m_score > -1.78: tags.append("舞弊高風險")
    if (rc / r) > 0.45 if r>0 else False: tags.append("掏空警訊")
    if ocf < ni * 0.15 and ni > 0: tags.append("龐氏吸金預警")
    return pd.Series([round(m_score, 2), " | ".join(tags) if tags else "經營穩健"])

def forecast_engine(co_df, future_years=3):
    """基於移動平均成長率進行預測"""
    last_val = co_df['營收'].iloc[-1]
    last_year = co_df['年度'].iloc[-1]
    growth_rate = co_df['營收'].pct_change().mean()
    
    forecast_list = []
    for i in range(1, future_years + 1):
        last_val *= (1 + growth_rate)
        forecast_list.append({
            '年度': f"{int(last_year) + i}(預測)",
            '營收': round(last_val, 2),
            '分析類型': '未來預測'
        })
    return pd.DataFrame(forecast_list)

# --- 3. 介面架構 ---
st.set_page_config(layout="wide", page_title="玄武旗艦鑑識與預測系統")

with st.sidebar:
    st.header("系統參數與數據上傳")
    auditor = st.text_input("主辦會計師", "張鈞翔會計師")
    firm = st.text_input("事務所名稱", "玄武聯合會計師事務所")
    st.divider()
    # 支援多檔案上傳
    files = st.file_uploader("批次上傳公司數據 (Excel)", type=["xlsx"], accept_multiple_files=True)

# --- 4. 數據聚合與運算 ---
if files:
    all_dfs = []
    for f in files:
        tmp = pd.read_excel(f)
        if '公司名稱' not in tmp.columns:
            tmp['公司名稱'] = f.name.replace(".xlsx", "")
        all_dfs.append(tmp)
    
    df = pd.concat(all_dfs, ignore_index=True)
    df[['M分數', '鑑定結論']] = df.apply(forensic_logic, axis=1)
    st.success(f"已聚合 {len(files)} 份檔案。")

    # --- 5. 多維度功能分析選單 ---
    mode = st.radio("功能模式選擇", ["單一公司：深度診斷與財務預測", "多間公司：橫向對比與產業分析"])

    if mode == "單一公司：深度診斷與財務預測":
        target = st.selectbox("選擇調查標的", df['公司名稱'].unique())
        sub_df = df[df['公司名稱'] == target].sort_values('年度')
        
        # 顯示歷史趨勢與預測圖表
        st.subheader(f"{target} 財務趨勢與 3 年預測")
        f_df = forecast_engine(sub_df)
        
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.plot(sub_df['年度'].astype(str), sub_df['營收'], marker='o', label='歷史營收')
        ax.plot(f_df['年度'], f_df['營收'], marker='s', linestyle='--', color='gray', label='模型預測營收')
        ax.set_title("營收歷史軌跡與未來推估", fontproperties=font_prop)
        ax.legend(prop=font_prop)
        st.pyplot(fig)
        st.dataframe(sub_df)

    else:
        st.subheader("多公司同產業風險對比")
        year = st.selectbox("比較年度", sorted(df['年度'].unique(), reverse=True))
        year_df = df[df['年度'] == year]
        
        col1, col2 = st.columns(2)
        with col1:
            fig2, ax2 = plt.subplots()
            ax2.bar(year_df['公司名稱'], year_df['M分數'], color='orange')
            ax2.axhline(y=-1.78, color='red', linestyle='--')
            ax2.set_title(f"{year} 年度各公司舞弊風險評比", fontproperties=font_prop)
            st.pyplot(fig2)
        with col2:
            st.write("### 異常標的警告名單")
            st.dataframe(year_df[year_df['M分數'] > -1.78][['公司名稱', 'M分數', '鑑定結論']])

    # --- 6. 報告匯出與法律聲明 ---
    st.divider()
    st.caption("法律聲明：本鑑定報告與財務預測內容僅供專業查核參考，不具備最終司法判決效力。")

    c1, c2 = st.columns(2)
    with c1:
        out_ex = io.BytesIO()
        df.to_excel(out_ex, index=False)
        st.download_button("匯出聚合鑑定底稿 (Excel)", out_ex.getvalue(), "聚合底稿.xlsx")
    with c2:
        if st.button("產生綜合鑑定意見書 (Word)"):
            doc = Document()
            doc.add_heading("鑑識會計鑑定與預測報告", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"事務所：{firm}\n會計師：{auditor}")
            doc.add_heading("一、 重大舞弊風險發現", level=1)
            for _, r in df[df['M分數'] > -1.78].iterrows():
                doc.add_paragraph(f"{r['公司名稱']} ({r['年度']})：{r['鑑定結論']}")
            
            doc.add_heading("二、 未來營運成長預測", level=1)
            doc.add_paragraph("本報告內建之預測模型已針對個別標的進行趨勢分析，詳見附件底稿。")
            
            doc.add_paragraph("\n\n會計師簽署：____________________")
            buf_wd = io.BytesIO()
            doc.save(buf_word := buf_wd)
            st.download_button("下載綜合報告 (Word)", buf_word.getvalue(), "鑑定報告.docx")

else:
    st.info("系統就緒。請上傳一個或多個 Excel 檔案以啟動多維度分析。")
