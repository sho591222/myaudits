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

# --- 2. 核心邏輯引擎 ---
def forensic_engine(row):
    r = row.get('營收', 0)
    rc = row.get('應收帳款', 0)
    inv = row.get('存貨', 0)
    m_score = -3.2 + (0.15 * (rc/r if r>0 else 0)) + (0.1 * (inv/r if r>0 else 0))
    status = "風險預警" if m_score > -1.78 else "經營穩健"
    return pd.Series([round(m_score, 2), status])

def get_forecast_data(df, years=3):
    """財務預測：基於營收平均成長率"""
    if len(df) < 2: return pd.DataFrame()
    last_year = df['年度'].max()
    growth = df['營收'].pct_change().mean()
    curr_rev = df['營收'].iloc[-1]
    f_list = []
    for i in range(1, years + 1):
        curr_rev *= (1 + growth)
        f_list.append({'年度': f"{int(last_year)+i}(預測)", '營收': round(curr_rev, 2), '分析類型': '未來推估'})
    return pd.DataFrame(f_list)

# --- 3. 介面架構 ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計旗艦平台")

with st.sidebar:
    st.header("功能導航中心")
    # 提供四種分析視角
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
    files = st.file_uploader("批次上傳數據 (Excel/TEJ)", type=["xlsx"], accept_multiple_files=True)

# --- 4. 數據整合與執行 ---
if files:
    all_data = []
    for f in files:
        tmp = pd.read_excel(f)
        if '公司名稱' not in tmp.columns:
            tmp['公司名稱'] = f.name.replace(".xlsx", "")
        all_data.append(tmp)
    
    df = pd.concat(all_data, ignore_index=True)
    df[['M分數', '結論']] = df.apply(forensic_engine, axis=1)

    # --- 5. 根據選單切換介面 ---
    if analysis_mode == "單一公司：歷年趨勢與財務預測":
        target = st.selectbox("選擇公司", df['公司名稱'].unique())
        sub = df[df['公司名稱'] == target].sort_values('年度')
        st.header(f"{target} 歷年趨勢與未來預測")
        
        f_df = get_forecast_data(sub)
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='歷史實際營收')
        if not f_df.empty:
            ax.plot(f_df['年度'], f_df['營收'], '--', marker='s', color='gray', label='模型預測趨勢')
        ax.set_title("營收歷史軌跡與 3 年未來推估", fontproperties=font_prop)
        ax.legend(prop=font_prop)
        st.pyplot(fig)
        st.dataframe(sub)

    elif analysis_mode == "單一公司：特定單年深度診斷":
        target = st.selectbox("選擇公司", df['公司名稱'].unique(), key="single_co")
        year = st.selectbox("選擇年份", df[df['公司名稱'] == target]['年度'].unique())
        st.header(f"{target} - {year} 年度鑑定明細")
        st.write(df[(df['公司名稱'] == target) & (df['年度'] == year)])

    elif analysis_mode == "對比模式：單一公司 vs 多家同業":
        year = st.selectbox("比較基準年度", sorted(df['年度'].unique(), reverse=True))
        target = st.selectbox("主選公司 (標色)", df['公司名稱'].unique())
        year_df = df[df['年度'] == year]
        
        st.header(f"{year} 年度同業風險評比")
        fig2, ax2 = plt.subplots(figsize=(10, 5))
        colors = ['red' if c == target else 'skyblue' for c in year_df['公司名稱']]
        ax2.bar(year_df['公司名稱'], year_df['M分數'], color=colors)
        ax2.axhline(y=-1.78, color='black', linestyle='--')
        ax2.set_title(f"{target} 與同業之舞弊指標對照 (紅色為主選)", fontproperties=font_prop)
        st.pyplot(fig2)

    else: # 多公司多年掃描
        st.header("全體標的多年度風險掃描")
        risk_only = df[df['M分數'] > -1.78].sort_values(['公司名稱', '年度'])
        st.warning("偵測到以下年度數據具備高舞弊風險：")
        st.dataframe(risk_only)

    # --- 6. 報告產出區 ---
    st.divider()
    st.caption("【法律聲明】本報告包含自動化鑑定與財務預測。最終結論應以會計師執行實質查核後之簽署報告為準。")
    
    col_ex, col_wd = st.columns(2)
    with col_ex:
        out_ex = io.BytesIO()
        df.to_excel(out_ex, index=False)
        st.download_button("下載整合鑑定底稿 (Excel)", out_ex.getvalue(), "鑑定底稿.xlsx")
    
    with col_wd:
        if st.button("產生綜合分析報告 (Word)"):
            doc = Document()
            doc.add_heading("鑑識會計多維度分析報告書", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"主辦會計師：{auditor}\n報告產出日期：{datetime.now().strftime('%Y/%m/%d')}")
            
            doc.add_heading("一、 重大風險異常名單", level=1)
            for _, r in df[df['M分數'] > -1.78].iterrows():
                doc.add_paragraph(f"標的：{r['公司名稱']} ({r['年度']}) - M分數：{r['M分數']} (判定：{r['結論']})")
            
            doc.add_heading("二、 鑑定意見與聲明", level=1)
            doc.add_paragraph("本鑑定報告由自動化系統產出。財務預測數據係基於歷史成長率推估，不保證未來獲利實現。")
            
            buf_wd = io.BytesIO()
            doc.save(buf_wd)
