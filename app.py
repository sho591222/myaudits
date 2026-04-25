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

# --- 2. 鑑定與預測引擎 ---
def forensic_engine_v4(row):
    r, rc, inv = row.get('營收', 0), row.get('應收帳款', 0), row.get('存貨', 0)
    m_score = -3.2 + (0.15 * (rc/r if r>0 else 0)) + (0.1 * (inv/r if r>0 else 0))
    status = "風險預警" if m_score > -1.78 else "經營穩健"
    return pd.Series([round(m_score, 2), status])

def get_forecast(df, years=3):
    last_year = df['年度'].max()
    growth = df['營收'].pct_change().mean()
    f_list = []
    curr_rev = df['營收'].iloc[-1]
    for i in range(1, years + 1):
        curr_rev *= (1 + growth)
        f_list.append({'年度': f"{int(last_year)+i}(預測)", '營收': round(curr_rev, 2), '類型': '預測'})
    return pd.DataFrame(f_list)

# --- 3. 側邊欄：進階功能選單 ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計大數據平台")

with st.sidebar:
    st.header("鑑定參數與功能選擇")
    # 模式切換分界
    analysis_mode = st.radio(
        "選擇分析模式",
        [
            "單一公司：單年/多年診斷與預測",
            "單一公司 vs 其他公司對比",
            "多公司：多年橫向風險評比"
        ]
    )
    
    st.divider()
    auditor = st.text_input("主辦會計師", "張鈞翔會計師")
    files = st.file_uploader("批次上傳 TEJ/Excel 資料", type=["xlsx"], accept_multiple_files=True)

# --- 4. 數據整合與分析顯示 ---
if files:
    all_dfs = []
    for f in files:
        tmp = pd.read_excel(f)
        if '公司名稱' not in tmp.columns:
            tmp['公司名稱'] = f.name.replace(".xlsx", "")
        all_dfs.append(tmp)
    
    df = pd.concat(all_dfs, ignore_index=True)
    df[['M分數', '結論']] = df.apply(forensic_engine_v4, axis=1)

    if "單一公司：單年/多年診斷與預測" in analysis_mode:
        target_co = st.selectbox("選擇受調查公司", df['公司名稱'].unique())
        sub_df = df[df['公司名稱'] == target_co].sort_values('年度')
        
        st.header(f"{target_co} 深度診斷報告")
        
        # 單年 vs 多年切換
        view_type = st.radio("檢視範疇", ["多年趨勢與未來預測", "特定單年詳細數據"], horizontal=True)
        
        if view_type == "多年趨勢與未來預測":
            f_df = get_forecast(sub_df)
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(sub_df['年度'].astype(str), sub_df['營收'], marker='o', label='歷史營收')
            ax.plot(f_df['年度'], f_df['營收'], '--', marker='s', color='gray', label='預測模型')
            ax.set_title("歷年營收軌跡與未來 3 年推估", fontproperties=font_prop)
            ax.legend(prop=font_prop)
            st.pyplot(fig)
            st.dataframe(sub_df)
        else:
            sel_year = st.selectbox("選擇年份", sub_df['年度'].unique())
            st.write(sub_df[sub_df['年度'] == sel_year])

    elif "單一公司 vs 其他公司對比" in analysis_mode:
        st.header("同業競爭力與風險對比")
        target_co = st.selectbox("主選公司", df['公司名稱'].unique())
        comp_year = st.selectbox("比較基準年度", sorted(df['年度'].unique(), reverse=True))
        
        comp_df = df[df['年度'] == comp_year]
        
        fig2, ax2 = plt.subplots(figsize=(10, 5))
        colors = ['red' if c == target_co else 'teal' for c in comp_df['公司名稱']]
        ax2.bar(comp_df['公司名稱'], comp_df['M分數'], color=colors)
        ax2.axhline(y=-1.78, color='black', linestyle='--')
        ax2.set_title(f"{comp_year} 年度：{target_co} 與同業之風險分佈 (紅色為主選)", fontproperties=font_prop)
        st.pyplot(fig2)

    else: # 多公司多年橫向評比
        st.header("全體標的多年風險掃描")
        st.write("自動偵測所有上傳標的中具備舞弊嫌疑之年度數據：")
        risk_summary = df[df['M分數'] > -1.78].sort_values(['公司名稱', '年度'])
        st.dataframe(risk_summary)
        
        # 散佈圖分析
        fig3, ax3 = plt.subplots()
        scatter = ax3.scatter(df['營收'], df['應收帳款'], c=df['M分數'], cmap='Reds', s=100)
        plt.colorbar(scatter, label='M-Score Risk')
        ax3.set_xlabel("營收規模", fontproperties=font_prop)
        ax3.set_ylabel("應收帳款", fontproperties=font_prop)
        ax3.set_title("全體標的風險分佈散佈圖", fontproperties=font_prop)
        st.pyplot(fig3)

    # --- 5. 報告導出模組 ---
    st.divider()
    st.caption("法律聲明：本報告包含鑑定數據與財務預測。鑑定結論之最終法律效力以會師簽證之紙本報告為準。")
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        out_ex = io.BytesIO()
        df.to_excel(out_ex, index=False)
        st.download_button("下載整合鑑定底稿 (Excel)", out_ex.getvalue(), "聚合鑑底稿.xlsx")
    with
