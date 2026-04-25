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

# --- 2. 核心鑑識邏輯 ---
def forensic_engine(row):
    r, rc, inv = row.get('營收', 0), row.get('應收帳款', 0), row.get('存貨', 0)
    m_score = -3.2 + (0.15 * (rc/r if r>0 else 0)) + (0.1 * (inv/r if r>0 else 0))
    status = "財報舞弊高風險" if m_score > -1.78 else "經營狀態尚屬穩健"
    return pd.Series([round(m_score, 2), status])

# --- 3. 介面與多檔案上傳功能 ---
st.set_page_config(layout="wide", page_title="玄武多檔案鑑識聚合系統")

with st.sidebar:
    st.header("數據採集設定")
    auditor = st.text_input("主辦會計師", "張鈞翔會計師")
    firm = st.text_input("事務所名稱", "玄武聯合會計師事務所")
    
    st.divider()
    # 關鍵功能：多檔案上傳
    uploaded_files = st.file_uploader(
        "請上傳多份公司數據底稿 (Excel)", 
        type=["xlsx"], 
        accept_multiple_files=True
    )

# --- 4. 數據聚合處理 ---
if uploaded_files:
    all_data_list = []
    
    for f in uploaded_files:
        try:
            temp_df = pd.read_excel(f)
            # 標記來源檔案名稱，方便追蹤數據來自哪間公司
            if '公司名稱' not in temp_df.columns:
                temp_df['公司名稱'] = f.name.replace(".xlsx", "")
            all_data_list.append(temp_df)
        except Exception as e:
            st.error(f"檔案 {f.name} 讀取失敗: {e}")

    if all_data_list:
        # 將所有上傳的檔案「合併」成一個 master dataframe
        df = pd.concat(all_data_list, ignore_index=True)
        df[['M分數', '鑑定結論']] = df.apply(forensic_engine, axis=1)

        st.success(f"成功聚合 {len(uploaded_files)} 份檔案，共計 {len(df)} 筆年度數據。")

        # --- 5. 多維度動態分析選單 ---
        analysis_mode = st.radio(
            "請選擇分析視角", 
            ["全體公司橫向評比", "特定公司歷年深度診斷", "同產業對比分析"]
        )

        if analysis_mode == "全體公司橫向評比":
            st.subheader("跨公司舞弊風險矩陣")
            fig1, ax1 = plt.subplots(figsize=(12, 5))
            df_latest = df.sort_values('年度').groupby('公司名稱').tail(1)
            ax1.bar(df_latest['公司名稱'], df_latest['M分數'], color='darkred')
            ax1.axhline(y=-1.78, color='black', linestyle='--')
            ax1.set_title("各公司最新年度風險指標對比", fontproperties=font_prop)
            st.pyplot(fig1)

        elif analysis_mode == "特定公司歷年深度診斷":
            target = st.selectbox("選擇要深入調查的公司", df['公司名稱'].unique())
            target_df = df[df['公司名稱'] == target].sort_values('年度')
            st.line_chart(target_df.set_index('年度')[['營收', '應收帳款']])
            st.write(f"### {target} 鑑定結論：{target_df['鑑定結論'].iloc[-1]}")

        # --- 6. 報告產出與法律聲明 ---
        st.divider()
        st.caption("法律聲明：本報告係由多源數據聚合模組自動產出，僅供專業審計參考。")

        col1, col2 = st.columns(2)
        with col1:
            output_ex = io.BytesIO()
            df.to_excel(output_ex, index=False)
            st.download_button("下載聚合鑑定底稿 (Excel)", output_ex.getvalue(), "聚合底稿.xlsx")
        
        with col2:
            if st.button("產生綜合鑑定 Word 報告"):
                doc = Document()
                doc.add_heading("多源數據鑑識聚合鑑定報告", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph(f"主辦會計師：{auditor}\n事務所：{firm}")
                
                doc.add_heading("一、 重大風險警告名單", level=1)
                risk_df = df[df['鑑定結論'] != "經營狀態尚屬穩健"]
                for _, r in risk_df.iterrows():
                    doc.add_paragraph(f"標的：{r['公司名稱']} ({r['年度']}年) - 結論：{r['鑑定結論']}")

                buf = io.BytesIO()
                doc.save(buf)
                st.download_button("點此下載 Word 報告", buf.getvalue(), "綜合鑑定報告.docx")

else:
    st.info("系統就緒。請上傳一個或多個 Excel 檔案以開始聚合分析。")
