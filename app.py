import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
import io
import matplotlib.font_manager as fm
import os
import requests
from datetime import datetime
import pdfplumber

# --- 1. 環境設定：中文字體 (解決圖表亂碼) ---
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

# --- 2. 強化版 PDF 解析引擎 (解決 Column 不匹配與欄位偏移) ---
def parse_pdf_robustly(file):
    try:
        extracted_data = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if len(table) < 2: continue
                    df_tmp = pd.DataFrame(table)
                    df_tmp = df_tmp.dropna(how='all').dropna(axis=1, how='all')
                    extracted_data.append(df_tmp)
        
        if not extracted_data: return pd.DataFrame()
        full_pdf_df = pd.concat(extracted_data, ignore_index=True)
        
        # 初始化標準列
        final_row = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0, "應收帳款": 0, "存貨": 0}
        
        for _, row in full_pdf_df.iterrows():
            row_str = "".join([str(x) for x in row.values])
            # 搜尋關鍵字並提取該行中的第一個數字
            if any(k in row_str for k in ["年度", "Year"]):
                for val in row.values:
                    if str(val).isdigit(): final_row["年度"] = int(val)
            if any(k in row_str for k in ["營收", "營業收入"]):
                nums = [pd.to_numeric(str(x).replace(",",""), errors='coerce') for x in row.values if str(x).replace(",","").replace(".","").isdigit()]
                if nums: final_row["營收"] = nums[0]
            if "應收帳款" in row_str:
                nums = [pd.to_numeric(str(x).replace(",",""), errors='coerce') for x in row.values if str(x).replace(",","").replace(".","").isdigit()]
                if nums: final_row["應收帳款"] = nums[0]
            if "存貨" in row_str:
                nums = [pd.to_numeric(str(x).replace(",",""), errors='coerce') for x in row.values if str(x).replace(",","").replace(".","").isdigit()]
                if nums: final_row["存貨"] = nums[0]
        
        return pd.DataFrame([final_row])
    except Exception as e:
        st.warning(f"檔案 {file.name} 解析失敗: {e}")
        return pd.DataFrame()

# --- 3. 鑑識與預測運算核心 ---
def forensic_engine(row):
    r, rc, inv = row.get('營收', 0), row.get('應收帳款', 0), row.get('存貨', 0)
    if r == 0: return pd.Series([0, "數據不足"])
    # 簡化 Beneish M-Score 模型
    m_score = -3.2 + (0.15 * (rc/r)) + (0.1 * (inv/r))
    status = "風險預警" if m_score > -1.78 else "經營穩健"
    return pd.Series([round(m_score, 2), status])

def get_forecast(df, years=3):
    if len(df) < 2: return pd.DataFrame()
    df = df.sort_values('年度')
    growth = df['營收'].pct_change().mean()
    curr_rev = df['營收'].iloc[-1]
    last_year = df['年度'].iloc[-1]
    f_list = []
    for i in range(1, years + 1):
        curr_rev *= (1 + (growth if not np.isnan(growth) else 0))
        f_list.append({'年度': f"{int(last_year)+i}(預測)", '營收': round(curr_rev, 2)})
    return pd.DataFrame(f_list)

# --- 4. 側邊欄與介面配置 ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計旗艦平台")

with st.sidebar:
    st.header("功能導航中心")
    analysis_mode = st.radio(
        "選擇分析模式", 
        ["單一公司：趨勢與預測", "單一公司：單年診斷", "對比：單一 vs 同業", "掃描：多公司多年"]
    )
    st.divider()
    auditor = st.text_input("主辦會計師", "張鈞翔會計師")
    files = st.file_uploader("批次上傳數據 (Excel/PDF)", type=["xlsx", "pdf"], accept_multiple_files=True)

# --- 5. 數據處理 ---
if files:
    all_data = []
    for f in files:
        if f.name.endswith('.xlsx'):
            tmp = pd.read_excel(f)
            if '公司名稱' not in tmp.columns: tmp['公司名稱'] = f.name.replace(".xlsx", "")
            all_data.append(tmp)
        elif f.name.endswith('.pdf'):
            tmp = parse_pdf_robustly(f)
            all_data.append(tmp)
    
    if all_data:
        df = pd.concat(all_data, ignore_index=True).fillna(0)
        # 確保數值欄位正確
        for col in ['營收', '應收帳款', '存貨', '年度']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        df[['M分數', '結論']] = df.apply(forensic_engine, axis=1)

        # --- 介面呈現與邏輯分界 ---
        if "趨勢與預測" in analysis_mode:
            target = st.selectbox("選擇調查標的", df['公司名稱'].unique())
            sub = df[df['公司名稱'] == target].sort_values('年度')
            st.header(f"{target} 鑑定與成長預測")
            
            f_df = get_forecast(sub)
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='歷史實際')
            if not f_df.empty:
                ax.plot(f_df['年度'], f_df['營收'], '--', marker='s', color='gray', label='預測趨勢')
            ax.set_title("營收歷史軌跡與未來推估", fontproperties=font_prop)
            ax.legend(prop=font_prop)
            st.pyplot(fig)
            st.dataframe(sub)

        elif "單一 vs 同業" in analysis_mode:
            year = st.selectbox("選擇比較年度", sorted(df['年度'].unique(), reverse=True))
            target = st.selectbox("主選公司 (標色)", df['公司名稱'].unique())
            year_df = df[df['年度'] == year]
            st.header(f"{year} 年度同業風險評比")
            
            fig2, ax2 = plt.subplots(figsize=(10, 5))
            colors = ['red' if c == target else 'skyblue' for c in year_df['公司名稱']]
            ax2.bar(year_df['公司名稱'], year_df['M分數'], color=colors)
            ax2.axhline(y=-1.78, color='black', linestyle='--')
            ax2.set_title("舞弊指標分佈 (紅色為主選標的)", fontproperties=font_prop)
            st.pyplot(fig2)

        # --- 報告生成 ---
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            out_ex = io.BytesIO()
            df.to_excel(out_ex, index=False)
            st.download_button("匯出整合底稿 (Excel)", out_ex.getvalue(), "鑑定底稿.xlsx")
        with c2:
            if st.button("生成 Word 鑑定報告"):
                doc = Document()
                doc.add_heading("鑑識會計鑑定報告書", 0)
                doc.add_paragraph(f"主辦會計師：{auditor}\n產出日期：{datetime.now().strftime('%Y/%m/%d')}")
                doc.add_heading("一、 重大風險異常摘要", level=1)
                for _, r in df[df['M分數'] > -1.78].iterrows():
                    doc.add_paragraph(f"公司：{r['公司名稱']} ({r['年度']}) - 指標：{r['M分數']}")
                buf_word = io.BytesIO()
                doc.save(buf_word)
                st.download_button("下載 Word 報告", buf_word.getvalue(), "鑑定報告.docx")
else:
    st.info("系統就緒。請上傳 Excel 或 PDF 數據檔案。")
