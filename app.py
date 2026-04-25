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
import pdfplumber

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

# --- 2. 核心：智慧 PDF 數據採集 ---
def parse_pdf_robustly(file):
    try:
        raw_tables = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if len(table) < 2: continue
                    df_tmp = pd.DataFrame(table).dropna(how='all').dropna(axis=1, how='all')
                    raw_tables.append(df_tmp)
        
        if not raw_tables: return pd.DataFrame()
        master_df = pd.concat(raw_tables, ignore_index=True)
        
        res = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0, "應收帳款": 0, "存貨": 0}
        for _, row in master_df.iterrows():
            row_str = "".join([str(x) for x in row.values])
            # 年度辨識
            if any(k in row_str for k in ["年度", "Year"]):
                for val in row.values:
                    if str(val).isdigit() and len(str(val)) >= 3: res["年度"] = int(val)
            # 數值辨識
            def get_first_num(r):
                nums = [pd.to_numeric(str(x).replace(",",""), errors='coerce') for x in r.values if str(x).replace(",","").replace(".","").isdigit()]
                return nums[0] if nums else 0

            if any(k in row_str for k in ["營收", "營業收入", "Revenue"]): res["營營"] = get_first_num(row)
            if "應收帳款" in row_str: res["應收帳款"] = get_first_num(row)
            if "存貨" in row_str: res["存貨"] = get_first_num(row)
        
        # 修正欄位名稱映射
        if "營營" in res: res["營收"] = res.pop("營營")
        return pd.DataFrame([res])
    except Exception as e:
        st.warning(f"檔案 {file.name} 解析受阻: {e}")
        return pd.DataFrame()

# --- 3. 核心：鑑識與預測模型 ---
def forensic_engine(row):
    r, rc, inv = row.get('營收', 0), row.get('應收帳款', 0), row.get('存貨', 0)
    if r <= 0: return pd.Series([0, "數據不足"])
    m_score = -3.2 + (0.15 * (rc/r)) + (0.1 * (inv/r))
    status = "風險預警" if m_score > -1.78 else "經營穩健"
    return pd.Series([round(m_score, 2), status])

def get_forecast(df, years=3):
    if len(df) < 2: return pd.DataFrame()
    df = df.sort_values('年度')
    growth = df['營收'].pct_change().mean()
    curr_rev, last_year = df['營收'].iloc[-1], df['年度'].iloc[-1]
    f_list = []
    for i in range(1, years + 1):
        curr_rev *= (1 + (growth if not np.isnan(growth) else 0))
        f_list.append({'年度': f"{int(last_year)+i}(預測)", '營收': round(curr_rev, 2), '類型': '預測'})
    return pd.DataFrame(f_list)

# --- 4. 側邊欄與功能切換 ---
st.set_page_config(layout="wide", page_title="玄武鑑識會計平台")

with st.sidebar:
    st.header("鑑定管理中心")
    auditor = st.text_input("主辦會計師", "張鈞翔會計師")
    mode = st.radio("功能選單", ["單一公司：趨勢與預測", "對比模式：單一 vs 同業", "掃描：多公司多年風險"])
    st.divider()
    files = st.file_uploader("上傳 Excel 或 PDF", type=["xlsx", "pdf"], accept_multiple_files=True)

# --- 5. 數據處理與顯示 ---
if files:
    data_list = []
    for f in files:
        if f.name.endswith('.xlsx'):
            tmp = pd.read_excel(f)
            if '公司名稱' not in tmp.columns: tmp['公司名稱'] = f.name.replace(".xlsx", "")
            data_list.append(tmp)
        else:
            data_list.append(parse_pdf_robustly(f))
    
    if data_list:
        df = pd.concat(data_list, ignore_index=True).fillna(0)
        # 數據清洗：確保年度為整數並移除重複
        df['年度'] = pd.to_numeric(df['年度'], errors='coerce').fillna(0).astype(int)
        df = df.drop_duplicates(subset=['公司名稱', '年度'], keep='last')
        
        df[['M分數', '結論']] = df.apply(forensic_engine, axis=1)

        st.title("專業鑑識分析看板")
        
        if "趨勢與預測" in mode:
            target = st.selectbox("選擇調查對象", df['公司名稱'].unique())
            sub = df[df['公司名稱'] == target].sort_values('年度')
            st.header(f"{target} 歷史鑑定軌跡與成長預估")
            
            f_df = get_forecast(sub)
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(sub['年度'].astype(str), sub['營收'], label='實際營收', marker='o')
            if not f_df.empty:
                ax.plot(f_df['年度'], f_df['營收'], '--', label='模型預測', marker='s', color='gray')
            ax.set_title("營收歷史與未來預測圖", fontproperties=font_prop)
            ax.legend(prop=font_prop)
            st.pyplot(fig)
            st.dataframe(sub)

        elif "單一 vs 同業" in mode:
            year = st.selectbox("比較基準年度", sorted(df['年度'].unique(), reverse=True))
            target = st.selectbox("主選公司", df['公司名稱'].unique())
            year_df = df[df['年度'] == year]
            st.header(f"{year} 年度產業風險對照")
            fig2, ax2 = plt.subplots(figsize=(10, 5))
            colors = ['red' if c == target else 'skyblue' for c in year_df['公司名稱']]
            ax2.bar(year_df['公司名稱'], year_df['M分數'], color=colors)
            ax2.axhline(y=-1.78, color='black', linestyle='--')
            ax2.set_title("產業 M-Score 分佈圖 (紅色為主選標的)", fontproperties=font_prop)
            st.pyplot(fig2)

        # --- 6. 報告下載區 ---
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            out_ex = io.BytesIO()
            df.to_excel(out_ex, index=False)
            st.download_button("下載整合底稿 (Excel)", out_ex.getvalue(), "鑑定底稿.xlsx")
        with c2:
            if st.button("生成鑑定報告 (Word)"):
                doc = Document()
                doc.add_heading("鑑識會計鑑定報告書", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph(f"主辦會計師：{auditor}\n報告產出日期：{datetime.now().strftime('%Y/%m/%d')}")
                doc.add_heading("一、 異常預警名單", level=1)
                for _, r in df[df['M分數'] > -1.78].iterrows():
                    doc.add_paragraph(f"標的：{r['公司名稱']} ({r['年度']}) - M分數：{r['M分數']} ({r['結論']})")
                buf_word = io.BytesIO()
                doc.save(buf_word)
                st.download_button("下載 Word 報告書", buf_word.getvalue(), "鑑定報告.docx")
else:
    st.info("系統就緒。請上傳檔案以啟動大數據鑑識引擎。")
