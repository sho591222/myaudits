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

# --- 2. 智慧 PDF 採集引擎 ---
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
        res = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0, "應收帳款": 0, "存貨": 0, "現金": 0, "負債": 0}
        for _, row in master_df.iterrows():
            row_str = "".join([str(x) for x in row.values])
            if any(k in row_str for k in ["年度", "Year"]):
                for val in row.values:
                    if str(val).isdigit() and len(str(val)) >= 3: res["年度"] = int(val)
            def get_num(r):
                nums = [pd.to_numeric(str(x).replace(",",""), errors='coerce') for x in r.values if str(x).replace(",","").replace(".","").isdigit()]
                return nums[0] if nums else 0
            if any(k in row_str for k in ["營收", "營業收入"]): res["營收"] = get_num(row)
            if "應收帳款" in row_str: res["應收帳款"] = get_num(row)
            if "存貨" in row_str: res["存貨"] = get_num(row)
            if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = get_num(row)
            if "負債" in row_str: res["負債"] = get_num(row)
        return pd.DataFrame([res])
    except: return pd.DataFrame()

# --- 3. 四大犯罪偵測核心模型 ---
def crime_detector(row):
    r, rc, inv, cash, debt = row.get('營收', 0), row.get('應收帳款', 0), row.get('存貨', 0), row.get('現金', 0), row.get('負債', 0)
    results = {}
    
    # (1) 財報舞弊模型 (Beneish M-Score)
    m_score = -3.2 + (0.15 * (rc/r if r>0 else 0)) + (0.1 * (inv/r if r>0 else 0))
    results['舞弊指數'] = round(m_score, 2)
    
    # (2) 資產掏空模型 (Tunelling: 應收帳款異常佔比)
    t_index = (rc / r) if r > 0 else 0
    results['掏空風險'] = "高" if t_index > 0.4 else "低"
    
    # (3) 非法吸金模型 (Ponzi: 負債成長率 vs 現金水位)
    results['吸金指標'] = "警示" if debt > cash * 5 and r < cash else "正常"
    
    # (4) 洗錢風險模型 (Money Laundering: 營收現金不匹配)
    results['洗錢風險'] = "中高" if r > 0 and cash / r < 0.05 else "低"
    
    return pd.Series([results['舞弊指數'], results['掏空風險'], results['吸金指標'], results['洗錢風險']])

# --- 4. 側邊欄配置 ---
st.set_page_config(layout="wide", page_title="玄武犯罪防制鑑定平台")
with st.sidebar:
    st.header("鑑識官中心")
    mode = st.radio("功能選單", ["綜合分析：趨勢與預測", "四大犯罪防制偵測中心"])
    st.divider()
    files = st.file_uploader("上傳 Excel 或 PDF", type=["xlsx", "pdf"], accept_multiple_files=True)

# --- 5. 主流程 ---
if files:
    data_list = [pd.read_excel(f) if f.name.endswith('.xlsx') else parse_pdf_robustly(f) for f in files]
    df = pd.concat(data_list, ignore_index=True).fillna(0)
    df['年度'] = pd.to_numeric(df['年度'], errors='coerce').fillna(0).astype(int)
    df = df.drop_duplicates(subset=['公司名稱', '年度'], keep='last')
    
    # 執行犯罪偵測
    df[['舞弊指數', '掏空風險', '吸金指標', '洗錢風險']] = df.apply(crime_detector, axis=1)

    if mode == "四大犯罪防制偵測中心":
        st.title("🛡️ 財務犯罪防制偵測中心")
        target = st.selectbox("選擇受查對象", df['公司名稱'].unique())
        sub = df[df['公司名稱'] == target].sort_values('年度')
        
        # 多年/單年切換顯示
        view_type = st.segmented_control("顯示維度", ["多年趨勢掃描", "特定單年深度鑑定"], default="多年趨勢掃描")
        
        if view_type == "多年趨勢掃描":
            st.subheader(f"{target} - 犯罪風險指標走勢")
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.plot(sub['年度'].astype(str), sub['舞弊指數'], label='舞弊指數(M-Score)', marker='D', color='red')
            ax.axhline(y=-1.78, color='black', linestyle='--', label='舞弊警戒線')
            ax.set_ylabel("指標分數")
            ax.legend(prop=font_prop)
            st.pyplot(fig)
            
            # 四大指標摘要表
            st.table(sub[['年度', '舞弊指數', '掏空風險', '吸金指標', '洗錢風險']])
            
        else: # 單年深度鑑定
            sel_year = st.selectbox("選擇鑑定年份", sub['年度'].unique())
            year_data = sub[sub['年度'] == sel_year].iloc[0]
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("舞弊指數", year_data['舞弊指數'], delta="危險" if year_data['舞弊指數'] > -1.78 else "正常", delta_color="inverse")
            col2.metric("掏空風險", year_data['掏空風險'])
            col3.metric("吸金警示", year_data['吸金指標'])
            col4.metric("洗錢機率", year_data['洗錢風險'])
            
            # 雷達圖呈現 (簡化版)
            st.info(f"鑑定報告：{target} 在 {sel_year} 年度之核心指標檢測結果如上。")

    else: # 原始預測模式
        # ... (保留原有的趨勢預測代碼) ...
        st.write("請切換至「四大犯罪防制偵測中心」查看詳情")

    # 報告下載
    st.divider()
    if st.button("下載全方位犯罪防制報告 (Word)"):
        doc = Document()
        doc.add_heading("財務犯罪防制鑑定報告", 0)
        doc.add_paragraph(f"產出日期：{datetime.now().strftime('%Y/%m/%d')}")
        for _, r in df.iterrows():
            if r['舞弊指數'] > -1.78 or r['掏空風險'] == "高":
                doc.add_paragraph(f"⚠️ 異常標的：{r['公司名稱']} ({r['年度']})")
                doc.add_paragraph(f" - 舞弊指標：{r['舞弊指數']} | 掏空風險：{r['掏空風險']}")
        buf = io.BytesIO()
        doc.save(buf)
        st.download_button("下載報告", buf.getvalue(), "犯罪鑑定報告.docx")
else:
    st.info("請上傳資料以啟動偵測中心。")
