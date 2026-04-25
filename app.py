import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
import matplotlib.font_manager as fm
import os
import requests
from datetime import datetime
import pdfplumber

# --- 1. 字體設定 ---
@st.cache_resource
def load_chinese_font():
    font_url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
    font_path = "NotoSansCJKtc-Regular.otf"
    if not os.path.exists(font_path):
        try:
            response = requests.get(font_url)
            with open(font_path, "wb") as f: f.write(response.content)
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

# --- 2. 數據採集 ---
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
        res = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0, "應收帳款": 0, "存貨": 0, "現金": 0, "負債總額": 0}
        
        for _, row in master_df.iterrows():
            row_str = "".join([str(x) for x in row.values if x])
            for val in row.values:
                val_s = str(val).strip()
                if val_s.isdigit() and 1900 <= int(val_s) <= 2100: res["年度"] = int(val_s)
            
            def extract_num(r):
                nums = [pd.to_numeric(str(x).replace(",","").replace("(","-").replace(")",""), errors='coerce') for x in r.values if x]
                nums = [n for n in nums if not np.isnan(n)]
                return nums[0] if nums else 0

            if any(k in row_str for k in ["營收", "營業收入"]): res["營收"] = extract_num(row)
            if "應收帳款" in row_str: res["應收帳款"] = extract_num(row)
            if "存貨" in row_str: res["存貨"] = extract_num(row)
            if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = extract_num(row)
            if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = extract_num(row)
        return pd.DataFrame([res])
    except: return pd.DataFrame()

# --- 3. 預測與鑑定核心 (修正回傳長度) ---
def forensic_analyze(row):
    r = row.get('營收', 0)
    rc = row.get('應收帳款', 0)
    inv = row.get('存貨', 0)
    cash = row.get('現金', 0)
    debt = row.get('負債總額', 0)
    
    # 確保無論如何都回傳 4 個值
    if r <= 0: 
        return pd.Series([0.0, "數據不足", "正常", "正常"])
    
    m_score = -3.2 + (0.15 * (rc/r)) + (0.1 * (inv/r))
    t_risk = "高風險" if (rc/r) > 0.4 else "正常"
    p_risk = "警示" if cash > 0 and (debt/cash) > 5 else "正常"
    
    return pd.Series([round(m_score, 2), "危險" if m_score > -1.78 else "正常", t_risk, p_risk])

def get_forecast(df, years=2):
    if len(df) < 2: return pd.DataFrame()
    df = df.sort_values('年度')
    avg_growth = df['營收'].pct_change().mean()
    if np.isnan(avg_growth): avg_growth = 0
    last_year = int(df['年度'].iloc[-1])
    last_rev = df['營收'].iloc[-1]
    f_results = []
    for i in range(1, years + 1):
        last_rev *= (1 + avg_growth)
        f_results.append({'年度': last_year + i, '營收': round(last_rev, 2), '類型': '預測'})
    return pd.DataFrame(f_results)

# --- 4. 介面 ---
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財務鑑識與成長預測系統</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("事務所中心")
    st.text("Slogan: 玄武鑑定，真偽分明")
    mode = st.radio("功能選單", ["單一公司深度鑑定", "多公司競爭力PK"])
    auditor = st.text_input("主辦會計師", "會計師")
    files = st.file_uploader("批次上傳資料", type=["pdf", "xlsx"], accept_multiple_files=True)

if files:
    all_dfs = []
    for f in files:
        if f.name.endswith('.xlsx'):
            all_dfs.append(pd.read_excel(f))
        else:
            all_dfs.append(parse_pdf_robustly(f))
    
    if all_dfs:
        df = pd.concat(all_dfs, ignore_index=True)
        # 排除年度為 0 的資料
        df = df[df['年度'] > 0].copy()
        
        if not df.empty:
            df = df.drop_duplicates(subset=['公司名稱', '年度']).sort_values('年度')
            # 關鍵修正點：確保 apply 產出的 Series 長度永遠與欄位清單一致
            analysis_res = df.apply(forensic_analyze, axis=1)
            df[['M分數', '舞弊狀態', '掏空風險', '吸金指標']] = analysis_res

            if mode == "單一公司深度鑑定":
                target = st.selectbox("選擇受查對象", df['公司名稱'].unique())
                sub = df[df['公司名稱'] == target]
                
                st.subheader("營收歷史軌跡與成長預估")
                f_df = get_forecast(sub)
                fig, ax = plt.subplots(figsize=(12, 5))
                ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='歷史實際營收', linewidth=3, color='#268bd2')
                if not f_df.empty:
                    ax.plot(f_df['年度'].astype(str), f_df['營收'], '--', marker='s', label='AI 成長預測', color='#cb4b16')
                ax.set_title(f"{target} 財務動能分析", fontproperties=font_prop)
                ax.legend(prop=font_prop)
                st.pyplot(fig)
                
                c1, c2, c3 = st.columns(3)
                latest = sub.iloc[-1]
                c1.metric("舞弊指標 (M-Score)", latest['M分數'], latest['舞弊狀態'], delta_color="inverse")
                c2.metric("資產掏空風險", latest['掏空風險'])
                c3.metric("非法吸金警示", latest['吸金指標'])
                st.dataframe(sub)

            elif mode == "多公司競爭力PK":
                st.subheader("產業風險與成長 PK")
                col1, col2 = st.columns(2)
                with col1:
                    fig2, ax2 = plt.subplots()
                    ax2.bar(df['公司名稱'], df['M分數'], color='#2aa198')
                    ax2.axhline(y=-1.78, color='red', linestyle='--', label='警戒線')
                    ax2.set_title("跨公司舞弊指標評比", fontproperties=font_prop)
                    st.pyplot(fig2)
                with col2:
                    fig3, ax3 = plt.subplots()
                    for co in df['公司名稱'].unique():
                        co_d = df[df['公司名稱'] == co]
                        f_co = get_forecast(co_d)
                        full_view = pd.concat([co_d[['年度','營收']], f_co[['年度','營收']]])
                        ax3.plot(full_view['年度'].astype(str), full_view['營營' if '營營' in full_view else '營收'], label=co, marker='.')
                    ax3.set_title("成長動能對比線", fontproperties=font_prop)
                    ax3.legend(prop=font_prop)
                    st.pyplot(fig3)

            if st.button("匯出正式鑑定報告"):
                doc = Document()
                doc.add_heading("玄武會計師事務所 - 財務鑑定報告書", 0)
                doc.add_paragraph(f"主辦會計師：{auditor}")

                doc.add_heading("一、 營運成長與未來預測敘述", level=1)
                for co in df['公司名稱'].unique():
                    co_sub = df[df['公司名稱'] == co]
                    if len(co_sub) >= 2:
                        growth = co_sub['營收'].pct_change().mean()
                        f_data = get_forecast(co_sub)
                        doc.add_paragraph(f"受查對象 {co} 平均年增率為 {growth:.2%}。未來營收預計可達 {f_data['營收'].iloc[0]:,.0f} 元。")
                
                buf = io.BytesIO()
                doc.save(buf)
                st.download_button("下載報告檔案", buf.getvalue(), "鑑定報告.docx")
        else:
            st.warning("未能從上傳檔案中提取到有效的年度數據。")
else:
    st.info("系統就緒。請上傳兩年份以上的財報檔案以啟動分析。")
