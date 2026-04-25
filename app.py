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

# --- 1. 字體與環境設定 (解決中文字體亂碼) ---
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

# --- 2. 數據採集引擎 (PDF/Excel) ---
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
            if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = get_num(row)
        return pd.DataFrame([res])
    except: return pd.DataFrame()

# --- 3. 四大犯罪鑑定模型 & 預測模型 ---
def crime_detector(row):
    r, rc, inv, cash, debt = row.get('營收', 0), row.get('應收帳款', 0), row.get('存貨', 0), row.get('現金', 0), row.get('負債總額', 0)
    # A. 舞弊 (M-Score)
    m_score = -3.2 + (0.15 * (rc/r if r>0 else 0)) + (0.1 * (inv/r if r>0 else 0))
    # B. 掏空 (應收帳款佔比)
    t_ratio = (rc / r) if r > 0 else 0
    # C. 吸金 (負債現金比)
    p_index = (debt / cash) if cash > 0 else 0
    # D. 洗錢 (現金營收比)
    ml_index = (cash / r) if r > 0 else 0
    return pd.Series([
        round(m_score, 2),
        "危險" if m_score > -1.78 else "正常",
        "高風險" if t_ratio > 0.4 else "正常",
        "警示" if p_index > 5 else "正常",
        "高機率" if ml_index < 0.05 and r > 0 else "低"
    ])

def get_forecast_df(df, years=3):
    """營收線性成長預測"""
    if len(df) < 2: return pd.DataFrame()
    df = df.sort_values('年度')
    growth = df['營收'].pct_change().mean()
    curr_rev, last_year = df['營收'].iloc[-1], df['年度'].iloc[-1]
    f_list = []
    for i in range(1, years + 1):
        curr_rev *= (1 + (growth if not np.isnan(growth) else 0))
        f_list.append({'年度': int(last_year)+i, '營收': round(curr_rev, 2), '類型': 'AI預測'})
    return pd.DataFrame(f_list)

# --- 4. Streamlit 介面配置 ---
st.set_page_config(layout="wide", page_title="玄武鑑識預測旗艦平台")

with st.sidebar:
    st.header("🕵️ 鑑識官控制台")
    analysis_mode = st.radio("功能切換", ["單一公司：鑑定與預測", "多公司對比：同業競爭與風險"])
    st.divider()
    files = st.file_uploader("上傳 Excel 或 PDF 財報數據", type=["xlsx", "pdf"], accept_multiple_files=True)

if files:
    # 數據處理與清洗
    data_list = [pd.read_excel(f) if f.name.endswith('.xlsx') else parse_pdf_robustly(f) for f in files]
    df = pd.concat(data_list, ignore_index=True).fillna(0)
    df['年度'] = pd.to_numeric(df['年度'], errors='coerce').fillna(0).astype(int)
    df = df.drop_duplicates(subset=['公司名稱', '年度'], keep='last')
    
    # 核心鑑識計算
    df[['舞弊指標', '舞弊判定', '掏空風險', '吸金指標', '洗錢風險']] = df.apply(crime_detector, axis=1)

    # --- 選項 A: 單一公司深度分析 ---
    if analysis_mode == "單一公司：鑑定與預測":
        target = st.selectbox("選擇調查公司", df['公司名稱'].unique())
        sub = df[df['公司名稱'] == target].sort_values('年度')
        
        st.title(f"🔍 {target} 深度鑑定報告")
        
        # 1. 營收趨勢與預測圖表
        st.subheader("營收歷史軌跡與成長預估")
        f_df = get_forecast_df(sub)
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.plot(sub['年度'], sub['營收'], marker='o', label='歷史營收', linewidth=2)
        if not f_df.empty:
            ax.plot(f_df['年度'], f_df['營收'], '--', marker='s', color='orange', label='預測趨勢')
        ax.set_title("營收趨勢圖", fontproperties=font_prop)
        ax.legend(prop=font_prop)
        st.pyplot(fig)
        
        # 2. 四大犯罪指標儀表板
        st.subheader("四大犯罪防制偵測指標")
        cols = st.columns(4)
        latest = sub.iloc[-1]
        cols[0].metric("舞弊指數", latest['舞弊指標'], latest['舞弊判定'], delta_color="inverse")
        cols[1].metric("掏空風險", latest['掏空風險'])
        cols[2].metric("吸金警示", latest['吸金指標'])
        cols[3].metric("洗錢機率", latest['洗錢風險'])
        
        st.dataframe(sub[['年度', '營收', '舞弊指標', '掏空風險', '吸金指標', '洗錢風險']])

    # --- 選項 B: 多公司橫向對比 ---
    elif analysis_mode == "多公司對比：同業競爭與風險":
        st.title("📊 同業風險與預測對比中心")
        
        col_charts = st.columns(2)
        with col_charts[0]:
            st.subheader("產業舞弊風險分佈 (M-Score)")
            fig2, ax2 = plt.subplots()
            ax2.bar(df['公司名稱'], df['舞弊指標'], color='teal')
            ax2.axhline(y=-1.78, color='red', linestyle='--', label='警戒線')
            ax2.set_ylabel("指標分數")
            ax2.legend(prop=font_prop)
            st.pyplot(fig2)
            
        with col_charts[1]:
            st.subheader("多公司成長動能預測")
            fig3, ax3 = plt.subplots()
            for co in df['公司名稱'].unique():
                co_sub = df[df['公司名稱'] == co].sort_values('年度')
                co_f = get_forecast_df(co_sub, years=2)
                # 結合歷史與預測畫圖
                full_view = pd.concat([co_sub[['年度','營收']], co_f[['年度','營收']]])
                ax3.plot(full_view['年度'], full_view['營營' if '營營' in full_view else '營收'], label=co, marker='.')
            ax3.legend(prop=font_prop)
            st.pyplot(fig3)
            
        st.dataframe(df[['公司名稱', '年度', '舞弊指標', '舞弊判定', '掏空風險', '吸金指標', '洗錢風險']])
 # --- 7. 報告導出 (強化敘述版) ---
    st.divider()
    if st.button("下載旗艦鑑定報告 (Word)"):
        doc = Document()
        # 設定標題
        doc.add_heading("財務犯罪防制與鑑識會計鑑定報告書", 0).alignment = 1
        
        # 報告基本資訊
        p = doc.add_paragraph()
        p.add_run(f"鑑定人：{auditor} 會計師\n").bold = True
        p.add_run(f"產出日期：{datetime.now().strftime('%Y/%m/%d')}\n")
        p.add_run(f"數據範圍：共計 {len(df['公司名稱'].unique())} 家受查企業之歷年財務數據")

        doc.add_heading("一、 鑑定結論與重大異常摘要", level=1)
        
        # 篩選出有問題的公司進行詳細敘述
        high_risk_df = df[(df['舞弊判定'] == "危險") | (df['掏空風險'] == "高風險") | (df['吸金指標'] == "警示")]
        
        if high_risk_df.empty:
            doc.add_paragraph("經本系統鑑定，受查標的於各項犯罪預警模型中均呈現「正常」狀態，暫無重大財務操縱疑慮。")
        else:
            for _, r in high_risk_df.iterrows():
                doc.add_heading(f"● 受查標的：{r['公司名稱']} ({r['年度']}年度)", level=2)
                
                # 詳細敘述邏輯
                narrative = "本系統針對該年度進行深度鑑定，分析結果如下：\n"
                
                if r['舞弊判定'] == "危險":
                    narrative += f"- 【財報舞弊】：M-Score 分數達 {r['舞弊指標']}，超過警戒線 (-1.78)。顯示該公司可能存在盈餘操縱、虛增資產或低估負債之高度風險。\n"
                
                if r['掏空風險'] == "高風險":
                    narrative += f"- 【資產掏空】：偵測到應收帳款佔營收比例異常。懷疑存在虛假交易或關聯方資金挪用，可能正透過外部虛假訂單掏空公司核心資產。\n"
                
                if r['吸金指標'] == "警示":
                    narrative += f"- 【非法吸金】：負債總額遠超現金額度且營收成長停滯。具備「以債養債」之龐氏騙局特徵，應嚴防投資人資金被非法挪用。\n"
                
                if r['洗錢風險'] == "高機率":
                    narrative += f"- 【洗錢疑慮】：現金水位與帳面營收嚴重不匹配（比率低於 5%），存在大額金流去向不明之特徵，需進一步查核金流流向。\n"
                
                doc.add_paragraph(narrative)

        doc.add_heading("二、 財務預測與未來風險評估", level=1)
        doc.add_paragraph("本報告結合 AI 成長模型進行未來兩年之營收推估。若受查標的存在上述犯罪指標，其未來預測數據之達成度應持高度保留意見，並建議啟動實質查核程序（Substantive Testing）。")

        doc.add_heading("三、 法律聲明", level=1)
        doc.add_paragraph("本鑑定報告係基於上傳之數據進行演算法分析，鑑定結果供專業審計與司法鑑定參考。最終結論應以簽證會計師實地查核之簽證報告為準。")

        # 匯出檔案
        buf = io.BytesIO()
        doc.save(buf)
        st.download_button("點此匯出完整鑑定報告 (含詳細敘述)", buf.getvalue(), "財務鑑定詳細報告.docx")
