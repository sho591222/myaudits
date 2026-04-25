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

# --- 1. 字體系統重整 ---
@st.cache_resource
def load_chinese_font():
    font_url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
    font_path = "NotoSansCJKtc-Regular.otf"
    if not os.path.exists(font_path):
        try:
            response = requests.get(font_url)
            if response.status_code == 200:
                with open(font_path, "wb") as f:
                    f.write(response.content)
            else:
                return None
        except:
            return None
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

# --- 2. 數據採集引擎 ---
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
                if val_s.isdigit() and 1900 <= int(val_s) <= 2100:
                    res["年度"] = int(val_s)
            
            def extract_num(r):
                # 處理括號負數與千分位
                nums = [pd.to_numeric(str(x).replace(",","").replace("(","-").replace(")",""), errors='coerce') for x in r.values if x]
                nums = [n for n in nums if not np.isnan(n)]
                return nums[0] if nums else 0

            if any(k in row_str for k in ["營收", "營業收入"]): res["營收"] = extract_num(row)
            if "應收帳款" in row_str: res["應收帳款"] = extract_num(row)
            if "存貨" in row_str: res["存貨"] = extract_num(row)
            if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = extract_num(row)
            if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = extract_num(row)
        return pd.DataFrame([res])
    except:
        return pd.DataFrame()

# --- 3. 鑑識核心 (嚴格限制回傳結構) ---
def forensic_analyze(row):
    # 初始化回傳 Series，確保長度永遠為 4
    results = pd.Series([0.0, "數據不足", "正常", "正常"], index=['M分數', '舞弊狀態', '掏空風險', '吸金指標'])
    
    try:
        r = float(row.get('營收', 0))
        rc = float(row.get('應收帳款', 0))
        inv = float(row.get('存貨', 0))
        cash = float(row.get('現金', 0))
        debt = float(row.get('負債總額', 0))
        
        if r <= 0: return results
        
        m_score = -3.2 + (0.15 * (rc/r)) + (0.1 * (inv/r))
        results['M分數'] = round(m_score, 2)
        results['舞弊狀態'] = "危險" if m_score > -1.78 else "正常"
        results['掏空風險'] = "高風險" if (rc/r) > 0.4 else "正常"
        results['吸金指標'] = "警示" if cash > 0 and (debt/cash) > 5 else "正常"
    except:
        pass
        
    return results

def get_forecast(df, years=2):
    if len(df) < 2: return pd.DataFrame()
    df_sorted = df.sort_values('年度')
    avg_growth = df_sorted['營收'].pct_change().mean()
    if np.isnan(avg_growth) or np.isinf(avg_growth): avg_growth = 0
    
    last_year = int(df_sorted['年度'].iloc[-1])
    last_rev = float(df_sorted['營收'].iloc[-1])
    
    f_list = []
    for i in range(1, years + 1):
        last_rev *= (1 + avg_growth)
        f_list.append({'年度': last_year + i, '營營' if '營營' in df else '營收': round(last_rev, 2), '類型': '預測'})
    return pd.DataFrame(f_list)

# --- 4. 使用者介面 ---
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

# 移除表情符號的標題
st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財務鑑識與成長預測系統</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("事務所控制台")
    st.text("品牌口號: 玄武鑑定，真偽分明")
    mode = st.radio("模式切換", ["單一標的深度分析", "多標的風險對比"])
    auditor_name = st.text_input("主辦會計師簽署", "張鈞翔會計師")
    st.divider()
    uploaded_files = st.file_uploader("批次上傳數據檔案", type=["pdf", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    data_frames = []
    for f in uploaded_files:
        if f.name.endswith('.xlsx'):
            data_frames.append(pd.read_excel(f))
        else:
            data_frames.append(parse_pdf_robustly(f))
    
    if data_frames:
        main_df = pd.concat(data_frames, ignore_index=True)
        main_df = main_df[main_df['年度'] > 0].drop_duplicates(subset=['公司名稱', '年度']).sort_values('年度').copy()
        
        if not main_df.empty:
            # 修正後的 apply 調用
            analysis_data = main_df.apply(forensic_analyze, axis=1)
            main_df[['M分數', '舞弊狀態', '掏空風險', '吸金指標']] = analysis_data

            if mode == "單一標的深度分析":
                target_company = st.selectbox("選擇受查公司", main_df['公司名稱'].unique())
                sub_df = main_df[main_df['公司名稱'] == target_company]
                
                st.subheader("營收歷史軌跡與成長預估趨勢圖")
                forecast_df = get_forecast(sub_df)
                
                fig, ax = plt.subplots(figsize=(12, 5))
                ax.plot(sub_df['年度'].astype(str), sub_df['營收'], marker='o', label='歷史實際值', linewidth=3, color='#268bd2')
                if not forecast_df.empty:
                    ax.plot(forecast_df['年度'].astype(str), forecast_df['營收'], '--', marker='s', label='AI 預測值', color='#cb4b16')
                
                ax.set_title(f"{target_company} 財務動能與預測圖表", fontproperties=font_prop)
                ax.legend(prop=font_prop)
                st.pyplot(fig)
                
                # 指標看板
                stat_cols = st.columns(3)
                latest_data = sub_df.iloc[-1]
                stat_cols[0].metric("舞弊指標 (M-Score)", latest_data['M分數'], latest_data['舞弊狀態'], delta_color="inverse")
                stat_cols[1].metric("資產掏空風險", latest_data['掏空風險'])
                stat_cols[2].metric("非法吸金警示", latest_data['吸金指標'])
                
                st.write("詳細數據列表")
                st.dataframe(sub_df)

            elif mode == "多標的風險對比":
                st.subheader("跨公司風險與成長 PK 視覺化")
                comp_cols = st.columns(2)
                with comp_cols[0]:
                    fig2, ax2 = plt.subplots()
                    ax2.bar(main_df['公司名稱'], main_df['M分數'], color='#2aa198')
                    ax2.axhline(y=-1.78, color='red', linestyle='--', label='舞弊警戒線')
                    ax2.set_title("各標的舞弊風險評比", fontproperties=font_prop)
                    st.pyplot(fig2)
                with comp_cols[1]:
                    fig3, ax3 = plt.subplots()
                    for company in main_df['公司名稱'].unique():
                        c_data = main_df[main_df['公司名稱'] == company]
                        f_data = get_forecast(c_data)
                        combined = pd.concat([c_data[['年度','營收']], f_data[['年度','營收']]])
                        ax3.plot(combined['年度'].astype(str), combined['營收'], label=company, marker='.')
                    ax3.set_title("產業成長動能 PK 線", fontproperties=font_prop)
                    ax3.legend(prop=font_prop)
                    st.pyplot(fig3)

            # 報告匯出
            if st.button("點此產出正式鑑定報告檔案"):
                report = Document()
                report.add_heading("玄武會計師事務所 財務鑑定報告書", 0)
                report.add_paragraph(f"主辦會計師：{auditor_name}")
                report.add_paragraph(f"產出日期：{datetime.now().strftime('%Y-%m-%d')}")

                report.add_heading("一、 營運成長與未來預測深度敘述", level=1)
                for company in main_df['公司名稱'].unique():
                    c_sub = main_df[main_df['公司名稱'] == company]
                    if len(c_sub) >= 2:
                        g_rate = c_sub['營收'].pct_change().mean()
                        f_info = get_forecast(c_sub)
                        desc = (f"受查對象 {company} 之歷史平均年成長率為 {g_rate:.2%}。 "
                                f"經 AI 模組推估，未來一期之營收目標預計可達 {f_info['營收'].iloc[0]:,.0f} 元。")
                        report.add_paragraph(desc)

                report.add_heading("二、 財務犯罪防制鑑定結論", level=1)
                abnormal = main_df[main_df['舞弊狀態'] == "危險"]
                if abnormal.empty:
                    report.add_paragraph("經查核，目前所有受查對象之各項財務指標均在正常範圍內。")
                else:
                    for _, row in abnormal.iterrows():
                        p = report.add_paragraph(f"異常標的：{row['公司名稱']} ({row['年度']}年度)")
                        p.add_run(f"\n舞弊鑑定結論：M-Score 為 {row['M分數']}，屬危險級別。").bold = True
                        p.add_run(f"\n其他風險：掏空評級為 {row['掏空風險']}，吸金警示為 {row['吸金指標']}。")

                stream = io.BytesIO()
                report.save(stream)
                st.download_button("下載 Word 報告", stream.getvalue(), "玄武鑑定報告.docx")
        else:
            st.warning("檔案解析成功但未發現有效數據。")
else:
    st.info("系統就緒。請上傳至少兩年份財報以執行自動化鑑定。")
