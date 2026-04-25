import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt
import io
import re
import pdfplumber

# 1. 系統環境與佈局
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>終極鑑識系統：舞弊、掏空、洗錢與財報不實深度分析</p>
    </div>
""", unsafe_allow_html=True)

# 2. 數值提取核心 (強化版：處理千分位、括號、及碎裂文字)
def clean_num(text):
    if not text: return 0.0
    # 處理會計格式 (1,234.56) -> -1234.56
    s = str(text).strip().replace(',', '')
    if '(' in s and ')' in s:
        s = '-' + s.replace('(', '').replace(')', '')
    # 提取純數字部分
    match = re.search(r'[-+]?\d*\.\d+|\d+', s)
    return float(match.group()) if match else 0.0

def parse_pdf_advanced(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, 
            "應收帳款": 0.0, "存貨": 0.0, "現金": 0.0, "負債總額": 0.0, 
            "其他應收款": 0.0, "預付款項": 0.0, "股份酬勞": 0.0, "淨利": 0.0
        }
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:10]: # 擴大掃描範圍至附註
                text = page.extract_text() or ""
                # 1. 抓取年度 (民國/西元)
                if res["年度"] == 0:
                    y_match = re.search(r"(\d{3,4})\s*年度", text)
                    if y_match:
                        y = int(y_match.group(1))
                        res["年度"] = y + 1911 if y < 1000 else y

                # 2. 萬能表格掃描 (針對抓不到數據的優化)
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        row_str = "".join([str(x) for x in row if x])
                        # 只要該行出現關鍵字，就嘗試提取該行所有數字
                        if any(k in row_str for k in ["營", "收", "資產", "淨利", "應收"]):
                            nums = [clean_num(x) for x in row if x and any(c.isdigit() for c in str(x))]
                            if not nums: continue
                            val = nums[-1] # 假設最後一欄為當期數
                            
                            if any(k in row_str for k in ["營業收入", "營收"]): res["營收"] = val
                            if "應收帳款" in row_str and "其他" not in row_str: res["應收帳款"] = val
                            if "存貨" in row_str: res["存貨"] = val
                            if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = val
                            if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = val
                            if "其他應收款" in row_str: res["其他應收款"] = val
                            if "預付款項" in row_str: res["預付款項"] = val
                            if any(k in row_str for k in ["股份酬勞", "認股權"]): res["股份酬勞"] = val
                            if any(k in row_str for k in ["本期淨利", "本期損益"]): res["淨利"] = val
                            
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except: return pd.DataFrame()

# 3. 鑑識與分析邏輯 (包含預測)
def run_analytics(df, method):
    for c in ['M分數', '掏空指數', '洗錢風險', '財報不實']: df[c] = ""
    for i in df.index:
        r = df.at[i, '營收']
        if r > 0:
            # 舞弊 M-Score
            m = -3.2 + (0.15 * (df.at[i, '應收帳款']/r)) + (0.1 * (df.at[i, '存貨']/r))
            df.at[i, 'M分數'] = round(m, 2)
            # 掏空分析
            t_idx = (df.at[i, '其他應收款'] + df.at[i, '預付款項']) / r
            df.at[i, '掏空指數'] = round(t_idx, 3)
            # 洗錢與不實 (邏輯判斷)
            df.at[i, '洗錢風險'] = "注意" if df.at[i, '現金'] > r * 0.8 else "正常"
            df.at[i, '財報不實'] = "高風險" if m > -1.78 or t_idx > 0.2 else "低風險"
    return df

# 4. 側邊欄控制
with st.sidebar:
    st.header("⚙️ 鑑識功能中心")
    view_mode = st.radio("模式選擇", ["🔍 深度鑑定報告", "⚔️ 同年度橫向 PK"])
    st.divider()
    forecast_method = st.selectbox("財務預測模型", ["線性成長 (5%)", "保守估計 (0%)", "歷史平均成長"])
    st.divider()
    uploaded_files = st.file_uploader("上傳財報 PDF", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "會計師")

# 5. 主程式執行
if uploaded_files:
    all_dfs = [parse_pdf_advanced(f) for f in uploaded_files]
    df_pool = pd.concat([d for d in all_dfs if not d.empty], ignore_index=True)
    
    if not df_pool.empty:
        df_pool = run_analytics(df_pool, forecast_method)
        
        if view_mode == "🔍 深度鑑定報告":
            target = st.selectbox("選擇公司", df_pool['公司名稱'].unique())
            sub = df_pool[df_pool['公司名稱'] == target].sort_values('年度')
            
            # --- 視覺化圖表 (含預測) ---
            st.subheader(f"財務鑑識看板：{target}")
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='歷史營收', linewidth=2)
            # 預測虛線
            last_val = sub['營收'].iloc[-1]
            growth = 1.05 if "5%" in forecast_method else 1.0
            ax.plot([sub['年度'].astype(str).iloc[-1], "Forecast"], [last_val, last_val * growth], '--', marker='s', label='預測路徑')
            ax.legend()
            st.pyplot(fig)
            
            st.dataframe(sub)

            # --- 生成詳細 Word 報告 ---
            st.divider()
            if st.button("📥 生成詳盡鑑定報告書"):
                doc = Document()
                doc.add_heading(f"玄武鑑識鑑定報告 - {target}", 0)
                doc.add_heading("一、舞弊與財報不實分析", level=1)
                doc.add_paragraph(f"經偵測該公司 M 分數為 {sub.iloc[-1]['M分數']}，財報不實風險：{sub.iloc[-1]['財報不實']}。")
                doc.add_heading("二、掏空與洗錢風險分析", level=1)
                doc.add_paragraph(f"資產掏空指數：{sub.iloc[-1]['掏空指數']} / 洗錢風險：{sub.iloc[-1]['洗錢風險']}。")
                doc.add_heading("三、會計師鑑定建議", level=1)
                doc.add_paragraph(f"主辦會計師：{auditor}")
                
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.download_button("下載 Word 報告", buf, f"Forensic_Report_{target}.docx")
        
        else: # PK 模式
            st.subheader("多公司同年度橫向 PK")
            year = st.selectbox("PK年度", sorted(df_pool['年度'].unique(), reverse=True))
            pk_df = df_pool[df_pool['年度'] == year]
            st.bar_chart(pk_df.set_index('公司名稱')[['M分數', '掏空指數']])
            st.dataframe(pk_df)

    else:
        st.error("【錯誤】未能讀取數值。請確認 PDF 內容是否為可選取的文字（非圖片掃描件）。")
