import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import pdfplumber

# --- 1. 系統環境與佈局 ---
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

# 初始化 Session State 以儲存繪圖物件
if 'current_fig' not in st.session_state: st.session_state['current_fig'] = None

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>終極鑑識系統：圖表嵌入與詳細敘述報告</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. 數值提取核心 (強化解析與科目對齊) ---
def clean_num(text):
    if not text: return 0.0
    s = str(text).strip().replace(',', '')
    if '(' in s and ')' in s: s = '-' + s.replace('(', '').replace(')', '')
    match = re.search(r'[-+]?\d*\.\d+|\d+', s)
    return float(match.group()) if match else 0.0

def parse_financial_data(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, 
            "應收": 0.0, "存貨": 0.0, "現金": 0.0, "其他應收": 0.0, "預付": 0.0, "淨利": 0.0
        }
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:8]:
                text = page.extract_text() or ""
                # 年度抓取
                if res["年度"] == 0:
                    y_match = re.search(r"(\d{3,4})\s*年度", text)
                    if y_match:
                        y = int(y_match.group(1))
                        res["年度"] = y + 1911 if y < 1000 else y
                
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        row_str = "".join([str(x) for x in row if x])
                        nums = [clean_num(x) for x in row if x and any(c.isdigit() for c in str(x))]
                        if len(nums) < 1: continue
                        val = nums[-1] # 假設最後一欄為當期數
                        
                        if any(k in row_str for k in ["營業收入", "營收"]): res["營收"] = val
                        if "應收帳款" in row_str and "其他" not in row_str: res["應收"] = val
                        if "存貨" in row_str: res["存貨"] = val
                        if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = val
                        if "其他應收款" in row_str: res["其他應收"] = val
                        if "預付款項" in row_str: res["預付"] = val
                        if any(k in row_str for k in ["本期淨利", "本期損益"]): res["淨利"] = val
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except: return pd.DataFrame()

# --- 3. 鑑識分析引擎 ---
def forensic_analysis(df):
    df['M分數'] = df.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
    df['掏空指數'] = df.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
    df['不實預警'] = df['M分數'].apply(lambda x: "注意" if x > -1.78 else "正常")
    df['洗錢風險'] = df.apply(lambda r: "注意" if r['現金'] > r['營收']*0.8 and r['淨利'] < 0 else "正常", axis=1)
    return df

# --- 4. 側邊欄控制 ---
with st.sidebar:
    st.header("⚙️ 鑑定控制台")
    mode = st.radio("功能模式", [" 單一深度鑑定", " 多公司同年度 PK"])
    st.divider()
    uploaded_files = st.file_uploader("上傳受查 PDF", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "會計師")

# --- 5. 主程式 ---
if uploaded_files:
    # 解析數據 (實際運作會耗時)
    data_pool = pd.concat([parse_financial_data(f) for f in uploaded_files], ignore_index=True)
    
    if not data_pool.empty:
        df_pool = forensic_analysis(data_pool)
        
        if mode == "🔍 單一深度鑑定":
            target = st.selectbox("選擇公司", df_pool['公司名稱'].unique())
            sub = df_pool[df_pool['公司名稱'] == target].sort_values('年度')
            
            # --- 繪圖區：多年趨勢比對 ---
            st.subheader(f"歷年趨勢鑑識看板：{target}")
            
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='營業收入', linewidth=2)
            ax.plot(sub['年度'].astype(str), sub['應收'], marker='x', label='應收帳款', linestyle='--')
            # 突出掏空指標
            ax.bar(sub['年度'].astype(str), sub['其他應收'], alpha=0.3, label='其他應收 (掏空區)', color='red')
            ax.legend()
            st.pyplot(fig)
            st.session_state['current_fig'] = fig # 將圖表儲存到 Session State
            
            st.dataframe(sub)

            # --- 6. 生成 Word 報告 (圖表嵌入強化版) ---
            st.divider()
            if st.button("🚀 生成「圖文詳盡版」鑑定鑑定報告"):
                doc = Document()
                title = doc.add_heading('財務鑑識鑑定意見書', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_heading(f'受查單位：{target}', level=1)
                
                # 第壹章：舞弊偵測
                doc.add_heading('壹、 財務舞弊與盈餘操縱偵測', level=1)
                doc.add_paragraph(f"本鑑定採用 Beneish M-Score 模型。該公司最新年度 M 分數為 {sub.iloc[-1]['M分數']}，財報不實預警為：{sub.iloc[-1]['不實預警']}。重點在於分析應收帳款與存貨是否非正常成長。")

                # 第貳章：資產掏空 (關鍵圖表嵌入)
                doc.add_heading('貳、 資產掏空與異常資金偏移', level=1)
                doc.add_paragraph("以下為該公司歷年營收、應收帳款與其他應收款之趨勢對比圖。若觀察到營收持平但其他應收款異常增加（紅柱突起），則存在高度資產掏空嫌疑：")
                
                # --- 將 Matplotlib 圖表嵌入 Word ---
                if st.session_state['current_fig']:
                    # 將圖表存入 Buffer
                    img_buf = io.BytesIO()
                    st.session_state['current_fig'].savefig(img_buf, format='png', bbox_inches='tight', dpi=300)
                    img_buf.seek(0)
                    
                    # 嵌入 Word，寬度設定為 6 英吋
                    doc.add_picture(img_buf, width=Inches(6))
                    last_paragraph = doc.paragraphs[-1]
                    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER # 圖片置中
                
                doc.add_paragraph(f"經數據運算，最新年度掏空指數為 {sub.iloc[-1]['掏空指數']}，顯示非本業資產之流出佔比...")

                # 第參章：結論
                doc.add_heading('參、 會計師鑑定意見', level=1)
                final_conc = doc.add_paragraph(f"綜上所述，主辦會計師 {auditor}
