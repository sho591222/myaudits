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

# --- 1. 系統環境設定 ---
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

if 'current_fig' not in st.session_state:
    st.session_state['current_fig'] = None

# 介面美化
st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財報鑑識系統：圖表嵌入與深沈比對報告</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. 核心數值解析引擎 ---
def clean_num(text):
    if not text: return 0.0
    s = str(text).strip().replace(',', '')
    if '(' in s and ')' in s:
        s = '-' + s.replace('(', '').replace(')', '')
    match = re.search(r'[-+]?\d*\.\d+|\d+', s)
    return float(match.group()) if match else 0.0

def parse_financial_data(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, 
            "應收": 0.0, "存貨": 0.0, "現金": 0.0, "其他應收": 0.0, "預付": 0.0, "淨利": 0.0
        }
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:10]:
                text = page.extract_text() or ""
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
                        if not nums: continue
                        val = nums[-1]
                        
                        if any(k in row_str for k in ["營業收入", "營收"]): res["營收"] = val
                        if "應收帳款" in row_str and "其他" not in row_str: res["應收"] = val
                        if "存貨" in row_str: res["存貨"] = val
                        if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = val
                        if "其他應收款" in row_str: res["其他應收"] = val
                        if "預付款項" in row_str: res["預付"] = val
                        if any(k in row_str for k in ["本期淨利", "本期損益"]): res["淨利"] = val
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except:
        return pd.DataFrame()

# --- 3. 鑑識分析引擎 ---
def forensic_analysis(df):
    df['M分數'] = df.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
    df['掏空指數'] = df.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
    df['不實預警'] = df['M分數'].apply(lambda x: "注意" if x > -1.78 else "正常")
    df['洗錢風險'] = df.apply(lambda r: "注意" if r['現金'] > r['營收']*0.8 and r['淨利'] < 0 else "正常", axis=1)
    return df

# --- 4. 側邊欄與上傳控制 ---
with st.sidebar:
    st.header("⚙️ 鑑定控制台")
    mode = st.radio("功能模式", ["🔍 單一深度鑑定", "⚔️ 多公司橫向 PK"])
    st.divider()
    uploaded_files = st.file_uploader("上傳財報 PDF (可多選)", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "會計師")

# --- 5. 主程式執行 ---
if uploaded_files:
    dfs = [parse_financial_data(f) for f in uploaded_files]
    df_pool = pd.concat([d for d in dfs if not d.empty], ignore_index=True)
    
    if not df_pool.empty:
        df_pool = forensic_analysis(df_pool)
        
        if mode == "🔍 單一深度鑑定":
            target = st.selectbox("選擇公司", df_pool['公司名稱'].unique())
            sub = df_pool[df_pool['公司名稱'] == target].sort_values('年度')
            
            # --- 視覺化圖表 ---
            st.subheader(f"歷年趨勢鑑識看板：{target}")
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='營業收入', linewidth=2)
            ax.plot(sub['年度'].astype(str), sub['應收'], marker='x', label='應收帳款', linestyle='--')
            ax.bar(sub['年度'].astype(str), sub['其他應收'], alpha=0.3, label='其他應收 (掏空區)', color='red')
            ax.legend()
            st.pyplot(fig)
            st.session_state['current_fig'] = fig 
            
            st.write("### 歷年鑑定明細數據")
            st.dataframe(sub)

            # --- 6. 生成 Word 報告 (已修正縮排) ---
            st.divider()
            if st.button("🚀 生成「深沈比對」詳盡鑑定報告"):
                doc = Document()
                
                # 報告開頭
                title = doc.add_heading('財務鑑識鑑定意見書', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_heading(f'受查單位：{target}', level=1)
                doc.add_paragraph(f"主辦會計師：{auditor}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
                
                # 第一章：盈餘操縱分析
                doc.add_heading('壹、 財務舞弊與盈餘操縱偵測', level=1)
                latest = sub.iloc[-1]
                m_score = latest['M分數']
                if m_score > -1.78:
                    m_desc = f"【警示】M-Score 為 {m_score}，超過風險門檻。顯示該公司存在顯著盈餘操縱風險。"
                else:
                    m_desc = f"【正常】M-Score 為 {m_score}，利潤品質尚屬穩定。"
                doc.add_paragraph(m_desc)
                
                # 第二章：趨勢圖表嵌入 (核心比對)
                doc.add_heading('貳、 歷年趨勢與資產掏空比對', level=1)
                doc.add_paragraph("透過縱向數據對比，分析營收與異常科目之連動性：")
                
                if st.session_state['current_fig']:
                    img_buf = io.BytesIO()
                    st.session_state['current_fig'].savefig(img_buf, format='png', dpi=300)
                    img_buf.seek(0)
                    doc.add_picture(img_buf, width=Inches(5.5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_paragraph(f"深沈分析：最新年度掏空指數為 {latest['掏空指數']}，反映非本業資金偏離程度。")

                # 第三章：結論
                doc.add_heading('參、 會計師鑑定總結意見', level=1)
                doc.add_paragraph(f"綜上所述，針對 {target} 之財務風險指標，建議進一步抽查傳票...")
                
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.download_button("📥 下載詳盡鑑定報告 (.docx)", buf, f"Forensic_Report_{target}.docx")
        
        else: # PK 模式
            st.subheader("多公司同年度橫向 PK")
            target_year = st.selectbox("選擇年度", sorted(df_pool['年度'].unique(), reverse=True))
            pk_df = df_pool[df_pool['年度'] == target_year]
            st.bar_chart(pk_df.set_index('公司名稱')[['M分數', '掏空指數']])
            st.dataframe(pk_df)
            
    else:
        st.info("請上傳財報 PDF 文件以開始分析。")
