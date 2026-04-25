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

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>AI 財報鑑識旗艦系統：自動化深沈敘述與圖表鑑定</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. 強力數值提取引擎 (解決數據為 0 的問題) ---
def clean_num(text):
    if not text: return 0.0
    s = str(text).strip().replace(',', '').replace('$', '')
    if '(' in s and ')' in s:
        s = '-' + s.replace('(', '').replace(')', '')
    # 提取第一個看起來像數字的字串
    match = re.search(r'[-+]?\d*\.\d+|\d+', s)
    return float(match.group()) if match else 0.0

def parse_financial_data(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, 
            "應收": 0.0, "存貨": 0.0, "現金": 0.0, "其他應收": 0.0, "預付": 0.0, "淨利": 0.0
        }
        with pdfplumber.open(file) as pdf:
            full_text = ""
            for page in pdf.pages[:10]:
                text = page.extract_text() or ""
                full_text += text
                
                # 抓年度
                if res["年度"] == 0:
                    y_match = re.search(r"(\d{3,4})\s*年度", text)
                    if y_match:
                        y = int(y_match.group(1))
                        res["年度"] = y + 1911 if y < 1000 else y

                # 遍歷頁面中的表格行，進行關鍵字精準對位
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row: continue
                        row_str = "".join([str(x) for x in row if x])
                        nums = [clean_num(x) for x in row if x and any(c.isdigit() for c in str(x))]
                        if not nums: continue
                        
                        # 邏輯：只要行內有關鍵字，就取該行最後一個有效數字
                        val = nums[-1]
                        if any(k in row_str for k in ["營業收入", "營收"]): res["營收"] = val
                        if "應收帳款" in row_str and "其他" not in row_str: res["應收"] = val
                        if "存貨" in row_str: res["存貨"] = val
                        if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = val
                        if "其他應收款" in row_str: res["其他應收"] = val
                        if "預付款項" in row_str: res["預付"] = val
                        if any(k in row_str for k in ["本期淨利", "本期損益"]): res["淨利"] = val
        
        # 如果透過表格抓不到，則進行全文正則掃描 (備援機制)
        if res["營收"] == 0:
            rev_match = re.search(r"營業收入.*?([\d,]+)", full_text)
            if rev_match: res["營收"] = clean_num(rev_match.group(1))

        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except:
        return pd.DataFrame()

# --- 3. 鑑識分析引擎 ---
def forensic_analysis(df):
    # Beneish M-Score 簡化版與掏空指標
    df['M分數'] = df.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
    df['掏空指數'] = df.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
    df['舞弊風險'] = df['M分數'].apply(lambda x: "注意" if x > -1.78 else "正常")
    return df

# --- 4. 側邊欄控制 ---
with st.sidebar:
    st.header("⚙️ 鑑定功能中心")
    mode = st.radio("模式選擇", [" 單一深度比較", " 多公司橫向 PK"])
    st.divider()
    uploaded_files = st.file_uploader("上傳受查 PDF", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "會計師")

# --- 5. 主程式 ---
if uploaded_files:
    data_list = [parse_financial_data(f) for f in uploaded_files]
    df_pool = pd.concat([d for d in data_list if not d.empty], ignore_index=True)
    
    if not df_pool.empty:
        df_pool = forensic_analysis(df_pool)
        
        if mode == "🔍 單一深度比較":
            target = st.selectbox("選擇受查公司", df_pool['公司名稱'].unique())
            sub = df_pool[df_pool['公司名稱'] == target].sort_values('年度')
            
            st.subheader(f"歷年趨勢鑑識看板：{target}")
            
            # 修正圖表：確保數據不為 0
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='營業收入', linewidth=2)
            ax.plot(sub['年度'].astype(str), sub['應收'], marker='x', label='應收帳款', linestyle='--')
            ax.bar(sub['年度'].astype(str), sub['其他應收'], alpha=0.3, label='其他應收 (掏空指標)', color='red')
            ax.legend()
            st.pyplot(fig)
            st.session_state['current_fig'] = fig 
            
            st.write("### 歷年鑑定數據清單")
            st.dataframe(sub)

            # --- 6. 生成 Word 報告 (深沈敘述 + 圖表嵌入) ---
            st.divider()
            if st.button(" 生成「深成比對」詳盡鑑定報告書"):
                doc = Document()
                title = doc.add_heading('財務鑑識鑑定意見書', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_heading(f'壹、 受查單位與鑑定聲明', level=1)
                doc.add_paragraph(f"受查單位：{target}\n簽證會計師：{auditor}\n鑑定日期：2026年4月")
                
                doc.add_heading(f'貳、 盈餘操縱與財報不實分析', level=1)
                latest = sub.iloc[-1]
                m_score = latest['M分數']
                risk = latest['舞弊風險']
                
                # 深沈比對文字敘述
                narrative = (f"經系統偵測，{target} 之最新年度 M-Score 為 {m_score}（門檻值 -1.78）。"
                             f"鑑定結果為「{risk}」。分析顯示該公司之營收成長與應收帳款之變動對位關係...")
                doc.add_paragraph(narrative)
                
                doc.add_heading(f'參、 歷年趨勢與資產掏空比對圖', level=1)
                doc.add_paragraph("以下為系統自動產出之趨勢對比圖，用於偵測營收與其他應收款之異常背離：")
                
                # 嵌入圖表
                if st.session_state['current_fig']:
                    img_buf = io.BytesIO()
                    st.session_state['current_fig'].savefig(img_buf, format='png', dpi=300)
                    img_buf.seek(0)
                    doc.add_picture(img_buf, width=Inches(5.5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_heading(f'肆、 綜合鑑定建議', level=1)
                doc.add_paragraph(f"主辦會計師 {auditor} 建議針對異常年度進行原始憑證之抽盤...")

                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.download_button("📥 下載深沈敘述鑑定報告 (.docx)", buf, f"Forensic_Report_{target}.docx")
    else:
        st.warning("⚠️ 未能讀取到有效財務數值，請確認 PDF 是否為掃描檔（圖片）。")
