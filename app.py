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

# --- 1. 介面美化與配置 ---
st.set_page_config(layout="wide", page_title="玄武極速鑑識系統")

st.markdown("""
    <div style='background-color:#073642; padding:25px; border-radius:15px; border-left: 10px solid #268bd2;'>
        <h1 style='color:#eee8d5; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#93a1a1; margin:0; font-size: 1.1em;'>專業級財報鑑識：高效數據對位、多維趨勢分析、自動化鑑定意見</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. 專業級數值提取引擎 ---
def professional_parse(file):
    res = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, "應收": 0.0, 
           "存貨": 0.0, "現金": 0.0, "其他應收": 0.0, "預付": 0.0, "淨利": 0.0}
    try:
        with pdfplumber.open(file) as pdf:
            # 抓取包含資產負債與損益表的核心頁面
            text_pool = ""
            for i in range(min(15, len(pdf.pages))):
                text_pool += (pdf.pages[i].extract_text() or "")
            
            # 抓取年度 (支援民國與西元)
            y_match = re.search(r"(\d{3,4})\s*年度", text_pool)
            if y_match:
                y = int(y_match.group(1))
                res["年度"] = y + 1911 if y < 1000 else y
            
            # 定位科目與數值 (改進：尋找科目後最接近的金額格式字串)
            maps = {
                "營收": ["營業收入", "營收合計", "Operating revenue"],
                "應收": ["應收帳款淨額", "應收帳款", "Accounts receivable"],
                "存貨": ["存貨", "Inventories"],
                "現金": ["現金及約當現金", "Cash and cash equivalents"],
                "其他應收": ["其他應收款", "Other receivables"],
                "預付": ["預付款項", "Prepayments"],
                "淨利": ["本期淨利", "本期損益", "Net income"]
            }

            for key, kw_list in maps.items():
                for kw in kw_list:
                    # 匹配關鍵字後方 30 字元內的金額格式 (處理逗號、括號、小數點)
                    pattern = rf"{kw}.{{0,30}}?([\d,]{{2,}}|\([\d,]{{2,}}\))"
                    m = re.search(pattern, text_pool)
                    if m:
                        s = m.group(1).replace(',', '').replace('$', '')
                        if '(' in s and ')' in s: s = '-' + s.replace('(', '').replace(')', '')
                        try:
                            res[key] = float(s)
                            break
                        except: continue
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except Exception as e:
        st.error(f"解析 {file.name} 時發生錯誤: {e}")
        return pd.DataFrame()

# --- 3. 側邊欄控制 ---
with st.sidebar:
    st.header("⚙️ 鑑定功能中心")
    mode = st.radio("分析模式", ["🔍 單一公司多年比較", "⚔️ 多公司橫向 PK"])
    st.divider()
    uploaded_files = st.file_uploader("批次上傳受查 PDF", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "張鈞翔會計師")

# --- 4. 數據處理與視覺化分析 ---
if uploaded_files:
    dfs = [professional_parse(f) for f in uploaded_files]
    data_pool = pd.concat([d for d in dfs if not d.empty], ignore_index=True)
    
    if not data_pool.empty:
        # 指標運算 (Beneish M-Score 核心參數)
        data_pool['M分數'] = data_pool.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
        data_pool['掏空指數'] = data_pool.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
        
        if mode == "🔍 單一公司多年比較":
            target = st.selectbox("選擇受查對象", data_pool['公司名稱'].unique())
            sub = data_pool[data_pool['公司名稱'] == target].sort_values('年度')
            
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("📈 營收與應收對位趨勢")
                fig1, ax1 = plt.subplots(figsize=(8, 5))
                ax1.plot(sub['年度'].astype(str), sub['營收'], marker='o', label='營收', color='#268bd2', linewidth=2)
                ax1.plot(sub['年度'].astype(str), sub['應收'], marker='x', label='應收', color='#cb4b16', linestyle='--')
                ax1.legend()
                st.pyplot(fig1)
            with c2:
                st.subheader("🚨 異常資金流出監測")
                fig2, ax2 = plt.subplots(figsize=(8, 5))
                ax2.bar(sub['年度'].astype(str), sub['掏空指數'], color='#859900', alpha=0.7, label='掏空指標')
                ax2.axhline(y=0.2, color='red', linestyle=':', label='風險門檻 (20%)')
                ax2.legend()
                st.pyplot(fig2)
            
            st.write("### 歷年鑑識明細數據")
            st.dataframe(sub.style.background_gradient(subset=['M分數', '掏空指數'], cmap='YlOrRd'))

        # --- 5. 專業自動化報告生成 ---
        st.divider()
        if st.button("🚀 生成專業鑑定意見書"):
            doc = Document()
            # 報告標題與排版
            head = doc.add_heading('財務鑑識鑑定意見書', 0)
            head.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph(f"受查單位：{target if mode == '🔍 單一公司多年比較' else '多樣本聯合鑑定'}")
            doc.add_paragraph(f"簽證會計師：{auditor}\n報告日期：2026年4月")
            
            # 專業敘述邏輯
            doc.add_heading('壹、 鑑定結論與舞弊風險評估', level=1)
            latest = sub.iloc[-1]
            risk_lvl = "【高度預警】" if latest['M分數'] > -1.78 else "【品質尚屬穩定】"
            
            narrative = (f"經由本系統之財務模組鑑定，受查單位最新年度之 M-Score 為 {latest['M分數']}，顯示盈餘品質{risk_lvl}。"
                         f"觀察其掏空指數（其他應收與預付之佔比）為 {latest['掏空指數']*100:.2f}%。"
                         "若指標顯著偏離產業均值，審計人員應擴大對關係人交易與非常規資金調度之抽查比率...")
            doc.add_paragraph(narrative)
            
            # 圖表嵌入
            doc.add_heading('貳、 歷年財務趨勢鑑識圖表', level=1)
            img_buf = io.BytesIO()
            plt.savefig(img_buf, format='png', dpi=200)
            img_buf.seek(0)
            doc.add_picture(img_buf, width=Inches(5.5))
            
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 下載專業鑑定報告 (.docx)", buf, f"Forensic_Audit_{target}.docx")
    else:
        st.error("⚠️ 無法從 PDF 提取數據。請確認文件內容是否為可選取文字。")
