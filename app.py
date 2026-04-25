import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import pdfplumber

# --- 1. 介面極簡化與效能配置 ---
st.set_page_config(layout="wide", page_title="玄武極速鑑識")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>【極速版】全自動語意對位：數據提取、趨勢比對、深沈報告</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. 高效暴力解析引擎 ---
def fast_parse(file):
    res = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, "應收": 0.0, 
           "存貨": 0.0, "現金": 0.0, "其他應收": 0.0, "預付": 0.0, "淨利": 0.0}
    try:
        with pdfplumber.open(file) as pdf:
            # 只取前 12 頁（資產負債表與損益表通常在此），節省解析時間
            text_pool = "".join([p.extract_text() or "" for p in pdf.pages[:12]])
            
            # 抓取年度 (民國/西元自動換算)
            y_match = re.search(r"(\d{3,4})\s*年度", text_pool)
            if y_match:
                y = int(y_match.group(1))
                res["年度"] = y + 1911 if y < 1000 else y
            
            # 數值對位邏輯 (科目關鍵字庫)
            maps = {
                "營收": ["營業收入", "營收合計"],
                "應收": ["應收帳款淨額", "應收帳款"],
                "存貨": ["存貨", "存貨淨額"],
                "現金": ["現金及約當現金", "現金及流動資產"],
                "其他應收": ["其他應收款"],
                "預付": ["預付款項"],
                "淨利": ["本期淨利", "本期損益"]
            }

            for key, kw_list in maps.items():
                for kw in kw_list:
                    # 匹配關鍵字後方 25 字元內的數字
                    m = re.search(rf"{kw}.{{0,25}}?([\d,]{{2,}}|\([\d,]{{2,}}\))", text_pool)
                    if m:
                        val = m.group(1).replace(',', '').replace('(', '-').replace(')', '')
                        res[key] = float(val)
                        break
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except:
        return pd.DataFrame()

# --- 3. 側邊欄與鑑定設置 ---
with st.sidebar:
    st.header(" 鑑識控制台")
    mode = st.radio("分析模式", [" 單一公司多年比較", " 多公司橫向 PK"])
    uploaded_files = st.file_uploader("批次上傳 PDF", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "張鈞翔會計師")

# --- 4. 主程序執行 ---
if uploaded_files:
    # 平行化解析概念：直接批次處理
    data_all = pd.concat([fast_parse(f) for f in uploaded_files], ignore_index=True)
    
    if not data_all.empty:
        # 指標運算
        data_all['M分數'] = data_all.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
        data_all['掏空指數'] = data_all.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
        
        if mode == "🔍 單一公司多年比較":
            target = st.selectbox("選擇受查對象", data_all['公司名稱'].unique())
            sub = data_all[data_all['公司名稱'] == target].sort_values('年度')
            
            # 視覺化：雙看板對比
            c1, c2 = st.columns(2)
            with c1:
                fig_v, ax_v = plt.subplots()
                ax_v.plot(sub['年度'].astype(str), sub['營收'], 'o-', label='營收')
                ax_v.plot(sub['年度'].astype(str), sub['應收'], 'x--', label='應收')
                ax_v.set_title("營收與應收對位趨勢")
                ax_v.legend()
                st.pyplot(fig_v)
            with c2:
                fig_t, ax_t = plt.subplots()
                ax_t.bar(sub['年度'].astype(str), sub['掏空指數'], color='salmon', label='掏空指標')
                ax_t.set_title("資產流出(掏空)壓力測試")
                ax_t.legend()
                st.pyplot(fig_t)
            
            st.dataframe(sub)

        # --- 5. 快速生成深沈敘述 Word ---
        st.divider()
        if st.button(" 生成極速鑑定報告"):
            doc = Document()
            doc.add_heading('財務鑑識鑑定意見書', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"受查單位：{target}\n簽證會計師：{auditor}")
            
            doc.add_heading('壹、 深度數據比對與風險敘述', level=1)
            row = sub.iloc[-1]
            risk_text = "【高風險】" if row['M分數'] > -1.78 else "【正常】"
            doc.add_paragraph(f"經本所系統鑑定，該公司最新年度 M-Score 為 {row['M分數']}，判定結果為 {risk_text}。"
                              f"特別注意：掏空指數已達 {row['掏空指數']}，建議加強抽查「其他應收款」之真實性...")
            
            # 嵌入當前畫面的圖表
            img_buf = io.BytesIO()
            plt.savefig(img_buf, format='png', dpi=200)
            img_buf.seek(0)
            doc.add_picture(img_buf, width=Inches(5.5))
            
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 領取 Word 報告", buf, f"Fast_Report_{target}.docx")
    else:
        st.warning("請上傳財報 PDF 以開始運作。")
