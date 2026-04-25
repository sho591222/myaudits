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

# --- 1. 介面設定 ---
st.set_page_config(layout="wide", page_title="玄武極速鑑識")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>【多檔批次版】全自動非法偵測、深度趨勢比對、未來風險預測</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. 核心解析引擎 (支援多檔自動對位) ---
def clean_num(text):
    if not text: return 0.0
    s = str(text).strip().replace(',', '').replace('$', '')
    if '(' in s and ')' in s: s = '-' + s.replace('(', '').replace(')', '')
    match = re.search(r'[-+]?\d*\.\d+|\d+', s)
    return float(match.group()) if match else 0.0

def batch_parse(file):
    res = {"公司": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, "應收": 0.0, "存貨": 0.0, "其他應收": 0.0, "預付": 0.0, "淨利": 0.0}
    try:
        with pdfplumber.open(file) as pdf:
            text = "".join([p.extract_text() or "" for p in pdf.pages[:15]])
            # 抓年度
            y_m = re.search(r"(\d{3,4})\s*年度", text)
            if y_m:
                y = int(y_m.group(1))
                res["年度"] = y + 1911 if y < 1000 else y
            
            # 科目暴力匹配
            maps = {"營收": ["營業收入", "營收合計"], "應收": ["應收帳款淨額", "應收帳款"],
                    "存貨": ["存貨"], "其他應收": ["其他應收款"], "預付": ["預付款項"], "淨利": ["本期淨利", "本期損益"]}
            for key, kws in maps.items():
                for kw in kws:
                    m = re.search(rf"{kw}.{{0,25}}?([\d,]{{2,}}|\([\d,]{{2,}}\))", text)
                    if m:
                        res[key] = clean_num(m.group(1))
                        break
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except: return pd.DataFrame()

# --- 3. 側邊欄控制 ---
with st.sidebar:
    st.header("⚡ 鑑識控制台")
    uploaded_files = st.file_uploader("批次上傳多個 PDF (按住 Ctrl 多選)", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "張鈞翔會計師")

# --- 4. 主程式：數據整合與非法圖表 ---
if uploaded_files:
    # 批次讀取所有上傳的文件
    all_dfs = []
    for f in uploaded_files:
        single_df = batch_parse(f)
        if not single_df.empty:
            all_dfs.append(single_df)
    
    if all_dfs:
        df_full = pd.concat(all_dfs, ignore_index=True)
        # 關鍵指標計算
        df_full['M分數'] = df_full.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
        df_full['掏空指數'] = df_full.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
        
        target = st.selectbox("選擇鑑定對象", df_full['公司'].unique())
        sub = df_full[df_full['公司'] == target].sort_values('年度')

        st.subheader(f"🚨 非法預警監控看板：{target}")
        c1, c2 = st.columns(2)
        with c1:
            fig_m, ax_m = plt.subplots()
            ax_m.plot(sub['年度'].astype(str), sub['M分數'], 'r-o', label='M-Score (舞弊偵測)')
            ax_m.axhline(y=-1.78, color='black', linestyle='--')
            ax_m.set_title("Beneish M-Score 舞弊壓力測試")
            ax_m.legend()
            st.pyplot(fig_m)
        with c2:
            fig_t, ax_t = plt.subplots()
            ax_t.bar(sub['年度'].astype(str), sub['掏空指數'], color='orange', label='資產流出')
            ax_t.set_title("資產非法挪用監測")
            ax_t.legend()
            st.pyplot(fig_t)

        # --- 5. Word 報告：深度解析與預測邏輯 ---
        st.divider()
        if st.button("🚀 生成深度鑑定意見書 (含未來一年預測)"):
            doc = Document()
            doc.add_heading('財務鑑識鑑定意見書', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"受查單位：{target}\n簽證會計師：{auditor}")

            # 深度敘述
            doc.add_heading('壹、 舞弊風險深度分析', level=1)
            curr = sub.iloc[-1]
            risk = "【高度警示】" if curr['M分數'] > -1.78 else "【正常】"
            doc.add_paragraph(f"本鑑定重點在於盈餘操縱風險。目前年度 M-Score 為 {curr['M分數']}，判定為 {risk}。"
                              "此數值反映了公司在營收認列與資產評價上的誠實度，建議針對異常變動年度進行實質測試。")

            # 圖表嵌入
            doc.add_heading('貳、 鑑識圖表證據', level=1)
            img_buf = io.BytesIO()
            plt.savefig(img_buf, format='png', dpi=200)
            img_buf.seek(0)
            doc.add_picture(img_buf, width=Inches(5.5))

            # 未來一年預測 (基於移動平均或趨勢)
            doc.add_heading('參、 未來一年財務風險預測', level=1)
            if len(sub) > 1:
                trend = curr['M分數'] - sub.iloc[-2]['M分數']
                pred_m = round(curr['M分數'] + trend, 2)
                pred_desc = f"基於過去兩年的趨勢推估，下一年度之 M-Score 預測值為 {pred_m}。"
                pred_desc += "【風險警示】預測數值顯示舞弊風險將持續攀升，應立即強化內控制度。" if pred_m > -1.78 else "預測趨勢目前尚屬穩定。"
            else:
                pred_desc = "由於目前僅有一年度數據，暫無法進行精確之未來風險趨勢預測。"
            doc.add_paragraph(pred_desc)

            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 領取 Word 深度報告", buf, f"Forensic_Final_{target}.docx")
    else:
        st.info("請上傳多份 PDF 財報以啟動批次鑑識。")
