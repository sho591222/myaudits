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

# --- 0. 圖表顯示修正 (解決空白問題) ---
# 強制設定：優先使用中文，若無則使用通用字體，避免渲染失敗
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Arial Unicode MS', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False 

# --- 1. 介面設定 ---
st.set_page_config(layout="wide", page_title="玄武極速鑑識")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #d9534f;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>【圖表強制修復版】非法預警線圖、深度趨勢分析、未來風險預測</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. 暴力解析引擎 ---
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
            y_m = re.search(r"(\d{3,4})\s*年度", text)
            if y_m:
                y = int(y_m.group(1))
                res["年度"] = y + 1911 if y < 1000 else y
            
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

# --- 3. 側邊欄 ---
with st.sidebar:
    st.header("⚡ 鑑識控制台")
    uploaded_files = st.file_uploader("批次上傳多個 PDF", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "會計師")

# --- 4. 主程式：非法看板 ---
if uploaded_files:
    dfs = [batch_parse(f) for f in uploaded_files]
    df_full = pd.concat([d for d in dfs if not d.empty], ignore_index=True)
    
    if not df_full.empty:
        # 計算指標
        df_full['M-Score'] = df_full.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
        df_full['掏空指數'] = df_full.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
        
        target = st.selectbox("選擇鑑定對象", df_full['公司'].unique())
        sub = df_full[df_full['公司'] == target].sort_values('年度')

        st.subheader(f"🚨 非法預警監控看板：{target}")
        
        # --- 線圖與圖表渲染區 (核心修正) ---
        c1, c2 = st.columns(2)
        
        with c1:
            st.write("#### 舞弊預警 M-Score 線圖")
            fig1 = plt.figure(figsize=(8, 5))
            plt.plot(sub['年度'].astype(str), sub['M-Score'], marker='o', color='red', linewidth=2, label='M-Score')
            plt.axhline(y=-1.78, color='black', linestyle='--', label='風險門檻 (-1.78)')
            plt.fill_between(sub['年度'].astype(str), -1.78, max(sub['M-Score'])+0.5, color='red', alpha=0.1)
            plt.xlabel("年度 (Year)")
            plt.ylabel("舞弊分數 (M-Score)")
            plt.legend()
            st.pyplot(fig1) # 使用獨立 fig 物件渲染
            
        with c2:
            st.write("#### 資產挪用(掏空)壓力圖")
            fig2 = plt.figure(figsize=(8, 5))
            plt.bar(sub['年度'].astype(str), sub['掏空指數'], color='orange', alpha=0.8, label='掏空指標')
            plt.xlabel("年度 (Year)")
            plt.ylabel("掏空比率 (Expropriation Index)")
            plt.legend()
            st.pyplot(fig2)

        st.write("### 鑑識原始數據清單")
        st.dataframe(sub)

        # --- 5. Word 深度鑑定報告 ---
        if st.button("🚀 生成深度鑑定意見書 (含風險預測)"):
            doc = Document()
            doc.add_heading('財務鑑識鑑定意見書', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_heading('壹、 舞弊與非法預警鑑定', level=1)
            curr = sub.iloc[-1]
            prev = sub.iloc[-2] if len(sub) > 1 else curr
            risk = "【高度警示】" if curr['M-Score'] > -1.78 else "【正常】"
            doc.add_paragraph(f"當前 M-Score 為 {curr['M-Score']}，判定為 {risk}。趨勢較前一年度呈現 {'惡化' if curr['M-Score'] > prev['M-Score'] else '改善'}。")

            # 插入圖表到 Word
            img_buf = io.BytesIO()
            fig1.savefig(img_buf, format='png', dpi=200) # 儲存線圖
            img_buf.seek(0)
            doc.add_picture(img_buf, width=Inches(5.0))

            doc.add_heading('貳、 未來一年財務風險預測', level=1)
            pred_m = round(curr['M-Score'] + (curr['M-Score'] - prev['M-Score']), 2)
            doc.add_paragraph(f"基於趨勢模型推估，明年度預測值為 {pred_m}。建議針對「應收帳款」科目執行實質性抽查。")
            
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 下載深度鑑定意見書", buf, f"Forensic_Report_{target}.docx")
    else:
        st.info("請上傳 PDF 開始解析數據。")
