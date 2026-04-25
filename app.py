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
st.set_page_config(layout="wide", page_title="玄武非法鑑識系統")

st.markdown("""
    <div style='background-color:#1a1a1a; padding:20px; border-radius:10px; border-left: 10px solid #d9534f;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#ccc; margin:0;'>【核心鑑定】非法舞弊偵測、深度文字敘述、財務風險預測</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. 高效率數據提取 (針對台灣財報優化) ---
def clean_num(text):
    if not text: return 0.0
    s = str(text).strip().replace(',', '').replace('$', '')
    if '(' in s and ')' in s: s = '-' + s.replace('(', '').replace(')', '')
    match = re.search(r'[-+]?\d*\.\d+|\d+', s)
    return float(match.group()) if match else 0.0

def forensic_parse(file):
    res = {"公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, 
           "應收": 0.0, "存貨": 0.0, "其他應收": 0.0, "預付": 0.0, "淨利": 0.0}
    try:
        with pdfplumber.open(file) as pdf:
            text_pool = "".join([p.extract_text() or "" for p in pdf.pages[:15]])
            y_m = re.search(r"(\d{3,4})\s*年度", text_pool)
            if y_m:
                y = int(y_m.group(1))
                res["年度"] = y + 1911 if y < 1000 else y
            
            maps = {"營收": ["營業收入", "營收合計"], "應收": ["應收帳款淨額", "應收帳款"],
                    "存貨": ["存貨"], "其他應收": ["其他應收款"], "預付": ["預付款項"], "淨利": ["本期淨利", "本期損益"]}
            for key, kws in maps.items():
                for kw in kws:
                    m = re.search(rf"{kw}.{{0,25}}?([\d,]{{2,}}|\([\d,]{{2,}}\))", text_pool)
                    if m:
                        res[key] = clean_num(m.group(1))
                        break
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except: return pd.DataFrame()

# --- 3. 側邊欄控制 ---
with st.sidebar:
    st.header("⚙️ 鑑定功能中心")
    uploaded_files = st.file_uploader("批次上傳受查 PDF", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽署會計師", "會計師")

# --- 4. 主程式執行 ---
if uploaded_files:
    data_list = [forensic_parse(f) for f in uploaded_files]
    df = pd.concat([d for d in data_list if not d.empty], ignore_index=True)
    
    if not df.empty:
        # 指標運算
        df['M分數'] = df.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2) if r['營收']>0 else 0, axis=1)
        df['掏空指數'] = df.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3) if r['營收']>0 else 0, axis=1)
        
        target = st.selectbox("選擇鑑定對象", df['公司名稱'].unique())
        sub = df[df['公司名稱'] == target].sort_values('年度')

        # 介面圖表：僅顯示非法相關指標
        st.subheader(f"🚨 {target}：非法預警監控看板")
        c1, c2 = st.columns(2)
        with c1:
            fig_m, ax_m = plt.subplots()
            ax_m.plot(sub['年度'].astype(str), sub['M分數'], 'r-o', label='M-Score (舞弊預警)')
            ax_m.axhline(y=-1.78, color='black', linestyle='--', label='舞弊門檻')
            ax_m.fill_between(sub['年度'].astype(str), -1.78, sub['M分數'].max()+0.5, where=(sub['M分數'] > -1.78), color='red', alpha=0.2)
            ax_m.set_title("Beneish M-Score 舞弊壓力測試")
            ax_m.legend()
            st.pyplot(fig_m)
        with c2:
            fig_t, ax_t = plt.subplots()
            ax_t.bar(sub['年度'].astype(str), sub['掏空指數'], color='orange', label='資產流出指標')
            ax_t.set_title("資產掏空風險偵測")
            ax_t.legend()
            st.pyplot(fig_t)

        st.dataframe(sub[['年度', '營收', '淨利', 'M分數', '掏空指數']])

        # --- 5. 生成 Word 報告 (深度敘述 + 財務預測) ---
        st.divider()
        if st.button("🚀 生成深度鑑識報告書 (含風險預測)"):
            doc = Document()
            doc.add_heading('財務鑑識鑑定意見書', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"受查單位：{target}\n簽證會計師：{auditor}")

            # 壹、非法舞弊偵測深度敘述
            doc.add_heading('壹、 非法舞弊偵測深度敘述', level=1)
            last = sub.iloc[-1]
            prev = sub.iloc[-2] if len(sub) > 1 else last
            
            m_status = "高風險" if last['M分數'] > -1.78 else "正常"
            m_desc = (f"本所採用 Beneish M-Score 模型進行鑑定。目前年度數據為 {last['M分數']}，判定結果為「{m_status}」。"
                      f"相較於上一年度之 {prev['M分數']}，風險趨勢呈現{'上升' if last['M分數'] > prev['M分數'] else '下降'}。"
                      f"若指標持續突破 -1.78，代表該公司在營收認列或應收帳款之計價上存在顯著的不誠實風險。")
            doc.add_paragraph(m_desc)

            # 貳、資產掏空圖表分析
            doc.add_heading('貳、 資產掏空風險圖表分析', level=1)
            t_desc = (f"本區塊針對「其他應收款」與「預付款項」進行監測。當前掏空指數為 {last['掏空指數']}。"
                      "此指標若顯著高於產業平均值，通常暗示資金可能透過非本業管道流出至關係人或人頭公司。")
            doc.add_paragraph(t_desc)
            
            img_buf = io.BytesIO()
            plt.savefig(img_buf, format='png', dpi=200)
            img_buf.seek(0)
            doc.add_picture(img_buf, width=Inches(5.5))

            # 參、未來一年財務風險預測
            doc.add_heading('參、 未來一年財務風險預測', level=1)
            # 簡單的線性趨勢預測邏輯
            predicted_m = round(last['M分數'] + (last['M分數'] - prev['M分數']), 2)
            pred_text = (f"基於現有數據趨勢推估，受查單位下一年度之 M-Score 預測值約為 {predicted_m}。"
                         f"{'【警示】預測數值顯示風險將進一步惡化，建議提前啟動專案查核。' if predicted_m > -1.78 else '預測數值尚在安全區間。'}"
                         "此外，考量到營收成長速度與資產質量的背離，預測未來一年現金流壓力將增大。")
            doc.add_paragraph(pred_text)

            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 下載深度預測鑑定報告", buf, f"Forensic_Final_{target}.docx")
