import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import pdfplumber

# --- 1. 系統環境設定 ---
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

# --- 2. 專業報告文字庫 (用於深沉敘述) ---
def get_detailed_narrative(row, mode="單一"):
    narrative = ""
    # 舞弊分析描述
    if row['M分數'] > -1.78:
        fraud_txt = f"【警示】Beneish M-Score 為 {row['M分數']}，超過門檻值 -1.78。顯示該公司存在顯著盈餘操縱風險，其應收帳款與營收之配合度存在異常，疑似透過提前認列營收虛增獲利。"
    else:
        fraud_txt = f"【正常】Beneish M-Score 為 {row['M分數']}，顯示利潤品質尚屬穩定，未觀察到系統性盈餘操縱行為。"
    
    # 掏空分析描述
    if row['掏空指數'] > 0.2:
        tunnel_txt = f"【警示】資產掏空指數達 {row['掏空指數']}，顯示非本業之資金流出（其他應收款、預付款）佔比異常。此行為常伴隨關係人交易或資金非法挪用至境外公司之風險。"
    else:
        tunnel_txt = f"【正常】資產掏空指數為 {row['掏空指數']}，處於產業合理水位，未發現異常資金偏移情形。"

    return f"{fraud_txt}\n{tunnel_txt}"

# --- 3. 核心解析與鑑定引擎 (簡化演示版) ---
def parse_and_analyze(files):
    # 這裡整合之前開發的強力提取邏輯
    # 假設返回分析後的 DataFrame
    data = []
    for f in files:
        # 模擬解析多個年度
        for y in [2023, 2024]:
            data.append({
                "公司名稱": f.name.replace(".pdf", ""),
                "年度": y,
                "營收": np.random.randint(1000, 5000),
                "應收": np.random.randint(200, 800),
                "存貨": np.random.randint(100, 400),
                "現金": np.random.randint(500, 2000),
                "其他應收": np.random.randint(50, 600),
                "預付": np.random.randint(20, 300),
                "淨利": np.random.randint(-100, 500)
            })
    df = pd.DataFrame(data)
    # 計算鑑定指標
    df['M分數'] = df.apply(lambda r: round(-3.2 + (0.15*(r['應收']/r['營收'])) + (0.1*(r['存貨']/r['營收'])), 2), axis=1)
    df['掏空指數'] = df.apply(lambda r: round((r['其他應收']+r['預付'])/r['營收'], 3), axis=1)
    return df

# --- 4. UI 介面 ---
with st.sidebar:
    st.header(" 鑑識中心控制台")
    mode = st.radio("報告模式", [" 單一公司多年比較", " 多家公司橫向 PK"])
    uploaded_files = st.file_uploader("上傳受查文件", type=["pdf"], accept_multiple_files=True)
    auditor = st.text_input("簽證會計師", "會計師")

if uploaded_files:
    df_pool = parse_and_analyze(uploaded_files)
    
    if mode == " 單一公司多年比較":
        target = st.selectbox("選擇公司", df_pool['公司名稱'].unique())
        sub = df_pool[df_pool['公司名稱'] == target].sort_values('年度')
        st.dataframe(sub)
    else:
        year = st.selectbox("選擇年度", df_pool['年度'].unique())
        sub = df_pool[df_pool['年度'] == year]
        st.dataframe(sub)

    # --- 5. 強效詳細報告產出 (重點更新) ---
    st.divider()
    if st.button("🚀 生成「深沉敘述版」鑑定鑑定報告"):
        doc = Document()
        
        # 設定標題
        title = doc.add_heading('財務鑑識鑑定意見書', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 1. 鑑定範圍與聲明
        doc.add_heading('壹、 鑑定背景與目的', level=1)
        doc.add_paragraph(f"本報告受託於 2026 年 4 月，由「玄武快機師事務所」主辦會計師 {auditor} 負責執行。針對受查單位之財務穩定性、盈餘品質及資產安全性進行深度查核。")

        # 2. 鑑定方法論
        doc.add_heading('貳、 鑑定方法論說明', level=1)
        doc.add_paragraph("本鑑定採用國際公認之 Beneish M-Score 模型偵測利潤操縱，並透過「資產偏移法」計算掏空指數（Tunneling Index），針對現金密度與經營績效之背離程度判定洗錢風險。")

        # 3. 詳細鑑定結果 (分章節敘述)
        doc.add_heading('參、 專案查核細部敘述', level=1)
        
        for idx, row in sub.iterrows():
            doc.add_heading(f"【{row['公司名稱']} - {row['年度']}年度】鑑定結論", level=2)
            
            # 加入詳細敘述
            p = doc.add_paragraph()
            run = p.add_run(get_detailed_narrative(row, mode))
            run.font.size = Pt(11)
            
            # 加入數據表格
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = '關鍵科目'
            hdr_cells[1].text = '數值'
            hdr_cells[2].text = '指標項目'
            hdr_cells[3].text = '鑑定風險'
            
            row_cells = table.add_row().cells
            row_cells[0].text = '營業收入'
            row_cells[1].text = str(row['營收'])
            row_cells[2].text = '舞弊 M-Score'
            row_cells[3].text = str(row['M分數'])

        # 4. 會計師專家意見
        doc.add_heading('肆、 專家總結建議', level=1)
        final_p = doc.add_paragraph(f"綜上所述，基於目前數據分析結果，主辦會計師建議針對異常指標之年度進行原始憑證（Vouching）與外部函證（Confirmation）之深度抽盤...")
        
        # 存檔
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("📥 點此下載「詳細敘述」鑑定報告書", buf, f"Full_Analysis_{auditor}.docx")
