import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import io
import re
import pdfplumber

# 1. 系統環境設定
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>旗艦鑑定平台：舞弊、掏空與洗錢偵測模組</p>
    </div>
""", unsafe_allow_html=True)

# 2. 強效數值提取器：能處理 (1,234), 1.234,56 以及格式混亂的文字
def force_extract_numbers(row_list):
    vals = []
    for item in row_list:
        if not item: continue
        s = str(item).strip().replace(',', '')
        # 處理會計負數格式 (xxx)
        if '(' in s and ')' in s:
            s = '-' + s.replace('(', '').replace(')', '')
        # 提取純數字、負號與小數點
        clean_s = "".join(re.findall(r'[0-9\.-]+', s))
        try:
            val = float(clean_s)
            vals.append(val)
        except:
            continue
    return vals

# 3. 終極 PDF/Excel 解析引擎
def parse_financial_data(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, 
            "應收帳款": 0.0, "存貨": 0.0, "現金": 0.0, "負債總額": 0.0, 
            "其他應收款": 0.0, "預付款項": 0.0
        }
        
        if file.name.endswith('.xlsx'):
            df = pd.read_excel(file)
            # 簡單處理 Excel 欄位對齊
            for col in df.columns:
                c_str = str(col)
                if "年" in c_str: res["年度"] = df[col].iloc[0]
                if "營" in c_str: res["營收"] = df[col].iloc[0]
            return pd.DataFrame([res])

        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:6]:
                # 抓取年度
                text = page.extract_text() or ""
                if res["年度"] == 0:
                    years = re.findall(r"(\d{3,4})\s*年度", text)
                    if years:
                        y = int(years[0])
                        res["年度"] = y + 1911 if y < 1000 else y

                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        row_str = "".join([str(x) for x in row if x])
                        nums = force_extract_numbers(row)
                        if not nums: continue
                        
                        # 根據關鍵字鎖定數值（取該行最後一個數字，通常是當期數）
                        val = nums[-1]
                        if any(k in row_str for k in ["營業收入", "營收"]): res["營收"] = val
                        if "應收帳款" in row_str and "其他" not in row_str: res["應收帳款"] = val
                        if "存貨" in row_str: res["存貨"] = val
                        if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = val
                        if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = val
                        if "其他應收款" in row_str: res["其他應收款"] = val
                        if "預付款項" in row_str: res["預付款項"] = val
        
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except:
        return pd.DataFrame()

# 4. 鑑識邏輯計算
def run_forensics(df):
    for c in ['M分數', '舞弊風險', '掏空指數', '結論']: df[c] = ""
    for i in df.index:
        r = float(df.at[i, '營收'])
        if r > 0:
            m = -3.2 + (0.15 * (df.at[i, '應收帳款']/r)) + (0.1 * (df.at[i, '存貨']/r))
            df.at[i, 'M分數'] = round(m, 2)
            df.at[i, '舞弊風險'] = "危險" if m > -1.78 else "正常"
            t_idx = (df.at[i, '其他應收款'] + df.at[i, '預付款項']) / r
            df.at[i, '掏空指數'] = round(t_idx, 3)
            df.at[i, '結論'] = "建議抽核" if (m > -1.78 or t_idx > 0.2) else "無明顯異常"
    return df

# 5. 主介面
with st.sidebar:
    st.header("控制台")
    uploaded_files = st.file_uploader("批次上傳財報", type=["pdf", "xlsx"], accept_multiple_files=True)
    st.divider()
    auditor = st.text_input("主辦會計師簽署", "張鈞翔會計師")

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        data = parse_financial_data(f)
        if not data.empty: all_data.append(data)
    
    if all_data:
        df_final = pd.concat(all_data, ignore_index=True).sort_values('年度')
        df_final = run_forensics(df_final)
        
        target = st.selectbox("受查公司切換", df_final['公司名稱'].unique())
        sub = df_final[df_final['公司名稱'] == target]
        
        # 視覺化看板
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Revenue Trend (營收動能)")
            fig1, ax1 = plt.subplots(figsize=(6, 4))
            ax1.plot(sub['年度'].astype(str), sub['營收'], marker='o', linewidth=2, color='#268bd2')
            if sub['營收'].max() > 0: ax1.set_ylim(0, sub['營收'].max() * 1.2)
            st.pyplot(fig1)

        with c2:
            st.subheader("Asset Risk (資產偏移)")
            fig2, ax2 = plt.subplots(figsize=(6, 4))
            ax2.bar(sub['年度'].astype(str), sub['其他應收款'], label='Other RCV', color='#d33682')
            ax2.bar(sub['年度'].astype(str), sub['預付款項'], bottom=sub['其他應收款'], label='Prepay', color='#b58900')
            ax2.legend()
            st.pyplot(fig2)

        st.write("鑑定數據清單")
        st.dataframe(sub)

        # --- WORD 報告生成按鈕 (確保出現在最下方) ---
        st.divider()
        st.subheader("報告產出中心")
        
        doc = Document()
        doc.add_heading(f"玄武鑑識鑑定報告 - {target}", 0)
        doc.add_paragraph(f"主辦會計師：{auditor}")
        doc.add_paragraph(f"鑑定結論：{sub.iloc[-1]['結論']}")
        
        # 將數據表加入 Word
        t = doc.add_table(rows=1, cols=4)
        t.style = 'Table Grid'
        hdr_cells = t.rows[0].cells
        hdr_cells[0].text = '年度'
        hdr_cells[1].text = '營收'
        hdr_cells[2].text = 'M分數'
        hdr_cells[3].text = '掏空指數'
        
        for _, row in sub.iterrows():
            row_cells = t.add_row().cells
            row_cells[0].text = str(row['年度'])
            row_cells[1].text = str(row['營收'])
            row_cells[2].text = str(row['M分數'])
            row_cells[3].text = str(row['掏空指數'])

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        
        st.download_button(
            label="生成並下載鑑定報告 (.docx)",
            data=buf,
            file_name=f"Forensic_Report_{target}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("檔案讀取中或未能抓取有效數值。請確認 PDF 內容是否為可選取文字。")
else:
    st.info("系統就緒，請於左側控制台批次上傳受查財報。")
