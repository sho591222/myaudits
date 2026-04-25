import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
import io
import re
import pdfplumber

# 1. 系統環境設定
st.set_page_config(layout="wide", page_title="玄武鑑識中心")

st.markdown("""
    <div style='background-color:#002b36; padding:20px; border-radius:10px; border-left: 10px solid #b58900;'>
        <h1 style='color:white; margin:0;'>玄武快機師事務所</h1>
        <p style='color:#839496; margin:0;'>旗艦鑑識平台：舞弊、掏空、洗錢與財報不實自動化分析系統</p>
    </div>
""", unsafe_allow_html=True)

# 2. 強效數值提取與科目擴充
def force_extract_numbers(row_list):
    vals = []
    for item in row_list:
        if not item: continue
        s = str(item).strip().replace(',', '')
        if '(' in s and ')' in s: s = '-' + s.replace('(', '').replace(')', '')
        clean_s = "".join(re.findall(r'[0-9\.-]+', s))
        try:
            val = float(clean_s)
            vals.append(val)
        except: continue
    return vals

def parse_financial_data(file):
    try:
        res = {
            "公司名稱": file.name.replace(".pdf", ""), "年度": 0, "營收": 0.0, 
            "應收帳款": 0.0, "存貨": 0.0, "現金": 0.0, "負債總額": 0.0, 
            "其他應收款": 0.0, "預付款項": 0.0, "股份酬勞": 0.0, "總資產": 0.0, "淨利": 0.0
        }
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages[:8]:
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
                        val = nums[-1]
                        if any(k in row_str for k in ["營業收入", "營收"]): res["營收"] = val
                        if "應收帳款" in row_str and "其他" not in row_str: res["應收帳款"] = val
                        if "存貨" in row_str: res["存貨"] = val
                        if any(k in row_str for k in ["現金", "約當現金"]): res["現金"] = val
                        if any(k in row_str for k in ["負債總額", "負債合計"]): res["負債總額"] = val
                        if "其他應收款" in row_str: res["其他應收款"] = val
                        if "預付款項" in row_str: res["預付款項"] = val
                        if any(k in row_str for k in ["股份酬勞", "認股權"]): res["股份酬勞"] = val
                        if any(k in row_str for k in ["資產總額", "資產總計"]): res["總資產"] = val
                        if any(k in row_str for k in ["本期淨利", "本期損益"]): res["淨利"] = val
        return pd.DataFrame([res]) if res["年度"] > 0 else pd.DataFrame()
    except: return pd.DataFrame()

# 3. 鑑識核心邏輯 (四大維度)
def run_forensic_engine(df):
    for c in ['M分數', '掏空指數', '洗錢風險', '不實預警', '綜合結論']: df[c] = ""
    for i in df.index:
        r = float(df.at[i, '營收'])
        ca = float(df.at[i, '現金'])
        inv = float(df.at[i, '存貨'])
        ar = float(df.at[i, '應收帳款'])
        other_r = float(df.at[i, '其他應收款'])
        prepay = float(df.at[i, '預付款項'])
        ni = float(df.at[i, '淨利'])
        
        if r > 0:
            # A. 舞弊 M-Score (簡化版)
            m = -3.2 + (0.15 * (ar/r)) + (0.1 * (inv/r))
            df.at[i, 'M分數'] = round(m, 2)
            
            # B. 掏空指數 (非本業資金流出)
            t_idx = (other_r + prepay) / r
            df.at[i, '掏空指數'] = round(t_idx, 3)
            
            # C. 洗錢風險偵測 (現金密度異常)
            df.at[i, '洗錢風險'] = "高" if (ca > r * 0.8 and ni < 0) else "低"
            
            # D. 財報不實預警 (應計項目脫節)
            accrual_gap = abs(ni - (r * 0.1)) # 假設正常利潤率 10%
            df.at[i, '不實預警'] = "顯著" if accrual_gap / r > 0.3 else "輕微"
            
            # E. 綜合結論
            if m > -1.78 or t_idx > 0.2 or df.at[i, '洗錢風險'] == "高":
                df.at[i, '綜合結論'] = "重大異常：建議立即啟動專案查核"
            else:
                df.at[i, '綜合結論'] = "尚無重大顯著異常"
    return df

# 4. 側邊欄與預測設定
with st.sidebar:
    st.header("⚙️ 鑑定功能選單")
    view_mode = st.radio("模式選擇", ["🔍 單一深度報告", "⚔️ 多公司同年度 PK"])
    st.divider()
    
    st.subheader("📈 財務預測設定")
    forecast_method = st.selectbox("預測模型", ["線性成長 (5%)", "保守估計 (0%)", "歷史平均成長"])
    
    st.subheader("🛡️ 診斷範圍")
    f_fraud = st.toggle("舞弊偵測 (M-Score)", value=True)
    f_tunnel = st.toggle("掏空分析", value=True)
    f_aml = st.toggle("洗錢風險掃描", value=True)
    f_misstate = st.toggle("財報不實預警", value=True)
    
    st.divider()
    uploaded_files = st.file_uploader("上傳受查文件", type=["pdf", "xlsx"], accept_multiple_files=True)
    auditor_name = st.text_input("簽署會計師", "張鈞翔會計師")

# 5. 主程式與報告產出
if uploaded_files:
    df_pool = pd.concat([parse_financial_data(f) for f in uploaded_files], ignore_index=True)
    if not df_pool.empty:
        df_pool = run_forensic_engine(df_pool)
        
        if view_mode == "🔍 單一深度報告":
            target = st.selectbox("選擇受查公司", df_pool['公司名稱'].unique())
            sub = df_pool[df_pool['公司名稱'] == target].sort_values('年度')
            
            st.subheader(f"鑑識看板：{target}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("舞弊 M-Score", sub.iloc[-1]['M分數'])
            c2.metric("掏空指數", sub.iloc[-1]['掏空指數'])
            c3.metric("洗錢風險", sub.iloc[-1]['洗錢風險'])
            c4.metric("不實預警", sub.iloc[-1]['不實預警'])
            
            st.line_chart(sub.set_index('年度')[['營收', '應收帳款', '其他應收款']])
            st.dataframe(sub)

            # --- 超詳細 WORD 報告生成 ---
            st.divider()
            if st.button("生成「詳細鑑識鑑定報告書」"):
                doc = Document()
                doc.add_heading(f'財務鑑識鑑定報告書 - {target}', 0)
                
                # 第一章：基本資料
                doc.add_heading('壹、 鑑定基本資料', level=1)
                p = doc.add_paragraph()
                p.add_run(f'受查公司：{target}\n').bold = True
                p.add_run(f'鑑定年度：{sub["年度"].min()} - {sub["年度"].max()}\n')
                p.add_run(f'主辦會計師：{auditor_name}\n')
                p.add_run(f'報告日期：2026年4月25日\n')

                # 第二章：四大維度深度分析
                doc.add_heading('貳、 專案鑑定分析', level=1)
                
                if f_fraud:
                    doc.add_heading('一、 財務舞弊偵測 (Beneish M-Score)', level=2)
                    doc.add_paragraph(f"經偵測，該公司 M 分數為 {sub.iloc[-1]['M分數']}。若分數大於 -1.78，代表具有高度盈餘操縱傾向。")
                
                if f_tunnel:
                    doc.add_heading('二、 資產掏空分析 (Tunneling Index)', level=2)
                    doc.add_paragraph(f"該公司掏空指數為 {sub.iloc[-1]['掏空指數']}。主要分析其他應收款與預付款項是否脫輯營業常規。")
                
                if f_aml:
                    doc.add_heading('三、 洗錢風險掃描 (AML Screening)', level=2)
                    doc.add_paragraph(f"系統偵測結果：{sub.iloc[-1]['洗錢風險']}。分析重點在於現金密度與淨利之背離程度。")

                if f_misstate:
                    doc.add_heading('四、 財報不實預警 (Accounting Misstatement)', level=2)
                    doc.add_paragraph(f"該年度財報不實風險為：{sub.iloc[-1]['不實預警']}。針對應計項目與實際獲利能力進行交叉比對。")

                # 第三章：結論
                doc.add_heading('參、 鑑定結論與建議', level=1)
                conc = doc.add_paragraph(sub.iloc[-1]['綜合結論'])
                conc.runs[0].font.size = Pt(14)
                conc.runs[0].font.bold = True
                
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.download_button(f"📥 下載 {target} 詳盡鑑定報告 (.docx)", buf, f"Forensic_Report_{target}.docx")
        
        else: # 多公司 PK
            st.subheader("多公司同年度橫向 PK")
            year = st.selectbox("選擇 PK 年度", sorted(df_pool['年度'].unique(), reverse=True))
            pk_df = df_pool[df_pool['年度'] == year]
            st.bar_chart(pk_df.set_index('公司名稱')[['M分數', '掏空指數']])
            st.dataframe(pk_df)

    else: st.error("未能讀取有效數據。")
