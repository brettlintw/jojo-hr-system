import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl

# 化石先生(JoJo) 雲端穩定版
st.set_page_config(page_title="化石先生(JoJo)：工時分析系統", layout="wide")

def process_data(file):
    try:
        file.seek(0)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
        month_data_dict = {}
        for sheet_name, df in all_sheets.items():
            header_idx = -1
            for i, row in df.iterrows():
                if '人員' in str(row.values) and '日期' in str(row.values):
                    header_idx = i
                    break
            if header_idx == -1: continue
            
            dates_raw = [str(d).split(' ')[0] if pd.notnull(d) else "" for d in df.iloc[header_idx]]
            rows = df.iloc[header_idx + 1:].reset_index(drop=True)
            all_records = []
            
            idx = 0
            while idx < len(rows):
                person = str(rows.iloc[idx, 0]).strip()
                if person not in ["", "nan", "None"]:
                    person_records = []
                    for col_idx in range(2, len(dates_raw)):
                        target_date = dates_raw[col_idx]
                        if len(target_date) < 5: continue
                        try:
                            work_v = rows.iloc[idx+3, col_idx]
                            work_h = round(float(work_v), 1) if pd.notnull(work_v) else 0.0
                            if str(rows.iloc[idx, col_idx]).strip() != "nan" or work_h > 0:
                                dt_obj = pd.to_datetime(target_date)
                                person_records.append({
                                    '人員': person, '日期': target_date, '星期': f"週{['一','二','三','四','五','六','日'][dt_obj.weekday()]}",
                                    '班次': str(rows.iloc[idx, col_idx]).strip(),
                                    '上班': str(rows.iloc[idx+1, col_idx]).strip()[:5],
                                    '下班': str(rows.iloc[idx+2, col_idx]).strip()[:5],
                                    '當日工時': work_h,
                                    '用餐': round(float(rows.iloc[idx+4, col_idx]), 1) if pd.notnull(rows.iloc[idx+4, col_idx]) else 0.0,
                                    '備註': str(rows.iloc[idx+5, col_idx]).strip() if pd.notnull(rows.iloc[idx+5, col_idx]) else "",
                                    '正常(8h)': round(min(work_h, 8.0), 1),
                                    '加班': round(max(work_h - 8.0, 0), 1)
                                })
                        except: pass
                    total_m = sum(r['當日工時'] for r in person_records)
                    for r in person_records: r['月總工時'] = round(total_m, 1)
                    all_records.extend(person_records)
                    idx += 6 
                else: idx += 1
            if all_records: month_data_dict[str(sheet_name)] = pd.DataFrame(all_records)
        return month_data_dict
    except: return None

st.title("🛡️ 化石先生(JoJo)：工時分析系統")
uploaded_file = st.file_uploader("導入原始班表 Excel", type=["xlsx"])

if uploaded_file and st.button("🚀 啟動分析"):
    month_dict = process_data(uploaded_file)
    if month_dict:
        st.success("分析完成")
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            wb = writer.book
            for month, data in month_dict.items():
                summary = data.groupby('人員').agg({'當日工時':'sum', '正常(8h)':'sum', '加班':'sum', '月總工時':'max'}).reset_index()
                data.to_excel(writer, index=False, sheet_name=f"{month}_明細")
                summary.to_excel(writer, index=False, sheet_name=f"{month}_摘要")
                ws = writer.sheets[f"{month}_明細"]
                p_map = {p: ['#F0F2F6', '#E1F5FE', '#E8F5E9', '#FFFDE7', '#F3E5F5', '#EFEBE9'][i%6] for i, p in enumerate(data['人員'].unique())}
                for r_idx in range(len(data)):
                    p_bg = p_map.get(data.iloc[r_idx]['人員'])
                    fmt = wb.add_format({'bg_color': p_bg, 'border': 1, 'num_format': '0.0'})
                    red_fmt = wb.add_format({'bg_color': p_bg, 'border': 1, 'num_format': '0.0', 'font_color': 'red', 'bold': True})
                    for c_idx, col in enumerate(data.columns):
                        is_red = (col == '星期' and data.iloc[r_idx][col] in ['週六', '週日'])
                        ws.write(r_idx + 1, c_idx, data.iloc[r_idx][col], red_fmt if is_red else fmt)
                ws.set_column('A:L', 15)
        st.download_button("📥 下載整合報告", output_excel.getvalue(), "化石先生報告.xlsx")
