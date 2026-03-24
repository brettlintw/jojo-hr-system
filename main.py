import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
import openpyxl

# V13.8.7 雲端藍調標頭 + 全域靠左版
st.set_page_config(page_title="化石先生(JoJo)：雲端工時分析系統", layout="wide")

def process_data_v13_8_7(file):
    try:
        file.seek(0)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
        month_data_dict = {}
        for sheet_name, df in all_sheets.items():
            header_idx = -1
            for i, row in df.iterrows():
                row_str = "".join(str(v) for v in row.values)
                if '人員' in row_str and '日期' in row_str:
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
                            shift = str(rows.iloc[idx, col_idx]).strip()
                            if shift != "nan" or work_h > 0:
                                dt_obj = pd.to_datetime(target_date)
                                person_records.append({
                                    '人員': person, '日期': target_date, '星期': f"週{['一','二','三','四','五','六','日'][dt_obj.weekday()]}",
                                    '班次': shift if shift != "nan" else "",
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

# --- UI 介面 ---
st.title("🛡️ 化石先生(JoJo)：雲端工時分析系統")
st.info("系統校準完畢：標頭藍字鎖定，所有儲存格數據已強制靠左對齊。")

uploaded_file = st.file_uploader("導入原始班表 Excel", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 啟動衛星連線分析"):
        month_dict = process_data_v13_8_7(uploaded_file)
        if month_dict:
            st.success("數據掃描與靠左渲染完成。")
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                wb = writer.book
                
                # 1. 定義標頭格式 (粗體、藍字、靠左)
                header_fmt = wb.add_format({
                    'bold': True,
                    'font_color': 'blue',
                    'border': 1,
                    'align': 'left',
                    'valign': 'vcenter'
                })
                
                for month, data in month_dict.items():
                    safe_m = str(month)[:25]
                    summary = data.groupby('人員').agg({'當日工時':'sum', '正常(8h)':'sum', '加班':'sum', '月總工時':'max'}).reset_index()
                    
                    # --- 寫入明細頁 ---
                    sheet_detail = f"{safe_m}_明細"
                    data.to_excel(writer, index=False, sheet_name=sheet_detail)
                    ws_d = writer.sheets[sheet_detail]
                    
                    # 強制重新寫入標頭 (靠左藍字)
                    for c_idx, col_val in enumerate(data.columns):
                        ws_d.write(0, c_idx, col_val, header_fmt)
                    
                    # 設定內容格式
                    unique_p = data['人員'].unique()
                    colors = ['#F0F2F6', '#E1F5FE', '#E8F5E9', '#FFFDE7', '#F3E5F5', '#EFEBE9']
                    p_map = {p: colors[i%len(colors)] for i, p in enumerate(unique_p)}
                    
                    for r_idx in range(len(data)):
                        p_bg = p_map.get(data.iloc[r_idx]['人員'])
                        # 內容標準格式 (靠左)
                        std_f = wb.add_format({'bg_color': p_bg, 'border': 1, 'num_format': '0.0', 'align': 'left'})
                        # 內容紅字格式 (靠左)
                        red_f = wb.add_format({'bg_color': p_bg, 'border': 1, 'num_format': '0.0', 'font_color': 'red', 'bold': True, 'align': 'left'})
                        
                        for c_idx, col_name in enumerate(data.columns):
                            val = data.iloc[r_idx][col_name]
                            is_wk = (col_name == '星期' and val in ['週六', '週日'])
                            ws_d.write(r_idx + 1, c_idx, val, red_f if is_wk else std_f)
                    ws_d.set_column('A:L', 15)
                    
                    # --- 寫入摘要頁 ---
                    sheet_sum = f"{safe_m}_摘要"
                    summary.to_excel(writer, index=False, sheet_name=sheet_sum)
                    ws_s = writer.sheets[sheet_sum]
                    
                    # 摘要頁標頭 (靠左藍字)
                    for c_idx, col_val in enumerate(summary.columns):
                        ws_s.write(0, c_idx, col_val, header_fmt)
                    
                    # 摘要頁內容 (靠左)
                    sum_f = wb.add_format({'border': 1, 'num_format': '0.0', 'align': 'left'})
                    for r_idx in range(len(summary)):
                        for c_idx in range(len(summary.columns)):
                            ws_s.write(r_idx + 1, c_idx, summary.iloc[r_idx, c_idx], sum_f)
                    ws_s.set_column('A:E', 15)

            st.download_button("📥 下載靠左對齊報告", output_excel.getvalue(), "化石先生報告.xlsx")
