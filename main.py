import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl

# V13.9.0 雲端終極穩定版：修正格式衝突 + 精確出勤統計
st.set_page_config(page_title="化石先生(JoJo)：雲端工時分析系統", layout="wide")

def process_data_v13_9_0(file):
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
                            # 判定是否有效出勤 (非空、非休、且有工時)
                            is_working_day = 1 if (shift not in ["休", "nan", ""] and work_h > 0) else 0
                            
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
                                    '加班': round(max(work_h - 8.0, 0), 1),
                                    '出勤計算': is_working_day
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

# --- UI ---
st.title("🛡️ 化石先生(JoJo)：雲端工時分析系統")
st.info("系統校準：V13.9.0 穩定版已啟動。摘要移除當日工時，改計總工作天數；休假行自動染粉。")

uploaded_file = st.file_uploader("導入原始班表 Excel", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 啟動衛星連線分析"):
        month_dict = process_data_v13_9_0(uploaded_file)
        if month_dict:
            st.success("數據掃描與視覺格式校準完成。")
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                wb = writer.book
                
                # 預先定義所有格式物件 (避免迴圈內重複建立)
                head_f = wb.add_format({'bold': 1, 'font_color': 'blue', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
                sum_body_f = wb.add_format({'border': 1, 'num_format': '0.0', 'align': 'left'})
                pink_bg = "#FFC0CB"
                
                for month, data in month_dict.items():
                    safe_m = str(month)[:25]
                    
                    # 摘要計算：取消當日工時，增加總工作天數
                    summary = data.groupby('人員').agg({
                        '出勤計算': 'sum',
                        '正常(8h)': 'sum',
                        '加班': 'sum',
                        '月總工時': 'max'
                    }).reset_index()
                    summary.rename(columns={'出勤計算': '總工作天數'}, inplace=True)
                    
                    # --- 寫入明細頁 ---
                    sheet_d = f"{safe_m}_明細"
                    display_data = data.drop(columns=['出勤計算'])
                    display_data.to_excel(writer, index=False, sheet_name=sheet_d)
                    ws_d = writer.sheets[sheet_d]
                    
                    # 寫入明細標頭
                    for c_idx, col in enumerate(display_data.columns):
                        ws_d.write(0, c_idx, col, head_f)
                    
                    colors = ['#F0F2F6', '#E1F5FE', '#E8F5E9', '#FFFDE7', '#F3E5F5', '#EFEBE9']
                    p_map = {p: colors[i%6] for i, p in enumerate(display_data['人員'].unique())}
                    
                    for r_idx in range(len(display_data)):
                        row = display_data.iloc[r_idx]
                        is_off = (str(row['班次']) == "休")
                        base_bg = pink_bg if is_off else p_map.get(row['人員'])
                        
                        # 動態建立當前行格式
                        f_std = wb.add_format({'bg_color': base_bg, 'border': 1, 'num_format': '0.0', 'align': 'left'})
                        f_red = wb.add_format({'bg_color': base_bg, 'border': 1, 'num_format': '0.0', 'font_color': 'red', 'bold': 1, 'align': 'left'})
                        
                        for c_idx, col_n in enumerate(display_data.columns):
                            val = row[col_n]
                            is_wknd = (col_n == '星期' and val in ['週六', '週日'])
                            ws_d.write(r_idx + 1, c_idx, val, f_red if is_wknd else f_std)
                    ws_d.set_column('A:L', 15)
                    
                    # --- 寫入摘要頁 ---
                    sheet_s = f"{safe_m}_摘要"
                    summary.to_excel(writer, index=False, sheet_name=sheet_s)
                    ws_s = writer.sheets[sheet_s]
                    
                    for c_idx, col in enumerate(summary.columns):
                        ws_s.write(0, c_idx, col, head_f)
                    
                    for r_idx in range(len(summary)):
                        for c_idx in range(len(summary.columns)):
                            ws_s.write(r_idx + 1, c_idx, summary.iloc[r_idx, c_idx], sum_body_f)
                    ws_s.set_column('A:E', 15)

            st.download_button("📥 下載化石先生最終校準報表", output_excel.getvalue(), "化石先生整合報告.xlsx")
