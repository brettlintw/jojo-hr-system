import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl

# V13.9.8 雲端終極穩定版：自動休息判定 + 加班(扣除休息) + 多彩人員標籤
st.set_page_config(page_title="化石先生(JoJo)：雲端工時分析系統", layout="wide")

def process_data_v13_9_8(file):
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
                                # 1. 休息時間計算邏輯
                                rest_h = 0.0
                                if work_h >= 8.0:
                                    rest_h = 1.0
                                elif 4.0 < work_h < 8.0:
                                    rest_h = 0.5
                                
                                # 2. 加班計算邏輯：> 8.0 則加班 = (工時 - 8.0 - 休息時間)
                                over_h = 0.0
                                if work_h > 8.0:
                                    over_h = round(max(work_h - 8.0 - rest_h, 0.0), 1)
                                
                                dt_obj = pd.to_datetime(target_date)
                                person_records.append({
                                    '人員': person, '日期': target_date, '星期': f"週{['一','二','三','四','五','六','日'][dt_obj.weekday()]}",
                                    '班次': shift if shift != "nan" else "",
                                    '上班': str(rows.iloc[idx+1, col_idx]).strip()[:5],
                                    '下班': str(rows.iloc[idx+2, col_idx]).strip()[:5],
                                    '當日工時': work_h,
                                    '休息時間': rest_h,
                                    '用餐': round(float(rows.iloc[idx+4, col_idx]), 1) if pd.notnull(rows.iloc[idx+4, col_idx]) else 0.0,
                                    '備註': str(rows.iloc[idx+5, col_idx]).strip() if pd.notnull(rows.iloc[idx+5, col_idx]) else "",
                                    '正常(8h)': round(min(work_h, 8.0), 1),
                                    '加班': over_h,
                                    '出勤計算': 1 if (shift != "休" and work_h > 0) else 0
                                })
                        except: pass
                    
                    # 計算月總工時並同步至明細
                    total_m = sum(r['當日工時'] for r in person_records)
                    for r in person_records: r['月總工時'] = round(total_m, 1)
                    all_records.extend(person_records)
                    idx += 6 
                else: idx += 1
            if all_records: month_data_dict[str(sheet_name)] = pd.DataFrame(all_records)
        return month_data_dict
    except: return None

# --- UI ---
st.title("🛡️ 化石先生(JoJo)：雲端工時分析系統 (V13.8.8+)")
st.info("系統校準：V13.9.8 已啟動。休息時間自動判定，加班已扣除休息並紅字化，人員多彩識別。")

uploaded_file = st.file_uploader("導入原始班表 Excel", type=["xlsx"])

if uploaded_file and st.button("🚀 啟動衛星連線分析"):
    month_dict = process_data_v13_9_8(uploaded_file)
    if month_dict:
        st.success("數據掃描與視覺格式校準完成。")
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            wb = writer.book
            
            # 定義標頭與摘要格式
            head_f = wb.add_format({'bold': 1, 'font_color': 'blue', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
            sum_body_f = wb.add_format({'border': 1, 'num_format': '0.0', 'align': 'left'})
            
            for month, data in month_dict.items():
                safe_m = str(month)[:25]
                summary = data.groupby('人員').agg({'出勤計算': 'sum', '正常(8h)': 'sum', '加班': 'sum'}).reset_index()
                summary.rename(columns={'出勤計算': '總工作天數'}, inplace=True)
                summary['月總工時'] = summary['正常(8h)'] + summary['加班']
                
                # --- 明細頁 ---
                sheet_d = f"{safe_m}_明細"
                display_data = data.drop(columns=['出勤計算'])
                display_data.to_excel(writer, index=False, sheet_name=sheet_d)
                ws_d = writer.sheets[sheet_d]
                
                for c_idx, col in enumerate(display_data.columns):
                    ws_d.write(0, c_idx, col, head_f)
                
                # 不同人員分配不同顏色
                text_colors = ['#0000FF', '#008000', '#800080', '#FF8C00', '#008080', '#A52A2A', '#2F4F4F']
                p_unique = display_data['人員'].unique()
                p_text_map = {p: text_colors[i % len(text_colors)] for i, p in enumerate(p_unique)}
                
                for r_idx in range(len(display_data)):
                    row = display_data.iloc[r_idx]
                    is_off = (str(row['班次']) == "休")
                    p_name = row['人員']
                    
                    for c_idx, col_n in enumerate(display_data.columns):
                        val = row[col_n]
                        # 全域靠左對齊
                        fmt_p = {'border': 1, 'num_format': '0.0', 'align': 'left'}
                        
                        # 人員文字顏色
                        if col_n == '人員':
                            fmt_p['font_color'] = p_text_map.get(p_name)
                        
                        # 班次為「休」：整行紅字 (不含人員)
                        elif is_off:
                            fmt_p['font_color'] = '#FF0000'
                        
                        # 加班欄位紅字
                        elif col_n == '加班':
                            fmt_p['font_color'] = '#FF0000'
                        
                        # 週末紅字標示
                        elif col_n == '星期' and val in ['週六', '週日']:
                            fmt_p['font_color'] = '#FF0000'
                            fmt_p['bold'] = True

                        ws_d.write(r_idx + 1, c_idx, val, wb.add_format(fmt_p))
                
                ws_d.set_column('A:O', 15)
                
                # --- 摘要頁 ---
                sheet_s = f"{safe_m}_摘要"
                summary.to_excel(writer, index=False, sheet_name=sheet_s)
                ws_s = writer.sheets[sheet_s]
                for c_idx, col in enumerate(summary.columns):
                    ws_s.write(0, c_idx, col, head_f)
                for r_idx in range(len(summary)):
                    for c_idx in range(len(summary.columns)):
                        ws_s.write(r_idx + 1, c_idx, summary.iloc[r_idx, c_idx], sum_body_f)
                ws_s.set_column('A:E', 15)

        st.download_button("📥 下載 V13.9.8 終極校準報告", output_excel.getvalue(), "化石先生進階報告.xlsx")
