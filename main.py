import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl

# V14.0.1 雲端終極視覺版：摘要加班紅字 + 休假橫槓 + 紅底黑字警示
st.set_page_config(page_title="化石先生(JoJo)：雲端工時分析系統", layout="wide")

def process_data_v14_0_1(file):
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
                                # 1. 休息時間判定
                                rest_h = 0.0
                                if work_h >= 8.0: rest_h = 1.0
                                elif 4.0 < work_h < 8.0: rest_h = 0.5
                                
                                # 2. 加班計算
                                over_h = round(max(work_h - 8.0 - rest_h, 0.0), 1) if work_h > 8.0 else 0.0
                                
                                # 3. 休假符號化
                                is_off = (shift == "休")
                                start_t = "-" if is_off else str(rows.iloc[idx+1, col_idx]).strip()[:5]
                                end_t = "-" if is_off else str(rows.iloc[idx+2, col_idx]).strip()[:5]

                                dt_obj = pd.to_datetime(target_date)
                                person_records.append({
                                    '人員': person, '日期': target_date, '星期': f"週{['一','二','三','四','五','六','日'][dt_obj.weekday()]}",
                                    '班次': shift if shift != "nan" else "",
                                    '上班': start_t, '下班': end_t,
                                    '當日工時': work_h, '休息時間': rest_h,
                                    '用餐': round(float(rows.iloc[idx+4, col_idx]), 1) if pd.notnull(rows.iloc[idx+4, col_idx]) else 0.0,
                                    '備註': str(rows.iloc[idx+5, col_idx]).strip() if pd.notnull(rows.iloc[idx+5, col_idx]) else "",
                                    '正常(8h)': round(min(work_h, 8.0), 1),
                                    '加班': over_h,
                                    '出勤計算': 1 if (not is_off and work_h > 0) else 0
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
st.title("🛡️ 化石先生(JoJo)：雲端工時分析系統 (V14.0.1)")
st.info("系統校準：摘要頁「加班」欄位已改為紅字顯示，維持休假紅底與人員多彩識別。")

uploaded_file = st.file_uploader("導入原始班表 Excel", type=["xlsx"])

if uploaded_file and st.button("🚀 啟動衛星連線分析"):
    month_dict = process_data_v14_0_1(uploaded_file)
    if month_dict:
        st.success("數據掃描與視覺格式校準完成。")
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            wb = writer.book
            head_f = wb.add_format({'bold': 1, 'font_color': 'blue', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
            
            for month, data in month_dict.items():
                safe_m = str(month)[:25]
                summary = data.groupby('人員').agg({'出勤計算': 'sum', '正常(8h)': 'sum', '加班': 'sum'}).reset_index()
                summary.rename(columns={'出勤計算': '總工作天數'}, inplace=True)
                summary['月總工時'] = summary['正常(8h)'] + summary['加班']
                
                # --- A. 明細頁渲染 ---
                sheet_d = f"{safe_m}_明細"
                display_data = data.drop(columns=['出勤計算'])
                display_data.to_excel(writer, index=False, sheet_name=sheet_d)
                ws_d = writer.sheets[sheet_d]
                for c_idx, col in enumerate(display_data.columns): ws_d.write(0, c_idx, col, head_f)
                
                text_colors = ['#0000FF', '#008000', '#800080', '#FF8C00', '#008080', '#A52A2A', '#2F4F4F']
                p_text_map = {p: text_colors[i % len(text_colors)] for i, p in enumerate(display_data['人員'].unique())}
                
                for r_idx in range(len(display_data)):
                    row = display_data.iloc[r_idx]
                    is_off = (str(row['班次']).strip() == "休")
                    for c_idx, col_n in enumerate(display_data.columns):
                        fmt_p = {'border': 1, 'num_format': '0.0', 'align': 'left'}
                        if col_n == '人員': fmt_p['font_color'] = p_text_map.get(row['人員'], '#000000')
                        elif is_off:
                            fmt_p['bg_color'] = '#FF0000'
                            fmt_p['font_color'] = '#000000'
                        else:
                            if col_n == '加班': fmt_p['font_color'] = '#FF0000'
                            elif col_n == '星期' and row[col_n] in ['週六', '週日']: 
                                fmt_p['font_color'] = '#FF0000'
                                fmt_p['bold'] = True
                        ws_d.write(r_idx + 1, c_idx, row[col_n], wb.add_format(fmt_p))
                ws_d.set_column('A:O', 15)
                
                # --- B. 摘要頁渲染 ---
                sheet_s = f"{safe_m}_摘要"
                summary.to_excel(writer, index=False, sheet_name=sheet_s)
                ws_s = writer.sheets[sheet_s]
                for c_idx, col in enumerate(summary.columns): ws_s.write(0, c_idx, col, head_f)
                
                for r_idx in range(len(summary)):
                    row_s = summary.iloc[r_idx]
                    for c_idx, col_n in enumerate(summary.columns):
                        # 摘要格式：全域靠左，加班欄位紅字
                        sum_fmt = {'border': 1, 'num_format': '0.0', 'align': 'left'}
                        if col_n == '加班':
                            sum_fmt['font_color'] = '#FF0000' # 加班紅字
                        ws_s.write(r_idx + 1, c_idx, row_s[col_n], wb.add_format(sum_fmt))
                ws_s.set_column('A:E', 15)

        st.download_button("📥 下載 V14.0.1 全域強化報告", output_excel.getvalue(), "化石先生進階報告.xlsx")
