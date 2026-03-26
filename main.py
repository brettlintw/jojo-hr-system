import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl
from datetime import datetime
from PIL import Image

# V14.3.2 雲端品牌旗艦版：班次顏色同步校準 + 休假全紅底强化 + 週末加乘標註
st.set_page_config(page_title="化石先生：雲端工時分析系統", layout="wide")

# --- UI 品牌頭部設定 ---
def display_header():
    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        try:
            img = Image.open('rsz_mrfossillogo_20190422182824.png')
            st.image(img, width=150)
        except Exception:
            st.error("📷 Logo 檔案未讀取到")
    with col_title:
        st.title("化石先生：雲端工時分析系統 (V14.3.2)")
    st.markdown("---")

display_header()

def process_data_v14_3_2(file):
    # 定義班次規則 (對照附件)
    shift_rules = {
        'A': ('09:30', '17:30'),
        'B': ('13:00', '21:00'),
        'B2': ('14:00', '22:00'),
        'C': ('12:00', '20:30'),
        'All': ('09:30', '21:00'),
        'All2': ('09:30', '22:00')
    }

    try:
        file.seek(0)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
        month_data_dict = {}
        
        for sheet_name, df in all_sheets.items():
            header_idx = -1
            is_compatible = False
            for i, row in df.iterrows():
                row_str = "".join(str(v) for v in row.values)
                if '人員' in row_str and '日期' in row_str:
                    header_idx = i
                    is_compatible = True
                    break
            
            if not is_compatible: continue 
            
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
                                # 休息時間判定
                                if work_h < 4.0: rest_h = 0.0
                                elif 4.0 <= work_h <= 8.0: rest_h = 0.5
                                else: rest_h = 1.0
                                
                                over_h = round(max(work_h - 8.0 - rest_h, 0.0), 1) if work_h > 8.0 else 0.0
                                actual_h = round(work_h - rest_h, 1)
                                is_off = (shift == "休")
                                start_t = "-" if is_off else str(rows.iloc[idx+1, col_idx]).strip()[:5]
                                end_t = "-" if is_off else str(rows.iloc[idx+2, col_idx]).strip()[:5]
                                dt_obj = pd.to_datetime(target_date)
                                weekday_str = f"週{['一','二','三','四','五','六','日'][dt_obj.weekday()]}"
                                is_weekend = weekday_str in ['週六', '週日']
                                
                                # 班次核對
                                check_status = "-" if is_off else "班次錯誤"
                                if not is_off and shift in shift_rules:
                                    rule_start, rule_end = shift_rules[shift]
                                    if start_t == rule_start and end_t == rule_end:
                                        check_status = "班次正確"
                                
                                actual_h_str = str(actual_h)
                                if is_weekend and not is_off and work_h > 0:
                                    actual_h_str = f"{actual_h} (加乘)"

                                person_records.append({
                                    '人員': person, '日期': target_date, '星期': weekday_str,
                                    '班次': shift if shift != "nan" else "",
                                    '班次核對': check_status,
                                    '上班': start_t, '下班': end_t,
                                    '當日工時': work_h, '休息時間/用餐': rest_h,
                                    '實際產出工時': actual_h_str, '加班': over_h,
                                    '備註': str(rows.iloc[idx+5, col_idx]).strip() if pd.notnull(rows.iloc[idx+5, col_idx]) else "",
                                    '出勤計算': 1 if (not is_off and work_h > 0) else 0,
                                    '_is_weekend': is_weekend, '_is_off': is_off
                                })
                        except: pass
                    all_records.extend(person_records)
                    idx += 6 
                else: idx += 1
            if all_records: month_data_dict[str(sheet_name)] = pd.DataFrame(all_records)
        return month_data_dict, shift_rules if month_data_dict else (None, None)
    except: return None, None

# --- UI ---
uploaded_file = st.file_uploader("導入原始班表 Excel", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 啟動衛星連線分析"):
        month_dict, shift_rules = process_data_v14_3_2(uploaded_file)
        if not month_dict:
            st.error("❌ 檔案相容性異常。")
        else:
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                wb = writer.book
                head_f = wb.add_format({'bold': 1, 'font_color': 'blue', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
                
                # --- 班次對照表分頁 ---
                shift_df = pd.DataFrame([{'班次': k, '上班': v[0], '下班': v[1]} for k, v in shift_rules.items()])
                shift_df.to_excel(writer, index=False, sheet_name='班次對照表')
                ws_shift = writer.sheets['班次對照表']
                for c_idx, col in enumerate(shift_df.columns): ws_shift.write(0, c_idx, col, head_f)
                ws_shift.set_column('A:C', 15)

                p_colors = [{'text': '#0000FF', 'bg': '#E1F5FE'}, {'text': '#008000', 'bg': '#E8F5E9'}, {'text': '#800080', 'bg': '#F3E5F5'}, {'text': '#FF8C00', 'bg': '#FFF3E0'}, {'text': '#008080', 'bg': '#E0F2F1'}, {'text': '#A52A2A', 'bg': '#EFEBE9'}, {'text': '#2F4F4F', 'bg': '#ECEFF1'}]
                
                for month, data in month_dict.items():
                    safe_m = str(month)[:25]
                    summary = data.groupby('人員').agg({'出勤計算': 'sum', '當日工時': 'sum', '休息時間/用餐': 'sum', '加班': 'sum'}).reset_index()
                    summary.rename(columns={'出勤計算': '總工作天數', '當日工時': '當月工時'}, inplace=True)
                    p_color_map = {p: p_colors[i % len(p_colors)] for i, p in enumerate(data['人員'].unique())}
                    
                    # --- 明細頁 (含自動篩選與班次顏色同步) ---
                    sheet_d = f"{safe_m}_明細"
                    cols_d = ['人員', '日期', '星期', '班次', '班次核對', '上班', '下班', '當日工時', '休息時間/用餐', '實際產出工時', '加班', '備註']
                    data[cols_d].to_excel(writer, index=False, sheet_name=sheet_d)
                    ws_d = writer.sheets[sheet_d]
                    ws_d.autofilter(0, 0, len(data), len(cols_d)-1) 
                    for c_idx, col_n in enumerate(cols_d): ws_d.write(0, c_idx, col_n, head_f)
                    
                    for r_idx, row in data.iterrows():
                        c_sets = p_color_map.get(row['人員'])
                        is_boundary = (r_idx == len(data)-1) or (data.iloc[r_idx+1]['人員'] != row['人員'])
                        for c_idx, col_n in enumerate(cols_d):
                            val = row[col_n]
                            fmt = {'border': 1, 'align': 'left'}
                            if is_boundary: fmt['bottom'] = 2
                            
                            if col_n == '人員':
                                fmt['font_color'] = c_sets['text']; fmt['bg_color'] = c_sets['bg']
                            elif row['_is_off']:
                                # 休假全紅底，文字維持黑字
                                fmt['bg_color'] = '#FF0000'; fmt['font_color'] = '#000000'
                            elif col_n == '實際產出工時' and row['_is_weekend']:
                                fmt['bg_color'] = '#FFE0B2'; fmt['bold'] = True; fmt['font_color'] = '#E65100'
                            elif col_n == '加班':
                                fmt['font_color'] = '#FF0000'
                            elif col_n == '星期' and row['_is_weekend']:
                                fmt['font_color'] = '#FF0000'; fmt['bold'] = True
                            elif col_n == '班次核對':
                                # --- V14.3.2 新增：班次與核對文字顏色同步 (移除綠/紅字) ---
                                fmt['bold'] = True
                                # 採用一般文字顏色，不特別標示
                            
                            ws_d.write(r_idx + 1, c_idx, val, wb.add_format(fmt))
                    ws_d.set_column('A:L', 15)
                    
                    # --- 摘要頁 ---
                    sheet_s = f"{safe_m}_摘要"
                    summary.to_excel(writer, index=False, sheet_name=sheet_s)
                    ws_s = writer.sheets[sheet_s]
                    ws_s.autofilter(0, 0, len(summary), len(summary.columns)-1)
                    for c_idx, col in enumerate(summary.columns): ws_s.write(0, c_idx, col, head_f)
                    for r_idx, row_s in summary.iterrows():
                        c_sets_s = p_color_map.get(row_s['人員'])
                        for c_idx, col_n in enumerate(summary.columns):
                            val_s = row_s[col_n]
                            s_fmt = {'border': 1, 'num_format': '0.0', 'align': 'left'}
                            if col_n == '人員':
                                s_fmt['font_color'] = c_sets_s['text']; s_fmt['bg_color'] = c_sets_s['bg']
                            elif col_n == '加班':
                                if val_s > 20.0: s_fmt['bg_color'] = '#FF0000'; s_fmt['font_color'] = '#FFFFFF'; s_fmt['bold'] = True
                                else: s_fmt['font_color'] = '#FF0000'
                            ws_s.write(r_idx + 1, c_idx, val_s, wb.add_format(s_fmt))
                    ws_s.set_column('A:E', 15)

            st.download_button("📥 下載 V14.3.2 視覺美學版", output_excel.getvalue(), "化石先生報告.xlsx")
