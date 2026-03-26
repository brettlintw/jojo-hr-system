import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl
import re
from datetime import datetime
from PIL import Image

# V14.5.2 雲端品牌旗艦版：月份欄位純數字化 + 全月份整合 + 溯源資訊首行崁入
st.set_page_config(page_title="化石先生：雲端工時分析系統", layout="wide")

def display_header():
    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        try:
            img = Image.open('rsz_mrfossillogo_20190422182824.png')
            st.image(img, width=150)
        except Exception:
            st.error("📷 Logo 遺失")
    with col_title:
        st.title("化石先生：雲端工時分析系統 (V14.5.2)")
    st.markdown("---")

display_header()

def process_data_v14_5_2(file):
    shift_rules = {
        'A': ('09:30', '17:30'), 'B': ('13:00', '21:00'), 'B2': ('14:00', '22:00'),
        'C': ('12:00', '20:30'), 'All': ('09:30', '21:00'), 'All2': ('09:30', '22:00')
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
                                rest_h = 0.0 if work_h < 4.0 else (0.5 if 4.0 <= work_h <= 8.0 else 1.0)
                                over_h = round(max(work_h - 8.0 - rest_h, 0.0), 1) if work_h > 8.0 else 0.0
                                actual_h = round(work_h - rest_h, 1)
                                is_off = (shift == "休")
                                start_t = "-" if is_off else str(rows.iloc[idx+1, col_idx]).strip()[:5]
                                end_t = "-" if is_off else str(rows.iloc[idx+2, col_idx]).strip()[:5]
                                dt_obj = pd.to_datetime(target_date)
                                weekday_str = f"週{['一','二','三','四','五','六','日'][dt_obj.weekday()]}"
                                is_weekend = weekday_str in ['週六', '週日']
                                check_status = "-" if is_off else "班次錯誤"
                                if not is_off and shift in shift_rules:
                                    r_s, r_e = shift_rules[shift]
                                    if start_t == r_s and end_t == r_e: check_status = "班次正確"
                                actual_h_str = f"{actual_h} (加乘)" if (is_weekend and not is_off and work_h > 0) else str(actual_h)
                                person_records.append({
                                    '人員': person, '日期': target_date, '星期': weekday_str, '班次': shift if shift != "nan" else "",
                                    '班次核對': check_status, '上班': start_t, '下班': end_t, '當日工時': work_h, '休息時間/用餐': rest_h,
                                    '實際產出工時': actual_h_str, '加班': over_h, '備註': str(rows.iloc[idx+5, col_idx]).strip() if pd.notnull(rows.iloc[idx+5, col_idx]) else "",
                                    '出勤計算': 1 if (not is_off and work_h > 0) else 0, '_is_weekend': is_weekend, '_is_off': is_off
                                })
                        except: pass
                    all_records.extend(person_records)
                    idx += 6 
                else: idx += 1
            if all_records: month_data_dict[str(sheet_name)] = pd.DataFrame(all_records)
        return month_data_dict, shift_rules
    except: return None, None

uploaded_file = st.file_uploader("導入原始班表 Excel", type=["xlsx"])

if uploaded_file:
    f_name = uploaded_file.name
    try:
        uploaded_file.seek(0)
        wb_meta = openpyxl.load_workbook(uploaded_file, read_only=True)
        meta = wb_meta.properties
        m_time = meta.modified.strftime("%Y/%m/%d %H:%M:%S") if meta.modified else "無法讀取"
        e_time = meta.created.strftime("%Y/%m/%d %H:%M:%S") if meta.created else "無法讀取"
        
        st.markdown(f"""
        <div style="background-color: #F0F2F6; padding: 15px; border-radius: 10px; border-left: 5px solid #0000FF; margin-bottom: 20px;">
            <p style="color: #0000FF; font-size: 1.1em; margin-bottom: 5px;"><b>原始檔名：</b>{f_name}</p>
            <p style="color: #0000FF; font-size: 1.1em; margin-bottom: 5px;"><b>最後修改時間：</b>{m_time}</p>
            <p style="color: #0000FF; font-size: 1.1em; margin-bottom: 0;"><b>上次編輯時間：</b>{e_time}</p>
        </div>
        """, unsafe_allow_html=True)
    except:
        m_time, e_time = "無法讀取", "無法讀取"

    if st.button("🚀 啟動衛星連線分析"):
        month_dict, shift_rules = process_data_v14_5_2(uploaded_file)
        if not month_dict:
            st.error("❌ 檔案相容性異常。")
        else:
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                wb = writer.book
                head_f = wb.add_format({'bold': 1, 'font_color': 'blue', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
                info_f = wb.add_format({'bold': 1, 'font_color': '#FFFFFF', 'bg_color': '#0000FF', 'align': 'left'})
                yellow_head_f = wb.add_format({'bold': 1, 'font_color': 'blue', 'bg_color': 'yellow', 'border': 2, 'align': 'left'})
                bold_border_f = wb.add_format({'border': 2, 'align': 'left'})

                # 1. 班次對照表
                shift_df = pd.DataFrame([{'班次': k, '上班': v[0], '下班': v[1]} for k, v in shift_rules.items()])
                shift_df.to_excel(writer, index=False, sheet_name='班次對照表')
                ws_shift = writer.sheets['班次對照表']
                for c_idx, col in enumerate(shift_df.columns): ws_shift.write(0, c_idx, col, yellow_head_f)
                for r_idx, row_vals in enumerate(shift_df.values):
                    for c_idx, val in enumerate(row_vals): ws_shift.write(r_idx + 1, c_idx, val, bold_border_f)
                ws_shift.set_column(0, 2, 20)

                p_colors = [{'text': '#0000FF', 'bg': '#E1F5FE'}, {'text': '#008000', 'bg': '#E8F5E9'}, {'text': '#800080', 'bg': '#F3E5F5'}, {'text': '#FF8C00', 'bg': '#FFF3E0'}, {'text': '#008080', 'bg': '#E0F2F1'}, {'text': '#A52A2A', 'bg': '#EFEBE9'}, {'text': '#2F4F4F', 'bg': '#ECEFF1'}]
                all_confirm_data = []

                for month_name, data in month_dict.items():
                    # --- V14.5.2 更新：從分頁名稱擷取純數字月份 ---
                    month_num_match = re.search(r'(\d+)', month_name)
                    clean_month = month_num_match.group(1) if month_num_match else month_name
                    
                    safe_m = str(month_name)[:15]
                    sheet_main = f"{safe_m}_明細+摘要"
                    summary = data.groupby('人員').agg({'出勤計算': 'sum', '當日工時': 'sum', '休息時間/用餐': 'sum', '加班': 'sum'}).reset_index()
                    summary.rename(columns={'出勤計算': '總工作天數', '當日工時': '當月工時'}, inplace=True)
                    p_color_map = {p: p_colors[i % len(p_colors)] for i, p in enumerate(data['人員'].unique())}
                    
                    cols_d = ['人員', '日期', '星期', '班次', '班次核對', '上班', '下班', '當日工時', '休息時間/用餐', '實際產出工時', '加班', '備註']
                    start_col_sum = len(cols_d) + 2
                    
                    data[cols_d].to_excel(writer, index=False, sheet_name=sheet_main, startrow=1)
                    ws = writer.sheets[sheet_main]
                    ws.freeze_panes(2, 0)
                    ws.merge_range(0, 0, 0, start_col_sum + 4, f"原始檔名：{f_name}  |  最後修改時間：{m_time}  |  上次編輯時間：{e_time}", info_f)
                    ws.autofilter(1, 0, len(data)+1, len(cols_d)-1) 
                    
                    for c_idx, col_n in enumerate(cols_d): ws.write(1, c_idx, col_n, head_f)
                    sum_head_f = wb.add_format({'bold': 1, 'font_color': '#800000', 'border': 1, 'bg_color': '#FFEBEE', 'align': 'left'})
                    for c_idx, col_n in enumerate(summary.columns): ws.write(1, start_col_sum + c_idx, col_n, sum_head_f)
                    
                    for r_idx, row in data.iterrows():
                        c_sets = p_color_map.get(row['人員'])
                        is_bound = (r_idx == len(data)-1) or (data.iloc[r_idx+1]['人員'] != row['人員'])
                        c_color = '#008000' if row['班次核對'] == "班次正確" else ('#FF0000' if row['班次核對'] == "班次錯誤" else '#000000')
                        for c_idx, col_n in enumerate(cols_d):
                            val = row[col_n]
                            fmt = {'border': 1, 'align': 'left', 'valign': 'vcenter'}
                            if is_bound: fmt['bottom'] = 2
                            if col_n == '備註': fmt['text_wrap'] = True
                            if col_n == '人員': fmt['font_color'] = c_sets['text']; fmt['bg_color'] = c_sets['bg']
                            elif row['_is_off']: fmt['bg_color'] = '#FF0000'; fmt['font_color'] = '#000000'
                            elif col_n in ['班次', '班次核對']: fmt['bold'] = True; fmt['font_color'] = c_color
                            elif col_n == '實際產出工時' and row['_is_weekend']: fmt['bg_color'] = '#FFE0B2'; fmt['bold'] = True; fmt['font_color'] = '#E65100'
                            elif col_n == '加班' or (col_n == '星期' and row['_is_weekend']): fmt['font_color'] = '#FF0000'; fmt['bold'] = (col_n == '星期')
                            ws.write(r_idx + 2, c_idx, val, wb.add_format(fmt))

                    for r_idx, row_s in summary.iterrows():
                        c_sets_s = p_color_map.get(row_s['人員'])
                        for c_idx, col_n in enumerate(summary.columns):
                            v_s = row_s[col_n]
                            s_fmt = {'border': 1, 'num_format': '0.0', 'align': 'left', 'valign': 'vcenter'}
                            if col_n == '人員': s_fmt['font_color'] = c_sets_s['text']; s_fmt['bg_color'] = c_sets_s['bg']
                            elif col_n == '加班':
                                if v_s > 20.0: s_fmt['bg_color'] = '#FF0000'; s_fmt['font_color'] = '#FFFFFF'; s_fmt['bold'] = True
                                else: s_fmt['font_color'] = '#FF0000'
                            ws.write(r_idx + 2, start_col_sum + c_idx, v_s, wb.add_format(s_fmt))

                    # 3. 準備 排班確定表
                    active_duty = data[~data['_is_off']].copy()
                    confirm_df = active_duty.groupby('日期').agg({'星期': 'first', '人員': lambda x: ",".join(x), '出勤計算': 'sum'}).reset_index()
                    confirm_df['月份'] = clean_month # 填入純數字月份
                    confirm_df['排班人數確認'] = confirm_df.apply(lambda r: "正常" if (r['星期'] in ['週六', '週日'] and 3 <= r['出勤計算'] <= 4) or (r['星期'] not in ['週六', '週日'] and 2 <= r['出勤計算'] <= 3) else "異常", axis=1)
                    confirm_df['備註'] = ""
                    all_confirm_data.append(confirm_df)

                # 4. 寫入合併 排班確定表
                total_confirm_df = pd.concat(all_confirm_data, ignore_index=True)
                total_confirm_df.rename(columns={'出勤計算': '當天人數', '星期': '平日/假日'}, inplace=True)
                total_confirm_df = total_confirm_df.reindex(columns=['月份', '日期', '平日/假日', '當天人數', '人員', '排班人數確認', '備註'])
                
                total_confirm_df.to_excel(writer, index=False, sheet_name='排班確定表', startrow=1)
                ws_c = writer.sheets['排班確定表']
                ws_c.freeze_panes(2, 0)
                ws_c.merge_range(0, 0, 0, 6, f"人力密度監控儀表板  |  原始檔名：{f_name}  |  最後修改時間：{m_time}  |  上次編輯時間：{e_time}", info_f)
                for c_idx, col in enumerate(total_confirm_df.columns): ws_c.write(1, c_idx, col, head_f)
                
                for r_idx, row_c in total_confirm_df.iterrows():
                    is_we = row_c['平日/假日'] in ['週六', '週日']
                    is_m_end = (r_idx < len(total_confirm_df)-1) and (total_confirm_df.iloc[r_idx+1]['月份'] != row_c['月份'])
                    for c_idx, col_n in enumerate(total_confirm_df.columns):
                        val_c = row_c[col_n]
                        c_fmt = {'border': 1, 'align': 'left', 'valign': 'vcenter'}
                        if is_m_end: c_fmt['bottom'] = 5
                        if col_n == '平日/假日':
                            c_fmt['font_color'] = '#FF0000' if is_we else '#008000'
                            val_c = row_c['平日/假日'] if is_we else "平日"
                        elif col_n == '當天人數':
                            lim = (3, 4) if is_we else (2, 3)
                            c_fmt['font_color'] = '#008000' if lim[0] <= val_c <= lim[1] else '#FF0000'
                        elif col_n == '排班人數確認':
                            c_fmt['font_color'] = '#008000' if val_c == "正常" else '#FF0000'
                            c_fmt['bold'] = True
                        ws_c.write(r_idx + 2, c_idx, val_c, wb.add_format(c_fmt))
                
                ws_c.set_column('A:G', 15)
                ws_c.set_column('E:E', 40)

            st.download_button(f"📥 下載 V14.5.2 月份精煉版", output_excel.getvalue(), "化石先生報告.xlsx")
