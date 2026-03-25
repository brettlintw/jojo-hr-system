import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl
from datetime import datetime
from PIL import Image

# V14.2.5 雲端品牌校準穩定版：新增「週末加乘標記」+ 人員間粗底線條 + 階梯式休息判定
st.set_page_config(page_title="化石先生：雲端工時分析系統", layout="wide")

# --- UI 品牌頭部設定 (Logo 與 標題) ---
def display_header():
    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        try:
            img = Image.open('rsz_mrfossillogo_20190422182824.png')
            st.image(img, width=150)
        except Exception:
            st.error("📷 Logo 檔案未讀取到")
    with col_title:
        st.title("化石先生：雲端工時分析系統 (V14.2.5)")
    st.markdown("---")

display_header()

def process_data_v14_2_5(file):
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
                                # 休息時間階梯判定
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
                                
                                # --- V14.2.5 新增：週末加乘標註邏輯 ---
                                actual_h_str = str(actual_h)
                                is_weekend = weekday_str in ['週六', '週日']
                                if is_weekend and not is_off and work_h > 0:
                                    actual_h_str = f"{actual_h} (加乘)"

                                person_records.append({
                                    '人員': person, '日期': target_date, '星期': weekday_str,
                                    '班次': shift if shift != "nan" else "",
                                    '上班': start_t, '下班': end_t,
                                    '當日工時': work_h, '休息時間/用餐': rest_h,
                                    '實際產出工時': actual_h_str, 
                                    '加班': over_h,
                                    '備註': str(rows.iloc[idx+5, col_idx]).strip() if pd.notnull(rows.iloc[idx+5, col_idx]) else "",
                                    '出勤計算': 1 if (not is_off and work_h > 0) else 0,
                                    '_is_weekend': is_weekend # 內部標記用於格式化
                                })
                        except: pass
                    
                    total_work_h = sum(r['當日工時'] for r in person_records)
                    for r in person_records: r['月總工時'] = round(total_work_h, 1)
                    all_records.extend(person_records)
                    idx += 6 
                else: idx += 1
            if all_records: month_data_dict[str(sheet_name)] = pd.DataFrame(all_records)
        return month_data_dict if month_data_dict else "INCOMPATIBLE"
    except: return "INCOMPATIBLE"

# --- UI ---
uploaded_file = st.file_uploader("導入原始班表 Excel", type=["xlsx"])

if uploaded_file:
    # (時戳掃描邏輯維持不變)
    try:
        uploaded_file.seek(0)
        wb_prop = openpyxl.load_workbook(uploaded_file, read_only=True)
        props = wb_prop.properties
        time_fmt = "%Y年%m月%d日 %H:%M:%S"
        m_time = props.modified.strftime(time_fmt) if props.modified else "數據遺失"
        e_time = props.created.strftime(time_fmt) if props.created else "數據遺失"
        last_user = props.lastModifiedBy if props.lastModifiedBy else "未知成員"
        st.success(f"📡 檔案掃描成功！ 最後修改：{m_time}")
        uploaded_file.seek(0)
    except: pass

    if st.button("🚀 啟動衛星連線分析"):
        result = process_data_v14_2_5(uploaded_file)
        if result == "INCOMPATIBLE":
            st.error("❌ 偵測到不相容檔案。")
        else:
            month_dict = result
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                wb = writer.book
                head_f = wb.add_format({'bold': 1, 'font_color': 'blue', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
                p_colors = [{'text': '#0000FF', 'bg': '#E1F5FE'}, {'text': '#008000', 'bg': '#E8F5E9'}, {'text': '#800080', 'bg': '#F3E5F5'}, {'text': '#FF8C00', 'bg': '#FFF3E0'}, {'text': '#008080', 'bg': '#E0F2F1'}, {'text': '#A52A2A', 'bg': '#EFEBE9'}, {'text': '#2F4F4F', 'bg': '#ECEFF1'}]
                
                for month, data in month_dict.items():
                    safe_m = str(month)[:25]
                    summary = data.groupby('人員').agg({'出勤計算': 'sum', '當日工時': 'sum', '休息時間/用餐': 'sum', '加班': 'sum'}).reset_index()
                    summary.rename(columns={'出勤計算': '總工作天數', '當日工時': '當月工時'}, inplace=True)
                    p_color_map = {p: p_colors[i % len(p_colors)] for i, p in enumerate(data['人員'].unique())}
                    
                    # --- 明細頁 ---
                    sheet_d = f"{safe_m}_明細"
                    cols_d = ['人員', '日期', '星期', '班次', '上班', '下班', '當日工時', '休息時間/用餐', '實際產出工時', '加班', '月總工時', '備註']
                    data[cols_d].to_excel(writer, index=False, sheet_name=sheet_d)
                    ws_d = writer.sheets[sheet_d]
                    for c_idx, col_n in enumerate(cols_d): ws_d.write(0, c_idx, col_n, head_f)
                    
                    for r_idx, row in data.iterrows():
                        is_off = (str(row['班次']).strip() == "休")
                        c_sets = p_color_map.get(row['人員'])
                        is_person_boundary = (r_idx == len(data) - 1) or (data.iloc[r_idx + 1]['人員'] != row['人員'])
                        
                        for c_idx, col_n in enumerate(cols_d):
                            val = row[col_n]
                            fmt_dict = {'border': 1, 'align': 'left'}
                            if is_person_boundary: fmt_dict['bottom'] = 2
                            
                            if col_n == '人員':
                                fmt_dict['font_color'] = c_sets['text']; fmt_dict['bg_color'] = c_sets['bg']
                            elif is_off:
                                fmt_dict['bg_color'] = '#FF0000'; fmt_dict['font_color'] = '#000000'
                            elif col_n == '實際產出工時' and row['_is_weekend'] and not is_off:
                                # 週末加成高亮設定
                                fmt_dict['bg_color'] = '#FFE0B2'; fmt_dict['bold'] = True; fmt_dict['font_color'] = '#E65100'
                            elif col_n == '加班':
                                fmt_dict['font_color'] = '#FF0000'
                            elif col_n == '星期' and row['_is_weekend']:
                                fmt_dict['font_color'] = '#FF0000'; fmt_dict['bold'] = True
                            
                            ws_d.write(r_idx + 1, c_idx, val, wb.add_format(fmt_dict))
                    ws_d.set_column('A:L', 15)
                    
                    # --- 摘要頁 ---
                    sheet_s = f"{safe_m}_摘要"
                    summary.to_excel(writer, index=False, sheet_name=sheet_s)
                    ws_s = writer.sheets[sheet_s]
                    for c_idx, col_n in enumerate(summary.columns): ws_s.write(0, c_idx, col_n, head_f)
                    for r_idx, row_s in summary.iterrows():
                        c_sets_s = p_color_map.get(row_s['人員'])
                        for c_idx, col_n in enumerate(summary.columns):
                            val_s = row_s[col_n]
                            sum_fmt = {'border': 1, 'num_format': '0.0', 'align': 'left'}
                            if col_n == '人員':
                                sum_fmt['font_color'] = c_sets_s['text']; sum_fmt['bg_color'] = c_sets_s['bg']
                            elif col_n == '加班' and val_s > 20.0:
                                sum_fmt['bg_color'] = '#FF0000'; sum_fmt['font_color'] = '#FFFFFF'; sum_fmt['bold'] = True
                            ws_s.write(r_idx + 1, c_idx, val_s, wb.add_format(sum_fmt))
                    ws_s.set_column('A:E', 15)

            st.download_button("📥 下載 V14.2.5 週末標註報告", output_excel.getvalue(), "化石先生報告.xlsx")
