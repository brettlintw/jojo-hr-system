import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl
import re
from datetime import datetime
from PIL import Image

# V14.5.8 雲端品牌旗艦版：自動診斷排班異常原因 + 動態檔名生成 + 明細摘要緊縮間距
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
        st.title("化石先生：雲端工時分析系統 (V14.5.8)")
    st.markdown("---")

display_header()

def process_data_v14_5_8(file):
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
    f_name_raw = uploaded_file.name
    f_name_no_ext = f_name_raw.rsplit('.', 1)[0]
    date_str = datetime.now().strftime("%Y%m%d")
    final_filename = f"化石先生-{f_name_no_ext}-{date_str}.xlsx"

    try:
        uploaded_file.seek(0)
        wb_meta = openpyxl.load_workbook(uploaded_file, read_only=True)
        meta = wb_meta.properties
        m_time = meta.modified.strftime("%Y/%m/%d %H:%M:%S") if meta.modified else "無法讀取"
        st.markdown(f"""
        <div style="background-color: #F0F2F6; padding: 15px; border-radius: 10px; border-left: 5px solid #0000FF; margin-bottom: 20px;">
            <p style="color: #0000FF; font-size: 1.1em; margin-bottom: 5px;"><b>原始檔名：</b>{f_name_raw}</p>
            <p style="color: #0000FF; font-size: 1.1em; margin-bottom: 0;"><b>最後修改時間：</b>{m_time}</p>
        </div>
        """, unsafe_allow_html=True)
    except:
        m_time = "無法讀取"

    if st.button("🚀 啟動衛星連線分析"):
        month_dict, shift_rules = process_data_v14_4_0(uploaded_file) # 指向穩定核心
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
