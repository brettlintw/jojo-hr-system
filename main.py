import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
import openpyxl

# 雲端版專屬配置
st.set_page_config(page_title="化石先生(JoJo)：工時分析系統 V13.6.1", layout="wide")

def get_excel_save_time(file):
    try:
        # 雲端讀取必須確保 seek(0)
        file.seek(0)
        wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
        last_saved = wb.properties.modified
        wb.close()
        file.seek(0)
        return last_saved.strftime("%Y年%m月%d日 %H:%M:%S") if last_saved else "未知"
    except: return "無法讀取存檔時間"

def process_data_cloud(file):
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
                            if str(rows.iloc[idx, col_idx]).strip() != "nan" or work_h > 0:
                                dt_obj = pd.to_datetime(target_date)
                                week_cn = f"週{['一','二','三','四','五','六','日'][dt_obj.weekday()]}"
                                person_records.append({
                                    '人員': person, '日期': target_date, '星期': week_cn,
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

# --- UI ---
st.title("🛡️ 化石先生(JoJo)：工時分析系統 (V13.6.1 Cloud)")
st.markdown("---")
uploaded_file = st.file_uploader("導入原始班表 Excel", type=["xlsx"])

if uploaded_file:
    st.write(f"💾 **班表存檔時間：{get_excel_save_time(uploaded_file)}**")
    if st.button("🚀 啟動雲端內容分析"):
        uploaded_file.seek(0)
        month_dict = process_data_cloud(uploaded_file)
        if month_dict:
            st.success("分析任務已完成。")
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                wb = writer.book
                for month, data in month_dict.items():
                    # 網頁預覽
                    with st.expander(f"📂 查看 {month} 明細預覽"):
                        st.dataframe(data.style.format(precision=1), use_container_width=True)
                    
                    safe_m = str(month)[:25]
                    summary = data.groupby('人員').agg({'當日工時':'sum', '正常(8h)':'sum', '加班':'sum', '月總工時':'max'}).reset_index()
                    data.to_excel(writer, index=False, sheet_name=f"{safe_m}_明細")
                    summary.to_excel(writer, index=False, sheet_name=f"{safe_m}_摘要")
                    
                    # 框線與色彩
                    ws = writer.sheets[f"{safe_m}_明細"]
                    unique_p = data['人員'].unique()
                    colors = ['#F0F2F6', '#E1F5FE', '#E8F5E9', '#FFFDE7', '#F3E5F5', '#EFEBE9']
                    p_map = {p: colors[i%6] for i, p in enumerate(unique_p)}
                    
                    for r_idx in range(len(data)):
                        row_v = data.iloc[r_idx]
                        p_bg = p_map.get(row_v['人員'])
                        std_f = wb.add_format({'bg_color': p_bg, 'border': 1, 'num_format': '0.0'})
                        red_f = wb.add_format({'bg_color': p_bg, 'border': 1, 'num_format': '0.0', 'font_color': 'red', 'bold': True})
                        for c_idx, col in enumerate(data.columns):
                            is_red = (col == '星期' and row_v[col] in ['週六', '週日'])
                            ws.write(r_idx + 1, c_idx, row_v[col], red_f if is_red else std_f)
                    ws.set_column('A:L', 15)

            st.download_button("📥 下載整合分析報告 (Excel)", output_excel.getvalue(), f"化石先生整合報告.xlsx")