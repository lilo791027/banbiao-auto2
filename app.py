import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import openpyxl

st.title("班表自動化系統（VBA 完整邏輯轉 Python）")

# 上傳班表 Excel（包含工作表「總表」）
uploaded_schedule = st.file_uploader("請上傳班表 Excel（總表）", type=['xlsx'])
# 上傳員工資料表 Excel
uploaded_employee = st.file_uploader("請上傳員工資料表 Excel", type=['xlsx'])

def unmerge_and_fill(ws):
    """解合併儲存格並填入原值"""
    for merged in list(ws.merged_cells.ranges):
        top_left = ws.cell(merged.min_row, merged.min_col).value
        ws.unmerge_cells(str(merged))
        for row in ws.iter_rows(min_row=merged.min_row, max_row=merged.max_row,
                                min_col=merged.min_col, max_col=merged.max_col):
            for cell in row:
                cell.value = top_left

def collect_shift_data(ws_total):
    """彙整排班資料"""
    output = []
    last_row = ws_total.max_row
    last_col = ws_total.max_column
    clinic_name = str(ws_total.cell(1,1).value)[:4]

    for r in range(1, last_row+1):
        for c in range(2, last_col+1):
            cell_val = ws_total.cell(r,c).value
            if isinstance(cell_val, datetime) or (isinstance(cell_val, str) and '/' in str(cell_val)):
                date_value = cell_val
                i = r + 3
                while i <= last_row:
                    shift_type = str(ws_total.cell(i,c).value).strip()
                    if isinstance(ws_total.cell(i,c).value, datetime) or shift_type == "":
                        break
                    if shift_type in ["早","午","晚"]:
                        i += 1
                        while i <= last_row:
                            if isinstance(ws_total.cell(i,c).value, datetime):
                                break
                            cell_value = str(ws_total.cell(i,c).value).strip()
                            if cell_value in ["早","午","晚"]:
                                break
                            row_data = [
                                clinic_name,
                                date_value.strftime("%Y/%m/%d") if isinstance(date_value, datetime) else date_value,
                                shift_type,
                                cell_value,
                                ws_total.cell(i,1).value,
                                ws_total.cell(i,21).value
                            ]
                            output.append(row_data)
                            i += 1
                        i -= 1
                    i += 1
    df_out = pd.DataFrame(output, columns=["診所","日期","班別","姓名","A欄資料","U欄資料"])
    return df_out

def format_shift_order(shift_str):
    result = ""
    if "早" in shift_str:
        result += "早"
    if "午" in shift_str:
        result += "午"
    if "晚" in shift_str:
        result += "晚"
    return result

def get_class_code(emp_title, clinic_name, shift_type):
    if not emp_title or emp_title.strip() == "":
        return ""
    # VBA 邏輯完全保留
    if emp_title in ["早班護理師","早班內視鏡助理","醫務專員","兼職早班內視鏡助理"]:
        return "【員工】純早班"

    if emp_title == "醫師":
        code = "★醫師★"
    elif emp_title in ["櫃臺","護理師","兼職護理師","兼職跟診助理","副店長"]:
        code = "【員工】"
    else:
        if "副店長" in emp_title:
            code = "【員工】"
        elif "店長" in emp_title or "護士" in emp_title:
            code = "◇主管◇"
        else:
            code = ""

    if shift_type != "早":
        if clinic_name in ["上吉診所","立吉診所","上承診所","立全診所","立竹診所","立順診所","上京診所"]:
            code += "板土中京"
        elif clinic_name == "立丞診所":
            code += "立丞"

    shift_map = {
        "早":"早班", "午晚":"午晚班", "早午晚":"全天班", "早晚":"早晚班",
        "午":"午班","晚":"晚班","早午":"早午班"
    }
    code += shift_map.get(shift_type,"")
    if code.endswith("早班早班"):
        code = code.replace("早班早班","早班")
    return code

def build_shift_analysis(df_out, df_employee):
    """建立班別分析表"""
    emp_dict = {row['姓名']:[row['員工編號'],row['部門'],row['職稱']] for _, row in df_employee.iterrows()}
    dict_data = {}
    records = []
    for _, row in df_out.iterrows():
        name = str(row['姓名']).strip()
        if name == "" or len(name) > 4:
            continue
        key = f"{name}|{row['日期']}|{row['診所']}|{row['A欄資料']}"
        if key not in dict_data:
            dict_data[key] = row['班別']
        else:
            dict_data[key] = dict_data[key] + " " + row['班別']

    for key, shift_str in dict_data.items():
        name, date_value, clinic_name, e_value = key.split("|")
        shift_type = format_shift_order(shift_str)
        if name in emp_dict:
            emp_id, emp_dept, emp_title = emp_dict[name]
        else:
            emp_id, emp_dept, emp_title = "", "", ""
        records.append({
            "診所":clinic_name,
            "員工編號":emp_id,
            "所屬部門":emp_dept,
            "姓名":name,
            "職稱":emp_title,
            "日期":date_value,
            "班別":shift_type,
            "E欄資料":e_value,
            "班別代碼":get_class_code(emp_title, clinic_name, shift_type)
        })
    df_analysis = pd.DataFrame(records)
    return df_analysis

def build_shift_summary(df_analysis):
    """建立班別總表"""
    df_analysis['日期'] = pd.to_datetime(df_analysis['日期'], errors='coerce')
    min_date = df_analysis['日期'].min()
    max_date = df_analysis['日期'].max()
    all_dates = pd.date_range(min_date, max_date).strftime("%Y-%m-%d").tolist()

    summary_dict = {}
    for _, row in df_analysis.iterrows():
        emp_id = str(row['員工編號']).strip()
        emp_name = row['姓名']
        class_code = row['班別代碼']
        key = (emp_id, emp_name)
        if key not in summary_dict:
            summary_dict[key] = {}
        summary_dict[key][row['日期'].strftime("%Y-%m-%d")] = class_code

    # 建立 dataframe
    df_summary = pd.DataFrame(columns=["員工編號","員工姓名"]+all_dates)
    for (emp_id, emp_name), date_dict in summary_dict.items():
        row = [emp_id, emp_name] + [date_dict.get(d,"") for d in all_dates]
        df_summary.loc[len(df_summary)] = row
    return df_summary

if uploaded_schedule is not None and uploaded_employee is not None:
    ws_total = openpyxl.load_workbook(uploaded_schedule)['總表']
    unmerge_and_fill(ws_total)
    df_out = collect_shift_data(ws_total)
    df_employee = pd.read_excel(uploaded_employee)
    df_analysis = build_shift_analysis(df_out, df_employee)
    df_summary = build_shift_summary(df_analysis)

    st.subheader("班別分析表")
    st.dataframe(df_analysis)

    st.subheader("班別總表")
    st.dataframe(df_summary)

    # 下載 Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_out.to_excel(writer, index=False, sheet_name="彙整結果")
        df_analysis.to_excel(writer, index=False, sheet_name="班別分析")
        df_summary.to_excel(writer, index=False, sheet_name="班別總表")
    st.download_button("下載完整班表", data=output.getvalue(), file_name="班表結果.xlsx")
