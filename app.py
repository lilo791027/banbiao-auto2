import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.title("班表自動化系統")

# --- 上傳 Excel ---
schedule_file = st.file_uploader("請上傳排班 Excel", type=["xlsx"])
employee_file = st.file_uploader("請上傳員工人事資料明細表 Excel", type=["xlsx"])

if schedule_file and employee_file:
    # 讀取班表
    wb_schedule = openpyxl.load_workbook(schedule_file)
    sheet_names = wb_schedule.sheetnames
    selected_sheet = st.selectbox("請選擇要使用的工作表", sheet_names)
    ws_total = wb_schedule[selected_sheet]
    
    # 解合併儲存格並填入原值
    for merged_cell in list(ws_total.merged_cells.ranges):
        min_row, min_col, max_row, max_col = merged_cell.bounds
        value_to_fill = ws_total.cell(min_row, min_col).value
        ws_total.unmerge_cells(str(merged_cell))
        for r in range(min_row, max_row+1):
            for c in range(min_col, max_col+1):
                ws_total.cell(r, c).value = value_to_fill

    df_total = pd.DataFrame(ws_total.values)
    df_total.columns = df_total.iloc[0]
    df_total = df_total[1:].reset_index(drop=True)
    
    st.subheader(f"工作表內容：{selected_sheet}")
    st.dataframe(df_total.head())

    # --- 彙整排班資料 ---
    df_out = pd.DataFrame(columns=["診所","日期","班別","姓名","A欄資料","U欄資料"])
    clinic_name = str(df_total.iloc[0,0])[:4]
    last_row, last_col = df_total.shape
    output_row = 0

    for r in range(last_row):
        for c in range(1, last_col):
            cell_val = df_total.iloc[r,c]
            if pd.api.types.is_datetime64_any_dtype(cell_val) or isinstance(cell_val, pd.Timestamp):
                date_value = cell_val
                i = r + 3
                while i < last_row:
                    shift_type = str(df_total.iloc[i,c]).strip()
                    if pd.api.types.is_datetime64_any_dtype(df_total.iloc[i,c]) or shift_type == "":
                        break
                    if shift_type in ["早","午","晚"]:
                        i += 1
                        while i < last_row:
                            if pd.api.types.is_datetime64_any_dtype(df_total.iloc[i,c]):
                                break
                            cell_value = str(df_total.iloc[i,c]).strip()
                            if cell_value in ["早","午","晚"]:
                                break
                            df_out.loc[output_row] = [
                                clinic_name,
                                date_value.strftime("%Y/%m/%d") if hasattr(date_value, "strftime") else date_value,
                                shift_type,
                                cell_value,
                                df_total.iloc[i,0],
                                df_total.iloc[i,20] if df_total.shape[1]>20 else ""
                            ]
                            output_row += 1
                            i += 1
                        i -= 1
                    i += 1

    st.subheader("彙整排班資料")
    st.dataframe(df_out.head())

    # --- 讀取員工明細 ---
    wb_emp = openpyxl.load_workbook(employee_file)
    ws_emp = wb_emp[wb_emp.sheetnames[0]]
    df_emp = pd.DataFrame(ws_emp.values)
    df_emp.columns = df_emp.iloc[0]
    df_emp = df_emp[1:].reset_index(drop=True)
    
    emp_dict = {}
    for idx, row in df_emp.iterrows():
        name = str(row[1]).strip()
        if name:
            emp_dict[name] = [str(row[0]), row[2], row[3]]  # empID, dept, title

    # --- 建立班別分析表 ---
    df_analysis = pd.DataFrame(columns=["診所","員工編號","所屬部門","姓名","職稱","日期","班別","E欄資料","班別代碼"])
    shift_dict = {}
    for idx, row in df_out.iterrows():
        name = str(row["姓名"]).strip()
        key = f"{name}|{row['日期']}|{row['診所']}"
        if key not in shift_dict:
            shift_dict[key] = row["班別"]
        else:
            shift_dict[key] += " " + row["班別"]

    for key, shift in shift_dict.items():
        name, date_value, clinic_name = key.split("|")
        e_value = df_out[df_out["姓名"]==name]["A欄資料"].values[0] if not df_out[df_out["姓名"]==name].empty else ""
        shift_type = shift.replace(" ","")
        if name in emp_dict:
            empID, empDept, empTitle = emp_dict[name]
        else:
            empID = empDept = empTitle = ""
        df_analysis.loc[len(df_analysis)] = [
            clinic_name,
            empID,
            empDept,
            name,
            empTitle,
            date_value,
            shift_type,
            e_value,
            ""  # 班別代碼
        ]

    # --- 計算班別代碼 ---
    def get_class_code(empTitle, clinicName, shiftType):
        if not empTitle:
            return ""
        if empTitle in ["早班護理師","早班內視鏡助理","醫務專員","兼職早班內視鏡助理"]:
            return "【員工】純早班"
        classCode = ""
        if empTitle=="醫師":
            classCode = "★醫師★"
        elif empTitle in ["櫃臺","護理師","兼職護理師","兼職跟診助理","副店長","護士","藥師"] or "副店長" in empTitle:
            classCode = "【員工】"
        elif "店長" in empTitle or "採購儲備組長" in empTitle:
            classCode = "◇主管◇"
        if shiftType!="早":
            if clinicName in ["上吉診所","立吉診所","上承診所","立全診所","立竹診所","立順診所","上京診所"]:
                classCode += "板土中京"
            elif clinicName=="立丞診所":
                classCode += "立丞"
        mapping = {"早":"早班","午晚":"午晚班","早午晚":"全天班","早晚":"早晚班","午":"午班","晚":"晚班","早午":"早午班"}
        classCode += mapping.get(shiftType,"")
        return classCode.replace("早班早班","早班")

    df_analysis['班別代碼'] = df_analysis.apply(lambda x: get_class_code(x['職稱'], x['診所'], x['班別']), axis=1)
    st.subheader("班別分析表")
    st.dataframe(df_analysis.head())

    # --- 建立班別總表 ---
    min_date = pd.to_datetime(df_analysis['日期']).min()
    max_date = pd.to_datetime(df_analysis['日期']).max()
    all_dates = pd.date_range(min_date, max_date).strftime("%Y-%m-%d")

    df_total_shift = pd.DataFrame(columns=["員工編號","員工姓名"]+list(all_dates))
    dict_shift = {}
    for idx, row in df_analysis.iterrows():
        empID = str(row["員工編號"]).strip()
        empName = str(row["姓名"]).strip()
        shiftDate = row["日期"]
        classCode = row["班別代碼"]
        empKey = f"{empID}|{empName}"
        if empKey not in dict_shift:
            dict_shift[empKey] = {}
        dict_shift[empKey][shiftDate] = classCode

    for empKey, dates in dict_shift.items():
        empID, empName = empKey.split("|")
        row_data = {"員工編號": empID, "員工姓名": empName}
        for d in all_dates:
            row_data[d] = dates.get(d,"")
        df_total_shift.loc[len(df_total_shift)] = row_data

    st.subheader("班別總表")
    st.dataframe(df_total_shift.head())

    # --- 下載 Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="彙整結果", index=False)
        df_analysis.to_excel(writer, sheet_name="班別分析", index=False)
        df_total_shift.to_excel(writer, sheet_name="班別總表", index=False)
    st.download_button(
        label="下載班表 Excel",
        data=output.getvalue(),
        file_name="班表結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
