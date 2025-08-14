import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.title("班表自動化系統")

# --- 上傳 Excel ---
uploaded_file = st.file_uploader("請上傳排班 Excel", type=["xlsx"])
if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file)
    sheet_names = wb.sheetnames

    # 下拉選擇工作表
    selected_sheet = st.selectbox("請選擇要使用的工作表", sheet_names)
    ws_total = wb[selected_sheet]
    df_total = pd.DataFrame(ws_total.values)

    st.write(f"目前使用工作表：{selected_sheet}")
    st.dataframe(df_total.head())

    # === 模組 1：解合併並填入原值 ===
    for merged_cell in ws_total.merged_cells.ranges:
        merged_area = ws_total[merged_cell.coord]
        value_to_fill = merged_area[0][0].value
        ws_total.unmerge_cells(str(merged_cell))
        for row in merged_area:
            for cell in row:
                cell.value = value_to_fill
    df_total = pd.DataFrame(ws_total.values)

    # === 模組 2：彙整排班資料 ===
    df_out = pd.DataFrame(columns=["診所","日期","班別","姓名","A欄資料","U欄資料"])
    clinic_name = str(df_total.iloc[0,0])[:4]
    last_row, last_col = df_total.shape
    output_row = 0

    for r in range(last_row):
        for c in range(1, last_col):
            if pd.api.types.is_datetime64_any_dtype(df_total.iloc[r,c]) or isinstance(df_total.iloc[r,c], pd.Timestamp):
                date_value = df_total.iloc[r,c]
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

    # === 模組 3：建立班別分析表 ===
    # 假設員工資料已上傳，這裡可改成 file_uploader
    uploaded_emp = st.file_uploader("請上傳員工人事資料明細表", type=["xlsx"])
    if uploaded_emp:
        wb_emp = openpyxl.load_workbook(uploaded_emp)
        ws_emp = wb_emp[wb_emp.sheetnames[0]]
        df_emp = pd.DataFrame(ws_emp.values)
        emp_dict = {}
        for idx, row in df_emp.iterrows():
            name = str(row[1]).strip()
            if name:
                emp_dict[name] = [str(row[0]), row[2], row[3]]

        # 建立班別分析表
        df_analysis = pd.DataFrame(columns=["診所","員工編號","所屬部門","姓名","職稱","日期","班別","E欄資料","班別代碼"])
        shift_dict = {}
        for idx, row in df_out.iterrows():
            name = str(row["姓名"]).strip()
            key = f"{name}|{row['日期']}|{row['診所']}|{row['A欄資料']}"
            if key not in shift_dict:
                shift_dict[key] = row["班別"]
            else:
                shift_dict[key] += " " + row["班別"]

        # 生成分析表 DataFrame
        for key, shift in shift_dict.items():
            name, date_value, clinic_name, e_value = key.split("|")
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
                ""  # 班別代碼，可用同 VBA 的 GetClassCode 轉邏輯
            ]

        st.subheader("班別分析表")
        st.dataframe(df_analysis.head())

        # === 模組 4：建立班別總表 ===
        # 依 VBA 條件生成總表
        all_dates = pd.date_range("2025-08-01","2025-08-31").strftime("%Y-%m-%d")
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
