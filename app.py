import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="班表自動化", layout="wide")
st.title("班表任務自動化系統 (Python 版本)")

# 上傳檔案
shift_file = st.file_uploader("上傳班表 Excel", type=["xlsx"])
emp_file = st.file_uploader("上傳員工明細 Excel", type=["xlsx"])

if shift_file and emp_file:
    # 讀取 Excel
    shift_df = pd.read_excel(shift_file, header=None)
    emp_df = pd.read_excel(emp_file)

    st.write("### 班表原始資料")
    st.dataframe(shift_df.head())

    st.write("### 員工明細資料")
    st.dataframe(emp_df.head())

    # =========================
    # 模組1：自動解合併
    # =========================
    def unmerge_fill_values(df):
        """
        模擬 VBA 解合併並填入原值
        """
        df_filled = df.ffill(axis=0)
        return df_filled

    shift_df_filled = unmerge_fill_values(shift_df)

    st.write("### 模組1：解合併後班表")
    st.dataframe(shift_df_filled.head())

    # =========================
    # 模組2：整理班別資料（對應 VBA 邏輯）
    # =========================
    def extract_shift_with_blanks(shift_df):
        output_cols = ["診所", "日期", "班別", "姓名", "A欄資料", "U欄資料"]
        output = []

        clinic_name = str(shift_df.iat[0, 0])[:4]
        n_rows, n_cols = shift_df.shape

        for r in range(n_rows):
            for c in range(1, n_cols):
                cell = shift_df.iat[r, c]
                # 嘗試轉日期
                try:
                    date_value = pd.to_datetime(cell, errors='coerce')
                except:
                    continue
                if pd.isna(date_value):
                    continue

                i = r + 3
                while i < n_rows:
                    shift_type = str(shift_df.iat[i, c]).strip()
                    try:
                        if pd.notna(pd.to_datetime(shift_df.iat[i, c], errors='coerce')):
                            break
                    except:
                        pass
                    if shift_type == "":
                        break

                    if shift_type in ["早", "午", "晚"]:
                        i += 1
                        while i < n_rows:
                            next_cell = str(shift_df.iat[i, c]).strip()
                            try:
                                if pd.notna(pd.to_datetime(shift_df.iat[i, c], errors='coerce')):
                                    break
                            except:
                                pass
                            if next_cell in ["早", "午", "晚"]:
                                break

                            date_str = date_value.strftime("%Y/%m/%d")
                            output.append([
                                clinic_name,
                                date_str,
                                shift_type,
                                next_cell,
                                shift_df.iat[i, 0],    # A欄
                                shift_df.iat[i, 20]    # U欄
                            ])
                            i += 1
                        i -= 1
                    i += 1

        out_df = pd.DataFrame(output, columns=output_cols)
        return out_df

    shift_clean_df = extract_shift_with_blanks(shift_df_filled)

    st.write("### 模組2：彙整結果")
    st.dataframe(shift_clean_df.head())

    # =========================
    # 模組3：班別分析表
    # =========================
    def format_shift_order(shift_str):
        result = ''
        for s in ['早','午','晚']:
            if s in shift_str:
                result += s
        return result

    def get_class_code(empTitle, clinicName, shiftType):
        if not empTitle or pd.isna(empTitle):
            return ''
        if empTitle in ["早班護理師", "早班內視鏡助理", "醫務專員", "兼職早班內視鏡助理"]:
            return "【員工】純早班"

        if empTitle == "醫師":
            classCode = "★醫師★"
        elif empTitle in ["櫃臺", "護理師", "兼職護理師", "兼職跟診助理", "副店長", "護士", "藥師"]:
            classCode = "【員工】"
        elif "副店長" in empTitle:
            classCode = "【員工】"
        elif "店長" in empTitle or "採購儲備組長" in empTitle:
            classCode = "◇主管◇"
        else:
            classCode = ""

        if shiftType != "早":
            if clinicName in ["上吉診所","立吉診所","上承診所","立全診所","立竹診所","立順診所","上京診所"]:
                classCode += "板土中京"
            elif clinicName == "立丞診所":
                classCode += "立丞"

        shift_map = {
            "早": "早班",
            "午晚": "午晚班",
            "早午晚": "全天班",
            "早晚": "早晚班",
            "午": "午班",
            "晚": "晚班",
            "早午": "早午班"
        }
        classCode += shift_map.get(shiftType, shiftType)
        if classCode.endswith("早班早班"):
            classCode = classCode.replace("早班早班", "早班")
        return classCode

    # 建立員工字典
    emp_dict = {str(row[1]).strip(): (str(row[0]), row[2], row[3]) for idx, row in emp_df.iterrows()}

    # 合併班別資料
    shift_dict = {}
    for idx, row in shift_clean_df.iterrows():
        name = str(row['姓名']).strip()
        dateValue = row['日期']
        clinicName = row['診所']
        shiftType = row['班別']
        eValue = row['A欄資料']

        if not name or len(name) > 4:
            continue

        key = f"{name}|{dateValue}|{clinicName}"
        if key not in shift_dict:
            shift_dict[key] = shiftType
        else:
            shift_dict[key] += " " + shiftType

    # 輸出班別分析表
    analysis_rows = []
    for key, shifts in shift_dict.items():
        name, dateValue, clinicName = key.split("|")
        shiftType = format_shift_order(shifts)
        if name in emp_dict:
            empID, empDept, empTitle = emp_dict[name]
        else:
            empID = empDept = empTitle = ''
        classCode = get_class_code(empTitle, clinicName, shiftType)
        analysis_rows.append([clinicName, empID, empDept, name, empTitle, dateValue, shiftType, '', classCode])

    analysis_df = pd.DataFrame(analysis_rows, columns=[
        "診所","員工編號","所屬部門","姓名","職稱","日期","班別","E欄資料","班別代碼"
    ])

    st.write("### 模組3：班別分析表")
    st.dataframe(analysis_df.head())

    # =========================
    # 模組4：班別總表
    # =========================
    analysis_df['日期'] = pd.to_datetime(analysis_df['日期'])
    first_date = analysis_df['日期'].min()
    yearInput = first_date.year
    monthInput = first_date.month
    days_in_month = pd.date_range(start=f"{yearInput}-{monthInput}-01",
                                  end=f"{yearInput}-{monthInput}-{pd.Period(f'{yearInput}-{monthInput}').days_in_month}")

    summary_dict = {}
    for idx, row in analysis_df.iterrows():
        empID = str(row['員工編號'])
        empName = row['姓名']
        dateKey = row['日期'].strftime("%Y-%m-%d")
        classCode = row['班別代碼']
        empKey = f"{empID}|{empName}"
        if empKey not in summary_dict:
            summary_dict[empKey] = {}
        summary_dict[empKey][dateKey] = classCode

    summary_rows = []
    for empKey, dates in summary_dict.items():
        empID, empName = empKey.split("|")
        row = [empID, empName] + [dates.get(day.strftime("%Y-%m-%d"), "") for day in days_in_month]
        summary_rows.append(row)

    summary_df = pd.DataFrame(summary_rows, columns=["員工編號","員工姓名"] + [day.strftime("%Y-%m-%d") for day in days_in_month])

    st.write("### 模組4：班別總表")
    st.dataframe(summary_df.head())

    # =========================
    # 下載檔案
    # =========================
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        analysis_df.to_excel(writer, sheet_name="班別分析", index=False)
        summary_df.to_excel(writer, sheet_name="班別總表", index=False)
        writer.save()
    processed_data = output.getvalue()

    st.download_button(
        label="下載整理後班表 Excel",
        data=processed_data,
        file_name="班表分析結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



