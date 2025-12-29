import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime
import re 
from itertools import cycle # 新增：用來做 {sta}/{res} 的循環填補

# --------------------
# 模組 1：解除合併儲存格並填入原值 (維持原樣)
# --------------------
def unmerge_and_fill(ws):
    for merged in list(ws.merged_cells.ranges):
        value = ws.cell(merged.min_row, merged.min_col).value
        ws.unmerge_cells(str(merged))
        for row in ws[merged.coord]:
            for cell in row:
                cell.value = value

# --------------------
# 模組 2：整理班表資料 (維持原樣)
# --------------------
def consolidate_selected_sheets(wb, sheet_names):
    all_data = []
    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        unmerge_and_fill(ws)
        try:
            clinic_name = str(ws.cell(row=1, column=1).value).strip()[:4] 
        except:
            clinic_name = ""
            
        max_row = ws.max_row
        max_col = ws.max_column
        for r in range(1, max_row + 1):
            for c in range(2, max_col + 1):
                cell_value = ws.cell(r, c).value
                if isinstance(cell_value, datetime):
                    date_val = cell_value
                    i = r + 3
                    while i <= max_row:
                        shift_type = str(ws.cell(i, c).value).strip()
                        if isinstance(ws.cell(i, c).value, datetime) or shift_type == "":
                            break
                        if shift_type in ["早", "午", "晚"]:
                            i += 1
                            while i <= max_row:
                                if isinstance(ws.cell(i, c).value, datetime):
                                    break
                                val = str(ws.cell(i, c).value).strip()
                                if val in ["早", "午", "晚"]:
                                    break
                                all_data.append([
                                    clinic_name,
                                    date_val.strftime("%Y/%m/%d"),
                                    shift_type,
                                    val
                                ])
                                i += 1
                            i -= 1
                        i += 1
    df = pd.DataFrame(all_data, columns=["診所", "日期", "班別", "姓名"])
    return df

# --------------------
# 模組 3：建立班別分析表 (維持原樣)
# --------------------
def create_shift_analysis(df_shift: pd.DataFrame, df_emp: pd.DataFrame, shift_map: dict) -> pd.DataFrame:
    df_shift = df_shift.copy()
    df_emp = df_emp.copy()
    df_shift.columns = [str(c).strip() for c in df_shift.columns]
    df_emp.columns = [str(c).strip() for c in df_emp.columns]

    # 建立員工資訊字典
    emp_dict = {}
    for _, row in df_emp.iterrows():
        name = str(row.get("員工姓名", "")).strip()
        if name:
            emp_dict[name] = [
                str(row.get("員工編號", "")).strip(),
                str(row.get("所屬部門", "")).strip(),
                str(row.get("職稱", "")).strip(), # 這裡會讀取職稱
                str(row.get("分類", "")).strip(),
                str(row.get("特殊早班", "")).strip()
            ]

    # 組合每日班別
    shift_dict = {}
    for _, row in df_shift.iterrows():
        name = str(row.get("姓名", "")).strip()
        clinic = str(row.get("診所", "")).strip()
        date_val = row.get("日期", "")
        shift_type = str(row.get("班別", "")).strip()
        if not name or pd.isna(date_val):
            continue
        key = f"{name}|{date_val}|{clinic}"
        if key not in shift_dict:
            shift_dict[key] = set()
        shift_dict[key].add(shift_type)

    data_out = []
    for key, shifts in shift_dict.items():
        name, date_val, clinic = key.split("|")
        if name not in emp_dict:
            continue
        
        # 組合班別：排序後字串 (例如 "早午")
        shift_parts = [s for s in ["早", "午", "晚"] if s in shifts]
        shift_type_for_code = "".join(sorted(shift_parts, key=lambda x: {"早": 1, "午": 2, "晚": 3}.get(x, 9)))

        emp_info = emp_dict.get(name, ["", "", "", "", ""])
        emp_id, emp_dept, emp_title, emp_category, emp_early_special = emp_info
        
        class_code = get_class_code(emp_category, emp_early_special, clinic, shift_type_for_code, shift_map)
        original_shift_type = shift_type_for_code

        data_out.append([clinic, emp_id, emp_dept, name, emp_title, date_val, original_shift_type, class_code])

    df_analysis = pd.DataFrame(
        data_out,
        columns=["診所", "員工編號", "所屬部門", "姓名", "職稱", "日期", "班別", "班別代碼"]
    )

    invalid_names = ["None", "nan", "義診", "單診", "盤點", "電打"]
    df_analysis = df_analysis[~df_analysis["姓名"].astype(str).str.strip().isin(invalid_names)].copy()
    return df_analysis


def get_class_code(emp_category, emp_early_special, clinic_name, shift_type, shift_map):
    """
    (維持原樣) 根據員工類別、特殊旗標、診所和班別類型計算排班代碼。
    """
    region = "立丞" if re.search(r"立丞", str(clinic_name), re.IGNORECASE) else "板土中京"
    is_early_special = str(emp_early_special).strip().lower() in ["是", "true"]

    # 1. 特殊早班
    if is_early_special and "早" in shift_type:
        if shift_type == "早": return "【員工】純早班"
        elif shift_type == "早午": return f"【員工】{region}純早、午班"
        elif shift_type == "早晚": return f"【員工】{region}純早、晚班"
        elif shift_type == "早午晚": return f"【員工】{region}純早午晚班"
    
    # 2. 一般單一早班
    if shift_type == "早":
        if emp_category == "★醫師★": return "★醫師★早班"
        elif emp_category == "◇主管◇": return "◇主管◇早班"
        elif emp_category == "【員工】": return "【員工】早班"

    # 3. 全天班轉換
    if shift_type == "早午晚":
        return f"{emp_category}{region}全天班"
    
    # 4. 預設分類
    base_shift = shift_map.get(shift_type)
    if base_shift is None: base_shift = shift_type
    
    if not str(base_shift).strip().endswith("班"):
        base_shift += "班" 
    
    class_code = emp_category + region + base_shift
    return class_code

# --------------------
# 模組 4：建立班別總表 (★修改處：新增自動填補邏輯)
# --------------------
def create_shift_summary(df_analysis: pd.DataFrame) -> pd.DataFrame:
    if df_analysis.empty:
        return pd.DataFrame()
    df_analysis = df_analysis.copy()
    df_analysis["日期"] = pd.to_datetime(df_analysis["日期"], errors="coerce")
    df_analysis = df_analysis.dropna(subset=["日期"])
    all_dates = sorted(df_analysis["日期"].dt.strftime("%Y-%m-%d").unique())

    # 建立一個職稱對照表 (用於判斷是否排除填補)
    # 以 (員工編號, 姓名) 為 Key，抓取對應的職稱
    emp_title_map = df_analysis[["員工編號", "姓名", "職稱"]].drop_duplicates().set_index(["員工編號", "姓名"])["職稱"].to_dict()

    summary_dict = {}
    for _, row in df_analysis.iterrows():
        emp_id = str(row["員工編號"])
        emp_name = str(row["姓名"])
        if not emp_name or emp_name.strip() in ["None", "nan"]:
            continue
        shift_date = row["日期"].strftime("%Y-%m-%d")
        class_code = row["班別代碼"]
        key = (emp_id, emp_name)
        if key not in summary_dict:
            summary_dict[key] = {}
        summary_dict[key][shift_date] = class_code

    data_out = []
    # 遍歷每一位員工
    for (emp_id, emp_name), shifts in summary_dict.items():
        
        # --- 步驟 3 新增邏輯：判斷是否需要自動填補 ---
        # 取得職稱
        title = str(emp_title_map.get((emp_id, emp_name), "")).strip()
        
        # 設定排除關鍵字：只要職稱包含 "醫師" 或 "兼職" 就不填
        is_excluded = ("醫師" in title) or ("兼職" in title)
        
        # 準備循環填補器：{sta} -> {res} -> {sta} ...
        leave_cycle = cycle(["{sta}", "{res}"])
        # ----------------------------------------

        row = [emp_id, emp_name]
        for d in all_dates:
            val = shifts.get(d, "") # 取得當天班別(可能為空)
            
            # 如果格子是空的，且該員工【不是】被排除的對象(不是醫師也不是兼職)
            if val == "" and not is_excluded:
                # 執行自動填補
                val = next(leave_cycle)
                
            row.append(val)
        data_out.append(row)

    columns = ["員工編號", "員工姓名"] + all_dates
    return pd.DataFrame(data_out, columns=columns)

# --------------------
# Streamlit 主程式 (維持原樣)
# --------------------
st.title("班表處理器")
st.markdown("""
**規則說明**：
1. 依據上傳的班表與員工資料進行代碼轉換。
2. **自動填補**：產出結果後，針對空白的班別進行自動填補 `{sta}` 與 `{res}`。
   * **排除對象**：職稱包含「醫師」或「兼職」的人員，空白班維持空白，不進行填補。
""")

shift_file = st.file_uploader("上傳班表 Excel 檔案", type=["xlsx", "xlsm"])
employee_file = st.file_uploader("上傳員工資料 Excel 檔案", type=["xlsx", "xlsm"])

if shift_file and employee_file:
    wb_shift = load_workbook(shift_file)
    wb_emp = load_workbook(employee_file)

    selectable_sheets = [s for s in wb_shift.sheetnames if s not in ["彙整結果", "班別分析", "班別總表"]]
    selected_sheets = st.multiselect("選擇要處理的工作表", selectable_sheets)
    employee_sheet_name = st.selectbox("選擇員工資料工作表", wb_emp.sheetnames)

    if st.button("開始處理"):
        if not selected_sheets:
            st.warning("請至少選擇一個工作表！")
        else:
            df_shift = consolidate_selected_sheets(wb_shift, selected_sheets)
            
            ws_emp = wb_emp[employee_sheet_name]
            data_emp = ws_emp.values
            cols_emp = [str(c).strip() for c in next(data_emp)]
            df_emp = pd.DataFrame(data_emp, columns=cols_emp)

            # shift_map 僅包含單班別的映射，複雜組合會在 get_class_code 中處理
            shift_map = {"早": "早", "午": "午", "晚": "晚"} 

            df_analysis = create_shift_analysis(df_shift, df_emp, shift_map)
            df_summary = create_shift_summary(df_analysis)

            st.success("班別總表已生成完成！")
            st.subheader("班別總表（含自動填補結果）")
            st.dataframe(df_summary)

            # --------------------
            # 下載 Excel（僅班別總表）
            # --------------------
            with BytesIO() as output:
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_summary.to_excel(writer, sheet_name="班別總表", index=False)
                st.download_button(
                    "下載班別總表Excel",
                    data=output.getvalue(),
                    file_name="班別總表_自動填補.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
