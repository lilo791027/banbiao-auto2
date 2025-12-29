import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime
import re
from itertools import cycle

# --------------------
# 模組 1：解除合併儲存格並填入原值
# --------------------
def unmerge_and_fill(ws):
    for merged in list(ws.merged_cells.ranges):
        value = ws.cell(merged.min_row, merged.min_col).value
        ws.unmerge_cells(str(merged))
        for row in ws[merged.coord]:
            for cell in row:
                cell.value = value

# --------------------
# 模組 2：整理班表資料
# --------------------
def consolidate_selected_sheets(wb, sheet_names):
    all_data = []
    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        unmerge_and_fill(ws)
        try:
            clinic_name = str(ws.cell(row=1, column=1).value).strip()[:4]
        except:
            clinic_name = "未知診所"

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
                        if shift_type in ["", "None"] or isinstance(ws.cell(i, c).value, datetime):
                            break
                        
                        if shift_type in ["早", "午", "晚"]:
                            i += 1
                            while i <= max_row:
                                cell_v = ws.cell(i, c).value
                                if isinstance(cell_v, datetime):
                                    break
                                val = str(cell_v).strip()
                                if val in ["早", "午", "晚"]:
                                    break
                                if val and val not in ["None", "nan", "="]:
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
# 模組 3：建立班別分析表
# --------------------
def create_shift_analysis(df_shift: pd.DataFrame, df_emp: pd.DataFrame, shift_map: dict) -> pd.DataFrame:
    df_shift = df_shift.copy()
    df_emp = df_emp.copy()
    
    # 清洗欄位名稱
    df_shift.columns = [str(c).replace(" ", "").replace("　", "").strip() for c in df_shift.columns]
    df_emp.columns = [str(c).replace(" ", "").replace("　", "").strip() for c in df_emp.columns]
    
    def get_col_name(df, keywords):
        for col in df.columns:
            for kw in keywords:
                if kw in col: return col
        return None

    col_map = {
        "姓名": get_col_name(df_emp, ["姓名"]),
        "編號": get_col_name(df_emp, ["編號", "工號"]), # 關鍵：抓取員編
        "職稱": get_col_name(df_emp, ["職稱", "職務", "職位"]),
        "部門": get_col_name(df_emp, ["部門", "單位"]),
        "分類": get_col_name(df_emp, ["分類", "類別"]),
        "特殊早班": get_col_name(df_emp, ["特殊早班", "特權"])
    }
    
    emp_dict = {}
    for _, row in df_emp.iterrows():
        name_col = col_map["姓名"]
        if not name_col: continue

        name = str(row.get(name_col, "")).strip()
        if name and name not in ["nan", "None"]:
            emp_dict[name] = [
                str(row.get(col_map["編號"], "")).strip(), # 這是最重要的判斷依據
                str(row.get(col_map["部門"], "")).strip(),
                str(row.get(col_map["職稱"], "")).strip(),
                str(row.get(col_map["分類"], "")).strip(),
                str(row.get(col_map["特殊早班"], "")).strip()
            ]

    shift_dict = {}
    for _, row in df_shift.iterrows():
        name = str(row.get("姓名", "")).strip()
        clinic = str(row.get("診所", "")).strip()
        date_val = row.get("日期", "")
        shift_type = str(row.get("班別", "")).strip()
        if not name or pd.isna(date_val): continue
        key = f"{name}|{date_val}|{clinic}"
        if key not in shift_dict: shift_dict[key] = set()
        shift_dict[key].add(shift_type)

    data_out = []
    for key, shifts in shift_dict.items():
        name, date_val, clinic = key.split("|")
        emp_info = emp_dict.get(name, ["", "", "", "", ""])
        emp_id, emp_dept, emp_title, emp_category, emp_early_special = emp_info
        
        shift_parts = [s for s in ["早", "午", "晚"] if s in shifts]
        shift_type_for_code = "".join(sorted(shift_parts, key=lambda x: {"早": 1, "午": 2, "晚": 3}.get(x, 9)))
        
        class_code = get_class_code(emp_category, emp_early_special, clinic, shift_type_for_code, shift_map)
        
        data_out.append([clinic, emp_id, emp_dept, name, emp_title, date_val, shift_type_for_code, class_code])

    df_analysis = pd.DataFrame(
        data_out,
        columns=["診所", "員工編號", "所屬部門", "姓名", "職稱", "日期", "班別", "班別代碼"]
    )
    
    invalid_names = ["None", "nan", "義診", "單診", "盤點", "電打", ""]
    df_analysis = df_analysis[~df_analysis["姓名"].astype(str).str.strip().isin(invalid_names)].copy()
    
    return df_analysis

def get_class_code(emp_category, emp_early_special, clinic_name, shift_type, shift_map):
    region = "立丞" if re.search(r"立丞", str(clinic_name), re.IGNORECASE) else "板土中京"
    is_early_special = str(emp_early_special).strip().lower() in ["是", "true", "1", "checked"]

    if is_early_special and "早" in shift_type:
        if shift_type == "早": return "【員工】純早班"
        elif shift_type == "早午": return f"【員工】{region}純早、午班"
        elif shift_type == "早晚": return f"【員工】{region}純早、晚班"
        elif shift_type == "早午晚": return f"【員工】{region}純早午晚班"
    
    if shift_type == "早":
        if "醫師" in emp_category: return "★醫師★早班"
        elif "主管" in emp_category: return "◇主管◇早班"
        elif "員工" in emp_category: return "【員工】早班"

    if shift_type == "早午晚":
        return f"{emp_category}{region}全天班"
    
    base = shift_map.get(shift_type, shift_type)
    if not str(base).strip().endswith("班"): base += "班"
    return str(emp_category) + str(region) + str(base)

# --------------------
# 模組 4：建立班別總表 (邏輯更新：只看員編)
# --------------------
def create_shift_summary(df_analysis: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if df_analysis.empty:
        return pd.DataFrame(), pd.DataFrame()
        
    df_analysis = df_analysis.copy()
    df_analysis["日期"] = pd.to_datetime(df_analysis["日期"], errors="coerce")
    df_analysis = df_analysis.dropna(subset=["日期"])
    all_dates = sorted(df_analysis["日期"].dt.strftime("%Y-%m-%d").unique())

    # 轉置資料
    summary_dict = {}
    for _, row in df_analysis.iterrows():
        emp_id = str(row["員工編號"])
        emp_name = str(row["姓名"])
        shift_date = row["日期"].strftime("%Y-%m-%d")
        summary_dict.setdefault((emp_id, emp_name), {})[shift_date] = row["班別代碼"]

    data_out = []
    debug_list = []

    for (emp_id, emp_name), shifts in summary_dict.items():
        # --- 判斷邏輯更新：只檢查是否有員工編號 ---
        clean_id = str(emp_id).strip()
        
        # 判斷是否為有效員編 (非空, 非None, 非nan)
        has_id = clean_id and clean_id.lower() not in ["nan", "none", ""]
        
        # 若【有員編】則【要填補】(should_fill = True)
        should_fill = has_id
        
        # 收集診斷資訊
        debug_list.append({
            "姓名": emp_name,
            "員工編號": clean_id if has_id else "(無)",
            "狀態": "✅ 自動填補" if should_fill else "❌ 不填補",
            "原因": "有員編" if should_fill else "無員編"
        })

        leave_cycle = cycle(["{sta}", "{res}"])
        
        row = [emp_id, emp_name]
        for d in all_dates:
            val = shifts.get(d, "")
            
            # 強力空值判斷
            is_empty = (val is None) or (str(val).strip() in ["", "nan", "None"])
            
            if is_empty:
                if should_fill:
                    val = next(leave_cycle) # 有員編就填
                else:
                    val = "" # 沒員編保持空白
            
            row.append(val)
        data_out.append(row)

    cols = ["員工編號", "員工姓名"] + all_dates
    return pd.DataFrame(data_out, columns=cols), pd.DataFrame(debug_list)

# --------------------
# Streamlit 主程式
# --------------------
st.set_page_config(page_title="班表處理器(員編判斷版)", layout="wide")
st.title("班表處理器 (員編判斷版)")
st.info("填補規則已更新：**只有具備「員工編號」的人員，空班才會自動填補**。")

col1, col2 = st.columns(2)
with col1:
    shift_file = st.file_uploader("1. 上傳班表", type=["xlsx", "xlsm"])
with col2:
    employee_file = st.file_uploader("2. 上傳員工資料", type=["xlsx", "xlsm"])

if shift_file and employee_file:
    try:
        wb_shift = load_workbook(shift_file, data_only=True)
        wb_emp = load_workbook(employee_file, data_only=True)
    except Exception as e:
        st.error(f"檔案讀取失敗: {e}")
        st.stop()

    sheets = [s for s in wb_shift.sheetnames if s not in ["彙整結果", "班別分析", "班別總表"]]
    selected_sheets = st.multiselect("選擇班表工作表", sheets)
    emp_sheet_name = st.selectbox("選擇員工資料工作表", wb_emp.sheetnames)

    if st.button("🚀 開始處理", type="primary"):
        if not selected_sheets:
            st.warning("請選擇至少一個工作表")
        else:
            with st.spinner("資料處理中..."):
                df_shift = consolidate_selected_sheets(wb_shift, selected_sheets)
                
                ws = wb_emp[emp_sheet_name]
                data = list(ws.values)
                if data:
                    cols = [str(c).strip() for c in data[0]]
                    df_emp = pd.DataFrame(data[1:], columns=cols)
                else:
                    st.error("員工資料表是空的！")
                    st.stop()

                shift_map = {"早": "早", "午": "午", "晚": "晚"}
                
                df_analysis = create_shift_analysis(df_shift, df_emp, shift_map)
                df_summary, df_debug = create_shift_summary(df_analysis)
            
            st.success("處理完成！")
            
            with st.expander("🕵️‍♀️ 診斷報告：檢查誰有員編？(點擊展開)", expanded=True):
                st.dataframe(df_debug, use_container_width=True)

            st.subheader("📊 班別總表")
            st.dataframe(df_summary, use_container_width=True)

            with BytesIO() as output:
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_summary.to_excel(writer, sheet_name="班別總表", index=False)
                st.download_button("📥 下載 Excel 結果", output.getvalue(), "班別總表_員編版.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
