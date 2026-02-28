import os
import pandas as pd
from openpyxl import load_workbook


# -------------------------------
# Visitor Counter
# -------------------------------

def get_count():
    if os.path.exists("counter.txt"):
        with open("counter.txt", "r") as f:
            count = int(f.read().strip())
    else:
        count = 0
    return count


def update_count():
    count = get_count() + 1
    with open("counter.txt", "w") as f:
        f.write(str(count))
    return count


# -------------------------------
# Excel Helper Functions
# -------------------------------

def clear_range_a1(ws, a1_range):
    for row in ws[a1_range]:
        for cell in row:
            cell.value = None


def write_df_at(ws, df, start_row=1, start_col=1):
    if df.empty:
        return

    # Write header
    for j, col in enumerate(df.columns):
        ws.cell(row=start_row, column=start_col + j).value = col

    # Write rows
    for i, row in enumerate(df.values):
        for j, value in enumerate(row):
            ws.cell(row=start_row + 1 + i, column=start_col + j).value = value
def get_institute_name(inst_type,inst_code):
    absolute_path = os.path.dirname(__file__)

    if inst_type == "BE":
        file_name = "DE_INST_CODE.xlsx"
    elif inst_type == "DI":
        file_name = "DI_INST_CODE.xlsx"
    else:
        return ""

    file_path = os.path.join(absolute_path, file_name)

    if not os.path.exists(file_path):
        return ""

    df_inst = pd.read_excel(file_path)

    match = df_inst[df_inst["inst_code"] == inst_code]

    if not match.empty:
        return match["inst_name"].iloc[0]

    return ""

# -------------------------------
# MAIN FUNCTION
# -------------------------------

def result_ana(df: pd.DataFrame, branch):

    visitor_count = update_count()

    absolute_path = os.path.dirname(__file__)
    file_path = os.path.join(absolute_path, 'GTU_RESULT_ANALYSIS.xlsx')

    wb = load_workbook(file_path)

    df = df.copy()

    if 'BR_CODE' in df.columns:
        df = df[df['BR_CODE'] == branch]

    if df.empty:
        wb.save(file_path)
        return file_path, visitor_count

    df = df.sort_values(by='MAP_NUMBER', ignore_index=True)
    inst_type = df['extype'].iloc[0]
    inst_code = df['instcode'].iloc[0]
    sem_exam = df['exam'].iloc[0]
    inst_name = get_institute_name(inst_type,inst_code)
    br_name = df['BR_NAME'].iloc[0]

    # Remove old subject sheets except template sheets
    template_sheets = ["exam", "list", "C_TO_D"]


    for sheet in wb.sheetnames:
        if sheet not in template_sheets:
            del wb[sheet]

    # ------------------------------------------
    # CORRECT SUBJECT-WISE SHEET CREATION
    # ------------------------------------------

    # Collect all unique subject codes from SUB1..SUB8
    all_subjects = set()

    for i in range(1, 9):
        col = f"SUB{i}"
        if col in df.columns:
            all_subjects.update(df[col].dropna().unique())

    # Now process each subject only once
    for subject in all_subjects:

        subject_rows = []

        for i in range(1, 9):

            col_code = f"SUB{i}"
            col_name = f"SUB{i}NA"
            col_grade = f"SUB{i}GR"

            if col_code not in df.columns:
                continue

            df_temp = df[df[col_code] == subject].copy()

            if df_temp.empty:
                continue

            required_cols = [
                "MAP_NUMBER",
                "name",
                col_code,
                col_name,
                col_grade,
                "SPI",
                "CPI",
                "CGPA",
                "RESULT"
            ]

            required_cols = [c for c in required_cols if c in df_temp.columns]
            df_temp = df_temp[required_cols]

            rename_map = {
                col_code: "SUB_CODE",
                col_name: "SUB_NAME",
                col_grade: "SUB_GRADE",
                "RESULT": "SEM_RESULT",
                "name": "NAME"
            }

            df_temp.rename(columns=rename_map, inplace=True)

            subject_rows.append(df_temp)

        if not subject_rows:
            continue

        df_sub = pd.concat(subject_rows, ignore_index=True)

        sheet_name = str(subject)[:30]

        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(sheet_name)

        clear_range_a1(ws, "A1:N3")
        clear_range_a1(ws, "B8:N22")
        write_df_at(ws, df_sub, start_row=1, start_col=1)
        # ------------------------------------------
        # EXAM SHEET SUBJECT ANALYSIS (WITH GRADES)
        # ------------------------------------------

        exam_ws = wb["exam"]

        exam_ws["A1"] = inst_code
        exam_ws["C1"] = inst_name
        exam_ws["C2"] = br_name
        exam_ws["A4"] = sem_exam

        # Clear previous exam summary rows
        for row in exam_ws["B8:N30"]:
            for cell in row:
                cell.value = None

        row_pointer = 8  # Start row

        for subject in sorted(all_subjects):

            subject_rows = []

            for i in range(1, 9):

                col_code = f"SUB{i}"
                col_name = f"SUB{i}NA"
                col_grade = f"SUB{i}GR"
                col_res = "RESULT"

                if col_code not in df.columns:
                    continue

                df_temp = df[df[col_code] == subject].copy()

                if df_temp.empty:
                    continue

                required_cols = [col_code, col_name, col_grade, col_res]
                required_cols = [c for c in required_cols if c in df_temp.columns]

                df_temp = df_temp[required_cols]

                rename_map = {
                    col_code: "SUB_CODE",
                    col_name: "SUB_NAME",
                    col_grade: "SUB_GRADE",
                    col_res:"SEM_RES"
                }

                df_temp.rename(columns=rename_map, inplace=True)

                subject_rows.append(df_temp)

            if not subject_rows:
                continue

            df_sub = pd.concat(subject_rows, ignore_index=True)

            TOTAL = len(df_sub)
            FAIL = len(df_sub[df_sub["SUB_GRADE"] == "FF"])
            S_FAIL = len(df_sub[df_sub["SEM_RES"] == "FAIL"])
            S_PASS = len(df_sub[df_sub["SEM_RES"] == "PASS"])
            PASS = TOTAL - FAIL
            PER = round((PASS / TOTAL) * 100, 2) if TOTAL > 0 else 0
            S_PER = round((S_PASS / TOTAL) * 100, 2) if TOTAL > 0 else 0
            exam_ws = wb["exam"]
            exam_ws["G4"] = TOTAL
            exam_ws["I4"] = S_PASS
            exam_ws["K4"] = S_PER
            # Grade Distribution
            grade_list = ["AA", "AB", "BB", "BC", "CC", "CD", "DD"]
            grade_count = {}

            for grade in grade_list:
                grade_count[grade] = len(df_sub[df_sub["SUB_GRADE"] == grade])

            subject_name = df_sub["SUB_NAME"].iloc[0] if "SUB_NAME" in df_sub.columns else ""

            # Write to exam sheet
            exam_ws.cell(row=row_pointer, column=2).value = subject
            exam_ws.cell(row=row_pointer, column=3).value = subject_name
            exam_ws.cell(row=row_pointer, column=4).value = TOTAL
            exam_ws.cell(row=row_pointer, column=5).value = PASS
            exam_ws.cell(row=row_pointer, column=6).value = FAIL

            exam_ws.cell(row=row_pointer, column=7).value = grade_count["AA"]
            exam_ws.cell(row=row_pointer, column=8).value = grade_count["AB"]
            exam_ws.cell(row=row_pointer, column=9).value = grade_count["BB"]
            exam_ws.cell(row=row_pointer, column=10).value = grade_count["BC"]
            exam_ws.cell(row=row_pointer, column=11).value = grade_count["CC"]
            exam_ws.cell(row=row_pointer, column=12).value = grade_count["CD"]
            exam_ws.cell(row=row_pointer, column=13).value = grade_count["DD"]

            exam_ws.cell(row=row_pointer, column=14).value = PER

            row_pointer += 1

    wb.save(file_path)

    return file_path, visitor_count





