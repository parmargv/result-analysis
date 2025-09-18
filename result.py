import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import math

# ---- Helpers ----------------------------------------------------------------
def clear_range(ws, start_row, start_col, end_row, end_col):
    """Clear cells in numeric coordinate range."""
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(row=r, column=c).value = None

def clear_range_a1(ws, a1_range):
    """Clear cells in A1-style range like 'A1:Z100'."""
    for row in ws[a1_range]:
        for cell in row:
            cell.value = None

def write_df_at(ws, df, start_row=1, start_col=1, include_index=False):
    """
    Write pandas DataFrame to worksheet starting at start_row/start_col.
    Writes header in start_row, data from start_row+1.
    """
    if df is None or df.empty:
        return
    cols = list(df.columns)
    # Header
    for j, col in enumerate(cols):
        ws.cell(row=start_row, column=start_col + j).value = col
    # Data rows
    rows = df.values.tolist()
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            ws.cell(row=start_row + 1 + i, column=start_col + j).value = (None if pd.isna(val) else val)


# ---- Main function ----------------------------------------------------------
def result_ana(df: pd.DataFrame, branch):
    """
    Convert original xlwings result_ana to openpyxl.
    This function:
    - loads template 'gtu_result_analysis.xlsx' from same folder
    - filters df by BR_CODE == branch
    - fills sheets: C_TO_D, sub1..sub8, exam, list
    - saves workbook back to same file and returns file_path
    """

    absolute_path = os.path.dirname(__file__)
    file_path = os.path.join(absolute_path, 'gtu_result_analysis.xlsx')

    # load workbook (template)
    wb = load_workbook(file_path)
    # convenience references
    def ws(name): return wb[name]

    # filter df by branch
    df = df.copy()
    if 'BR_CODE' in df.columns:
        df = df[df['BR_CODE'] == branch]
    else:
        # if column not present, return immediately
        wb.save(file_path)
        return file_path

    if df.empty:
        wb.save(file_path)
        return file_path

    df = df.sort_values(by='MAP_NUMBER', ignore_index=True)

    # safe extraction of sem and starting numbers
    try:
        sem = int(df['sem'].iloc[0])
    except Exception:
        sem = None

    try:
        reg_num_s = int(df['MAP_NUMBER'].iloc[0])
    except Exception:
        reg_num_s = 0
    reg_num_e = reg_num_s + 199
    cer_num_s = reg_num_s + 8000000
    cer_num_e = cer_num_s + 199

    # C_TO_D sheet handling (certificate pages)
    cer_ws = ws("C_TO_D")

    if sem in (1, 2):
        df_reg = df[(df['MAP_NUMBER'] >= reg_num_s) & (df['MAP_NUMBER'] <= reg_num_e)].copy()
        df_cer = df[(df['MAP_NUMBER'] >= cer_num_s) & (df['MAP_NUMBER'] <= cer_num_e)].copy()

        # select the big column block exactly as in original
        cols_block = ["MAP_NUMBER", "name",
                      'SUB1', 'SUB2', 'SUB3', 'SUB4', 'SUB5', 'SUB6', 'SUB7', 'SUB8',
                      'SUB1NA', 'SUB2NA', 'SUB3NA', 'SUB4NA', 'SUB5NA', 'SUB6NA', 'SUB7NA', 'SUB8NA',
                      'SUB1GR', 'SUB2GR', 'SUB3GR', 'SUB4GR', 'SUB5GR', 'SUB6GR', 'SUB7GR', 'SUB8GR',
                      'SUB1GRI', 'SUB2GRI', 'SUB3GRI', 'SUB4GRI', 'SUB5GRI', 'SUB6GRI', 'SUB7GRI', 'SUB8GRI',
                      'SUB1GRE', 'SUB2GRE', 'SUB3GRE', 'SUB4GRE', 'SUB5GRE', 'SUB6GRE', 'SUB7GRE', 'SUB8GRE',
                      'SUB1GRM', 'SUB2GRM', 'SUB3GRM', 'SUB4GRM', 'SUB5GRM', 'SUB6GRM', 'SUB7GRM', 'SUB8GRM',
                      'SUB1GRV', 'SUB2GRV', 'SUB3GRV', 'SUB4GRV', 'SUB5GRV', 'SUB6GRV', 'SUB7GRV', 'SUB8GRV',
                      'SPI', 'CPI', 'CGPA', 'RESULT']

        df2 = df_cer.loc[:, [c for c in cols_block if c in df_cer.columns]].copy()

        # clear A1:Z100 then write sub-blocks at A1, A10, A20, A30, A40, A50
        clear_range_a1(cer_ws, "A1:Z100")
        start_rows = [1, 10, 20, 30, 40, 50]
        for i in range(6):
            n = i + 1
            wanted_cols = ["MAP_NUMBER", "name",
                           f"SUB{n}", f"SUB{n}NA",
                           f"SUB{n}GR", f"SUB{n}GRI", f"SUB{n}GRM", f"SUB{n}GRE", f"SUB{n}GRV",
                           "SPI", "CPI", "CGPA", "RESULT"]
            wanted_cols = [c for c in wanted_cols if c in df2.columns]
            df_block = df2.loc[:, wanted_cols].copy()
            # rename grade columns to common names
            rename_map = {}
            if f"SUB{n}GR" in df_block.columns:
                rename_map[f"SUB{n}GR"] = "SUB_GRADE"
            if f"SUB{n}GRI" in df_block.columns:
                rename_map[f"SUB{n}GRI"] = "PA_PR"
            if f"SUB{n}GRE" in df_block.columns:
                rename_map[f"SUB{n}GRE"] = "ESE_TH"
            if f"SUB{n}GRM" in df_block.columns:
                rename_map[f"SUB{n}GRM"] = "PA_TH"
            if f"SUB{n}GRV" in df_block.columns:
                rename_map[f"SUB{n}GRV"] = "ESE_PR"
            if "RESULT" in df_block.columns:
                rename_map["RESULT"] = "SEM_RESULT"
            if rename_map:
                df_block = df_block.rename(columns=rename_map)

            write_df_at(cer_ws, df_block, start_row=start_rows[i], start_col=1)
    else:
        # sem not 1 or 2: use all rows as registration set, clear C_TO_D
        df_reg = df.copy()
        clear_range_a1(cer_ws, "A1:Z100")

    # Setup other sheets
    s_sheets = {}
    for i in range(1, 9):
        name = f"sub{i}"
        if name in wb.sheetnames:
            s_sheets[i] = ws(name)
        else:
            # create if missing to avoid KeyError
            s_sheets[i] = wb.create_sheet(title=name)

    ex = ws("exam")
    lst = ws("list")

    # exam value from last row of df
    exam = df['exam'].iloc[-1] if 'exam' in df.columns else None

    # Clear exam area B8:N30
    # convert B8:N30 to numeric coords: B=2, N=14
    clear_range(ex, 8, 2, 30, 14)

    # reduce df_reg to the big working block
    big_cols = ["MAP_NUMBER", "name",
                'SUB1', 'SUB2', 'SUB3', 'SUB4', 'SUB5', 'SUB6', 'SUB7', 'SUB8',
                'SUB1NA', 'SUB2NA', 'SUB3NA', 'SUB4NA', 'SUB5NA', 'SUB6NA', 'SUB7NA', 'SUB8NA',
                'SUB1GR', 'SUB2GR', 'SUB3GR', 'SUB4GR', 'SUB5GR', 'SUB6GR', 'SUB7GR', 'SUB8GR',
                'SUB1GRI', 'SUB2GRI', 'SUB3GRI', 'SUB4GRI', 'SUB5GRI', 'SUB6GRI', 'SUB7GRI', 'SUB8GRI',
                'SUB1GRE', 'SUB2GRE', 'SUB3GRE', 'SUB4GRE', 'SUB5GRE', 'SUB6GRE', 'SUB7GRE', 'SUB8GRE',
                'SUB1GRM', 'SUB2GRM', 'SUB3GRM', 'SUB4GRM', 'SUB5GRM', 'SUB6GRM', 'SUB7GRM', 'SUB8GRM',
                'SUB1GRV', 'SUB2GRV', 'SUB3GRV', 'SUB4GRV', 'SUB5GRV', 'SUB6GRV', 'SUB7GRV', 'SUB8GRV',
                'SPI', 'CPI', 'CGPA', 'RESULT', 'exam', 'sem', 'BR_CODE']
    available_cols = [c for c in big_cols if c in df_reg.columns]
    df_reg = df_reg.loc[:, available_cols].copy()

    # Clear list sheet area once (like original does before first write)
    clear_range_a1(lst, "A1:Z100")

    # function to write list failures at the correct start columns:
    def write_fail_list(df_fail, subj_idx):
        """Write df_fail (MAP_NUMBER,name) at row 2 and start_col (B/E/H/...)"""
        if df_fail is None or df_fail.empty:
            return
        start_col = 2 + (subj_idx - 1) * 3  # B=2, E=5 ...
        write_df_at(lst, df_fail, start_row=2, start_col=start_col)

    # Now process each subject sheet and exam summary cells
    for i in range(1, 9):
        # build subject-specific DataFrame
        subj_cols = ["MAP_NUMBER", "name",
                     f"SUB{i}", f"SUB{i}NA", f"SUB{i}GR", f"SUB{i}GRI", f"SUB{i}GRM", f"SUB{i}GRE", f"SUB{i}GRV",
                     "SPI", "CPI", "CGPA", "RESULT"]
        subj_cols = [c for c in subj_cols if c in df_reg.columns]
        df_sub = df_reg.loc[:, subj_cols].copy()
        # rename columns to uniform names
        rename_map = {}
        if f"SUB{i}GR" in df_sub.columns:
            rename_map[f"SUB{i}GR"] = "SUB_GRADE"
        if f"SUB{i}GRI" in df_sub.columns:
            rename_map[f"SUB{i}GRI"] = "PA_PR"
        if f"SUB{i}GRE" in df_sub.columns:
            rename_map[f"SUB{i}GRE"] = "ESE_TH"
        if f"SUB{i}GRM" in df_sub.columns:
            rename_map[f"SUB{i}GRM"] = "PA_TH"
        if f"SUB{i}GRV" in df_sub.columns:
            rename_map[f"SUB{i}GRV"] = "ESE_PR"
        if "RESULT" in df_sub.columns:
            rename_map["RESULT"] = "SEM_RESULT"
        if rename_map:
            df_sub = df_sub.rename(columns=rename_map)
        df_sub = df_sub.reset_index(drop=True)

        # find failing rows for list sheet
        df_fail = pd.DataFrame()
        if 'SUB_GRADE' in df_sub.columns:
            df_fail_rows = df_sub[df_sub['SUB_GRADE'] == "FF"]
            if not df_fail_rows.empty:
                df_fail = df_fail_rows[['MAP_NUMBER', 'name']].reset_index(drop=True)

        # write failure list at appropriate place
        write_fail_list(df_fail, i)

        # Compute stats safely
        TOTAL = len(df_sub)
        FF = int(df_sub[df_sub.get('SUB_GRADE', pd.Series()) == 'FF'].shape[0]) if TOTAL > 0 and 'SUB_GRADE' in df_sub.columns else 0
        PASS = TOTAL - FF if TOTAL > 0 else 0
        PER = (PASS / TOTAL) * 100 if TOTAL > 0 else 0.0
        RES = int(df_sub[df_sub.get('SEM_RESULT', pd.Series()) == 'PASS'].shape[0]) if TOTAL > 0 and 'SEM_RESULT' in df_sub.columns else 0
        R_PER = (RES / TOTAL) * 100 if TOTAL > 0 else 0.0

        # subject code and name: be defensive (original used index 1; we prefer index 0 then 1 fallback)
        code = ""
        name = ""
        col_code = f"SUB{i}"
        col_name = f"SUB{i}NA"
        if col_code in df_sub.columns:
            if len(df_sub) > 1:
                code = df_sub.at[1, col_code] if pd.notna(df_sub.at[1, col_code]) else df_sub.at[0, col_code]
            else:
                code = df_sub.at[0, col_code]
        if col_name in df_sub.columns:
            if len(df_sub) > 1:
                name = df_sub.at[1, col_name] if pd.notna(df_sub.at[1, col_name]) else df_sub.at[0, col_name]
            else:
                name = df_sub.at[0, col_name]

        # grade distributions
        grades = {}
        for grade in ['AA', 'AB', 'BB', 'BC', 'CC', 'CD', 'DD']:
            if 'SUB_GRADE' in df_sub.columns:
                grades[grade] = int(df_sub[df_sub['SUB_GRADE'] == grade].shape[0])
            else:
                grades[grade] = 0

        # write results to exam sheet - row mapping: sub1 -> row8, sub2 -> row9, ..., sub8 -> row15
        ex_row = 7 + i
        ex.cell(row=ex_row, column=2).value = code   # B
        ex.cell(row=ex_row, column=3).value = name   # C
        ex.cell(row=ex_row, column=7).value = grades.get('AA', 0)  # G
        ex.cell(row=ex_row, column=4).value = TOTAL  # D
        ex.cell(row=ex_row, column=5).value = PASS   # E
        ex.cell(row=ex_row, column=6).value = FF     # F
        ex.cell(row=ex_row, column=8).value = grades.get('AB', 0)  # H
        ex.cell(row=ex_row, column=9).value = grades.get('BB', 0)  # I
        ex.cell(row=ex_row, column=10).value = grades.get('BC', 0) # J
        ex.cell(row=ex_row, column=11).value = grades.get('CC', 0) # K
        ex.cell(row=ex_row, column=12).value = grades.get('CD', 0) # L
        ex.cell(row=ex_row, column=13).value = grades.get('DD', 0) # M
        ex.cell(row=ex_row, column=14).value = PER    # N

        # For the first subject (like original), set some summary cells (A4, G4, I4, M4)
        if i == 1:
            ex.cell(row=4, column=1).value = exam  # A4
            ex.cell(row=4, column=7).value = TOTAL  # G4
            ex.cell(row=4, column=9).value = RES    # I4
            ex.cell(row=4, column=11).value = R_PER # M4

        # write subject DataFrame into corresponding sub-sheet at A1 after clearing A1:Z100
        target_ws = s_sheets.get(i)
        if target_ws is not None:
            clear_range_a1(target_ws, "A1:Z100")
            # ensure column order as in df_sub
            write_df_at(target_ws, df_sub, start_row=1, start_col=1)

    # After all subjects, check certain cells in exam sheet and clear ranges if blank
    # If C13 blank -> clear D13:N13 (D=4,N=14 col numbers)
    b13_val = ex.cell(row=13, column=2).value
    if b13_val is None or (isinstance(b13_val, float) and math.isnan(b13_val)) or str(b13_val).strip() == "":
        clear_range(ex, 13, 4, 20, 14)

    b14_val = ex.cell(row=14, column=2).value
    if b14_val is None or (isinstance(b14_val, float) and math.isnan(b14_val)) or str(b14_val).strip() == "":
        clear_range(ex, 14, 4, 20, 14)

    b15_val = ex.cell(row=15, column=2).value
    if b15_val is None or (isinstance(b15_val, float) and math.isnan(b15_val)) or str(b15_val).strip() == "":
    # if b15_val is None or str(b15_val).strip() == "":
        clear_range(ex, 15, 4, 20, 14)

    b16_val = ex.cell(row=16, column=2).value
    if b16_val is None or (isinstance(b16_val, float) and math.isnan(b16_val)) or str(b16_val).strip() == "":
        clear_range(ex, 16, 4, 20, 14)

    b17_val = ex.cell(row=17, column=2).value
    if b17_val is None or (isinstance(b17_val, float) and math.isnan(b17_val)) or str(b17_val).strip() == "":
        clear_range(ex, 17, 4, 20, 14)

    b18_val = ex.cell(row=18, column=2).value
    if b18_val is None or (isinstance(b18_val, float) and math.isnan(b18_val)) or str(b18_val).strip() == "":
        clear_range(ex, 18, 4, 20, 14)

    # Save back to same template file (overwrites template). If you'd rather create new file, change file_path.
    wb.save(file_path)

    # Return path so web_data.py can open and download it
    return file_path





