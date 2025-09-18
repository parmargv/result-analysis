import pandas as pd
import os
import xlwings as xw


def result_ana(df,branch):
    # excel_file ="gtu_result_analysis.xlsx"
    # wb=xw.Book(excel_file)
    absolute_path = os.path.dirname(__file__)
    file_path = os.path.join(absolute_path, 'gtu_result_analysis.xlsx')
    wb=xw.Book(file_path)
    # df = pd.read_excel(data)
    # df1 = pd.read_excel("BRANCH_CODE.xlsx")

    df = df[df['BR_CODE'] == branch]

    df = df.sort_values(by='MAP_NUMBER')
    sem =df['sem'].iloc[0]
    reg_num_s =df['MAP_NUMBER'].iloc[0]
    reg_num_e = reg_num_s +199
    cer_num_s =reg_num_s +8000000
    cer_num_e =cer_num_s +199

    cer =wb.sheets("C_TO_D")

    if sem==1 or sem==2:
        df_reg = df[(df['MAP_NUMBER'] >= reg_num_s) & (df['MAP_NUMBER'] <= reg_num_e)]
        df_cer = df[(df['MAP_NUMBER'] >= cer_num_s) & (df['MAP_NUMBER'] <= cer_num_e)]
        df2 = df_cer[["MAP_NUMBER", "name", 'SUB1', 'SUB2', 'SUB3', 'SUB4', 'SUB5', 'SUB6', 'SUB7', 'SUB8',
                 'SUB1NA', 'SUB2NA', 'SUB3NA', 'SUB4NA', 'SUB5NA', 'SUB6NA', 'SUB7NA', 'SUB8NA',
                 'SUB1GR', 'SUB2GR', 'SUB3GR', 'SUB4GR', 'SUB5GR', 'SUB6GR', 'SUB7GR', 'SUB8GR',
                 'SUB1GRI', 'SUB2GRI', 'SUB3GRI', 'SUB4GRI', 'SUB5GRI', 'SUB6GRI', 'SUB7GRI', 'SUB8GRI',
                 'SUB1GRE', 'SUB2GRE', 'SUB3GRE', 'SUB4GRE', 'SUB5GRE', 'SUB6GRE', 'SUB7GRE', 'SUB8GRE',
                 'SUB1GRM', 'SUB2GRM', 'SUB3GRM', 'SUB4GRM', 'SUB5GRM', 'SUB6GRM', 'SUB7GRM', 'SUB8GRM',
                 'SUB1GRV', 'SUB2GRV', 'SUB3GRV', 'SUB4GRV', 'SUB5GRV', 'SUB6GRV', 'SUB7GRV', 'SUB8GRV',
                 'SPI', 'CPI', 'CGPA', 'RESULT']]
        cer.range("A1:Z100").clear()
        df_1 = df2[["MAP_NUMBER", "name", 'SUB1', 'SUB1NA', 'SUB1GR', 'SUB1GRI', 'SUB1GRM', 'SUB1GRE', 'SUB1GRV', 'SPI', 'CPI','CGPA', 'RESULT']]
        dict = {'SUB1GR': 'SUB_GRADE', 'SUB1GRI': 'PA_PR', 'SUB1GRE': 'ESE_TH', 'SUB1GRM': 'PA_TH', 'SUB1GRV': 'ESE_PR','RESULT': 'SEM_RESULT'}
        df_1.rename(columns=dict, inplace=True)
        cer.range("A1").options().value = df_1
        df_2 = df2[["MAP_NUMBER", "name", 'SUB2', 'SUB2NA', 'SUB2GR', 'SUB2GRI', 'SUB2GRM', 'SUB2GRE', 'SUB2GRV', 'SPI', 'CPI','CGPA', 'RESULT']]
        dict = {'SUB2GR': 'SUB_GRADE', 'SUB2GRI': 'PA_PR', 'SUB2GRE': 'ESE_TH', 'SUB2GRM': 'PA_TH', 'SUB2GRV': 'ESE_PR','RESULT': 'SEM_RESULT'}
        df_2.rename(columns=dict, inplace=True)
        cer.range("A10").options().value = df_2
        df_3 = df2[["MAP_NUMBER", "name", 'SUB3', 'SUB3NA', 'SUB3GR', 'SUB3GRI', 'SUB3GRM', 'SUB3GRE', 'SUB3GRV', 'SPI', 'CPI','CGPA', 'RESULT']]
        dict = {'SUB3GR': 'SUB_GRADE', 'SUB3GRI': 'PA_PR', 'SUB3GRE': 'ESE_TH', 'SUB3GRM': 'PA_TH', 'SUB3GRV': 'ESE_PR','RESULT': 'SEM_RESULT'}
        df_3.rename(columns=dict, inplace=True)
        cer.range("A20").options().value = df_3
        df_4 = df2[["MAP_NUMBER", "name", 'SUB4', 'SUB4NA', 'SUB4GR', 'SUB4GRI', 'SUB4GRM', 'SUB4GRE', 'SUB4GRV', 'SPI', 'CPI','CGPA', 'RESULT']]
        dict = {'SUB4GR': 'SUB_GRADE', 'SUB4GRI': 'PA_PR', 'SUB4GRE': 'ESE_TH', 'SUB4GRM': 'PA_TH', 'SUB4GRV': 'ESE_PR','RESULT': 'SEM_RESULT'}
        df_4.rename(columns=dict, inplace=True)
        cer.range("A30").options().value = df_4
        df_5 = df2[["MAP_NUMBER", "name", 'SUB5', 'SUB5NA', 'SUB5GR', 'SUB5GRI', 'SUB5GRM', 'SUB5GRE', 'SUB5GRV', 'SPI', 'CPI','CGPA', 'RESULT']]
        dict = {'SUB5GR': 'SUB_GRADE', 'SUB5GRI': 'PA_PR', 'SUB5GRE': 'ESE_TH', 'SUB5GRM': 'PA_TH', 'SUB5GRV': 'ESE_PR','RESULT': 'SEM_RESULT'}
        df_5.rename(columns=dict, inplace=True)
        cer.range("A40").options().value = df_5
        df_6 = df2[["MAP_NUMBER", "name", 'SUB6', 'SUB6NA', 'SUB6GR', 'SUB6GRI', 'SUB6GRM', 'SUB6GRE', 'SUB6GRV', 'SPI', 'CPI','CGPA', 'RESULT']]
        dict = {'SUB6GR': 'SUB_GRADE', 'SUB6GRI': 'PA_PR', 'SUB6GRE': 'ESE_TH', 'SUB6GRM': 'PA_TH', 'SUB6GRV': 'ESE_PR','RESULT': 'SEM_RESULT'}
        df_6.rename(columns=dict, inplace=True)
        cer.range("A50").options().value = df_6

    else:
        df_reg=df
        cer.range("A1:Z100").clear()


    # NAME1 =df['SUB1NA'].iloc[0]
    # NAME2 =df['SUB2NA'].iloc[0]
    # NAME3 =df['SUB3NA'].iloc[0]
    # NAME4 =df['SUB4NA'].iloc[0]
    # NAME5 =df['SUB5NA'].iloc[0]
    # NAME6 =df['SUB6NA'].iloc[0]
    # NAME7 =df['SUB7NA'].iloc[0]
    # NAME8 =df['SUB8NA'].iloc[0]

    s1 = wb.sheets['sub1']
    s2 = wb.sheets['sub2']
    s3 = wb.sheets['sub3']
    s4 = wb.sheets['sub4']
    s5 = wb.sheets['sub5']
    s6 = wb.sheets['sub6']
    s7 = wb.sheets['sub7']
    s8 = wb.sheets['sub8']
    ex = wb.sheets['exam']
    lst =wb.sheets['list']
    # lst =wb.sheets("LIST")

    # Folder path where your 21 Excel files are located

    exam =df['exam'].iloc[-1]
    ex.range("B8:N30").clear()

    df = df_reg[["MAP_NUMBER","name",'SUB1','SUB2','SUB3','SUB4','SUB5','SUB6','SUB7','SUB8',
             'SUB1NA','SUB2NA','SUB3NA','SUB4NA','SUB5NA','SUB6NA','SUB7NA','SUB8NA',
             'SUB1GR','SUB2GR','SUB3GR','SUB4GR','SUB5GR','SUB6GR','SUB7GR','SUB8GR',
             'SUB1GRI','SUB2GRI','SUB3GRI','SUB4GRI','SUB5GRI','SUB6GRI','SUB7GRI','SUB8GRI',
             'SUB1GRE','SUB2GRE','SUB3GRE','SUB4GRE','SUB5GRE','SUB6GRE','SUB7GRE','SUB8GRE',
             'SUB1GRM','SUB2GRM','SUB3GRM','SUB4GRM','SUB5GRM','SUB6GRM','SUB7GRM','SUB8GRM',
              'SUB1GRV','SUB2GRV','SUB3GRV','SUB4GRV','SUB5GRV','SUB6GRV','SUB7GRV','SUB8GRV',
              'SPI','CPI','CGPA','RESULT']]



    #------------------------------------------------------------------SUB-1-----------------------------------------------
    df1 =df[["MAP_NUMBER","name",'SUB1','SUB1NA','SUB1GR','SUB1GRI','SUB1GRM','SUB1GRE','SUB1GRV','SPI','CPI','CGPA','RESULT']]
    dict ={'SUB1GR':'SUB_GRADE','SUB1GRI':'PA_PR','SUB1GRE':'ESE_TH','SUB1GRM':'PA_TH','SUB1GRV':'ESE_PR','RESULT':'SEM_RESULT'}
    df1.rename(columns=dict,inplace=True)
    df1 = df1.reset_index(drop=True)
    # df1 =df1[df1['SUB1NA']==NAME1]

    df1_fail=df1[df1['SUB_GRADE']=="FF"]
    if len(df1_fail)==0:
        df_f1=[]
    else:
        df_f1=df1_fail[['MAP_NUMBER','name']]
        df_f1 = df_f1.reset_index(drop=True)
    lst.range("A1:Z100").clear()
    lst.range("B2").options().value=df_f1

    TOTAL =len(df1)
    FF = df1[df1['SUB_GRADE'] == 'FF'].shape[0]
    PASS=TOTAL-FF

    PER=(PASS/TOTAL)*100

    RES = df1[df1['SEM_RESULT'] == 'PASS'].shape[0]
    R_PASS=RES
    R_PER=(R_PASS/TOTAL)*100

    code = df1['SUB1'][1]
    name = df1['SUB1NA'][1]
    AA = df1[df1['SUB_GRADE'] == 'AA'].shape[0]
    AB = df1[df1['SUB_GRADE'] == 'AB'].shape[0]
    BB = df1[df1['SUB_GRADE'] == 'BB'].shape[0]
    BC = df1[df1['SUB_GRADE'] == 'BC'].shape[0]
    CC = df1[df1['SUB_GRADE'] == 'CC'].shape[0]
    CD = df1[df1['SUB_GRADE'] == 'CD'].shape[0]
    DD = df1[df1['SUB_GRADE'] == 'DD'].shape[0]
    ex.range("B8:N24").clear()

    ex.range("B8").options().value=code
    ex.range("C8").options().value=name

    ex.range("G8").options().value=AA
    ex.range("D8").options().value=TOTAL
    ex.range("E8").options().value=PASS
    ex.range("F8").options().value=FF
    ex.range("G8").options().value=AA
    ex.range("H8").options().value=AB
    ex.range("I8").options().value=BB
    ex.range("J8").options().value=BC
    ex.range("K8").options().value=CC
    ex.range("L8").options().value=CD
    ex.range("M8").options().value=DD
    ex.range("N8").options().value=PER
    ex.range("M4").options().value=R_PER
    ex.range("A4").options().value = exam
    ex.range("G4").options().value=TOTAL
    ex.range("I4").options().value=RES
    #-------------------------------------------------------------------------SUB-2-----------------------------------------
    s1.range("A1:Z100").clear()
    s1.range("A1").options().value=df1

    df2 =df[["MAP_NUMBER","name",'SUB2','SUB2NA','SUB2GR','SUB2GRI','SUB2GRM','SUB2GRE','SUB2GRV','SPI','CPI','CGPA','RESULT']]
    dict ={'SUB2GR':'SUB_GRADE','SUB2GRI':'PA_PR','SUB2GRE':'ESE_TH','SUB2GRM':'PA_TH','SUB2GRV':'ESE_PR','RESULT':'SEM_RESULT'}
    df2.rename(columns=dict,inplace=True)
    df2 = df2.reset_index(drop=True)
    # df2 =df2[df2['SUB2NA']==NAME2]
    df2_fail=df2[df2['SUB_GRADE']=="FF"]
    if len(df2_fail)==0:
        df_f2=[]
    else:
        df_f2=df2_fail[['MAP_NUMBER','name']]

    df_f2 = df_f2.reset_index(drop=True)
    lst.range("E2").options().value=df_f2

    df_fail=df2[df2['SUB_GRADE']=="FF"]
    TOTAL =len(df2)
    FF = df2[df2['SUB_GRADE'] == 'FF'].shape[0]
    PASS=TOTAL-FF
    PER=(PASS/TOTAL)*100
    code = df2['SUB2'][1]
    name = df2['SUB2NA'][1]
    AA = df2[df2['SUB_GRADE'] == 'AA'].shape[0]
    AB = df2[df2['SUB_GRADE'] == 'AB'].shape[0]
    BB = df2[df2['SUB_GRADE'] == 'BB'].shape[0]
    BC = df2[df2['SUB_GRADE'] == 'BC'].shape[0]
    CC = df2[df2['SUB_GRADE'] == 'CC'].shape[0]
    CD = df2[df2['SUB_GRADE'] == 'CD'].shape[0]
    DD = df2[df2['SUB_GRADE'] == 'DD'].shape[0]

    ex.range("B9").options().value=code
    ex.range("C9").options().value=name
    ex.range("G9").options().value=AA
    ex.range("D9").options().value=TOTAL
    ex.range("E9").options().value=PASS
    ex.range("F9").options().value=FF
    ex.range("G9").options().value=AA
    ex.range("H9").options().value=AB
    ex.range("I9").options().value=BB
    ex.range("J9").options().value=BC
    ex.range("K9").options().value=CC
    ex.range("L9").options().value=CD
    ex.range("M9").options().value=DD
    ex.range("N9").options().value=PER

    s2.range("A1:Z100").clear()
    s2.range("A1").options().value=df2
    #---------------------------------------------------------------------SUB-3------------------------------------
    df3 =df[["MAP_NUMBER","name",'SUB3','SUB3NA','SUB3GR','SUB3GRI','SUB3GRM','SUB3GRE','SUB3GRV','SPI','CPI','CGPA','RESULT']]
    dict ={'SUB3GR':'SUB_GRADE','SUB3GRI':'PA_PR','SUB3GRE':'ESE_TH','SUB3GRM':'PA_TH','SUB3GRV':'ESE_PR','RESULT':'SEM_RESULT'}
    df3.rename(columns=dict,inplace=True)
    df3 = df3.reset_index(drop=True)
    # df3 =df3[df3['SUB3NA']==NAME3]
    df3_fail=df3[df3['SUB_GRADE']=="FF"]
    if len(df3_fail)==0:
        df_f3=[]
    else:
        df_f3=df3_fail[['MAP_NUMBER','name']]
        df_f3 = df_f3.reset_index(drop=True)
    lst.range("H2").options().value=df_f3

    TOTAL =len(df3)
    FF = df3[df3['SUB_GRADE'] == 'FF'].shape[0]
    PASS=TOTAL-FF
    PER=(PASS/TOTAL)*100
    df_res=df3[df3['SUB_GRADE']=="FF"]
    code = df3['SUB3'][1]
    name = df3['SUB3NA'][1]
    AA = df3[df3['SUB_GRADE'] == 'AA'].shape[0]
    AB = df3[df3['SUB_GRADE'] == 'AB'].shape[0]
    BB = df3[df3['SUB_GRADE'] == 'BB'].shape[0]
    BC = df3[df3['SUB_GRADE'] == 'BC'].shape[0]
    CC = df3[df3['SUB_GRADE'] == 'CC'].shape[0]
    CD = df3[df3['SUB_GRADE'] == 'CD'].shape[0]
    DD = df3[df3['SUB_GRADE'] == 'DD'].shape[0]

    ex.range("B10").options().value=code
    ex.range("C10").options().value=name
    ex.range("G10").options().value=AA
    ex.range("D10").options().value=TOTAL
    ex.range("E10").options().value=PASS
    ex.range("F10").options().value=FF
    ex.range("G10").options().value=AA
    ex.range("H10").options().value=AB
    ex.range("I10").options().value=BB
    ex.range("J10").options().value=BC
    ex.range("K10").options().value=CC
    ex.range("L10").options().value=CD
    ex.range("M10").options().value=DD
    ex.range("N10").options().value=PER

    s3.range("A1:Z100").clear()
    s3.range("A1").options().value=df3
    #---------------------------------------------------------------------SUB-4--------------------------------
    df4 =df[["MAP_NUMBER","name",'SUB4','SUB4NA','SUB4GR','SUB4GRI','SUB4GRM','SUB4GRE','SUB4GRV','SPI','CPI','CGPA','RESULT']]
    dict ={'SUB4GR':'SUB_GRADE','SUB4GRI':'PA_PR','SUB4GRE':'ESE_TH','SUB4GRM':'PA_TH','SUB4GRV':'ESE_PR','RESULT':'SEM_RESULT'}
    df4.rename(columns=dict,inplace=True)
    df4 = df4.reset_index(drop=True)
    # df4 =df4[df4['SUB4NA']==NAME4]
    df4_fail=df4[df4['SUB_GRADE']=="FF"]

    if len(df4_fail)==0:
        df_f4=[]
    else:
        df_f4=df4_fail[['MAP_NUMBER','name']]
        df_f4 = df_f4.reset_index(drop=True)
    lst.range("K2").options().value=df_f4
    TOTAL =len(df4)
    FF = df4[df4['SUB_GRADE'] == 'FF'].shape[0]
    PASS=TOTAL-FF
    PER=(PASS/TOTAL)*100
    code = df4['SUB4'][1]
    name = df4['SUB4NA'][1]
    AA = df4[df4['SUB_GRADE'] == 'AA'].shape[0]
    AB = df4[df4['SUB_GRADE'] == 'AB'].shape[0]
    BB = df4[df4['SUB_GRADE'] == 'BB'].shape[0]
    BC = df4[df4['SUB_GRADE'] == 'BC'].shape[0]
    CC = df4[df4['SUB_GRADE'] == 'CC'].shape[0]
    CD = df4[df4['SUB_GRADE'] == 'CD'].shape[0]
    DD = df4[df4['SUB_GRADE'] == 'DD'].shape[0]

    ex.range("B11").options().value=code
    ex.range("C11").options().value=name
    ex.range("G11").options().value=AA
    ex.range("D11").options().value=TOTAL
    ex.range("E11").options().value=PASS
    ex.range("F11").options().value=FF
    ex.range("G11").options().value=AA
    ex.range("H11").options().value=AB
    ex.range("I11").options().value=BB
    ex.range("J11").options().value=BC
    ex.range("K11").options().value=CC
    ex.range("L11").options().value=CD
    ex.range("M11").options().value=DD
    ex.range("N11").options().value=PER

    s4.range("A1:Z100").clear()
    s4.range("A1").options().value=df4
    #-------------------------------------------------------------------------SUB-5----------------------------------------
    df12 =df[["MAP_NUMBER","name",'SUB5','SUB5NA','SUB5GR','SUB5GRI','SUB5GRM','SUB5GRE','SUB5GRV','SPI','CPI','CGPA','RESULT']]
    dict ={'SUB5GR':'SUB_GRADE','SUB5GRI':'PA_PR','SUB5GRE':'ESE_TH','SUB5GRM':'PA_TH','SUB5GRV':'ESE_PR','RESULT':'SEM_RESULT'}
    df12.rename(columns=dict,inplace=True)
    df12 = df12.reset_index(drop=True)
    # df12 =df12[df12['SUB5NA']==NAME5]
    df12_fail=df12[df12['SUB_GRADE']=="FF"]
    if len(df12_fail)==0:
        df_f12=[]
    else:
        df_f12=df12_fail[['MAP_NUMBER','name']]
        df_f12 = df_f12.reset_index(drop=True)
    lst.range("N2").options().value=df_f12
    TOTAL =len(df12)
    FF = df12[df12['SUB_GRADE'] == 'FF'].shape[0]
    PASS=TOTAL-FF
    PER=(PASS/TOTAL)*100
    code = df12['SUB5'][1]
    name = df12['SUB5NA'][1]
    AA = df12[df12['SUB_GRADE'] == 'AA'].shape[0]
    AB = df12[df12['SUB_GRADE'] == 'AB'].shape[0]
    BB = df12[df12['SUB_GRADE'] == 'BB'].shape[0]
    BC = df12[df12['SUB_GRADE'] == 'BC'].shape[0]
    CC = df12[df12['SUB_GRADE'] == 'CC'].shape[0]
    CD = df12[df12['SUB_GRADE'] == 'CD'].shape[0]
    DD = df12[df12['SUB_GRADE'] == 'DD'].shape[0]

    ex.range("B12").options().value=code
    ex.range("C12").options().value=name
    ex.range("G12").options().value=AA
    ex.range("D12").options().value=TOTAL
    ex.range("E12").options().value=PASS
    ex.range("F12").options().value=FF
    ex.range("G12").options().value=AA
    ex.range("H12").options().value=AB
    ex.range("I12").options().value=BB
    ex.range("J12").options().value=BC
    ex.range("K12").options().value=CC
    ex.range("L12").options().value=CD
    ex.range("M12").options().value=DD
    ex.range("N12").options().value=PER

    s5.range("A1:Z100").clear()
    s5.range("A1").options().value=df12
    #--------------------------------------------------------------------SUB-6----------------------------------------------------------
    df6 =df[["MAP_NUMBER","name",'SUB6','SUB6NA','SUB6GR','SUB6GRI','SUB6GRM','SUB6GRE','SUB6GRV','SPI','CPI','CGPA','RESULT']]
    dict ={'SUB6GR':'SUB_GRADE','SUB6GRI':'PA_PR','SUB6GRE':'ESE_TH','SUB6GRM':'PA_TH','SUB6GRV':'ESE_PR','RESULT':'SEM_RESULT'}
    df6.rename(columns=dict,inplace=True)
    df6 = df6.reset_index(drop=True)
    # df6 =df6[df6['SUB6NA']==NAME6]

    df6_fail=df6[df6['SUB_GRADE']=="FF"]
    if len(df6_fail)==0:
        df_f6=[]
    else:
        df_f6=df6_fail[['MAP_NUMBER','name']]
        df_f6 = df_f6.reset_index(drop=True)
    lst.range("Q2").options().value=df_f6
    TOTAL =len(df6)
    FF = df6[df6['SUB_GRADE'] == 'FF'].shape[0]
    PASS=TOTAL-FF
    PER=(PASS/TOTAL)*100
    code = df6['SUB6'][1]
    name = df6['SUB6NA'][1]
    AA = df6[df6['SUB_GRADE'] == 'AA'].shape[0]
    AB = df6[df6['SUB_GRADE'] == 'AB'].shape[0]
    BB = df6[df6['SUB_GRADE'] == 'BB'].shape[0]
    BC = df6[df6['SUB_GRADE'] == 'BC'].shape[0]
    CC = df6[df6['SUB_GRADE'] == 'CC'].shape[0]
    CD = df6[df6['SUB_GRADE'] == 'CD'].shape[0]
    DD = df6[df6['SUB_GRADE'] == 'DD'].shape[0]

    ex.range("B13").options().value=code
    ex.range("C13").options().value=name
    ex.range("G13").options().value=AA
    ex.range("D13").options().value=TOTAL
    ex.range("E13").options().value=PASS
    ex.range("F13").options().value=FF
    ex.range("G13").options().value=AA
    ex.range("H13").options().value=AB
    ex.range("I13").options().value=BB
    ex.range("J13").options().value=BC
    ex.range("K13").options().value=CC
    ex.range("L13").options().value=CD
    ex.range("M13").options().value=DD
    ex.range("N13").options().value=PER

    s6.range("A1:Z100").clear()
    s6.range("A1").options().value=df6
    #----------------------------------------------------------------------SUB-7-------------------------------------------
    df7 =df[["MAP_NUMBER","name",'SUB7','SUB7NA','SUB7GR','SUB7GRI','SUB7GRM','SUB7GRE','SUB7GRV','SPI','CPI','CGPA','RESULT']]
    dict ={'SUB7GR':'SUB_GRADE','SUB7GRI':'PA_PR','SUB7GRE':'ESE_TH','SUB7GRM':'PA_TH','SUB7GRV':'ESE_PR','RESULT':'SEM_RESULT'}
    df7.rename(columns=dict,inplace=True)
    df7.rename(columns=dict,inplace=True)
    df7 = df7.reset_index(drop=True)
    # df7 =df7[df7['SUB7NA']==NAME7]
    df7_fail=df7[df7['SUB_GRADE']=="FF"]
    if len(df7_fail)==0:
        df_f7=[]
    else:
        df_f7=df7_fail[['MAP_NUMBER','name']]
        df_f7 = df_f7.reset_index(drop=True)
    lst.range("T2").options().value=df_f7
    TOTAL =len(df7)
    FF = df7[df7['SUB_GRADE'] == 'FF'].shape[0]
    PASS=TOTAL-FF
    PER=(PASS/TOTAL)*100
    code = df7['SUB7'][1]
    name = df7['SUB7NA'][1]
    AA = df7[df7['SUB_GRADE'] == 'AA'].shape[0]
    AB = df7[df7['SUB_GRADE'] == 'AB'].shape[0]
    BB = df7[df7['SUB_GRADE'] == 'BB'].shape[0]
    BC = df7[df7['SUB_GRADE'] == 'BC'].shape[0]
    CC = df7[df7['SUB_GRADE'] == 'CC'].shape[0]
    CD = df7[df7['SUB_GRADE'] == 'CD'].shape[0]
    DD = df7[df7['SUB_GRADE'] == 'DD'].shape[0]

    ex.range("B14").options().value=code
    ex.range("C14").options().value=name
    ex.range("G14").options().value=AA
    ex.range("D14").options().value=TOTAL
    ex.range("E14").options().value=PASS
    ex.range("F14").options().value=FF
    ex.range("G14").options().value=AA
    ex.range("H14").options().value=AB
    ex.range("I14").options().value=BB
    ex.range("J14").options().value=BC
    ex.range("K14").options().value=CC
    ex.range("L14").options().value=CD
    ex.range("M14").options().value=DD
    ex.range("N14").options().value=PER

    s7.range("A1:Z100").clear()
    s7.range("A1").options().value=df7

    #----------------------------------------------------------------------------SUB-8-----------------------------------------------
    df8 =df[["MAP_NUMBER","name",'SUB8','SUB8NA','SUB8GR','SUB8GRI','SUB8GRM','SUB8GRE','SUB8GRV','SPI','CPI','CGPA','RESULT']]
    dict ={'SUB8GR':'SUB_GRADE','SUB8GRI':'PA_PR','SUB8GRE':'ESE_TH','SUB8GRM':'PA_TH','SUB8GRV':'ESE_PR','RESULT':'SEM_RESULT'}
    df8.rename(columns=dict,inplace=True)
    df8 = df8.reset_index(drop=True)
    # df8 =df8[df8['SUB8NA']==NAME8]
    df8_fail=df8[df8['SUB_GRADE']=="FF"]
    if len(df8_fail)==0:
        df_f8=[]
    else:
        df_f8=df8_fail[['MAP_NUMBER','name']]
        df_f8 = df_f8.reset_index(drop=True)
    lst.range("W2").options().value=df_f8

    TOTAL =len(df8)
    FF = df8[df8['SUB_GRADE'] == 'FF'].shape[0]
    PASS=TOTAL-FF
    PER=(PASS/TOTAL)*100
    code = df8['SUB8'][1]
    name = df8['SUB8NA'][1]
    AA = df8[df8['SUB_GRADE'] == 'AA'].shape[0]
    AB = df8[df8['SUB_GRADE'] == 'AB'].shape[0]
    BB = df8[df8['SUB_GRADE'] == 'BB'].shape[0]
    BC = df8[df8['SUB_GRADE'] == 'BC'].shape[0]
    CC = df8[df8['SUB_GRADE'] == 'CC'].shape[0]
    CD = df8[df8['SUB_GRADE'] == 'CD'].shape[0]
    DD = df8[df8['SUB_GRADE'] == 'DD'].shape[0]

    ex.range("B15").options().value=code
    ex.range("C15").options().value=name
    ex.range("G15").options().value=AA
    ex.range("D15").options().value=TOTAL
    ex.range("E15").options().value=PASS
    ex.range("F15").options().value=FF
    ex.range("G15").options().value=AA
    ex.range("H15").options().value=AB
    ex.range("I15").options().value=BB
    ex.range("J15").options().value=BC
    ex.range("K15").options().value=CC
    ex.range("L15").options().value=CD
    ex.range("M15").options().value=DD
    ex.range("N15").options().value=PER
    s8.range("A2:N100").clear()
    s8.range("A1").options().value=df8

    if ex["C13"].value is None or str(ex["C13"].value).strip() == "":
        # Clear values in range D14:N14
        for row in ex["D13:N13"]:
            for cell in row:
                cell.value = None
    if ex["C14"].value is None or str(ex["C14"].value).strip() == "":
        # Clear values in range D14:N14
        for row in ex["D14:N14"]:
            for cell in row:
                cell.value = None
    if ex["C15"].value is None or str(ex["C15"].value).strip() == "":
        # Clear values in range D14:N14
        for row in ex["D15:N15"]:
            for cell in row:
                cell.value = None

    wb.save()
    wb.close()
    return wb




