import numpy as np
import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
from _datetime import timedelta

def Export2Excel(df_list, sheet_Names, fileName='Lease IndAS 116.xlsx'):
    """ Sub for Excel Export with Format """
    writer = pd.ExcelWriter(fileName, engine='xlsxwriter')

    # df.to_excel(writer, sheet_name=sheetName)
    # workbook  = writer.book
    # worksheet = writer.sheets[sheetName]

    for df, sheetName in zip(df_list, sheet_Names):
        df.to_excel(writer, sheet_name=sheetName, index=False)
        workbook  = writer.book
        worksheet = writer.sheets[sheetName]

        cell_format = workbook.add_format({'bold': True, 'num_format': '#,##0'})
        cell_format.set_font_size(15)
        worksheet.set_row(row = 86, cell_format = cell_format)

        format1numbers = workbook.add_format({'num_format': '#,##0'})
        format2Percentage = workbook.add_format({'num_format': '0%'})
        format3date = workbook.add_format({'num_format': 'mmm yy'})

        worksheet.set_column('B:B', None, format1numbers)
        worksheet.set_column('D:D', None, format2Percentage)
        worksheet.set_column('E:J', None, format1numbers)

        
        if sheetName == "JV Entries":
            worksheet.set_column('B:K', 25)
            worksheet.set_column('D:D', None, format1numbers)
            worksheet.set_column('E:E', 50)
            # worksheet.deleteRows(86,1,True)

        else:
            worksheet.set_column('A:K', 18)
            worksheet.write_formula('A87', '="Grand Total"')
            worksheet.write_formula('B87', '=SUM(C2:C86)') 
            worksheet.write_formula('E87', '=SUM(F2:F86)') 
            worksheet.write_formula('F87', '=SUM(G2:G86)') 
            worksheet.write_formula('G87', '=SUM(H2:H86)') 
            worksheet.write_formula('H87', '=SUM(I2:I86)') 
            worksheet.write_formula('I87', '=SUM(J2:J86)') 


    writer.save()


"""[Read Inputs.txt and gather basic information]"""
fileObj = open("inputs.txt")
fileObj = fileObj.readlines()
_discRate = int(fileObj[0])
_startYear = datetime.datetime.strptime(fileObj[1][:-1], "%Y/%m/%d")
_endYear = _startYear + relativedelta(years=1, days=-1)
_discRate = _discRate / 100
""" x """

"""[Read InputSheet.xlsx and gather month wise information]"""
df = pd.read_excel("InputSheet.xlsx", sheet_name='Sheet1')
_startLease =df.iat[0,0]
""" x """


""" Find Out Last Row of the Table - this will be used to calculate Depreciation """
_lastRow = 0
for index, row in df.iterrows():
    if df.at[index,'Rent Pmt'] > 0:
        _lastRow = index
""" x """


""" Populate Fields  - PV, Lease Liab, Interest"""
df['PV Time'] = df.index #((df['Month'] - _startLease)/ timedelta(days=1))
df['PV Factor'] = 1/((1+_discRate/12) ** (df['PV Time']))
df['Present Value'] = df['PV Factor'] * df['Rent Pmt']
_TotalLeaseLiab = df['Present Value'].sum()

df['Lease Liab'] = 0.0
df['Interest'] = 0.0
df['Closing Liab'] = 0.0

# Calculating Values
jIndex = len(df.index)
for i in range(jIndex):
    # print(i)
    if i < 1:
        df.iloc[i, df.columns.get_loc('Lease Liab')] = _TotalLeaseLiab  - df.at[i,'Rent Pmt']
        df.iloc[i, df.columns.get_loc('Interest')] = df.at[i,'Lease Liab'] * _discRate /12 #  365 * (df.at[i,'PV Time'] - df.at[i+1,'PV Time'])
        df.iloc[i, df.columns.get_loc('Closing Liab')] = df.at[i,'Lease Liab'] + df.at[i,'Interest'] 
    else:
        df.iloc[i, df.columns.get_loc('Lease Liab')] = df.at[i-1, 'Closing Liab'] - df.at[i,'Rent Pmt']
        df.iloc[i, df.columns.get_loc('Interest')] = df.at[i,'Lease Liab'] * _discRate / 12 #365 * (df.at[i,'PV Time'] - df.at[i+1,'PV Time'])
        df.iloc[i, df.columns.get_loc('Closing Liab')] = df.at[i,'Lease Liab'] + df.at[i,'Interest']  

""" Validate whether Lease Liability is getting to Zero """
_CheckLeaseLiabZero = 1
if _lastRow < 100:
    if _lastRow > -100:
        _CheckLeaseLiabZero = 0
    
if _CheckLeaseLiabZero != 0:
   print('Lease Liability is not Zero. Please check inputs.') 

""" Monthly Depreciation Charge """
df['Depreciation'] = _TotalLeaseLiab / (_lastRow + 1) *-1

df = df.drop(df[df.index > _lastRow].index)

# Enable following after all coding done...
# Export2Excel(df)

df_prev = df.drop(df[df.Month >= _startYear].index)
# Export2Excel(df_prev,sheetName='Prior Period')

df_curyear = df.drop(df[df.Month < _startYear].index)
df_curyear = df_curyear.drop(df_curyear[df_curyear.Month >= _endYear].index)

mycolumns = ['Entry Number', 'Entry Date', 'GL Description', 'Amount', 'Narration']
df_entries = pd.DataFrame(columns=mycolumns)
rows = [
    [1,'01/04/2019','Right of Use Asset - BS (FA)',_TotalLeaseLiab,'Initial Recognition (including all adjustments)'],
    ['','','Lease Liabilities - BS (Liab)',_TotalLeaseLiab*-1,''],
    ['','','','',''],
    ['','','','',''],
    [2,'01/04/2019','Reserve & Surplus - BS',df_prev['Depreciation'].sum() *-1,'Transitional Depreciation Charge'],
    ['','','Acc Dep on RoU Assets - BS',df_prev['Depreciation'].sum() ,''],
    ['','','','',''],
    ['','','','',''],
    [3,'31/03/2020','Depreciation Charge - P&L',df_curyear['Depreciation'].sum() *-1,'Current Year Depreciation Charge'],
    ['','','Acc Dep on RoU Assets - BS',df_curyear['Depreciation'].sum() ,''],
    ['','','','',''],
    ['','','','',''],
    [4,'31/03/2020','Lease Liabilities - BS (Liab)',df_curyear['Rent Pmt'].sum(),'Adjustment of Lease recorded in Rent GL directly.'],
    ['','','Lease Rent GL - P&L',df_curyear['Rent Pmt'].sum() *-1 ,''],
    ['','','','',''],
    ['','','','',''],
    [5,'01/04/2019','Reserve & Surplus - BS',df_prev['Interest'].sum(),'Transitional Lease Cost recognition.'],
    ['','','Lease Liabilities - BS (Liab)',df_prev['Interest'].sum() *-1 ,''],
    ['','','','',''],
    ['','','','',''],
    [6,'31/03/2020','Interest Expenses - P&L',df_curyear['Interest'].sum(),'Annual Recognition of Lease Costs.'],
    ['','','Lease Liabilities - BS (Liab)',df_curyear['Interest'].sum() *-1 ,''],
    ['','','','',''],
    ['','','','',''],
    [7,'01/04/2019','Lease Liabilities - BS (Liab)',df_prev['Rent Pmt'].sum(),'Transitional Adjustment of Payments Done.'],
    ['','','Reserve & Surplus - BS',df_prev['Rent Pmt'].sum() *-1 ,''],
    ['','','','',''],
    ['','','','',''],
    ]

# rows.append[]
for row in rows:
    df_entries.loc[len(df_entries)] = row

# df_entries.to_excel('test.xlsx')
Export2Excel([df,df_curyear,df_prev, df_entries],['Complete Schedule','Current Year','Prior Period','JV Entries'])

print("-------------------")
print("-------------------")
# print(df_entries)
# print("")
print("-C-O-M-P-L-E-T-E-D-")
print("-------------------")
# print(df_curyear)
print("-------------------")