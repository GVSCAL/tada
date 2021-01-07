import win32com.client as win32
from pathlib import Path
import sys
import pandas as pd  # only used for synthetic data
import numpy as np  # only used for synthetic data
import random  # only used for synthetic data
from datetime import datetime  # only used for synthetic data


import pythoncom
win32c = win32.constants


def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list, pt_filters: list, pt_fields: list):
    """
    wb = workbook1 reference
    ws1 = worksheet1
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    """

    # pivot table location
    pt_loc = len(pt_filters) + 2
    
    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    
    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C1', TableName=pt_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in ((pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True


def run_excel(filename, sheet_name: str):
    pythoncom.CoInitialize()
    # create excel object
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    pythoncom.CoInitialize()

    # excel can be visible or not
    excel.Visible = False
    
    # try except for file / path
    # try:
    #     wb = excel.Workbooks.Open(filename)
    # except com_error as e:
    #     if e.excepinfo[5] == -2146827284:
    #         print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
    #     else:
    #         raise e
    #     sys.exit(1)

    wb = excel.Workbooks.Open(filename)

    # set worksheet
    ws1 = wb.Sheets('FEA')
    
    # Setup and call pivot_table
    ws2_name = 'FEA_pivot_table'
    wb.Sheets.Add(After=ws1,Count=1,Type=-4167).Name = ws2_name
    ws2 = wb.Sheets(ws2_name)

    pt_name = 'THC_pivot_table'  # must be a string
    pt_rows = ['loadcase_short_name']  # must be a list
    pt_cols = ['specs','integrity']  # must be a list
    pt_filters = ['seatversion','OEM','design_loop']  # must be a list
    # [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
    pt_fields = [['specs', 'specs: Count', win32c.xlCount, '#,#0.0'],  # must be a list of lists
                 ['integrity', 'integrity: Count', win32c.xlCount, '#,#0.0']]
    
    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
    ws2.Shapes.AddChart2(201,Left=250,Top=0,Width=600,Height=400,NewLayout=True)

    wb.Close(True)
    excel.Application.Quit()


def main():
    # sheet name for data
    sheet_name = 'FEA'  # update with sheet name from your file
    # file path
    f_path = Path(r'C:\Users\chenwang\Desktop\THC_output_file\THC_summary_regular_excel_07-10-2020_085838.xlsx')  # file located somewhere else
    
    # function calls
    # create_test_excel_file(f_path, f_name, sheet_name)  # remove when running your own file
    run_excel(f_path, sheet_name)



if __name__ == "__main__":
    main()