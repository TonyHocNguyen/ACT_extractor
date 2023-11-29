import time
import win32com.client as win32
import sys

win32c = win32.constants

def pivot_table(wb, ws1, pt_ws, ws_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields):
    """
    wb: Workbook reference
    ws1: Data worksheet
    pt_ws: Pivot table worksheet
    ws_name: Pivot table worksheet name
    pt_name: Name given to the pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: Lists of values for filling the pivot tables
    """

    # Pivot table location
    pt_loc = len(pt_filters) + 2

    # Grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)

    # Create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C1', TableName=pt_name)

    # Select the pivot table worksheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Set the rows, columns, and filters of the pivot table
    pt_table = pt_ws.PivotTables(pt_name)
    for field_list, field_r in ((pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_table.PivotFields(value).Orientation = field_r
            pt_table.PivotFields(value).Position = i + 1

    # Set the Values of the pivot table
    for field in pt_fields:
        field_name, column_name, calculation_method, number_format = field
        pt_field = pt_table.PivotFields(field_name)
        pt_field.Orientation = win32c.xlDataField
        pt_field.Function = calculation_method
        pt_field.NumberFormat = number_format
        pt_field.Caption = column_name

    # Visibility True or False
    pt_table.ShowValuesRow = True
    pt_table.ColumnGrand = True
    pt_table.RowAxisLayout(1)   # RowAxisLayout(1) for tabular form

def run_excel(file_path):
    try:
        # Create Excel object
        # excel = win32.dynamic.Dispatch('Excel.Application')

        # Create Excel object
        excel = win32.gencache.EnsureDispatch('Excel.Application')

        # Excel can be visible or not
        excel.Visible = True  # False

        # Try-except for file/path
        wb = excel.Workbooks.Open(file_path)

        # Data worksheet name
        ws1 = wb.Sheets('Sheet1')

        # Setup and call pivot_table
        ws2_name = 'Commonality_Analysis'
        wb.Sheets.Add().Name = ws2_name
        ws2 = wb.Sheets(ws2_name)

        pt_name = 'Commonality Analysis'  # Must be a string
        pt_rows = ['Functional Block', 'Package Net List', 'UUT Signal Name']  # Must be a list
        pt_cols = ['Unit #']  # Must be a list
        pt_filters = ['Result Type']  # Must be a list

        # [0]: field name [1]: pivot table column name [3]: calculation method [4]: number format
        pt_fields = [['UUT Pin', 'Count of UUT Pin', win32c.xlCount, '0']]  # Must be a list of lists

        pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
        wb.Close(SaveChanges=1)
        excel.Quit()
        time.sleep(0.4)

    except Exception as e:
        raise e
        # print("Failed to create pivot table.")
        # print("Please close any Excel file or delete temp files and try again.")
        # sys.exit(1)
    
def test():
    test_path = 'D:\\LabData\\FACR\\VIRAL Result'
    run_excel(test_path)

if __name__ == '__main__':
    test()