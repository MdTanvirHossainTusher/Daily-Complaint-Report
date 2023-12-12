import win32com.client
import openpyxl


def pivot_table_creation(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name):

    pt_cache = workbook.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt = pt_cache.CreatePivotTable(ws_report.Range(output_starting_cell), pivot_table_name)
    pt.TableStyle2 = "PivotStyleMedium9"
    insert_pt_field_set1(pt)

    pivot_table = ws_report.PivotTables(pivot_table_name)
    pivot_field_product = pivot_table.PivotFields("Team")
    return pivot_field_product


def filter_single_item(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, filtered_team_name):

    pivot_field_product = pivot_table_creation(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name)
    pivot_field_product.ClearAllFilters()
    pivot_field_product.CurrentPage = filtered_team_name


def filter_multiple_items(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, items_to_exclude):

    pivot_field_product = pivot_table_creation(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name)
    pivot_field_product.ClearAllFilters()
    pivot_field_product.EnableMultiplePageItems = True

    for item_name in items_to_exclude:
        pivot_field_product.PivotItems(item_name).Visible = False


def insert_pt_field_set1(pt):

    field_filters = {}
    field_filters['team'] = pt.PivotFields("Team")

    field_columns = {}
    field_columns['open_date'] = pt.PivotFields("OPEN_DATE")

    field_values = {}
    field_values['open_date_count'] = pt.PivotFields("OPEN_DATE")

    field_filters['team'].Orientation = 3  # hidden = 0, row = 1, column = 2, page = 3, data = 4,

    field_columns['open_date'].Orientation = 2

    field_values['open_date_count'].Orientation = 4
    field_values['open_date_count'].Function = -4112  # count = -4112
    field_values['open_date_count'].NumberFormat = "#,##0"


def filter_item():
    open_workbook = openpyxl.load_workbook(r'Raw Dump.xlsx')
    raw_pivot_sheet_name = "RAW Pivot"
    open_workbook.create_sheet(title=raw_pivot_sheet_name)
    open_workbook.save(r'Raw Dump.xlsx')

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # Optional: Set to True if you want to see Excel while the code is running

    workbook = excel.Workbooks.Open(r'I:\Openpyxl_tutorial\Parts\Raw Dump.xlsx')

    ws_data = workbook.Worksheets("RAW")
    ws_report = workbook.Worksheets("RAW Pivot")

    # Radio
    output_starting_cell = "A3"
    pivot_table_name = "PivotTable1"
    filtered_team_name = "Radio"
    filter_single_item(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, filtered_team_name)

    # BSM
    output_starting_cell = "A10"
    pivot_table_name = "PivotTable2"
    items_to_exclude = ["Radio", "Core", "toffee"]
    filter_multiple_items(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, items_to_exclude)

    # Core
    output_starting_cell = "A17"
    pivot_table_name = "PivotTable3"
    filtered_team_name = "Core"
    filter_single_item(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, filtered_team_name)

    workbook.SaveAs("I:\Openpyxl_tutorial\Parts\Raw Dump.xlsx")  # Optional: Save the changes
    workbook.Close()

    excel.Quit()


if __name__ == "__main__":
    filter_item()
