import win32com.client
import openpyxl
import os

raw_dump_file_name = r'Raw Dump.xlsx'


def pivot_table_creation(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name):

    pt_cache = workbook.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt = pt_cache.CreatePivotTable(ws_report.Range(output_starting_cell), pivot_table_name)
    pt.TableStyle2 = "PivotStyleMedium9"
    insert_pt_field_set1(pt)

    pivot_table = ws_report.PivotTables(pivot_table_name)
    pivot_field_product = pivot_table.PivotFields("Team")
    return pivot_field_product


def insert_pt_field_set2(pt):

    field_rows = {}
    field_rows['sub_category'] = pt.PivotFields("SUB_CATEGORY")

    field_columns = {}
    field_columns['assign_date'] = pt.PivotFields("ASSIGNED_DATE")

    field_values = {}
    field_values['assign_date_count'] = pt.PivotFields("ASSIGNED_DATE")

    field_rows['sub_category'].Orientation = 1  # hidden = 0, row = 1, column = 2, page = 3, data = 4,

    field_columns['assign_date'].Orientation = 2

    field_values['assign_date_count'].Orientation = 4
    field_values['assign_date_count'].Function = -4112  # count = -4112
    field_values['assign_date_count'].NumberFormat = "#,##0"


def pivot_table_creation_all(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name):

    pt_cache = workbook.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt = pt_cache.CreatePivotTable(ws_report.Range(output_starting_cell), pivot_table_name)
    pt.ColumnGrand = True
    pt.RowGrand = False
    pt.TableStyle2 = "PivotStyleMedium9"
    insert_pt_field_set2(pt)

    ws_report.PivotTables(pivot_table_name)


def filter_single_item(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, filtered_team_name):

    pivot_field_product = pivot_table_creation(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name)
    pivot_field_product.ClearAllFilters()
    pivot_field_product.CurrentPage = filtered_team_name


def filter_multiple_items(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, items_to_exclude):

    pivot_field_product = pivot_table_creation(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name)
    pivot_field_product.ClearAllFilters()
    pivot_field_product.EnableMultiplePageItems = True

    for item_name in items_to_exclude:
        try:
            pivot_field_product.PivotItems(item_name).Visible = False
        except Exception as e:
            print(f"Item '{item_name}' not found in the pivot table. Skipping...")


def filter_all(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name):
    pivot_table_creation_all(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name)


def insert_pt_field_set1(pt):

    field_filters = {}
    field_filters['team'] = pt.PivotFields("Team")

    field_columns = {}
    field_columns['assign_date'] = pt.PivotFields("ASSIGNED_DATE")

    field_values = {}
    field_values['sub_category_count'] = pt.PivotFields("SUB_CATEGORY")

    field_filters['team'].Orientation = 3  # hidden = 0, row = 1, column = 2, page = 3, data = 4,

    field_columns['assign_date'].Orientation = 2

    field_values['sub_category_count'].Orientation = 4
    field_values['sub_category_count'].Function = -4112  # count = -4112
    field_values['sub_category_count'].NumberFormat = "#,##0"


def filter_item(path_directory):

    open_workbook = openpyxl.load_workbook(raw_dump_file_name)
    assigned_pivot_sheet_name = "Assigned Pivot"
    open_workbook.create_sheet(title=assigned_pivot_sheet_name)
    open_workbook.save(raw_dump_file_name)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # Optional: Set to True if you want to see Excel while the code is running

    file_to_save = os.path.join(path_directory, raw_dump_file_name)
    workbook = excel.Workbooks.Open(file_to_save)
    ws_data = workbook.Worksheets("Assigned")
    ws_report = workbook.Worksheets("Assigned Pivot")

    # Radio
    output_starting_cell = "A3"
    pivot_table_name = "PivotTable1"
    filtered_team_name = "Radio"
    filter_single_item(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, filtered_team_name)

    # BSM
    output_starting_cell = "A10"
    pivot_table_name = "PivotTable2"
    items_to_exclude = ["Radio", "DSS-VAS", "Core", "toffee"]
    filter_multiple_items(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, items_to_exclude)

    # Core
    output_starting_cell = "A17"
    pivot_table_name = "PivotTable3"
    filtered_team_name = "Core"
    filter_single_item(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name, filtered_team_name)

    # All
    output_starting_cell = "A24"
    pivot_table_name = "PivotTable4"
    filter_all(workbook, ws_data, ws_report, output_starting_cell, pivot_table_name)

    workbook.SaveAs(file_to_save)
    workbook.Close()

    excel.Quit()
