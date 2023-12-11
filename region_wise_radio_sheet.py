import openpyxl
from datetime import datetime

# Load 'regional.xlsx'
technology_region_wb = openpyxl.load_workbook('I:\Openpyxl_tutorial\Parts\Technology 3 - Copy.xlsx')
print("done--")
regional_wb = openpyxl.load_workbook('I:\Openpyxl_tutorial\Parts\Regional_Pivot4.xlsx')
print("done+++++++++")

region_wise_radio_voice_sheet = technology_region_wb['Region wise Radio-Voice']
region_wise_radio_data_sheet = technology_region_wb['Region wise Radio-Data']
voice_regional_pivot_sheet = regional_wb['voice_pivot']
data_regional_pivot_sheet = regional_wb['data_pivot']

target_dates = ['3-Dec']
target_year = 2023

# Convert target dates to 'YYYY-MM-DD' format
formatted_dates = []
for target_date in target_dates:
    try:
        date_obj = datetime.strptime(f"{target_date}-{target_year}", '%d-%b-%Y')
        formatted_date = date_obj.strftime('%Y-%m-%d')
        formatted_dates.append(formatted_date)
    except ValueError:
        print(f"Invalid date format: {target_date}. Please use the '2-Dec' format.")

column_indices = []
# Iterate through target dates and find the column indices
for target_date in formatted_dates:
    column_index = None
    for col_num in range(1, region_wise_radio_voice_sheet.max_column + 1):
        # date time --> just take the date
        if str(region_wise_radio_voice_sheet.cell(row=2, column=col_num).value).split(' ')[0] == target_date:
            column_index = col_num
            column_indices.append(column_index)
            # print(column_index, end='---\n')
            break

    # Check if the column was found
    if column_index is not None:
        print(f"The column index for '{target_date}' is: {column_index}")
    else:
        print(f"Column '{target_date}' not found in the sheet.")


######################### region-wise-radio-Voice
for i in region_wise_radio_voice_sheet.iter_rows():  # iter_rows() --> 0 based index

    sales = i[1].value  # heading row = 1, Sales = 1
    row_number = i[1].row

    for j in voice_regional_pivot_sheet.iter_rows():  # min_row=3
        if sales is not None and j[0].value is not None and j[0].value.lower() == sales.lower():
            for k in range(len(column_indices)):
                region_wise_radio_voice_sheet.cell(row=row_number, column=column_indices[k]).value = j[k+1].value


# sei ghor gula #N/A or None segulare 0 diye fillup kortesi
for row in region_wise_radio_voice_sheet.iter_rows():
    sales = row[1].value
    row_number = row[1].row

    if sales is None or sales == ' ' or sales == '':
        continue

    for k in range(len(column_indices)):  # 5, 6
        if region_wise_radio_voice_sheet.cell(row=row_number, column=column_indices[k]).value is None:
            region_wise_radio_voice_sheet.cell(row=row_number, column=column_indices[k]).value = 0


######################### region-wise-radio-Data
for i in region_wise_radio_data_sheet.iter_rows():  # iter_rows() --> 0 based index

    sales = i[1].value  # heading row = 1, Sales = 1
    row_number = i[1].row

    for j in data_regional_pivot_sheet.iter_rows():
        if sales is not None and j[0].value is not None and j[0].value.lower() == sales.lower():
            for k in range(len(column_indices)):
                region_wise_radio_data_sheet.cell(row=row_number, column=column_indices[k]).value = j[k+1].value


# sei ghor gula #N/A or None segulare 0 diye fillup kortesi
for row in region_wise_radio_data_sheet.iter_rows():
    sales = row[1].value
    row_number = row[1].row

    if sales is None or sales == ' ' or sales == '':
        continue

    for k in range(len(column_indices)):  # 5, 6
        if region_wise_radio_data_sheet.cell(row=row_number, column=column_indices[k]).value is None:
            region_wise_radio_data_sheet.cell(row=row_number, column=column_indices[k]).value = 0

column_indices.clear()

technology_region_wb.save('region_wise_voice_data_complaint_file.xlsx')