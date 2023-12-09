import openpyxl

import pandas as pd
from openpyxl.styles import Font

# read_file = pd.read_csv(r'Daily_Dump(Updated).csv')
# read_file.to_excel(r'daily_dump.xlsx', index=None, header=True)

daily_dump = openpyxl.load_workbook("daily_dump.xlsx")
# print("done")
category_team = openpyxl.load_workbook(r"Category team.xlsx")
raw_dump = openpyxl.load_workbook(r'RAW Dump for daily trending - Copy.xlsx')
print("done!")

daily_dump_sheet = daily_dump.active
# print("done!!")
category_team_sheet = category_team.active

raw_sheet_of_raw_dump = raw_dump['RAW']
assigned_sheet_of_raw_dump = raw_dump['Assigned']
raw_pivot_sheet_of_raw_dump = raw_dump['RAW Pivot']
assigned_pivot_sheet_of_raw_dump = raw_dump['Assigned Pivot']

# print("done!!!")

daily_dump_sheet.insert_cols(8)  # 1 based index
daily_dump_sheet['H1'] = 'Team'
daily_dump_sheet['H1'].font = Font(bold=True)

heading_row = 1
# cnt = 1

for i in daily_dump_sheet.iter_rows():  # iter_rows() --> 0 based index
    if heading_row == 1:
        heading_row = 2
        continue

    sub_category = i[8].value  # 7 = Team, 8 = sub category, heading row = 1
    row_number = i[8].row

    for j in category_team_sheet.iter_rows():
        if sub_category is not None and j[0].value is not None and j[0].value.lower() == sub_category.lower():
            daily_dump_sheet.cell(row=row_number, column=8).value = j[2].value  # column=8 = Team

for cols in daily_dump_sheet.iter_rows():
    if cols[7].value is None:  # 7 = Team
        cols[7].value = "#N/A"


# Select those `#N/A` from `Team` col where sub_category has `short code` in it.

# cnt = 1

for i in daily_dump_sheet.iter_rows():  # tuple = iter_rows() --> 0 based index

    team_value = i[7].value  # 7 = Team, 8 = sub category, heading row = 1
    sub_category_value = i[8].value
    row_number = i[7].row

    if team_value == '#N/A' and sub_category_value[-1] == ')' and sub_category_value[-2].isnumeric():
        daily_dump_sheet.cell(row=row_number, column=8).value = 'VAS'  # column=8 = Team
#         cnt += 1
# print(cnt)
daily_dump.save("daily_dump.xlsx")

############################### RAW & Assigned Sheet ###############################

## RAW Sheet

from_date = "10/7/2023"
to_date = "10/8/2023"
cnt = 1

open_date_filtered_rows = []

def select_prv_days_from_open_date_col(cnt):

    for row in daily_dump_sheet.iter_rows(min_row=2, values_only=True):  # tuple = iter_rows() --> 0 based index

        team_value = row[7]  # 7 = Team, 8 = sub category, heading row = 1
        open_date_value = row[20]  # open date = 20

        if team_value != '#N/A' and from_date <= open_date_value <= to_date:
            open_date_filtered_rows.append(row)
    #         cnt += 1
    #         print(team_value, open_date_value)
    # print(cnt)


select_prv_days_from_open_date_col(1)


def paste_filtered_data_to_raw_sheet(count, c):

    for row in raw_sheet_of_raw_dump.iter_rows():
        # count += 1
        if row[0].value is None:
            for row_data in open_date_filtered_rows:
                # c += 1
                raw_sheet_of_raw_dump.append(row_data)

            break
    # print(count, c)
    # raw_dump.save(r'RAW Dump for daily trending - Copy2.xlsx')


paste_filtered_data_to_raw_sheet(0, 0)



## Assigned Sheet

assign_from_date = "7-Oct-23"
assign_to_date = "8-Oct-23"
assign_date_filtered_rows = []


def select_prv_assigned_date(cnt):
    for row in daily_dump_sheet.iter_rows(min_row=2, values_only=True):  # tuple = iter_rows() --> 0 based index

        team_value = row[7]  # 7 = Team, 8 = sub category, heading row = 1
        assign_date_value = row[21]  # assign date = 21

        if team_value != '#N/A' and assign_date_value is not None and assign_from_date <= assign_date_value <= assign_to_date:
            assign_date_filtered_rows.append(row)
            # cnt += 1
    #         print(team_value, assign_date_value)
    # print(cnt, len(assign_date_filtered_rows), end=' +++\n ')


select_prv_assigned_date(0)



def paste_assigned_date_to_assigned_sheet(count, c):

    for row_data in assign_date_filtered_rows:
        c += 1
        assigned_sheet_of_raw_dump.append(row_data)
    print(count, c)
    raw_dump.save(r'RAW Dump for daily trending - Copy2.xlsx')

paste_assigned_date_to_assigned_sheet(0, 0)




############# radio dump #######################

def select_prv_assigned_date(cnt):
    heading = True
    for row in daily_dump_sheet.iter_rows(min_row=1, values_only=True):  # tuple = iter_rows() --> 0 based index

        team_value = row[7]  # 7 = Team, 8 = sub category, heading row = 1
        assign_date_value = row[21]  # assign date = 21

        if heading:
            assign_date_filtered_rows.append(row)
            heading = False
            # print(team_value, assign_date_value)

        if team_value != '#N/A' and assign_date_value is not None and assign_from_date <= assign_date_value <= assign_to_date:
            assign_date_filtered_rows.append(row)
            # cnt += 1
            # print(team_value, assign_date_value)
    # print(cnt, len(assign_date_filtered_rows), end=' +++\n ')


select_prv_assigned_date(0)

def paste_to_radio_dump(c):

    radio_workbook = openpyxl.Workbook()
    radio_worksheet = radio_workbook.active
    # radio_worksheet.append(daily_dump_sheet[1])

    for row_data in assign_date_filtered_rows:
        # c += 1
        radio_worksheet.append(row_data)
    # print(c)

    radio_worksheet.insert_cols(19)  # 1 based index
    radio_worksheet.insert_cols(20)  # 1 based index
    radio_worksheet['S1'] = 'site id'
    radio_worksheet['T1'] = 'sales'

    radio_dump_file = "Radio Dump.xlsx"
    radio_workbook.save(radio_dump_file)

paste_to_radio_dump(0)


