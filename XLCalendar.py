""" Tiny Calendar Generator


-check number formats
-set print settings
-add command line arguments for month range
-add custom row-heights and column-widths?
-localize day-names and "week"?
-overrule default filename
"""


import calendar
from math import ceil
from datetime import date
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font


YEAR_START = 2019
MONTH_START = 1
YEAR_END = 2020
MONTH_END = 1
COLUMN_WIDTH = 3.5
ROW_HEIGHT = 12.7

weekday_dict = {}
i = 1
for day in range(24, 31):
    d = date(2023, 4, day)
    weekday_dict[i] = d.strftime("%a")
    i += 1
# weekday_dict[8] = "Week"

month_dict = {
    1: "Januari",
    2: "Februari",
    3: "Maart",
    4: "April",
    5: "Mei",
    6: "Juni",
    7: "Juli",
    8: "Augustus",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "December"
}

font_standard = Font(name="Arial", size=10, bold=False)
font_bold = Font(name="Arial", size=10, bold=True)

h_align_center = Alignment(horizontal="center")
h_align_right = Alignment(horizontal="right")

fill_white = PatternFill("solid", fgColor="FFFFFF")
fill_lightgrey = PatternFill("solid", fgColor="F2F2F2")

thin = Side(border_style="thin", color="000000")
medium = Side(border_style="medium", color="000000")

border_LrTb = Border(left=medium, right=thin, top=medium, bottom=thin)
border_r = Border(right=thin)
border_LrtB = Border(left=medium, right=thin, top=thin, bottom=medium)
border_lrTb = Border(left=thin, right=thin, top=medium, bottom=thin)
border_lRTb = Border(left=thin, right=medium, top=medium, bottom=thin)
border_lt = Border(left=thin, top=thin)
border_l = Border(left=thin)
border_R = Border(right=medium)
border_lrtB = Border(left=thin, right=thin, top=thin, bottom=medium)
border_lRtB = Border(left=thin, right=medium, top=thin, bottom=medium)


def get_week_number(year: int, month: int, day: int) -> int:
    # Returns the weeknumber for the given day
    if day == 0:
        return 0
    else:
        date_object = date(year, month, day)
        _, week_number, _ = date_object.isocalendar()
        return week_number


if __name__ == "__main__":
    MyCal = calendar.Calendar()
    # Get a list of days for the first requested month
    day_list = [[YEAR_START, MONTH_START, weekday, get_week_number(YEAR_START, MONTH_START, weekday)] for weekday in MyCal.itermonthdays(YEAR_START, MONTH_START)]

    # Remove trailing zero's but keep preceeding zero's
    for day in day_list[7:]:
        if day[2] == 0:
            day_list.pop()

    # Get the next month and year if more than one month is requested
    if not (YEAR_START == YEAR_END and MONTH_START == MONTH_END):
        current_year = YEAR_START
        if MONTH_START == 12:
            current_year += 1
            current_month = 1
        else:
            current_month = MONTH_START + 1

        # Extend the list of days with all following requested months
        while current_year <= YEAR_END:
            while current_month <= 12:
                # Extend the list of days with the current month, all zero's removed
                day_list.extend([[current_year, current_month, weekday, get_week_number(current_year, current_month, weekday)] for weekday in MyCal.itermonthdays(current_year, current_month) if weekday != 0])

                # Break out if the last requested month of the last requested year is added to the list of days
                if current_year == YEAR_END and current_month == MONTH_END:
                    break
                current_month += 1
            current_year += 1
            current_month = 1

    # Create an openpyxl worksheet and set some properties
    wb = Workbook()
    ws = wb.active
    ws.title = f"Calendar {MONTH_START}-{YEAR_START} to {MONTH_END}-{YEAR_END}"
    for i in range(1, 10):
        ws.row_dimensions[i].height = ROW_HEIGHT

    # Write row header values and styles
    c = ws.cell(row=1, column=1)
    c.value = YEAR_START
    c.font = font_bold
    c.alignment = h_align_center
    c.fill = fill_white
    c.border = border_LrTb
    for i in range(1, 8):
        c = ws.cell(row=i + 1, column=1)
        c.value = weekday_dict[i]
        c.font = font_standard
        c.alignment = h_align_right
        c.fill = fill_white
        c.border = border_r
    c = ws.cell(row=9, column=1)
    # c.value = weekday_dict[8]
    # c.font = font_standard
    # c.alignment = h_align_right
    c.fill = fill_white
    c.border = border_LrtB

    # Merge row header cells
    for i in range(1, 10):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=2)

    # Get some limits for further filling and styling the sheet
    len_day_list = len(day_list)
    last_column = ceil(len_day_list / 7) + 2

    # Write all month headers, dates and weeknumbers from the list of days to the worksheet
    day_index = 0
    for col in range(3, last_column + 1):
        if day_index == len_day_list:
            break
        # Write month headers
        c = ws.cell(row=1, column=col)
        c.value = month_dict[day_list[day_index][1]]
        c.font = font_bold
        c.alignment = h_align_center
        c.fill = fill_white

        # Write day- and weeknumber
        for row in range(2, 9):
            if day_list[day_index][2] != 0:
                c = ws.cell(row=row, column=col)
                c.value = day_list[day_index][2]
                c.font = font_standard
                c.alignment = h_align_center

                c = ws.cell(row=9, column=col)
                c.value = day_list[day_index][3]
                c.font = font_standard
                c.alignment = h_align_center

            day_index += 1
            if day_index == len_day_list:
                break

    # Merge cells of corresponding months and set border styles
    prev_cell_content = ws.cell(row=1, column=3).value
    merge_range_start = 3
    ws.cell(row=1, column=3).border = border_lrTb
    for col in range(4, last_column + 2):
        if ws.cell(row=1, column=col).value == prev_cell_content:
            merge_range_end = col
        else:
            if col == last_column + 1:
                ws.cell(row=1, column=merge_range_start).border = border_lRTb
                ws.merge_cells(start_row=1, start_column=merge_range_start, end_row=1, end_column=merge_range_end)
            else:
                ws.cell(row=1, column=col).border = border_lrTb
                ws.merge_cells(start_row=1, start_column=merge_range_start, end_row=1, end_column=merge_range_end)
                prev_cell_content = ws.cell(row=1, column=col).value
                merge_range_start = col

    # Apply custom width to all columns
    for i in range(1, last_column + 1):
        ws.column_dimensions[get_column_letter(i)].width = COLUMN_WIDTH

    # Background fill day- and weeknumber grid
    for col in range(3, last_column + 1):
        for row in range(2, 7):
            ws.cell(row=row, column=col).fill = fill_white
        ws.cell(row=7, column=col).fill = fill_lightgrey
        ws.cell(row=8, column=col).fill = fill_lightgrey
        ws.cell(row=9, column=col).fill = fill_white

    # Set borders for daynumber grid
    for col in range(3, last_column):
        for row in range(2, 9):
            c = ws.cell(row=row, column=col)
            if c.value:
                if c.value == 1:
                    c.border = border_lt
                elif 2 <= ws.cell(row=row, column=col).value <= 7:
                    c.border = border_l
    for row in range(2, 9):
        ws.cell(row=row, column=last_column).border = border_R

    # Set borders for weeknumbers
    for col in range(3, last_column):
        ws.cell(row=9, column=col).border = border_lrtB
    ws.cell(row=9, column=last_column).border = border_lRtB

    # Write changes to file
    wb.save(f"Calendar {MONTH_START}-{YEAR_START} to {MONTH_END}-{YEAR_END}.xlsx")
    print(f"File saved as: 'Calendar {MONTH_START}-{YEAR_START} to {MONTH_END}-{YEAR_END}.xlsx'")

