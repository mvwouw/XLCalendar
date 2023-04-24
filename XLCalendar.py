""" Tiny Calendar Generator


-add command line arguments for month range
-add custom row-heights and column-widths?
-localize day-names and "week"?
"""


import calendar
from math import ceil
from datetime import date
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


YEAR_START = 2023
MONTH_START = 1
YEAR_END = 2024
MONTH_END = 1
COLUMN_WIDTH = 3.5
ROW_HEIGHT = 12.7


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
    print(day_list)
    # Remove trailing zero's but keep preceeding zero's
    for day in day_list[7:]:
        if day[2] == 0:
            day_list.pop()
    print(day_list)

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

    print(day_list)
    print(len(day_list) / 30)

    # Create an openpyxl worksheet and set some properties
    wb = Workbook()
    ws = wb.active
    ws.title = f"Calendar {MONTH_START}-{YEAR_START} to {MONTH_END}-{YEAR_END}"
    for i in range(1, 10):
        ws.row_dimensions[i].height = ROW_HEIGHT

    # Write row headers
    ws.cell(row=1, column=1, value=YEAR_START)
    ws.cell(row=2, column=1, value="Ma")
    ws.cell(row=3, column=1, value="Di")
    ws.cell(row=4, column=1, value="Wo")
    ws.cell(row=5, column=1, value="Do")
    ws.cell(row=6, column=1, value="Vr")
    ws.cell(row=7, column=1, value="Za")
    ws.cell(row=8, column=1, value="Zo")
    ws.cell(row=9, column=1, value="Week")

    # Get some limits for filling and styling the sheet
    len_day_list = len(day_list)
    last_column = ceil(len_day_list / 7) + 3

    # Write all dates and weeknumbers from the list of days to the worksheet
    day_index = 0
    for col in range(3, last_column):
        for row in range(2, 9):
            if day_list[day_index][2] != 0:
                ws.cell(row=row, column=col, value=day_list[day_index][2])
                ws.cell(row=9, column=col, value=day_list[day_index][3])
            day_index += 1
            if day_index == len_day_list:
                break

    # Apply custom width to all columns
    for i in range(1, last_column):
        ws.column_dimensions[get_column_letter(i)].width = COLUMN_WIDTH






    wb.save(f"Calendar {MONTH_START}-{YEAR_START} to {MONTH_END}-{YEAR_END}.xlsx")

