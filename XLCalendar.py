""" Create a handy calendar to print from Excel

Michel van Wouw: I used to make these by hand every year, print them out and stick them underneath my monitor at work.
                 They give me an easy way to mark dates and see day types and weeknumbers at a glance. Much
                 over-engineered for what it does but it was a fun exercise.

Usage:
    XLCalendar [options]

General Options:
    -h, --help          Display help.
    -v, --version       Display version.
    -s <M> <YYYY>       Start calendar with month <M> of year <YYYY>. (Default: month 1 of current year)
    -e <M> <YYYY>       End calendar with month <M> of year <YYYY>. (Default: month 1 of next year)
    -o <output_file>    Use <output_file> as filename. (Default: 'Calendar.xlsx')
    -wr <%>             Resize column widths to <%> percent.
    -hr <%>             Resize row heights to <%> percent.
    -fnl                Force day and month names to NL. (Default: OS locale)
"""


import calendar
import re
from math import ceil
from datetime import date
from sys import argv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font


# Default settings
opt = {
    "Y_S": int(date.today().strftime("%Y")),
    "M_S": 1,
    "Y_E": int(date.today().strftime("%Y")) + 1,
    "M_E": 1,
    "COLUMN_WIDTH": 3.5,
    "ROW_HEIGHT": 12.7,
    "FORCE_NL": False,
    "OUTPUT_FILE": f"Calendar.xlsx",
    "OUTPUT_FILE_SET": False,
    "VERSION": 1.0,
    "HELP_TEXT": """
Tiny Calendar help

Usage:
    XLCalendar [options]

General Options:
    -h, --help          Display help.
    -v, --version       Display version.
    -s <M> <YYYY>       Start calendar with month <M> of year <YYYY>. (Default: month 1 of current year)
    -e <M> <YYYY>       End calendar with month <M> of year <YYYY>. (Default: month 1 of next year)
    -o <output_file>    Use <output_file> as filename. (Default: 'Calendar.xlsx')
    -wr <%>             Resize column widths to <%> percent.
    -hr <%>             Resize row heights to <%> percent.
    -fnl                Force day and month names to NL. (Default: OS locale)
"""
}


def main() -> None:
    # Process command line arguments
    if len(argv) == 1:
        # No command line arguments -> use default settings
        print(f"Generating Tiny Calendar with default settings. Use 'XLCalendar -h' for options.\n")
        create_calendar_file()
    else:
        # One or more command line arguments
        cl_args = argv[1:]
        cl_args = cl_args[::-1]
        while cl_args:
            # Go through provided arguments and validate
            arg = cl_args.pop()

            # Option: help
            if arg == "-h" or arg == "--help":
                hae()

            # Option: version
            if arg == "-v" or arg == "--version":
                print(f"\nXLCalendar version {opt['VERSION']}")
                exit()

            # Option: -s <M> <YYYY>
            elif arg == "-s":
                try:
                    sm = int(cl_args.pop())
                    if 1 <= sm <= 12:
                        opt['M_S'] = sm
                    else:
                        print(f"\nError: Option '-s' positional argument <M> should be a number from 1 to 12.")
                        hae()
                    sy = int(cl_args.pop())
                    if sy > 0:
                        opt['Y_S'] = sy
                    else:
                        print(f"\nError: Option '-s' positional argument <YYYY> should be a number > 0.")
                        hae()
                    if not opt['OUTPUT_FILE_SET']:
                        opt['OUTPUT_FILE'] = f"Calendar {opt['M_S']}-{opt['Y_S']} to {opt['M_E']}-{opt['Y_E']}.xlsx"
                except IndexError:
                    print(f"\nError: Option '-s' requires positional arguments <M> and <YYYY>.")
                    hae()
                except ValueError:
                    print(f"\nError: Option '-s' positional arguments <M> (1-12) and <YYYY> (1-inf) should be numbers.")
                    hae()

            # Option: -e <M> <YYYY>
            elif arg == "-e":
                try:
                    em = int(cl_args.pop())
                    if 1 <= em <= 12:
                        opt['M_E'] = em
                    else:
                        print(f"\nError: Option '-e' positional argument <M> should be a number from 1 to 12.")
                        hae()
                    ey = int(cl_args.pop())
                    if ey > 0:
                        opt['Y_E'] = ey
                    else:
                        print(f"\nError: Option '-e' positional argument <YYYY> should be a number > 0.")
                        hae()
                    if not opt['OUTPUT_FILE_SET']:
                        opt['OUTPUT_FILE'] = f"Calendar {opt['M_S']}-{opt['Y_S']} to {opt['M_E']}-{opt['Y_E']}.xlsx"
                except IndexError:
                    print(f"\nError: Option '-e' requires positional arguments <M> and <YYYY>.")
                    hae()
                except ValueError:
                    print(f"\nError: Option '-e' positional arguments <M> (1-12) and <YYYY> (1-inf) should be numbers.")
                    hae()

            # Option: -o <output_file>
            elif arg == "-o":
                try:
                    file_name = cl_args.pop()
                    if not re.fullmatch(r"^[0-9a-zA-Z_\-][0-9a-zA-Z_\-. ]*$", file_name):
                        print(f"\nError: Provided filename is not a valid Windows filename.")
                        hae()
                    opt['OUTPUT_FILE'] = file_name if re.fullmatch(r".xlsx$", file_name) else file_name + ".xlsx"
                    opt['OUTPUT_FILE_SET'] = True
                except IndexError:
                    print(f"\nError: Option '-o' requires positional argument <output_file>.")
                    hae()

            # Option: -wr <%>
            elif arg == "-wr":
                try:
                    wr = cl_args.pop()
                    if int(wr) > 0:
                        opt['COLUMN_WIDTH'] = float(opt['COLUMN_WIDTH'] * (int(wr) / 100))
                    else:
                        print(f"\nError: Option '-wr' positional argument <%> should be a number > 0.")
                        hae()
                except IndexError:
                    print(f"\nError: Option '-wr' requires positional argument <%>.")
                    hae()
                except ValueError:
                    print(f"\nError: Option '-wr' positional argument <%> should be a number > 0.")
                    hae()

            # Option: -hr <%>
            elif arg == "-hr":
                try:
                    hr = cl_args.pop()
                    if int(hr) > 0:
                        opt['ROW_HEIGHT'] = float(opt['ROW_HEIGHT'] * (int(hr) / 100))
                    else:
                        print(f"\nError: Option '-hr' positional argument <%> should be a number > 0.")
                        hae()
                except IndexError:
                    print(f"\nError: Option '-hr' requires positional argument <%>.")
                    hae()
                except ValueError:
                    print(f"\nError: Option '-hr' positional argument <%> should be a number > 0.")
                    hae()

            # Option: -fnl
            elif arg == "-fnl":
                opt['FORCE_NL'] = True

            # Unsupported argument provided
            else:
                print(f"\nNo such option: {arg}")
                hae()

        # Check whether end date is at least the same month or later
        if opt['Y_E'] < opt['Y_S']:
            print(f"\nError: End date is earlyer than start date.")
            hae()
        elif opt['Y_E'] == opt['Y_S']:
            if opt['M_E'] < opt['M_S']:
                print(f"\nError: End date is earlyer than start date.")
                hae()
        # Check maximum span of full calendar
        elif opt['Y_E'] - opt['Y_S'] > 100:
            print(f"\nError: Calendar cannot span more than 100 years.")
            hae()

        # All options validated -> create the file
        print(f"\nCreating calendar running from {opt['M_S']}-{opt['Y_S']} to {opt['M_E']}-{opt['Y_E']}.")
        create_calendar_file()


def hae() -> None:
    # Display help text and exit
    print(opt['HELP_TEXT'])
    exit()


def get_week_number(year: int, month: int, day: int) -> int:
    # Returns the weeknumber for the given day
    if day == 0:
        return 0
    else:
        date_object = date(year, month, day)
        _, week_number, _ = date_object.isocalendar()
        return week_number


def create_calendar_file() -> None:
    # Get localized day and month names
    if opt['FORCE_NL']:
        weekday_dict = {1: "Ma", 2: "Di", 3: "Wo", 4: "Do", 5: "Vr", 6: "Za", 7: "Zo", 8: "Week"}
        month_dict = {1: "Januari", 2: "Februari", 3: "Maart", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli",
                      8: "Augustus", 9: "September", 10: "Oktober", 11: "November", 12: "December"}
    else:
        weekday_dict = {}
        i = 1
        for day in range(24, 31):
            d = date(2023, 4, day)
            weekday_dict[i] = d.strftime("%a")
            i += 1
        # weekday_dict[8] = "Week"
        month_dict = {}
        for month in range(1, 13):
            d = date(2023, month, 1)
            month_dict[month] = d.strftime("%B")

    # Create styling templates
    font_standard = Font(name="Arial", size=10, bold=False)
    font_bold = Font(name="Arial", size=10, bold=True)

    h_align_center = Alignment(horizontal="center", vertical="center")
    h_align_right = Alignment(horizontal="right", vertical="center")

    fill_white = PatternFill("solid", fgColor="FFFFFF")
    fill_lightgrey = PatternFill("solid", fgColor="F2F2F2")
    # fill_mediumgrey = PatternFill("solid", fgColor="D9D9D9")

    thin = Side(border_style="thin", color="000000")
    medium = Side(border_style="medium", color="000000")

    border_LrTb = Border(left=medium, right=thin, top=medium, bottom=thin)
    border_Lr = Border(left=medium, right=thin)
    border_LrtB = Border(left=medium, right=thin, top=thin, bottom=medium)
    border_lrTb = Border(left=thin, right=thin, top=medium, bottom=thin)
    border_lRTb = Border(left=thin, right=medium, top=medium, bottom=thin)
    border_lt = Border(left=thin, top=thin)
    border_l = Border(left=thin)
    border_R = Border(right=medium)
    border_lrtB = Border(left=thin, right=thin, top=thin, bottom=medium)
    border_lRtB = Border(left=thin, right=medium, top=thin, bottom=medium)

    # Start calendar creation
    cal_obj = calendar.Calendar()

    # Get a list of days for the first requested month
    day_list = [[opt['Y_S'], opt['M_S'], weekday, get_week_number(opt['Y_S'], opt['M_S'], weekday)]
                for weekday in cal_obj.itermonthdays(opt['Y_S'], opt['M_S'])]

    # Remove trailing days with day number zero
    for day in day_list[7:]:
        if day[2] == 0:
            day_list.pop()

    # Get the next month and year if more than one month is requested
    if not (opt['Y_S'] == opt['Y_E'] and opt['M_S'] == opt['M_E']):
        current_year = opt['Y_S']
        if opt['M_S'] == 12:
            current_year += 1
            current_month = 1
        else:
            current_month = opt['M_S'] + 1

        # Extend the list of days with all following requested months
        while current_year <= opt['Y_E']:
            while current_month <= 12:
                # Extend the list of days with the current month, zero-days removed
                day_list.extend(
                    [[current_year, current_month, weekday, get_week_number(current_year, current_month, weekday)]
                     for weekday in cal_obj.itermonthdays(current_year, current_month) if weekday != 0]
                )

                # Break out if the last requested month of the last requested year is added to the list of days
                if current_year == opt['Y_E'] and current_month == opt['M_E']:
                    break
                current_month += 1
            current_year += 1
            current_month = 1

    # Create the openpyxl worksheet
    wb = Workbook()
    ws = wb.active

    # Set sheet title and apply row heights
    ws.title = f"Calendar {opt['M_S']}-{opt['Y_S']} to {opt['M_E']}-{opt['Y_E']}"
    for i in range(1, 10):
        ws.row_dimensions[i].height = opt['ROW_HEIGHT']

    # Write row header values and styles
    c = ws.cell(row=1, column=1)
    c.value = opt['Y_S']
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
        c.border = border_Lr
    c = ws.cell(row=9, column=1)
    # c.value = weekday_dict[8]
    c.font = font_standard
    c.alignment = h_align_right
    c.fill = fill_white
    c.border = border_LrtB

    # Merge the row header cells
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
        ws.column_dimensions[get_column_letter(i)].width = opt['COLUMN_WIDTH']

    # Fill the day- and weeknumber grid background
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

    # Apply page setup and set printing options
    ws.print_area = f"A1:{get_column_letter(last_column)}9"
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True

    # Save file to disk
    wb.save(opt['OUTPUT_FILE'])
    print(f"\nFile saved as: '{opt['OUTPUT_FILE']}'")


if __name__ == "__main__":
    main()

