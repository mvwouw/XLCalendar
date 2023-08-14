Create handy calendars to print from Excel

I used to make these by hand every year, print them out and stick them underneath my monitor at work. They give an easy
way to mark dates and see day types and weeknumbers at a glance. Much over-engineered for what it does, but it was a
fun exercise.

Executable generated with pyinstaller.

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
    -f <nl | fr>        Force day and month names to NL or FR. (Default: OS locale)  
    -mnl                Mark NL general holidays.  