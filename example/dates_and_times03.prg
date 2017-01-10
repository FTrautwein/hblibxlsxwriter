/*
 * Example of writing dates and times in Excel using different date formats.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local datetime, workbook, worksheet, row, col, i, date_formats, bold, format

    lxw_init() 

    /* A datetime to display. */
    datetime = HB_STOT("20130123123005123")
    row = 0
    col = 0

    /* Examples date and time formats. In the output file compare how changing
     * the format strings changes the appearance of the date.
     */
    date_formats = {;
        "dd/mm/yy",;
        "mm/dd/yy",;
        "dd m yy",;
        "d mm yy",;
        "d mmm yy",;
        "d mmmm yy",;
        "d mmmm yyy",;
        "d mmmm yyyy",;
        "dd/mm/yy hh:mm",;
        "dd/mm/yy hh:mm:ss",;
        "dd/mm/yy hh:mm:ss.000",;
        "hh:mm",;
        "hh:mm:ss",;
        "hh:mm:ss.000";
    }

    /* Create a new workbook and add a worksheet. */
    workbook  = lxw_workbook_new("date_and_times03.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Add a bold format. */
    bold      = lxw_workbook_add_format(workbook)
    lxw_format_set_bold(bold)

    /* Write the column headers. */
    lxw_worksheet_write_string(worksheet, row, col,     "Formatted date", bold)
    lxw_worksheet_write_string(worksheet, row, col + 1, "Format",         bold)

    /* Widen the first column to make the text clearer. */
    lxw_worksheet_set_column(worksheet, 0, 1, 20, NIL)

    /* Write the same date and time using each of the above formats. */
    for i = 0 TO 13
        row++

        /* Create a format for the date or time.*/
        format  = lxw_workbook_add_format(workbook)
        lxw_format_set_num_format(format, date_formats[i+1])
        lxw_format_set_align(format, LXW_ALIGN_LEFT)

        /* Write the datetime with each format. */
        lxw_worksheet_write_datetime(worksheet, row, col, datetime, format)

        /* Also write the format string for comparison. */
        lxw_worksheet_write_string(worksheet, row, col + 1, date_formats[i+1], NIL)
    next

    return lxw_workbook_close(workbook)

