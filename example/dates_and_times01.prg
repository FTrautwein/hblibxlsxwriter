/*
 * Example of writing a dates and time in Excel using a number with date
 * formatting. This demonstrates that dates and times in Excel are just
 * formatted real numbers.
 *
 * An easier approach using a lxw_datetime struct is shown in example
 * dates_and_times02.c.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local number, workbook, worksheet, format

    lxw_init() 

    /* A number to display as a date. */
    number = 41333.5

    /* Create a new workbook and add a worksheet. */
    workbook  = lxw_workbook_new("date_and_times01.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Add a format with date formatting. */
    format    = lxw_workbook_add_format(workbook)
    lxw_format_set_num_format(format, "mmm d yyyy hh:mm AM/PM")

    /* Widen the first column to make the text clearer. */
    lxw_worksheet_set_column(worksheet, 0, 0, 20, NIL)

    /* Write the number without formatting. */
    lxw_worksheet_write_number(worksheet, 0, 0, number, NIL   )  // 41333.5

    /* Write the number with formatting. Note: the worksheet_write_datetime()
     * function is preferable for writing dates and times. This is for
     * demonstration purposes only.
     */
    lxw_worksheet_write_number(worksheet, 1, 0, number, format)   // Feb 28 2013 12:00 PM

    return lxw_workbook_close(workbook)
