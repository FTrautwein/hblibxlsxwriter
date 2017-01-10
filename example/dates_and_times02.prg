/*
 * Example of writing dates and times in Excel using an lxw_datetime struct
 * and date formatting.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local datetime, workbook, worksheet, format

    lxw_init() 

    /* A datetime to display. */
    datetime = HB_STOT("201302281200")

    /* Create a new workbook and add a worksheet. */
    workbook  = lxw_workbook_new("date_and_times02.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Add a format with date formatting. */
    format    = lxw_workbook_add_format(workbook)
    lxw_format_set_num_format(format, "mmm d yyyy hh:mm AM/PM")

    /* Widen the first column to make the text clearer. */
    lxw_worksheet_set_column(worksheet, 0, 0, 20, NIL)

    /* Write the datetime without formatting. */
    lxw_worksheet_write_datetime(worksheet, 0, 0, datetime, NIL  )  // 41333.5

    /* Write the datetime with formatting. */
    lxw_worksheet_write_datetime(worksheet, 1, 0, datetime, format)  // Feb 28 2013 12:00 PM

    return lxw_workbook_close(workbook)

