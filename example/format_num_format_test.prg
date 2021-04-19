/*
 * Example of writing some data with numeric formatting to a simple Excel file
 * using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet, format01, format02, format03, format04, format05, format06, format07, format08, format09, format10, format11

    lxw_init() 

    /* Create a new workbook and add a worksheet. */
    workbook  = lxw_workbook_new("format_num_format_test.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Widen the first column to make the text clearer. */
    lxw_worksheet_set_column(worksheet, 0, 0, 30, NIL)

    /* Add some formats. */
    format01   = lxw_workbook_add_format(workbook)
    format02   = lxw_workbook_add_format(workbook)
    format03   = lxw_workbook_add_format(workbook)
    format04   = lxw_workbook_add_format(workbook)
    format05   = lxw_workbook_add_format(workbook)
    format06   = lxw_workbook_add_format(workbook)
    format07   = lxw_workbook_add_format(workbook)
    format08   = lxw_workbook_add_format(workbook)
    format09   = lxw_workbook_add_format(workbook)
    format10   = lxw_workbook_add_format(workbook)
    format11   = lxw_workbook_add_format(workbook)

    /* Set some example number formats. */
    lxw_format_set_num_format(format01, "#,##0.00")
    lxw_format_set_num_format(format02, "#,##0")
    lxw_format_set_num_format(format03, "R$* #,##0.00")
    lxw_format_set_num_format(format04, "0.00")
    lxw_format_set_num_format(format05, "R$* #,##0.00")
    lxw_format_set_num_format(format06, "R$* #,##0.00")
    lxw_format_set_num_format(format07, "R$* #,##0.00")
    lxw_format_set_num_format(format08, "#,##0.00")
    lxw_format_set_num_format(format09, "#,##0.00")

    /* Write data using the formats. */
    //lxw_worksheet_write_number(worksheet, 0, 0, 3.1415926, NIL)      // 3.1415926
    lxw_worksheet_write_number(worksheet, 1, 0, 3.1415926, format01)  // 3.142
    lxw_worksheet_write_number(worksheet, 2, 0, 1234.56,   format02)  // 1,235
    lxw_worksheet_write_number(worksheet, 3, 0, 1234.56,   format03)  // 1,234.56
    lxw_worksheet_write_number(worksheet, 4, 0, 49.99,     format04)  // 49.99
    lxw_worksheet_write_number(worksheet, 5, 0, 36892.521, format05)  // 01/01/01
    lxw_worksheet_write_number(worksheet, 6, 0, 36892.521, format06)  // Jan 1 2001
    lxw_worksheet_write_number(worksheet, 7, 0, 36892.521, format07)  // 1 January 2001
    lxw_worksheet_write_number(worksheet, 8, 0, 36892.521, format08)  // 01/01/2001 12:30 AM
    lxw_worksheet_write_number(worksheet, 9, 0, 1.87,      format09)  // 1 dollar and .87 cents

    /* Show limited conditional number formats. */
    lxw_format_set_num_format(format10, "[Green]General;[Red]-General;General")
    lxw_worksheet_write_number(worksheet, 10, 0, 123, format10)  // > 0 Green
    lxw_worksheet_write_number(worksheet, 11, 0, -45, format10)  // < 0 Red
    lxw_worksheet_write_number(worksheet, 12, 0,   0, format10)  // = 0 Default color

    /* Format a Zip code. */
    lxw_format_set_num_format(format11, "00000")
    lxw_worksheet_write_number(worksheet, 13, 0, 1209, format11)
    
    /* Close the workbook, save the file and free any memory. */
    return lxw_workbook_close(workbook)

