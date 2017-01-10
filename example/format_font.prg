/*
 * Example of writing some data with font formatting to a simple Excel
 * file using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet, format1, format2, format3

    lxw_init() 

    /* Create a new workbook. */
    workbook  = lxw_workbook_new("format_font.xlsx")

    /* Add a worksheet. */
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Widen the first column to make the text clearer. */
    lxw_worksheet_set_column(worksheet, 0, 0, 20, NIL)

    /* Add some formats. */
    format1   = lxw_workbook_add_format(workbook)
    format2   = lxw_workbook_add_format(workbook)
    format3   = lxw_workbook_add_format(workbook)

    /* Set the bold property for format 1. */
    lxw_format_set_bold(format1)

    /* Set the italic property for format 2. */
    lxw_format_set_italic(format2)

    /* Set the bold and italic properties for format 3. */
    lxw_format_set_bold  (format3)
    lxw_format_set_italic(format3)

    /* Write some formatted strings. */
    lxw_worksheet_write_string(worksheet, 0, 0, "This is bold",    format1)
    lxw_worksheet_write_string(worksheet, 1, 0, "This is italic",  format2)
    lxw_worksheet_write_string(worksheet, 2, 0, "Bold and italic", format3)

    /* Close the workbook, save the file and free any memory. */
    lxw_workbook_close(workbook)

    return 0
