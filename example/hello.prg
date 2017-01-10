/*
 * Example of writing some data to a simple Excel file using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet

    lxw_init() 

    workbook  = lxw_workbook_new("hello_world.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    lxw_worksheet_write_string(worksheet, 0, 0, "Hello", NIL)
    lxw_worksheet_write_number(worksheet, 1, 0, 123, NIL)

    lxw_workbook_close(workbook)

    return 0

