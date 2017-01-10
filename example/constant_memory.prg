/*
 * Example of using libxlsxwriter for writing large files in constant memory
 * mode.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local row, col, max_row, max_col, options, workbook, worksheet

    lxw_init() 

    max_row = 1000
    max_col = 50

    /* Set the worksheet options. */
    options = { "constant_memory" => 1, "tmpdir" => NIL }

    /* Create a new workbook with options. */
    workbook  = lxw_workbook_new_opt("constant_memory.xlsx", options)
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    for row = 0 TO max_row - 1
        for col = 0 TO max_col - 1
            lxw_worksheet_write_number(worksheet, row, col, 123.45, NIL)
        next
    next

    return lxw_workbook_close(workbook)

