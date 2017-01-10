/*
 * A simple Unicode UTF-8 example using libxlsxwriter.
 *
 * Note: The source file must be UTF-8 encoded.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet

    lxw_init() 

    workbook  = lxw_workbook_new("utf8.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    lxw_worksheet_write_string(worksheet, 2, 1, "Это фраза на русском!", NIL)

    return lxw_workbook_close(workbook)

