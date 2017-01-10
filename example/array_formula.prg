/*
 * Example of how to use the libxlsxwriter library to write simple
 * array formulas.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet

    lxw_init() 

    /* Create a new workbook and add a worksheet. */
    workbook  = lxw_workbook_new("array_formula.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Write some data for the formulas. */
    lxw_worksheet_write_number(worksheet, 0, 1, 500, NIL)
    lxw_worksheet_write_number(worksheet, 1, 1, 10, NIL)
    lxw_worksheet_write_number(worksheet, 4, 1, 1, NIL)
    lxw_worksheet_write_number(worksheet, 5, 1, 2, NIL)
    lxw_worksheet_write_number(worksheet, 6, 1, 3, NIL)

    lxw_worksheet_write_number(worksheet, 0, 2, 300, NIL)
    lxw_worksheet_write_number(worksheet, 1, 2, 15, NIL)
    lxw_worksheet_write_number(worksheet, 4, 2, 20234, NIL)
    lxw_worksheet_write_number(worksheet, 5, 2, 21003, NIL)
    lxw_worksheet_write_number(worksheet, 6, 2, 10000, NIL)

    /* Write an array formula that returns a single value. */
    lxw_worksheet_write_array_formula(worksheet, 0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", NIL)

    /* Similar to above but using the RANGE macro. */
    lxw_worksheet_write_array_formula(worksheet, LXW_RANGE("A2:A2"), "{=SUM(B1:C1*B2:C2)}", NIL)

    /* Write an array formula that returns a range of values. */
    lxw_worksheet_write_array_formula(worksheet, 4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", NIL)

    lxw_workbook_close(workbook)

    return 0

