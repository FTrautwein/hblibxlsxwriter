/*
 * Example of cell locking and formula hiding in an Excel worksheet using
 * libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet, unlocked, hidden

    lxw_init() 

    workbook  = lxw_workbook_new("protection.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    unlocked = lxw_workbook_add_format(workbook)
    lxw_format_set_unlocked(unlocked)

    hidden = lxw_workbook_add_format(workbook)
    lxw_format_set_hidden(hidden)

    /* Widen the first column to make the text clearer. */
    lxw_worksheet_set_column(worksheet, 0, 0, 40, NIL)

    /* Turn worksheet protection on without a password. */
    lxw_worksheet_protect(worksheet, NIL, NIL)


    /* Write a locked, unlocked and hidden cell. */
    lxw_worksheet_write_string(worksheet, 0, 0, "B1 is locked. It cannot be edited.",       NIL)
    lxw_worksheet_write_string(worksheet, 1, 0, "B2 is unlocked. It can be edited.",        NIL)
    lxw_worksheet_write_string(worksheet, 2, 0, "B3 is hidden. The formula isn't visible.", NIL)

    lxw_worksheet_write_formula(worksheet, 0, 1, "=1+2", NIL)     /* Locked by default. */
    lxw_worksheet_write_formula(worksheet, 1, 1, "=1+2", unlocked)
    lxw_worksheet_write_formula(worksheet, 2, 1, "=1+2", hidden)

    lxw_workbook_close(workbook)

    return 0

