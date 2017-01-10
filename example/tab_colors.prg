/*
 * Example of how to set Excel worksheet tab colors using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet1, worksheet2, worksheet3, worksheet4

    lxw_init() 

    workbook   = lxw_workbook_new("tab_colors.xlsx")

    /* Set up some worksheets. */
    worksheet1 = lxw_workbook_add_worksheet(workbook, NIL)
    worksheet2 = lxw_workbook_add_worksheet(workbook, NIL)
    worksheet3 = lxw_workbook_add_worksheet(workbook, NIL)
    worksheet4 = lxw_workbook_add_worksheet(workbook, NIL)


    /* Set the tab colors. */
    lxw_worksheet_set_tab_color(worksheet1, LXW_COLOR_RED)
    lxw_worksheet_set_tab_color(worksheet2, LXW_COLOR_GREEN)
    lxw_worksheet_set_tab_color(worksheet3, 0xFF9900) /* Orange. */

    /* worksheet4 will have the default color. */
    lxw_worksheet_write_string(worksheet4, 0, 0, "Hello", NIL)

    lxw_workbook_close(workbook)

    return 0

