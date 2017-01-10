/*
 * An example of merging cells using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet, merge_format

    lxw_init() 

    workbook     = lxw_workbook_new("merge_range.xlsx")
    worksheet    = lxw_workbook_add_worksheet(workbook, NIL)
    merge_format = lxw_workbook_add_format(workbook)

    /* Configure a format for the merged range. */
    lxw_format_set_align(merge_format, LXW_ALIGN_CENTER)
    lxw_format_set_align(merge_format, LXW_ALIGN_VERTICAL_CENTER)
    lxw_format_set_bold(merge_format)
    lxw_format_set_bg_color(merge_format, LXW_COLOR_YELLOW)
    lxw_format_set_border(merge_format, LXW_BORDER_THIN)

    /* Increase the cell size of the merged cells to highlight the formatting. */
    lxw_worksheet_set_column(worksheet, 1, 3, 12, NIL)
    lxw_worksheet_set_row(worksheet, 3, 30, NIL)
    lxw_worksheet_set_row(worksheet, 6, 30, NIL)
    lxw_worksheet_set_row(worksheet, 7, 30, NIL)

    /* Merge 3 cells. */
    lxw_worksheet_merge_range(worksheet, 3, 1, 3, 3, "Merged Range", merge_format)

    /* Merge 3 cells over two rows. */
    lxw_worksheet_merge_range(worksheet, 6, 1, 7, 3, "Merged Range", merge_format)

    lxw_workbook_close(workbook)

    return 0

