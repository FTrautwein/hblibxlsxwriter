/*
 * An example of a simple Excel chart using the libxlsxwriter library. This
 * example is used in the "Working with Charts" section of the docs.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

/*
 * Create a worksheet with examples charts.
 */
FUNCTION main()
LOCAL workbook, worksheet, chart, series
    
    lxw_init() 

    workbook  = lxw_new_workbook("chart_line.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Write some data for the chart. */
    lxw_worksheet_write_number(worksheet, 0, 0, 10, NIL)
    lxw_worksheet_write_number(worksheet, 1, 0, 40, NIL)
    lxw_worksheet_write_number(worksheet, 2, 0, 50, NIL)
    lxw_worksheet_write_number(worksheet, 3, 0, 20, NIL)
    lxw_worksheet_write_number(worksheet, 4, 0, 10, NIL)
    lxw_worksheet_write_number(worksheet, 5, 0, 50, NIL)

    /* Create a chart object. */
    chart = lxw_workbook_add_chart(workbook, LXW_CHART_LINE)

    /* Configure the chart. */
    series = lxw_chart_add_series(chart, NIL, "Sheet1!$A$1:$A$6")

    //series; /* Do something with series in the real examples. */

     /* Insert the chart into the worksheet. */
    lxw_worksheet_insert_chart(worksheet, LXW_CELL("C1"), chart)

    return lxw_workbook_close(workbook)

