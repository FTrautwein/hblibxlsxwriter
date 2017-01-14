/*
 * An example of creating an Excel pie chart using the libxlsxwriter library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

/*
 * Create a worksheet with examples charts.
 */
FUNCTION main()
LOCAL workbook, worksheet, chart, series, bold
    
    lxw_init() 

    workbook  = lxw_new_workbook("chart_pie.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Add a bold format to use to highlight the header cells. */
    bold = lxw_workbook_add_format(workbook)
    lxw_format_set_bold(bold)

    /* Write some data for the chart. */
    write_worksheet_data(worksheet, bold)


    /*
     * Create a pie chart.
     */
    chart = lxw_workbook_add_chart(workbook, LXW_CHART_PIE)

    /* Add the first series to the chart. */
    series = lxw_chart_add_series(chart, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4")

    /* Set the name for the series instead of the default "Series 1". */
    lxw_chart_series_set_name(series, "Pie sales data")

    /* Add a chart title. */
    lxw_chart_title_set_name(chart, "Popular Pie Types")

    /* Set an Excel chart style. */
    lxw_chart_set_style(chart, 10)

    /* Insert the chart into the worksheet. */
    lxw_worksheet_insert_chart(worksheet, LXW_CELL("D2"), chart)


    return lxw_workbook_close(workbook)



/*
 * Write some data to the worksheet.
 */
FUNCTION write_worksheet_data(worksheet, bold)

    lxw_worksheet_write_string(worksheet, LXW_CELL("A1"), "Category", bold)
    lxw_worksheet_write_string(worksheet, LXW_CELL("A2"), "Apple",    NIL)
    lxw_worksheet_write_string(worksheet, LXW_CELL("A3"), "Cherry",   NIL)
    lxw_worksheet_write_string(worksheet, LXW_CELL("A4"), "Pecan",    NIL)

    lxw_worksheet_write_string(worksheet, LXW_CELL("B1"), "Values",   bold)
    lxw_worksheet_write_number(worksheet, LXW_CELL("B2"), 60,         NIL)
    lxw_worksheet_write_number(worksheet, LXW_CELL("B3"), 30,         NIL)
    lxw_worksheet_write_number(worksheet, LXW_CELL("B4"), 10,         NIL)
    
    return nil

