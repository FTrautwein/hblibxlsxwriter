/*
 * An example of creating an Excel line chart using the libxlsxwriter library.
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

    workbook  = lxw_new_workbook("chart_line.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Add a bold format to use to highlight the header cells. */
    bold = lxw_workbook_add_format(workbook)
    lxw_format_set_bold(bold)

    /* Write some data for the chart. */
    write_worksheet_data(worksheet, bold)


    /*
     * Create a line chart.
     */
    chart = lxw_workbook_add_chart(workbook, LXW_CHART_LINE)

    /* Add the first series to the chart. */
    series = lxw_chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7")

    /* Set the name for the series instead of the default "Series 1". */
    lxw_chart_series_set_name(series, "=Sheet1!$B1$1")

    /* Add a second series but leave the categories and values undefined. They
     * can be defined later using the alternative syntax shown below.  */
    series = lxw_chart_add_series(chart, NIL, NIL)

    /* Configure the series using a syntax that is easier to define programmatically. */
    lxw_chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0) /* "=Sheet1!$A$2:$A$7" */
    lxw_chart_series_set_values(series,     "Sheet1", 1, 2, 6, 2) /* "=Sheet1!$C$2:$C$7" */
    lxw_chart_series_set_name_range(series, "Sheet1", 0, 2)       /* "=Sheet1!$C$1"      */

    /* Add a chart title and some axis labels. */
    lxw_chart_title_set_name(chart,        "Results of sample analysis")
    lxw_chart_axis_set_name(LXW_GET_X_AXYS(chart), "Test number")
    lxw_chart_axis_set_name(LXW_GET_Y_AXYS(chart), "Sample length (mm)")

    /* Set an Excel chart style. */
    lxw_chart_set_style(chart, 10)

    /* Insert the chart into the worksheet. */
    lxw_worksheet_insert_chart(worksheet, LXW_CELL("E2"), chart)


    return lxw_workbook_close(workbook)



/*
 * Write some data to the worksheet.
 */
FUNCTION write_worksheet_data( worksheet, bold) 
LOCAL row, col, data
    data:= {;
        {2, 10, 30},;
        {3, 40, 60},;
        {4, 50, 70},;
        {5, 20, 50},;
        {6, 10, 40},;
        {7, 50, 30};
    }

    lxw_worksheet_write_string(worksheet, LXW_CELL("A1"), "Number",  bold)
    lxw_worksheet_write_string(worksheet, LXW_CELL("B1"), "Batch 1", bold)
    lxw_worksheet_write_string(worksheet, LXW_CELL("C1"), "Batch 2", bold)

    for row = 0 to 5
        for col = 0 to 2
            lxw_worksheet_write_number(worksheet, row + 1, col, data[row+1,col+1] , NIL)
        next
    next      
    
    return NIL
