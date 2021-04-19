/*
 * An example of a simple Excel chart using the libxlsxwriter library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

/* Create a worksheet with a chart. */
FUNCTION main()
LOCAL workbook, worksheet //, chart, chart2
    
    lxw_init() 
    
    workbook  = lxw_new_workbook("chart_opt.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Write some data for the chart. */
    write_worksheet_data(worksheet)

    make_chart( workbook, worksheet, 6, 1 )

    make_chart( workbook, worksheet, 30, 1 )

    make_chart( workbook, worksheet, 54, 1 )

    return lxw_workbook_close(workbook)

FUNCTION make_chart( workbook, worksheet, row, col )
LOCAL options:= { => }, chart
    /* Create a chart object. */
    chart = lxw_workbook_add_chart(workbook, LXW_CHART_COLUMN)

    /* Configure the chart. In simplest case we just add some value data
     * series. The NIL categories will default to 1 to 5 like in Excel.
     */
    lxw_chart_add_series(chart, NIL, "Sheet1!$A$1:$A$5")
    lxw_chart_add_series(chart, NIL, "Sheet1!$B$1:$B$5")
    lxw_chart_add_series(chart, NIL, "Sheet1!$C$1:$C$5")

    //lxw_chart_legend_delete_series(chart, { 0, 1, 2, -1 } )
    lxw_chart_legend_set_position(chart, LXW_CHART_LEGEND_NONE)

    /* Insert the chart into the worksheet. */
    
    options[ "x_offset" ]:= 4
    options[ "y_offset" ]:= 5
    options[ "x_scale"  ]:= 1.44
    options[ "y_scale"  ]:= 1.39
    options[ "object_position" ]:= 1

    lxw_worksheet_insert_chart_opt(worksheet, row, col, chart, options)

RETURN Nil

/* Write some data to the worksheet. */
FUNCTION write_worksheet_data(worksheet) 
LOCAL data, row, col
    /* Three columns of data. */
    data:= {;
        {1,   2,   3},;
        {2,   4,   6},;
        {3,   6,   9},;
        {4,   8,  12},;
        {5,  10,  15};
    }

    for row:= 0 TO 4
        for col:= 0 TO 2
            lxw_worksheet_write_number(worksheet, row, col, data[row+1,col+1], NIL)
        next
    next
return NIL

