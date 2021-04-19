/*
 * An example of a simple Excel chart with user defined fonts using the
 * libxlsxwriter library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

/*
 * Create a worksheet with examples charts.
 */
FUNCTION main()
LOCAL workbook, worksheet, chart, font1, font2, font3, font4, font5, font6
    
    lxw_init() 

    workbook  = lxw_new_workbook("chart_fonts.xlsx")
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
    lxw_chart_add_series(chart, NIL, "Sheet1!$A$1:$A$6")

    /* Create some fonts to use in the chart.  */
    font1 = {"name" => "Calibri", "color" => LXW_COLOR_BLUE, "size" => 10 }
    font2 = {"name" => "Courier", "color" => 0x92D050, "size" => 10 }
    font3 = {"name" => "Arial",   "color" => 0x00B0F0, "size" => 10 }
    font4 = {"name" => "Century", "color" => LXW_COLOR_RED, "size" => 10, "rotation" => -30 }
    font5 = {"rotation" => -30 }
    font6 = {"name"      => "Calibri",;
             "bold"      => LXW_TRUE,;
             "italic"    => LXW_TRUE,;
             "underline" => LXW_TRUE,;
             "color"     => 0x7030A0,;
             "size"      => 10      ,;
             "rotation"  => -30      }

    /* Write the chart title with a font. */
    lxw_chart_title_set_name(chart, "Test Results")
    //lxw_chart_title_set_name_font(chart, font1)

    /* Write the Y axis with a font. */
    lxw_chart_axis_set_name(LXW_GET_Y_AXYS(chart), "Units")
    //lxw_chart_axis_set_name_font(LXW_GET_Y_AXYS(chart), font2)
    //lxw_chart_axis_set_num_font(LXW_GET_Y_AXYS(chart), font3)

    /* Write the X axis with a font. */
    lxw_chart_axis_set_name(LXW_GET_X_AXYS(chart), "Month")
    lxw_chart_axis_set_name_font(LXW_GET_X_AXYS(chart), font4)
    lxw_chart_axis_set_num_font(LXW_GET_X_AXYS(chart), font5)

    /* Display the chart legend at the bottom of the chart. */
    lxw_chart_legend_set_position(chart, LXW_CHART_LEGEND_BOTTOM)
    lxw_chart_legend_set_font(chart, font6 ) //LXW_GET_CHART_FONTS('Arial', 10, LXW_FALSE, LXW_FALSE, LXW_FALSE, 30, 0x7030A0 ) )

    /* Insert the chart into the worksheet. */
    lxw_worksheet_insert_chart(worksheet, LXW_CELL("C1"), chart)

    return lxw_workbook_close(workbook)

