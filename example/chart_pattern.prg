/*
 * An example of a simple Excel chart with patterns using the libxlsxwriter
 * library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

/*
 * Create a worksheet with examples charts.
 */
FUNCTION main()
LOCAL workbook, worksheet, chart, series1, series2, bold, pattern1, pattern2, line1, line2
    
    lxw_init() 

    workbook  = lxw_new_workbook("chart_pattern.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)


    /* Add a bold format to use to highlight the header cells. */
    bold = lxw_workbook_add_format(workbook)
    lxw_format_set_bold(bold)

    /* Write some data for the chart. */
    lxw_worksheet_write_string(worksheet, 0, 0, "Shingle", bold)
    lxw_worksheet_write_number(worksheet, 1, 0, 105,       NIL)
    lxw_worksheet_write_number(worksheet, 2, 0, 150,       NIL)
    lxw_worksheet_write_number(worksheet, 3, 0, 130,       NIL)
    lxw_worksheet_write_number(worksheet, 4, 0, 90,        NIL)
    lxw_worksheet_write_string(worksheet, 0, 1, "Brick",   bold)
    lxw_worksheet_write_number(worksheet, 1, 1, 50,        NIL)
    lxw_worksheet_write_number(worksheet, 2, 1, 120,       NIL)
    lxw_worksheet_write_number(worksheet, 3, 1, 100,       NIL)
    lxw_worksheet_write_number(worksheet, 4, 1, 110,       NIL)

    /* Create a chart object. */
    chart = lxw_workbook_add_chart(workbook, LXW_CHART_COLUMN)

    /* Configure the chart. */
    series1 = lxw_chart_add_series(chart, NIL, "Sheet1!$A$2:$A$5")
    series2 = lxw_chart_add_series(chart, NIL, "Sheet1!$B$2:$B$5")

    lxw_chart_series_set_name(series1, "=Sheet1!$A1$1")
    lxw_chart_series_set_name(series2, "=Sheet1!$B1$1")

    lxw_chart_title_set_name(chart,        "Cladding types")
    lxw_chart_axis_set_name(LXW_GET_X_AXYS(chart), "Region")
    lxw_chart_axis_set_name(LXW_GET_Y_AXYS(chart), "Number of houses")


    /* Configure an add the chart series patterns. */
    pattern1 = {"type" => LXW_CHART_PATTERN_SHINGLE,;
                "fg_color" => 0x804000,;
                "bg_color" => 0xC68C53}

    pattern2 = {"type" => LXW_CHART_PATTERN_HORIZONTAL_BRICK,;
                "fg_color" => 0xB30000,;
                "bg_color" => 0xFF6666}

    lxw_chart_series_set_pattern(series1, pattern1)
    lxw_chart_series_set_pattern(series2, pattern2)

    /* Configure and set the chart series borders. */
    line1 = {"color" => 0x804000, "none" => 0}
    line2 = {"color" => 0xB30000, "none" => 0}

    lxw_chart_series_set_line(series1, line1)
    lxw_chart_series_set_line(series2, line2)

    /* Insert the chart into the worksheet. */
    lxw_worksheet_insert_chart(worksheet, LXW_CELL("D2"), chart)

    return lxw_workbook_close(workbook)

