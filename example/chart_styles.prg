/*
 * An example showing all 48 default chart styles available in Excel 2007
 * using the libxlsxwriter library. Note, these styles are not the same as the
 * styles available in Excel 2013.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

FUNCTION main()
LOCAL workbook, worksheet, chart, row_num, col_num, chart_num, style_num, chart_types, chart_names, chart_title
    
    lxw_init() 

    chart_types:= {LXW_CHART_COLUMN, LXW_CHART_AREA, LXW_CHART_LINE, LXW_CHART_PIE}
    chart_names:= {"Column", "Area", "Line", "Pie"}
    chart_title:= {}

    workbook  = lxw_new_workbook("chart_styles.xlsx")


    for chart_num = 0 to 3

        /* Add a worksheet for each chart type. */
        worksheet = lxw_workbook_add_worksheet(workbook, chart_names[chart_num+1])
        lxw_worksheet_set_zoom(worksheet, 30)


        /* Create 48 charts, each with a different style. */
        style_num = 1
        for row_num = 0 to 89 STEP 15

            for col_num = 0 to 63 STEP 8

                chart = lxw_workbook_add_chart(workbook, chart_types[chart_num+1])
                chart_title:= "Style "+HB_NTOS(style_num)

                lxw_chart_add_series(chart, NIL, "=Data!$A$1:$A$6")
                lxw_chart_title_set_name(chart, chart_title)
                lxw_chart_set_style(chart, style_num)

                lxw_worksheet_insert_chart(worksheet, row_num, col_num, chart)

                style_num++
            next
        next
    next

    /* Create a worksheet with data for the charts. */
    worksheet = lxw_workbook_add_worksheet(workbook, "Data")
    lxw_worksheet_write_number(worksheet, 0, 0, 10, NIL)
    lxw_worksheet_write_number(worksheet, 1, 0, 40, NIL)
    lxw_worksheet_write_number(worksheet, 2, 0, 50, NIL)
    lxw_worksheet_write_number(worksheet, 3, 0, 20, NIL)
    lxw_worksheet_write_number(worksheet, 4, 0, 10, NIL)
    lxw_worksheet_write_number(worksheet, 5, 0, 50, NIL)

    return lxw_workbook_close(workbook)

