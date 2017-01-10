/*
 * Example of setting custom document properties for an Excel spreadsheet
 * using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet, datetime

    lxw_init() 

    workbook   = lxw_workbook_new("doc_custom_properties.xlsx")
    worksheet  = lxw_workbook_add_worksheet(workbook, NIL)
    datetime   = HB_STOT("20161212000000.0")

    /* Set some custom document properties in the workbook. */
    lxw_workbook_set_custom_property_string  (workbook, "Checked by",      "Eve")
    lxw_workbook_set_custom_property_datetime(workbook, "Date completed",   datetime)
    lxw_workbook_set_custom_property_number  (workbook, "Document number",  12345)
    lxw_workbook_set_custom_property_number  (workbook, "Reference number", 1.2345)
    lxw_workbook_set_custom_property_boolean (workbook, "Has Review",       1)
    lxw_workbook_set_custom_property_boolean (workbook, "Signed off",       0)


    /* Add some text to the file. */
    lxw_worksheet_set_column(worksheet, 0, 0, 50, NIL)
    lxw_worksheet_write_string(worksheet, 0, 0,;
                           "Select 'Workbook Properties' to see properties." , NIL)

    lxw_workbook_close(workbook)

    return 0

