/*
 * Example of setting document properties such as Author, Title, etc., for an
 * Excel spreadsheet using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet, properties

    lxw_init() 

    workbook   = lxw_workbook_new("doc_properties.xlsx")
    worksheet  = lxw_workbook_add_worksheet(workbook, NIL)

    /* Create a properties structure and set some of the fields. */
    properties = {;
        "title"    => "This is an example spreadsheet",;
        "subject"  => "With document properties",;
        "author"   => "John McNamara",;
        "manager"  => "Dr. Heinz Doofenshmirtz",;
        "company"  => "of Wolves",;
        "category" => "Example spreadsheets",;
        "keywords" => "Sample, Example, Properties",;
        "comments" => "Created with libxlsxwriter",;
        "status"   => "Quo";
    }

    /* Set the properties in the workbook. */
    lxw_workbook_set_properties(workbook, properties)

    /* Add some text to the file. */
    lxw_worksheet_set_column(worksheet, 0, 0, 50, NIL)
    lxw_worksheet_write_string(worksheet, 0, 0,;
                           "Select 'Workbook Properties' to see properties." , NIL)

    lxw_workbook_close(workbook)

    return 0

