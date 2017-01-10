/*
 * Anatomy of a simple libxlsxwriter program.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet1, worksheet2, myformat1, myformat2, error

    lxw_init() 
    
    /* Create a new workbook. */
    workbook   = lxw_workbook_new("anatomy.xlsx")

    /* Add a worksheet with a user defined sheet name. */
    worksheet1 = lxw_workbook_add_worksheet(workbook, "Demo")

    /* Add a worksheet with Excel's default sheet name: Sheet2. */
    worksheet2 = lxw_workbook_add_worksheet(workbook, NIL)

    /* Add some cell formats. */
    myformat1    = lxw_workbook_add_format(workbook)
    myformat2    = lxw_workbook_add_format(workbook)

    /* Set the bold property for the first format. */
    lxw_format_set_bold(myformat1)

    /* Set a number format for the second format. */
    lxw_format_set_num_format(myformat2, "$#,##0.00")

    /* Widen the first column to make the text clearer. */
    lxw_worksheet_set_column(worksheet1, 0, 0, 20, NIL)

    /* Write some unformatted data. */
    lxw_worksheet_write_string(worksheet1, 0, 0, "Peach", NIL)
    lxw_worksheet_write_string(worksheet1, 1, 0, "Plum",  NIL)

    /* Write formatted data. */
    lxw_worksheet_write_string(worksheet1, 2, 0, "Pear",  myformat1)

    /* Formats can be reused. */
    lxw_worksheet_write_string(worksheet1, 3, 0, "Persimmon",  myformat1)


    /* Write some numbers. */
    lxw_worksheet_write_number(worksheet1, 5, 0, 123,       NIL)
    lxw_worksheet_write_number(worksheet1, 6, 0, 4567.555,  myformat2)


    /* Write to the second worksheet. */
    lxw_worksheet_write_string(worksheet2, 0, 0, "Some text", myformat1)


    /* Close the workbook, save the file and free any memory. */
    error = lxw_workbook_close(workbook)

    /* Check if there was any error creating the xlsx file. */
    if !EMPTY(error)
        sprintf("Error in workbook_close().\n"+;
               "Error %d = %s\n", error, HB_NTOS(error))
    endif

    return error
