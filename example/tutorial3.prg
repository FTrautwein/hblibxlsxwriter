/*
 * A simple program to write some data to an Excel file using the
 * libxlsxwriter library.
 *
 * This program is shown, with explanations, in Tutorial 3 of the
 * libxlsxwriter documentation.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

#define item 1
#define cost 2
#define datetime 3

function main()
    local expenses, workbook, worksheet, row, col, bold, money, date_format, i

    /* Some data we want to write to the worksheet. */
    expenses:= {;
       {"Rent", 1000, STOD("20130113") },;
       {"Gas",   100, STOD("20130114") },;
       {"Food",  300, STOD("20130116") },;
       {"Gym",    50, STOD("20130120") };
    }

    lxw_init() 

    /* Create a workbook and add a worksheet. */
    workbook  = lxw_workbook_new("tutorial03.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)
    row = 0
    col = 0

    /* Add a bold format to use to highlight cells. */
    bold = lxw_workbook_add_format(workbook)
    lxw_format_set_bold(bold)

    /* Add a number format for cells with money. */
    money = lxw_workbook_add_format(workbook)
    lxw_format_set_num_format(money, "$#,##0")

    /* Add an Excel date format. */
    date_format = lxw_workbook_add_format(workbook)
    lxw_format_set_num_format(date_format, "mmmm d yyyy");

    /* Adjust the column width. */
    lxw_worksheet_set_column(worksheet, 0, 0, 15, NIL)

    /* Write some data header. */
    lxw_worksheet_write_string(worksheet, row, col,     "Item", bold)
    lxw_worksheet_write_string(worksheet, row, col + 1, "Cost", bold)

    /* Iterate over the data and write it out element by element. */
    for i = 0 TO 3
        /* Write from the first cell below the headers. */
        row = i + 1
        lxw_worksheet_write_string  (worksheet, row, col,     expenses[i+1,item],     NIL)
        lxw_worksheet_write_datetime(worksheet, row, col + 1, expenses[i+1,datetime], date_format)
        lxw_worksheet_write_number  (worksheet, row, col + 2, expenses[i+1,cost],     money)
    next

    /* Write a total using a formula. */
    lxw_worksheet_write_string (worksheet, row + 1, col,     "Total",       bold)
    lxw_worksheet_write_formula(worksheet, row + 1, col + 2, "=SUM(C2:C5)", money)

    /* Save the workbook and free any allocated memory. */
    return lxw_workbook_close(workbook)

