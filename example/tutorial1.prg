/*
 * A simple program to write some data to an Excel file using the
 * libxlsxwriter library.
 *
 * This program is shown, with explanations, in Tutorial 1 of the
 * libxlsxwriter documentation.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

#define item 1
#define cost 2

function main() 
   local expenses, workbook, worksheet, row, col

    /* Some data we want to write to the worksheet. */
    expenses:= {;
       {"Rent", 1000},;
       {"Gas",   100},;
       {"Food",  300},;
       {"Gym",    50};
    }
    
    lxw_init() 
    
    /* Create a workbook and add a worksheet. */
    workbook := lxw_workbook_new("tutorial01.xlsx")
    worksheet:= lxw_workbook_add_worksheet(workbook, NIL)

    /* Start from the first cell. Rows and columns are zero indexed. */
    row:= 0
    col:= 0

    /* Iterate over the data and write it out element by element. */
    for row:= 0 to 3
        lxw_worksheet_write_string(worksheet, row, col,     expenses[row+1,item], NIL)
        lxw_worksheet_write_number(worksheet, row, col + 1, expenses[row+1,cost], NIL)
    next

    /* Write a total using a formula. */
    lxw_worksheet_write_string (worksheet, row, col,     "Total",       NIL)
    lxw_worksheet_write_formula(worksheet, row, col + 1, "=SUM(B1:B4)", NIL)

    /* Save the workbook and free any allocated memory. */
    return lxw_workbook_close(workbook)


