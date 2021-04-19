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

#define item      1
#define cost      2
#define datetime  3

FUNCTION main()
LOCAL expenses, workbook, worksheet, row, col, bold, i, money, center, date_format, col_name

/* Some data we want to write to the worksheet. */
expenses:= {;
    { "Rent", 1000, STOD("20130113") },;
    { "Gas",   100, STOD("20130114") },;
    { "Food",  300, STOD("20130116") },;
    { "Gym",    50, STOD("20130120") };
}

lxw_init() 

/* Create a workbook and add a worksheet. */
workbook  = lxw_workbook_new("testedll.xlsx")
worksheet = lxw_workbook_add_worksheet(workbook, NIL)
row = 0
col = 0

lxw_workbook_set_properties( workbook,  {;
           "title"    => "This is an example spreadsheet" ,;
           "subject"  => "With document properties" ,;
           "author"   => "John McNamara" ,;
           "manager"  => "Dr. Heinz Doofenshmirtz" ,;
           "company"  => "of Wolves" ,;
           "category" => "Example spreadsheets" ,;
           "keywords" => "Sample, Example, Properties" ,;
           "comments" => "Created with libxlsxwriter" ,;
           "status"   => "Quo"  } )

/* Add a bold format to use to highlight cells. */
bold = lxw_workbook_add_format(workbook)
lxw_format_set_bold(bold)

/* Add a number format for cells with money. */
money = lxw_workbook_add_format(workbook)
lxw_format_set_num_format(money, "$#,##0")

/* Add an Excel date format. */
//date_format = workbook_add_format(workbook)
//format_set_num_format(date_format, "mmmm d yyyy")

/* Adjust the column width. */
lxw_worksheet_set_column(worksheet, 0, 0, 15, NIL)
//worksheet_set_column(worksheet, 1, 1, 15, NIL)

/* Write some data header. */
//worksheet_set_row(worksheet, row, 30, NIL)
lxw_worksheet_write_string(worksheet, row, col,     "Item", bold)
lxw_worksheet_write_string(worksheet, row, col + 1, "Date", bold)
lxw_worksheet_write_string(worksheet, row, col + 2, "Cost", bold)

/* Iterate over the data and write it out element by element. */
for i:= 1 TO 4
    /* Write from the first cell below the headers. */
    row:= i
    lxw_worksheet_write_string  (worksheet, row, col,     expenses[i,item],     NIL)
    lxw_worksheet_write_datetime(worksheet, row, col + 1, expenses[i,datetime], date_format)
    lxw_worksheet_write_number  (worksheet, row, col + 2, expenses[i,cost],     money)
next

/* Write a total using a formula. */
lxw_worksheet_write_string (worksheet, ++row, col,     "Total",       bold)
lxw_worksheet_write_formula(worksheet,   row, col + 2, "=SUM(C2:C5)", money)

center:= lxw_workbook_add_format(workbook)
lxw_format_set_align(center, LXW_ALIGN_CENTER)
lxw_format_set_align(center, LXW_ALIGN_VERTICAL_CENTER)
lxw_format_set_bold(center)
lxw_format_set_bg_color(center, LXW_COLOR_ORANGE)
lxw_format_set_font_name(center, "Arial")
lxw_format_set_border(center, LXW_BORDER_THIN)

lxw_worksheet_write_string(worksheet, LXW_CELL("A10"), "Foo", NIL)
lxw_worksheet_print_area(worksheet, LXW_RANGE("A1:F12") )

lxw_worksheet_merge_range(worksheet, ++row, col, row+1, col+3, "Worksheet Merge Range Test", center )

/* Save the workbook and free any allocated memory. */
lxw_workbook_close(workbook)

QOUT( lxw_name_to_row("B2"), lxw_name_to_col("B2") )
QOUT( lxw_name_to_col("C:C"), lxw_name_to_col_2("C:D") )
QOUT( lxw_name_to_row("A1:K42"), lxw_name_to_col("A1:K42"),;
      lxw_name_to_row_2("A1:K42"), lxw_name_to_col_2("A1:K42") )
col_name:= SPAC(10)
lxw_col_to_name(@col_name, 100, 0 )
QOUT( col_name )
RETURN NIL
