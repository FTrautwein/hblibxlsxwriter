/*
 * This program shows several examples of how to set up headers and
 * footers with libxlsxwriter.
 *
 * The control characters used in the header/footer strings are:
 *
 *     Control             Category            Description
 *     =======             ========            ===========
 *     &L                  Justification       Left
 *     &C                                      Center
 *     &R                                      Right
 *     &P                  Information         Page number
 *     &N                                      Total number of pages
 *     &D                                      Date
 *     &T                                      Time
 *     &F                                      File name
 *     &A                                      Worksheet name
 *     &fontsize           Font                Font size
 *     &"font,style"                           Font name and style
 *     &U                                      Single underline
 *     &E                                      Double underline
 *     &S                                      Strikethrough
 *     &X                                      Superscript
 *     &Y                                      Subscript
 *     &[Picture]          Images              Image placeholder
 *     &G                                      Same as &[Picture]
 *     &&                  Miscellaneous       Literal ampersand &
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, preview, worksheet1, header1, footer1, worksheet2, header2, footer2, worksheet3, header3, footer3, worksheet4, header4, worksheet5, header5, breaks

    lxw_init() 

    workbook = lxw_workbook_new("headers_footers.xlsx")

    preview = "Select Print Preview to see the header and footer"

    /*
     * A simple example to start
     */
    worksheet1 = lxw_workbook_add_worksheet(workbook, "Simple")
    header1 = "&CHere is some centered text."
    footer1 = "&LHere is some left aligned text."

    lxw_worksheet_set_header(worksheet1, header1)
    lxw_worksheet_set_footer(worksheet1, footer1)

    lxw_worksheet_set_column(worksheet1, 0, 0, 50, NIL)
    lxw_worksheet_write_string(worksheet1, 0, 0, preview, NIL)


    /*
     * This is an example of some of the header/footer variables.
     */
    worksheet2 = lxw_workbook_add_worksheet(workbook, "Variables")
    header2 = '&LPage &P of &N &CFilename: &F &RSheetname: &A'
    footer2 = '&LCurrent date: &D &RCurrent time: &T'
    breaks = {20, 30, 0} //{20, 40, 60, 80, 0} //

    lxw_worksheet_set_header(worksheet2, header2)
    lxw_worksheet_set_footer(worksheet2, footer2)

    lxw_worksheet_set_column(worksheet2, 0, 0, 50, NIL)
    lxw_worksheet_write_string(worksheet2, 0, 0, preview, NIL)

    lxw_worksheet_set_h_pagebreaks(worksheet2, breaks)
    lxw_worksheet_write_string(worksheet2, 20, 0, "Next page", NIL)
    lxw_worksheet_write_string(worksheet2, 30, 0, "Other Next page", NIL)


    /*
     * This example shows how to use more than one font.
     */
    worksheet3 = lxw_workbook_add_worksheet(workbook, "Mixed fonts")
    header3 = '&C&"Courier New,Bold"Hello &"Arial,Italic"World'
    footer3 = '&C&"Symbol"e&"Arial" = mc&X2'

    lxw_worksheet_set_header(worksheet3, header3)
    lxw_worksheet_set_footer(worksheet3, footer3)

    lxw_worksheet_set_column(worksheet3, 0, 0, 50, NIL)
    lxw_worksheet_write_string(worksheet3, 0, 0, preview, NIL)


    /*
     * Example of line wrapping.
     */
    worksheet4 = lxw_workbook_add_worksheet(workbook, "Word wrap")
    header4 = "&CHeading 1\nHeading 2"

    lxw_worksheet_set_header(worksheet4, header4)

    lxw_worksheet_set_column(worksheet4, 0, 0, 50, NIL)
    lxw_worksheet_write_string(worksheet4, 0, 0, preview, NIL)


    /*
     * Example of inserting a literal ampersand &
     */
    worksheet5 = lxw_workbook_add_worksheet(workbook, "Ampersand")
    header5 = "&CCuriouser && Curiouser - Attorneys at Law"

    lxw_worksheet_set_header(worksheet5, header5)

    lxw_worksheet_set_column(worksheet5, 0, 0, 50, NIL)
    lxw_worksheet_write_string(worksheet5, 0, 0, preview, NIL)


    lxw_workbook_close(workbook)

    return 0

