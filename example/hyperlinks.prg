/*
 * Example of writing urls/hyperlinks with the libxlsxwriter library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

function main() 
    local workbook, worksheet, url_format, red_format

    lxw_init() 

    /* Create a new workbook. */
    workbook   = lxw_workbook_new("hyperlinks.xlsx")

    /* Add a worksheet. */
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    /* Add some cell formats for the hyperlinks. */
    url_format   = lxw_workbook_add_format(workbook)
    red_format   = lxw_workbook_add_format(workbook)

    /* Create the standard url link format. */
    lxw_format_set_underline (url_format, LXW_UNDERLINE_SINGLE)
    lxw_format_set_font_color(url_format, LXW_COLOR_BLUE)

    /* Create another sample format. */
    lxw_format_set_underline (red_format, LXW_UNDERLINE_SINGLE)
    lxw_format_set_font_color(red_format, LXW_COLOR_RED)

    /* Widen the first column to make the text clearer. */
    lxw_worksheet_set_column(worksheet, 0, 0, 30, NIL)

    /* Write a hyperlink. */
    lxw_worksheet_write_url(worksheet,    0, 0, "http://libxlsxwriter.github.io", url_format)

    /* Write a hyperlink but overwrite the displayed string. */
    lxw_worksheet_write_url   (worksheet, 2, 0, "http://libxlsxwriter.github.io", url_format)
    lxw_worksheet_write_string(worksheet, 2, 0, "Read the documentation.",        url_format)

    /* Write a hyperlink with a different format. */
    lxw_worksheet_write_url(worksheet,    4, 0, "http://libxlsxwriter.github.io", red_format)

    /* Write a mail hyperlink. */
    lxw_worksheet_write_url   (worksheet, 6, 0, "mailto:jmcnamara@cpan.org",      url_format)

    /* Write a mail hyperlink and overwrite the displayed string. */
    lxw_worksheet_write_url   (worksheet, 8, 0, "mailto:jmcnamara@cpan.org",      url_format)
    lxw_worksheet_write_string(worksheet, 8, 0, "Drop me a line.",                url_format)


    /* Close the workbook, save the file and free any memory. */
    lxw_workbook_close(workbook)

    return 0

