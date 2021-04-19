FUNCTION main(nLines)
LOCAL workbook, worksheet, row, col, bold, money, date_format, i, nL, nC, nS

nLines:= IIF( nLines == Nil, 10, VAL(nLines) )

/* Some data we want to write to the worksheet. */

    lxw_init() 

    /* Create a workbook and add a worksheet. */
    workbook  = lxw_workbook_new("tutorial04.xlsx")
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
    lxw_format_set_num_format(date_format, "mmmm d yyyy")

    /* Adjust the column width. */
    lxw_worksheet_set_column(worksheet, 0, 0, 40, NIL)
    lxw_worksheet_set_column(worksheet, 1, 1, 15, NIL)
    lxw_worksheet_set_column(worksheet, 2, 2, 15, NIL)
    lxw_worksheet_set_column(worksheet, 3, 3, 45, NIL)
    lxw_worksheet_set_column(worksheet, 4, 4, 45, NIL)

    /* Write some data header. */
    lxw_worksheet_write_string(worksheet, row, col,     "Item", bold)
    lxw_worksheet_write_string(worksheet, row, col + 1, "Date", bold)
    lxw_worksheet_write_string(worksheet, row, col + 2, "Cost", bold)

    /* Iterate over the data and write it out element by element. */
    QOUT()
    nL:= ROW()
    nC:= COL()
    nS:= SECONDS()

    FOR i := 1 TO nLines
        @ nL,nC SAY i/nLines*100        /* Write from the first cell below the headers. */
        row:= i
        lxw_worksheet_write_string( worksheet, row, col,     "STRING TEST / TEST STRING "+HB_NTOS(i), NIL)
        lxw_worksheet_write_string( worksheet, row, col + 1, DTOC(DATE()), NIL)
        lxw_worksheet_write_number( worksheet, row, col + 2, i,  money)
        lxw_worksheet_write_string( worksheet, row, col + 3, "OTHER STRING TEST / TEST STRING "+HB_NTOS(i), NIL)
        lxw_worksheet_write_string( worksheet, row, col + 4, "ANOTHER STRING TEST / TEST STRING "+HB_NTOS(i), NIL)
    next

    /* Write a total using a formula. */
    lxw_worksheet_write_string (worksheet, row + 1, col,     "Total",       bold)
    lxw_worksheet_write_formula(worksheet, row + 1, col + 2, "=SUM(C2:C5)", money)

    /* Save the workbook and free any allocated memory. */
    lxw_workbook_close(workbook)
    QOUT( SECONDS() - nS )
RETURN NIL
