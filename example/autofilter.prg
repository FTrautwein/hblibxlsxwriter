/*
 * Example of adding an autofilter to a worksheet in Excel using
 * libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 */

#include "hblibxlsxwriter.ch"

#define region  1
#define item    2
#define volume  3
#define month   4

function main() 
    local workbook, worksheet, data, i

    lxw_init() 

    workbook  = lxw_workbook_new("autofilter.xlsx")
    worksheet = lxw_workbook_add_worksheet(workbook, NIL)

    data:= {;
        {"East",  "Apple",   9000, "July"      },;
        {"East",  "Apple",   5000, "July"      },;
        {"South", "Orange",  9000, "September" },;
        {"North", "Apple",   2000, "November"  },;
        {"West",  "Apple",   9000, "November"  },;
        {"South", "Pear",    7000, "October"   },;
        {"North", "Pear",    9000, "August"    },;
        {"West",  "Orange",  1000, "December"  },;
        {"West",  "Grape",   1000, "November"  },;
        {"South", "Pear",   10000, "April"     },;
        {"West",  "Grape",   6000, "January"   },;
        {"South", "Orange",  3000, "May"       },;
        {"North", "Apple",   3000, "December"  },;
        {"South", "Apple",   7000, "February"  },;
        {"West",  "Grape",   1000, "December"  },;
        {"East",  "Grape",   8000, "February"  },;
        {"South", "Grape",  10000, "June"      },;
        {"West",  "Pear",    7000, "December"  },;
        {"South", "Apple",   2000, "October"   },;
        {"East",  "Grape",   7000, "December"  },;
        {"North", "Grape",   6000, "April"     },;
        {"East",  "Pear",    8000, "February"  },;
        {"North", "Apple",   7000, "August"    },;
        {"North", "Orange",  7000, "July"      },;
        {"North", "Apple",   6000, "June"      },;
        {"South", "Grape",   8000, "September" },;
        {"West",  "Apple",   3000, "October"   },;
        {"South", "Orange", 10000, "November"  },;
        {"West",  "Grape",   4000, "July"      },;
        {"North", "Orange",  5000, "August"    },;
        {"East",  "Orange",  1000, "November"  },;
        {"East",  "Orange",  4000, "October"   },;
        {"North", "Grape",   5000, "August"    },;
        {"East",  "Apple",   1000, "December"  },;
        {"South", "Apple",   10000, "March"    },;
        {"East",  "Grape",   7000, "October"   },;
        {"West",  "Grape",   1000, "September" },;
        {"East",  "Grape",  10000, "October"   },;
        {"South", "Orange",  8000, "March"     },;
        {"North", "Apple",   4000, "July"      },;
        {"South", "Orange",  5000, "July"      },;
        {"West",  "Apple",   4000, "June"      },;
        {"East",  "Apple",   5000, "April"     },;
        {"North", "Pear",    3000, "August"    },;
        {"East",  "Grape",   9000, "November"  },;
        {"North", "Orange",  8000, "October"   },;
        {"East",  "Apple",  10000, "June"      },;
        {"South", "Pear",    1000, "December"  },;
        {"North", "Grape",   10000, "July"     },;
        {"East",  "Grape",   6000, "February"  };
    }

    /* Write the column headers. */
    lxw_worksheet_write_string(worksheet, 0, 0, "Region", NIL)
    lxw_worksheet_write_string(worksheet, 0, 1, "Item",   NIL)
    lxw_worksheet_write_string(worksheet, 0, 2, "Volume", NIL)
    lxw_worksheet_write_string(worksheet, 0, 3, "Month",  NIL)

    /* Write the row data. */
    for i = 0 TO LEN(data) - 1
        lxw_worksheet_write_string(worksheet, i + 1, 0, data[i+1,region], NIL)
        lxw_worksheet_write_string(worksheet, i + 1, 1, data[i+1,item],   NIL)
        lxw_worksheet_write_number(worksheet, i + 1, 2, data[i+1,volume], NIL)
        lxw_worksheet_write_string(worksheet, i + 1, 3, data[i+1,month],  NIL)
    next

    /* Add the autofilter. */
    lxw_worksheet_autofilter(worksheet, 0, 0, 50, 3)

    return lxw_workbook_close(workbook)

