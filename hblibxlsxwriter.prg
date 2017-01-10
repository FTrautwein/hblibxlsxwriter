// -----------------------------------------------------------------------------
// hblibxlsxwriter
//
// libxlsxwriter wrappers for Harbour
//
// Copyright 2017 Fausto Di Creddo Trautwein <ftwein@gmail.com>
//
// -----------------------------------------------------------------------------

/*
 * libxlsxwriter
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

#include "hbdyn.ch" 
#include "CStruct.ch" // required for "typedef struct"
#include "Wintypes.ch" 

STATIC nHDll

/** @brief Struct to represent a date and time in Excel.
 *
 * Struct to represent a date and time in Excel. See @ref working_with_dates.
 */
    /** Year     : 1900 - 9999 */
    /** Month    : 1 - 12 */
    /** Day      : 1 - 31 */
    /** Hour     : 0 - 23 */
    /** Minute   : 0 - 59 */
    /** Seconds  : 0 - 59.999 */
pragma pack(8) 
typedef struct lxw_datetime {;
    int year;
    int month;
    int day;
    int hour;
    int min;
    DOUBLE sec;
} lxw_datetime;

/**
 * @brief Workbook options.
 *
 * Optional parameters when creating a new Workbook object via
 * workbook_new_opt().
 *
 * The following properties are supported:
 *
 * - `constant_memory`: Reduces the amount of data stored in memory so that
 *   large files can be written efficiently.
 *
 *   @note In this mode a row of data is written and then discarded when a
 *   cell in a new row is added via one of the `worksheet_write_*()`
 *   functions. Therefore, once this option is active, data should be written in
 *   sequential row order. For this reason the `worksheet_merge_range()`
 *   doesn't work in this mode. See also @ref ww_mem_constant.
 *
 * - `tmpdir`: libxlsxwriter stores workbook data in temporary files prior
 *   to assembling the final XLSX file. The temporary files are created in the
 *   system's temp directory. If the default temporary directory isn't
 *   accessible to your application, or doesn't contain enough space, you can
 *   specify an alternative location using the `tempdir` option.
 */
pragma pack(8) 
typedef struct lxw_workbook_options {;
    UCHAR constant_memory;
    LPCSTR tmpdir;
} lxw_workbook_options;

/**
 * Workbook document properties.
 */
    /** The title of the Excel Document. */
    /** The subject of the Excel Document. */
    /** The author of the Excel Document. */
    /** The manager field of the Excel Document. */
    /** The company field of the Excel Document. */
    /** The category of the Excel Document. */
    /** The keywords of the Excel Document. */
    /** The comment field of the Excel Document. */
    /** The status of the Excel Document. */
    /** The hyperlink base url of the Excel Document. */
pragma pack(8) 
typedef struct lxw_doc_properties {;
    LPCSTR title;
    LPCSTR subject;
    LPCSTR author;
    LPCSTR manager;
    LPCSTR company;
    LPCSTR category;
    LPCSTR keywords;
    LPCSTR comments;
    LPCSTR status;
    LPCSTR hyperlink_base;
    int created;
} lxw_doc_properties;

FUNCTION lxw_init() 
nHDll:= hb_libLoad( "libxlsxwriter.dll" )
RETURN Nil

FUNCTION CallDll( cProc, ... )         
LOCAL xRet
xRet:= hb_DynCall( { cProc, nHDll, HB_DYN_CALLCONV_SYSCALL }, ... )
RETURN xRet

FUNCTION ToDouble(n)
RETURN n+0.01-0.01

// WORKBOOK ===========================================================================================================
/**
 * @page workbook_page The Workbook object
 *
 * The Workbook is the main object exposed by the libxlsxwriter library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * See @ref workbook.h for full details of the functionality.
 *
 * @file workbook.h
 *
 * @brief Functions related to creating an Excel xlsx workbook.
 *
 * The Workbook is the main object exposed by the libxlsxwriter library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * @code
 *     #include "xlsxwriter.h"
 *
 *     int main() {
 *
 *         lxw_workbook  *workbook  = workbook_new("filename.xlsx")
 *         FUNCTION lxw_worksheet = workbook_add_worksheet(workbook, NULL)
 *
 *         worksheet_write_string(worksheet, 0, 0, "Hello Excel", NULL)
 *
 *         return workbook_close(workbook)
 *     }
 * @endcode
 *
 * @image html workbook01.png
 *
 */

/**
 * @brief Create a new workbook object.
 *
 * @param filename The name of the new Excel file to create.
 *
 * @return A lxw_workbook instance.
 *
 * The `%workbook_new()` constructor is used to create a new Excel workbook
 * with a given filename:
 *
 * @code
 *     FUNCTION lxw_workbook  = workbook_new("filename.xlsx")
 * @endcode
 *
 * When specifying a filename it is recommended that you use an `.xlsx`
 * extension or Excel will generate a warning when opening the file.
 *
 */
FUNCTION lxw_workbook_new( filename )
RETURN CallDll( "workbook_new", filename )

/**
 * @brief Create a new workbook object, and set the workbook options.
 *
 * @param filename The name of the new Excel file to create.
 * @param options  Workbook options.
 *
 * @return A lxw_workbook instance.
 *
 * This function is the same as the `workbook_new()` constructor but allows
 * additional options to be set.
 *
 * @code
 *    lxw_workbook_options options = {.constant_memory = 1,
 *                                    .tmpdir = "C:\\Temp"};
 *
 *    lxw_workbook  *workbook  = workbook_new_opt("filename.xlsx", &options)
 * @endcode
 *
 * The options that can be set via #lxw_workbook_options are:
 *
 * - `constant_memory`: Reduces the amount of data stored in memory so that
 *   large files can be written efficiently.
 *
 *   @note In this mode a row of data is written and then discarded when a
 *   cell in a new row is added via one of the `worksheet_write_*()`
 *   functions. Therefore, once this option is active, data should be written in
 *   sequential row order. For this reason the `worksheet_merge_range()`
 *   doesn't work in this mode. See also @ref ww_mem_constant.
 *
 * - `tmpdir`: libxlsxwriter stores workbook data in temporary files prior
 *   to assembling the final XLSX file. The temporary files are created in the
 *   system's temp directory. If the default temporary directory isn't
 *   accessible to your application, or doesn't contain enough space, you can
 *   specify an alternative location using the `tempdir` option.*
 *
 * See @ref working_with_memory for more details.
 *
 */
FUNCTION lxw_workbook_new_opt(filename, options)
LOCAL oOptions , pOptions
oOptions := hb_CStructure( "lxw_workbook_options" ):Buffer()
IF HB_HHASKEY( options, "constant_memory" )
   oOptions:constant_memory:= options[ "constant_memory" ]
ENDIF   
IF HB_HHASKEY( options, "tmpdir" )
   oOptions:tmpdir:= options[ "tmpdir" ]
ENDIF
pOptions:= oOptions:GetPointer()
RETURN CallDll( "workbook_new_opt", filename, pOptions )

/* Deprecated function name for backwards compatibility. */
FUNCTION lxw_new_workbook(filename)
RETURN CallDll( "new_workbook", filename )

/* Deprecated function name for backwards compatibility. */
FUNCTION lxw_new_workbook_opt(filename, options)
LOCAL oOptions , pOptions
oOptions := hb_CStructure( "lxw_workbook_options" ):Buffer()
IF HB_HHASKEY( options, "constant_memory" )
   oOptions:constant_memory:= options[ "constant_memory" ]
ENDIF   
IF HB_HHASKEY( options, "tmpdir" )
   oOptions:tmpdir:= options[ "tmpdir" ]
ENDIF
pOptions:= oOptions:GetPointer()
RETURN CallDll( "new_workbook_opt", filename, pOptions )

/**
 * @brief Add a new worksheet to a workbook.
 *
 * @param workbook  Pointer to a lxw_workbook instance.
 * @param sheetname Optional worksheet name, defaults to Sheet1, etc.
 *
 * @return A lxw_worksheet object.
 *
 * The `%workbook_add_worksheet()` function adds a new worksheet to a workbook:
 *
 * At least one worksheet should be added to a new workbook: The @ref
 * worksheet.h "Worksheet" object is used to write data and configure a
 * worksheet in the workbook.
 *
 * The `sheetname` parameter is optional. If it is `NULL` the default
 * Excel convention will be followed, i.e. Sheet1, Sheet2, etc.:
 *
 * @code
 *     worksheet = workbook_add_worksheet(workbook, NULL  )     // Sheet1
 *     worksheet = workbook_add_worksheet(workbook, "Foglio2")  // Foglio2
 *     worksheet = workbook_add_worksheet(workbook, "Data")     // Data
 *     worksheet = workbook_add_worksheet(workbook, NULL  )     // Sheet4
 *
 * @endcode
 *
 * @image html workbook02.png
 *
 * The worksheet name must be a valid Excel worksheet name, i.e. it must be
 * less than 32 character and it cannot contain any of the characters:
 *
 *     / \ [ ] : * ?
 *
 * In addition, you cannot use the same, case insensitive, `sheetname` for more
 * than one worksheet.
 *
 */
FUNCTION lxw_workbook_add_worksheet( workbook, sheetname)
RETURN CallDll( "workbook_add_worksheet", workbook, sheetname )

/**
 * @brief Create a new @ref format.h "Format" object to formats cells in
 *        worksheets.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 *
 * @return A lxw_format instance.
 *
 * The `workbook_add_format()` function can be used to create new @ref
 * format.h "Format" objects which are used to apply formatting to a cell.
 *
 * @code
 *    // Create the Format.
 *    lxw_format *format = workbook_add_format(workbook)
 *
 *    // Set some of the format properties.
 *    format_set_bold(format)
 *    format_set_font_color(format, LXW_COLOR_RED)
 *
 *    // Use the format to change the text format in a cell.
 *    worksheet_write_string(worksheet, 0, 0, "Hello", format)
 * @endcode
 *
 * See @ref format.h "the Format object" and @ref working_with_formats
 * sections for more details about Format properties and how to set them.
 *
 */
FUNCTION lxw_workbook_add_format( workbook )
RETURN CallDll( "workbook_add_format", workbook )

/**
 * @brief Create a new chart to be added to a worksheet:
 *
 * @param workbook   Pointer to a lxw_workbook instance.
 * @param chart_type The type of chart to be created. See #lxw_chart_type.
 *
 * @return A lxw_chart object.
 *
 * The `%workbook_add_chart()` function creates a new chart object that can
 * be added to a worksheet:
 *
 * @code
 *     // Create a chart object.
 *     lxw_chart *chart = workbook_add_chart(workbook, LXW_CHART_COLUMN)
 *
 *     // Add data series to the chart.
 *     chart_add_series(chart, NULL, "Sheet1!$A$1:$A$5")
 *     chart_add_series(chart, NULL, "Sheet1!$B$1:$B$5")
 *     chart_add_series(chart, NULL, "Sheet1!$C$1:$C$5")
 *
 *     // Insert the chart into the worksheet
 *     worksheet_insert_chart(worksheet, CELL("B7"), chart)
 * @endcode
 *
 * The available chart types are defined in #lxw_chart_type. The types of
 * charts that are supported are:
 *
 * | Chart type                               | Description                            |
 * | :--------------------------------------- | :------------------------------------  |
 * | #LXW_CHART_AREA                          | Area chart.                            |
 * | #LXW_CHART_AREA_STACKED                  | Area chart - stacked.                  |
 * | #LXW_CHART_AREA_STACKED_PERCENT          | Area chart - percentage stacked.       |
 * | #LXW_CHART_BAR                           | Bar chart.                             |
 * | #LXW_CHART_BAR_STACKED                   | Bar chart - stacked.                   |
 * | #LXW_CHART_BAR_STACKED_PERCENT           | Bar chart - percentage stacked.        |
 * | #LXW_CHART_COLUMN                        | Column chart.                          |
 * | #LXW_CHART_COLUMN_STACKED                | Column chart - stacked.                |
 * | #LXW_CHART_COLUMN_STACKED_PERCENT        | Column chart - percentage stacked.     |
 * | #LXW_CHART_DOUGHNUT                      | Doughnut chart.                        |
 * | #LXW_CHART_LINE                          | Line chart.                            |
 * | #LXW_CHART_PIE                           | Pie chart.                             |
 * | #LXW_CHART_SCATTER                       | Scatter chart.                         |
 * | #LXW_CHART_SCATTER_STRAIGHT              | Scatter chart - straight.              |
 * | #LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS | Scatter chart - straight with markers. |
 * | #LXW_CHART_SCATTER_SMOOTH                | Scatter chart - smooth.                |
 * | #LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS   | Scatter chart - smooth with markers.   |
 * | #LXW_CHART_RADAR                         | Radar chart.                           |
 * | #LXW_CHART_RADAR_WITH_MARKERS            | Radar chart - with markers.            |
 * | #LXW_CHART_RADAR_FILLED                  | Radar chart - filled.                  |
 *
 *
 *
 * See @ref chart.h for details.
 */
FUNCTION lxw_workbook_add_chart( workbook, chart_type)
RETURN CallDll( "workbook_add_chart", workbook, chart_type )

/**
 * @brief Close the Workbook object and write the XLSX file.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 *
 * @return A #lxw_error.
 *
 * The `%workbook_close()` function closes a Workbook object, writes the Excel
 * file to disk, frees any memory allocated internally to the Workbook and
 * frees the object itself.
 *
 * @code
 *     workbook_close(workbook)
 * @endcode
 *
 * The `%workbook_close()` function returns any #FUNCTION error codes
 * encountered when creating the Excel file. The error code can be returned
 * from the program main or the calling function:
 *
 * @code
 *     return workbook_close(workbook)
 * @endcode
 *
 */
FUNCTION lxw_workbook_close(workbook)
RETURN CallDll( "workbook_close", workbook )

/**
 * @brief Set the document properties such as Title, Author etc.
 *
 * @param workbook   Pointer to a lxw_workbook instance.
 * @param properties Document properties to set.
 *
 * @return A #lxw_error.
 *
 * The `%workbook_set_properties` function can be used to set the document
 * properties of the Excel file created by `libxlsxwriter`. These properties
 * are visible when you use the `Office Button -> Prepare -> Properties`
 * option in Excel and are also available to external applications that read
 * or index windows files.
 *
 * The properties that can be set are:
 *
 * - `title`
 * - `subject`
 * - `author`
 * - `manager`
 * - `company`
 * - `category`
 * - `keywords`
 * - `comments`
 * - `hyperlink_base`
 *
 * The properties are specified via a `lxw_doc_properties` struct. All the
 * members are `char *` and they are all optional. An example of how to create
 * and pass the properties is:
 *
 * @code
 *     // Create a properties structure and set some of the fields.
 *     lxw_doc_properties properties = {
 *         .title    = "This is an example spreadsheet",
 *         .subject  = "With document properties",
 *         .author   = "John McNamara",
 *         .manager  = "Dr. Heinz Doofenshmirtz",
 *         .company  = "of Wolves",
 *         .category = "Example spreadsheets",
 *         .keywords = "Sample, Example, Properties",
 *         .comments = "Created with libxlsxwriter",
 *         .status   = "Quo",
 *     };
 *
 *     // Set the properties in the workbook.
 *     workbook_set_properties(workbook, &properties)
 * @endcode
 *
 * @image html doc_properties.png
 *
 */
FUNCTION lxw_workbook_set_properties( workbook, properties)
LOCAL oProperties , pProperties
oProperties := hb_CStructure( "lxw_doc_properties" ):Buffer()
IF HB_HHaskey( properties, "title" )
   oProperties:title         := properties[ "title"          ]
ENDIF
IF HB_HHaskey( properties, "subject" )
   oProperties:subject       := properties[ "subject"        ]
ENDIF
IF HB_HHaskey( properties, "author" )
   oProperties:author        := properties[ "author"         ]
ENDIF
IF HB_HHaskey( properties, "manager" )
   oProperties:manager       := properties[ "manager"        ]
ENDIF
IF HB_HHaskey( properties, "company" )
   oProperties:company       := properties[ "company"        ]
ENDIF
IF HB_HHaskey( properties, "category" )
   oProperties:category      := properties[ "category"       ]
ENDIF
IF HB_HHaskey( properties, "keywords" )
   oProperties:keywords      := properties[ "keywords"       ]
ENDIF
IF HB_HHaskey( properties, "comments" )
   oProperties:comments      := properties[ "comments"       ]
ENDIF
IF HB_HHaskey( properties, "status" )
   oProperties:status        := properties[ "status"         ]
ENDIF
IF HB_HHaskey( properties, "hyperlink_base" )
   oProperties:hyperlink_base:= properties[ "hyperlink_base" ]
ENDIF
IF HB_HHaskey( properties, "created" )
   oProperties:created       := properties[ "created"        ]
ENDIF
pProperties:= oProperties:GetPointer()
RETURN CallDll( "workbook_set_properties", workbook, pProperties )

/**
 * @brief Set a custom document text property.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     The name of the custom property.
 * @param value    The value of the custom property.
 *
 * @return A #lxw_error.
 *
 * The `%workbook_set_custom_property_string()` function can be used to set one
 * or more custom document text properties not covered by the standard
 * properties in the `workbook_set_properties()` function above.
 *
 *  For example:
 *
 * @code
 *     workbook_set_custom_property_string(workbook, "Checked by", "Eve")
 * @endcode
 *
 * @image html custom_properties.png
 *
 * There are 4 `workbook_set_custom_property_string_*()` functions for each
 * of the custom property types supported by Excel:
 *
 * - text/string: `workbook_set_custom_property_string()`
 * - number:      `workbook_set_custom_property_number()`
 * - datetime:    `workbook_set_custom_property_datetime()`
 * - boolean:     `workbook_set_custom_property_boolean()`
 *
 * **Note**: the name and value parameters are limited to 255 characters
 * by Excel.
 *
 */
FUNCTION lxw_workbook_set_custom_property_string( workbook, name, value)
RETURN CallDll( "workbook_set_custom_property_string", workbook, name, value )

/**
 * @brief Set a custom document number property.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     The name of the custom property.
 * @param value    The value of the custom property.
 *
 * @return A #lxw_error.
 *
 * Set a custom document number property.
 * See `workbook_set_custom_property_string()` above for details.
 *
 * @code
 *     workbook_set_custom_property_number(workbook, "Document number", 12345)
 * @endcode
 */
FUNCTION lxw_workbook_set_custom_property_number( workbook, name, value)
RETURN CallDll( "workbook_set_custom_property_number", workbook, name, ToDouble(value) )

/* Undocumented since the user can use workbook_set_custom_property_number().
 * Only implemented for file format completeness and testing.
 */
FUNCTION lxw_workbook_set_custom_property_integer( workbook, name, value)
RETURN CallDll( "workbook_set_custom_property_integer", workbook, name, value )

/**
 * @brief Set a custom document boolean property.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     The name of the custom property.
 * @param value    The value of the custom property.
 *
 * @return A #lxw_error.
 *
 * Set a custom document boolean property.
 * See `workbook_set_custom_property_string()` above for details.
 *
 * @code
 *     workbook_set_custom_property_boolean(workbook, "Has Review", 1)
 * @endcode
 */
FUNCTION lxw_workbook_set_custom_property_boolean( workbook, name, value)
RETURN CallDll( "workbook_set_custom_property_boolean", workbook, name, value )

/**
 * @brief Set a custom document date or time property.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     The name of the custom property.
 * @param datetime The value of the custom property.
 *
 * @return A #lxw_error.
 *
 * Set a custom date or time number property.
 * See `workbook_set_custom_property_string()` above for details.
 *
 * @code
 *     lxw_datetime datetime  = {2016, 12, 1,  11, 55, 0.0};
 *
 *     workbook_set_custom_property_datetime(workbook, "Date completed", &datetime)
 * @endcode
 */
FUNCTION lxw_workbook_set_custom_property_datetime( workbook, name, datetime)
LOCAL oDateTime , pDatetime
oDatetime := hb_CStructure( "LXW_DATETIME" ):Buffer()
oDatetime:year:= YEAR(datetime)
oDatetime:month:= MONTH(datetime)
oDatetime:day:= DAY(datetime)
oDatetime:hour:= HB_HOUR(datetime)
oDatetime:min:= HB_MINUTE(datetime)
oDatetime:sec:= HB_SEC(datetime)
pDatetime:= oDatetime:GetPointer()
RETURN CallDll( "workbook_set_custom_property_datetime", workbook, name, pDatetime )

/**
 * @brief Create a defined name in the workbook to use as a variable.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     The defined name.
 * @param formula  The cell or range that the defined name refers to.
 *
 * @return A #lxw_error.
 *
 * This function is used to defined a name that can be used to represent a
 * value, a single cell or a range of cells in a workbook: These defined names
 * can then be used in formulas:
 *
 * @code
 *     workbook_define_name(workbook, "Exchange_rate", "=0.96")
 *     worksheet_write_formula(worksheet, 2, 1, "=Exchange_rate", NULL)
 *
 * @endcode
 *
 * @image html defined_name.png
 *
 * As in Excel a name defined like this is "global" to the workbook and can be
 * referred to from any worksheet:
 *
 * @code
 *     // Global workbook name.
 *     workbook_define_name(workbook, "Sales", "=Sheet1!$G$1:$H$10")
 * @endcode
 *
 * It is also possible to define a local/worksheet name by prefixing it with
 * the sheet name using the syntax `'sheetname!definedname'`:
 *
 * @code
 *     // Local worksheet name.
 *     workbook_define_name(workbook, "Sheet2!Sales", "=Sheet2!$G$1:$G$10")
 * @endcode
 *
 * If the sheet name contains spaces or special characters you must follow the
 * Excel convention and enclose it in single quotes:
 *
 * @code
 *     workbook_define_name(workbook, "'New Data'!Sales", "=Sheet2!$G$1:$G$10")
 * @endcode
 *
 * The rules for names in Excel are explained in the
 * [Microsoft Office
documentation](http://office.microsoft.com/en-001/excel-help/define-and-use-names-in-formulas-HA010147120.aspx).
 *
 */
FUNCTION lxw_workbook_define_name( workbook, name, formula)
RETURN CallDll( "workbook_define_name", workbook, name, formula )

/**
 * @brief Get a worksheet object from its name.
 *
 * @param workbook
 * @param name
 *
 * @return A lxw_worksheet object.
 *
 * This function returns a lxw_worksheet object reference based on its name:
 *
 * @code
 *     worksheet = workbook_get_worksheet_by_name(workbook, "Sheet1")
 * @endcode
 *
 */
FUNCTION lxw_workbook_get_worksheet_by_name( workbook, name)
RETURN CallDll( "workbook_get_worksheet_by_name", workbook, name )

/**
 * @brief Validate a worksheet name.
 *
 * @param workbook  Pointer to a lxw_workbook instance.
 * @param sheetname Worksheet name to validate.
 *
 * @return A #lxw_error.
 *
 * This function is used to validate a worksheet name according to the rules
 * used by Excel:
 *
 * - The name is less than or equal to 31 UTF-8 characters.
 * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
 * - The name isn't already in use.
 *
 * @code
 *     FUNCTION err = workbook_validate_worksheet_name(workbook, "Foglio")
 * @endcode
 *
 * This function is called by `workbook_add_worksheet()` but it can be
 * explicitly called by the user beforehand to ensure that the worksheet
 * name is valid.
 *
 */
FUNCTION lxw_workbook_validate_worksheet_name( workbook, sheetname)
RETURN CallDll( "workbook_validate_worksheet_name", workbook, sheetname )

// WORKSHEET ==========================================================================================================

/**
 * @brief Write a number to a worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param number    The number to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #FUNCTION code.
 *
 * The `worksheet_write_number()` function writes numeric types to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_number(worksheet, 0, 0, 123456, NULL)
 *     worksheet_write_number(worksheet, 1, 0, 2.3451, NULL)
 * @endcode
 *
 * @image html write_number01.png
 *
 * The native data type for all numbers in Excel is a IEEE-754 64-bit
 * double-precision floating point, which is also the default type used by
 * `%worksheet_write_number`.
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object.
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_num_format(format, "$#,##0.00")
 *
 *     worksheet_write_number(worksheet, 0, 0, 1234.567, format)
 * @endcode
 *
 * @image html write_number02.png
 *
 */
FUNCTION lxw_worksheet_write_number(worksheet, row, col, number, format)
RETURN CallDll( "worksheet_write_number", worksheet, row, col, ToDouble(number), format )


/**
 * @brief Write a string to a worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param string    String to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_write_string()` function writes a string to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_string(worksheet, 0, 0, "This phrase is English!", NULL)
 * @endcode
 *
 * @image html write_string01.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_bold(format)
 *
 *     worksheet_write_string(worksheet, 0, 0, "This phrase is Bold!", format)
 * @endcode
 *
 * @image html write_string02.png
 *
 * Unicode strings are supported in UTF-8 encoding. This generally requires
 * that your source file is UTF-8 encoded or that the data has been read from
 * a UTF-8 source:
 *
 * @code
 *    worksheet_write_string(worksheet, 0, 0, "Это фраза на русском!", NULL)
 * @endcode
 *
 * @image html write_string03.png
 *
 */
FUNCTION lxw_worksheet_write_string(worksheet, row, col, string, format)
RETURN CallDll( "worksheet_write_string", worksheet, row, col, string, format )
                                                                     
/**
 * @brief Write a formula to a worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_write_formula()` function writes a formula or function to
 * the cell specified by `row` and `column`:
 *
 * @code
 *  worksheet_write_formula(worksheet, 0, 0, "=B3 + 6",                    NULL)
 *  worksheet_write_formula(worksheet, 1, 0, "=SIN(PI()/4)",               NULL)
 *  worksheet_write_formula(worksheet, 2, 0, "=SUM(A1:A2)",                NULL)
 *  worksheet_write_formula(worksheet, 3, 0, "=IF(A3>1,\"Yes\", \"No\")",  NULL)
 *  worksheet_write_formula(worksheet, 4, 0, "=AVERAGE(1, 2, 3, 4)",       NULL)
 *  worksheet_write_formula(worksheet, 5, 0, "=DATEVALUE(\"1-Jan-2013\")", NULL)
 * @endcode
 *
 * @image html write_formula01.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object.
 *
 * Libxlsxwriter doesn't calculate the value of a formula and instead stores a
 * default value of `0`. The correct formula result is displayed in Excel, as
 * shown in the example above, since it recalculates the formulas when it loads
 * the file. For cases where this is an issue see the
 * `worksheet_write_formula_num()` function and the discussion in that section.
 *
 * Formulas must be written with the US style separator/range operator which
 * is a comma (not semi-colon). Therefore a formula with multiple values
 * should be written as follows:
 *
 * @code
 *     // OK.
 *     worksheet_write_formula(worksheet, 0, 0, "=SUM(1, 2, 3)", NULL)
 *
 *     // NO. Error on load.
 *     worksheet_write_formula(worksheet, 1, 0, "=SUM(1; 2; 3)", NULL)
 * @endcode
 *
 */
FUNCTION lxw_worksheet_write_formula(worksheet, row, col, formula, format)
RETURN CallDll( "worksheet_write_formula", worksheet, row, col, formula, format )

/**
 * @brief Write an array formula to a worksheet cell.
 *
 * @param worksheet
 * @param first_row   The first row of the range. (All zero indexed.)
 * @param first_col   The first column of the range.
 * @param last_row    The last row of the range.
 * @param last_col    The last col of the range.
 * @param formula     Array formula to write to cell.
 * @param format      A pointer to a Format instance or NULL.
 *
 * @return A #FUNCTION code.
 *
  * The `%worksheet_write_array_formula()` function writes an array formula to
 * a cell range. In Excel an array formula is a formula that performs a
 * calculation on a set of values.
 *
 * In Excel an array formula is indicated by a pair of braces around the
 * formula: `{=SUM(A1:B1*A2:B2)}`.
 *
 * Array formulas can return a single value or a range or values. For array
 * formulas that return a range of values you must specify the range that the
 * return values will be written to. This is why this function has `first_`
 * and `last_` row/column parameters. The RANGE() macro can also be used to
 * specify the range:
 *
 * @code
 *     worksheet_write_array_formula(worksheet, 4, 0, 6, 0,     "{=TREND(C5:C7,B5:B7)}", NULL)
 *
 *     // Same as above using the RANGE() macro.
 *     worksheet_write_array_formula(worksheet, RANGE("A5:A7"), "{=TREND(C5:C7,B5:B7)}", NULL)
 * @endcode
 *
 * If the array formula returns a single value then the `first_` and `last_`
 * parameters should be the same:
 *
 * @code
 *     worksheet_write_array_formula(worksheet, 1, 0, 1, 0,     "{=SUM(B1:C1*B2:C2)}", NULL)
 *     worksheet_write_array_formula(worksheet, RANGE("A2:A2"), "{=SUM(B1:C1*B2:C2)}", NULL)
 * @endcode
 *
 */
FUNCTION lxw_worksheet_write_array_formula(worksheet, first_row, first_col, last_row, last_col, formula, format)
RETURN CallDll( "worksheet_write_array_formula", worksheet, first_row, first_col, last_row, last_col, formula, format )

FUNCTION lxw_worksheet_write_array_formula_num(worksheet, first_row, first_col, last_row, last_col, formula, format, result)
RETURN CallDll( "worksheet_write_array_formula", worksheet, first_row, first_col, last_row, last_col, formula, format, ToDouble(result) )

/**
 * @brief Write a date or time to a worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param datetime  The datetime to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #FUNCTION code.
 *
 * The `worksheet_write_datetime()` function can be used to write a date or
 * time to the cell specified by `row` and `column`:
 *
 * @dontinclude dates_and_times02.c
 * @skip include
 * @until num_format
 * @skip Feb
 * @until }
 *
 * The `format` parameter should be used to apply formatting to the cell using
 * a @ref format.h "Format" object as shown above. Without a date format the
 * datetime will appear as a number only.
 *
 * See @ref working_with_dates for more information about handling dates and
 * times in libxlsxwriter.
 */
FUNCTION lxw_worksheet_write_datetime(worksheet, row, col, datetime, format )
LOCAL oDateTime , pDatetime
oDatetime := hb_CStructure( "LXW_DATETIME" ):Buffer()
oDatetime:year:= YEAR(datetime)
oDatetime:month:= MONTH(datetime)
oDatetime:day:= DAY(datetime)
oDatetime:hour:= HB_HOUR(datetime)
oDatetime:min:= HB_MINUTE(datetime)
oDatetime:sec:= HB_SEC(datetime)
pDatetime:= oDatetime:GetPointer()
RETURN CallDll( "worksheet_write_datetime", worksheet, row, col, pDatetime, format )


FUNCTION lxw_worksheet_write_url_opt(worksheet, row_num, col_num, url, format, string, tooltip)
RETURN CallDll( "worksheet_write_url_opt", worksheet, row_num, col_num, url, format, string, tooltip )

/**
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param url       The url to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #FUNCTION code.
 *
 *
 * The `%worksheet_write_url()` function is used to write a URL/hyperlink to a
 * worksheet cell specified by `row` and `column`.
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "http://libxlsxwriter.github.io", url_format)
 * @endcode
 *
 * @image html hyperlinks_short.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a @ref
 * format.h "Format" object. The typical worksheet format for a hyperlink is a
 * blue underline:
 *
 * @code
 *    url_format   = workbook_add_format(workbook)
 *
 *    format_set_underline (url_format, LXW_UNDERLINE_SINGLE)
 *    format_set_font_color(url_format, LXW_COLOR_BLUE)
 *
 * @endcode
 *
 * The usual web style URI's are supported: `%http://`, `%https://`, `%ftp://`
 * and `mailto:` :
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "ftp://www.python.org/",     url_format)
 *     worksheet_write_url(worksheet, 1, 0, "http://www.python.org/",    url_format)
 *     worksheet_write_url(worksheet, 2, 0, "https://www.python.org/",   url_format)
 *     worksheet_write_url(worksheet, 3, 0, "mailto:jmcnamara@cpan.org", url_format)
 *
 * @endcode
 *
 * An Excel hyperlink is comprised of two elements: the displayed string and
 * the non-displayed link. By default the displayed string is the same as the
 * link. However, it is possible to overwrite it with any other
 * `libxlsxwriter` type using the appropriate `worksheet_write_*()`
 * function. The most common case is to overwrite the displayed link text with
 * another string:
 *
 * @code
 *  // Write a hyperlink but overwrite the displayed string.
 *  worksheet_write_url   (worksheet, 2, 0, "http://libxlsxwriter.github.io", url_format)
 *  worksheet_write_string(worksheet, 2, 0, "Read the documentation.",        url_format)
 *
 * @endcode
 *
 * @image html hyperlinks_short2.png
 *
 * Two local URIs are supported: `internal:` and `external:`. These are used
 * for hyperlinks to internal worksheet references or external workbook and
 * worksheet references:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "internal:Sheet2!A1",                url_format)
 *     worksheet_write_url(worksheet, 1, 0, "internal:Sheet2!B2",                url_format)
 *     worksheet_write_url(worksheet, 2, 0, "internal:Sheet2!A1:B2",             url_format)
 *     worksheet_write_url(worksheet, 3, 0, "internal:'Sales Data'!A1",          url_format)
 *     worksheet_write_url(worksheet, 4, 0, "external:c:\\temp\\foo.xlsx",       url_format)
 *     worksheet_write_url(worksheet, 5, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format)
 *     worksheet_write_url(worksheet, 6, 0, "external:..\\foo.xlsx",             url_format)
 *     worksheet_write_url(worksheet, 7, 0, "external:..\\foo.xlsx#Sheet2!A1",   url_format)
 *     worksheet_write_url(worksheet, 8, 0, "external:\\\\NET\\share\\foo.xlsx", url_format)
 *
 * @endcode
 *
 * Worksheet references are typically of the form `Sheet1!A1`. You can also
 * link to a worksheet range using the standard Excel notation:
 * `Sheet1!A1:B2`.
 *
 * In external links the workbook and worksheet name must be separated by the
 * `#` character:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format)
 * @endcode
 *
 * You can also link to a named range in the target worksheet: For example say
 * you have a named range called `my_name` in the workbook `c:\temp\foo.xlsx`
 * you could link to it as follows:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:\\temp\\foo.xlsx#my_name", url_format)
 *
 * @endcode
 *
 * Excel requires that worksheet names containing spaces or non alphanumeric
 * characters are single quoted as follows:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "internal:'Sales Data'!A1", url_format)
 * @endcode
 *
 * Links to network files are also supported. Network files normally begin
 * with two back slashes as follows `\\NETWORK\etc`. In order to represent
 * this in a C string literal the backslashes should be escaped:
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:\\\\NET\\share\\foo.xlsx", url_format)
 * @endcode
 *
 *
 * Alternatively, you can use Windows style forward slashes. These are
 * translated internally to backslashes:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:/temp/foo.xlsx",     url_format)
 *     worksheet_write_url(worksheet, 1, 0, "external://NET/share/foo.xlsx", url_format)
 *
 * @endcode
 *
 *
 * **Note:**
 *
 *    libxlsxwriter will escape the following characters in URLs as required
 *    by Excel: `\s " < > \ [ ]  ^ { }` unless the URL already contains `%%xx`
 *    style escapes. In which case it is assumed that the URL was escaped
 *    correctly by the user and will by passed directly to Excel.
 *
 */
FUNCTION lxw_worksheet_write_url(worksheet, row, col, url, format)
RETURN CallDll( "worksheet_write_url", worksheet, row, col, url, format )

/**
 * @brief Write a formatted boolean worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param value     The boolean value to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #FUNCTION code.
 *
 * Write an Excel boolean to the cell specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_boolean(worksheet, 2, 2, 0, my_format)
 * @endcode
 *
 */
FUNCTION lxw_worksheet_write_boolean(worksheet, row, col, value, format)
RETURN CallDll( "worksheet_write_boolean", worksheet, row, col, value, format )

/**
 * @brief Write a formatted blank worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #FUNCTION code.
 *
 * Write a blank cell specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_blank(worksheet, 1, 1, border_format)
 * @endcode
 *
 * This function is used to add formatting to a cell which doesn't contain a
 * string or number value.
 *
 * Excel differentiates between an "Empty" cell and a "Blank" cell. An Empty
 * cell is a cell which doesn't contain data or formatting whilst a Blank cell
 * doesn't contain data but does contain formatting. Excel stores Blank cells
 * but ignores Empty cells.
 *
 * As such, if you write an empty cell without formatting it is ignored.
 *
 */
FUNCTION lxw_worksheet_write_blank(worksheet, row, col, format)
RETURN CallDll( "worksheet_write_blank", worksheet, row, col, format )

/**
 * @brief Write a formula to a worksheet cell with a user defined result.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 * @param result    A user defined result for a formula.
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_write_formula_num()` function writes a formula or Excel
 * function to the cell specified by `row` and `column` with a user defined
 * result:
 *
 * @code
 *     // Required as a workaround only.
 *     worksheet_write_formula_num(worksheet, 0, 0, "=1 + 2", NULL, 3)
 * @endcode
 *
 * Libxlsxwriter doesn't calculate the value of a formula and instead stores
 * the value `0` as the formula result. It then sets a global flag in the XLSX
 * file to say that all formulas and functions should be recalculated when the
 * file is opened.
 *
 * This is the method recommended in the Excel documentation and in general it
 * works fine with spreadsheet applications.
 *
 * However, applications that don't have a facility to calculate formulas,
 * such as Excel Viewer, or some mobile applications will only display the `0`
 * results.
 *
 * If required, the `%worksheet_write_formula_num()` function can be used to
 * specify a formula and its result.
 *
 * This function is rarely required and is only provided for compatibility
 * with some third party applications. For most applications the
 * worksheet_write_formula() function is the recommended way of writing
 * formulas.
 *
 */
FUNCTION lxw_worksheet_write_formula_num(worksheet, row, col, formula, format, result)
RETURN CallDll( "worksheet_write_formula_num", worksheet, row, col, formula, format, ToDouble(result) )

/**
 * @brief Set the properties for a row of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param height    The row height.
 * @param format    A pointer to a Format instance or NULL.
 *
 * The `%worksheet_set_row()` function is used to change the default
 * properties of a row. The most common use for this function is to change the
 * height of a row:
 *
 * @code
 *     // Set the height of Row 1 to 20.
 *     worksheet_set_row(worksheet, 0, 20, NULL)
 * @endcode
 *
 * The other common use for `%worksheet_set_row()` is to set the a @ref
 * format.h "Format" for all cells in the row:
 *
 * @code
 *     bold = workbook_add_format(workbook)
 *     format_set_bold(bold)
 *
 *     // Set the header row to bold.
 *     worksheet_set_row(worksheet, 0, 15, bold)
 * @endcode
 *
 * If you wish to set the format of a row without changing the height you can
 * pass the default row height of #LXW_DEF_ROW_HEIGHT = 15:
 *
 * @code
 *     worksheet_set_row(worksheet, 0, LXW_DEF_ROW_HEIGHT, format)
 *     worksheet_set_row(worksheet, 0, 15, format) // Same as above.
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the row that don't
 * have a format. As with Excel the row format is overridden by an explicit
 * cell format. For example:
 *
 * @code
 *     // Row 1 has format1.
 *     worksheet_set_row(worksheet, 0, 15, format1)
 *
 *     // Cell A1 in Row 1 defaults to format1.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL)
 *
 *     // Cell B1 in Row 1 keeps format2.
 *     worksheet_write_string(worksheet, 0, 1, "Hello", format2)
 * @endcode
 *
 */
FUNCTION lxw_worksheet_set_row(worksheet, row, height, format)
RETURN CallDll( "worksheet_set_row", worksheet, row, ToDouble(height), format )

/**
 * @brief Set the properties for a row of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param height    The row height.
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * The `%worksheet_set_row_opt()` function  is the same as
 *  `worksheet_set_row()` with an additional `options` parameter.
 *
 * The `options` parameter is a #lxw_row_col_options struct. It has the
 * following members but currently only the `hidden` property is supported:
 *
 * - `hidden`
 * - `level`
 * - `collapsed`
 *
 * The `"hidden"` option is used to hide a row. This can be used, for
 * example, to hide intermediary steps in a complicated calculation:
 *
 * @code
 *     lxw_row_col_options options = {.hidden = 1, .level = 0, .collapsed = 0};
 *
 *     // Hide the fourth row.
 *     worksheet_set_row(worksheet, 3, 20, NULL, &options)
 * @endcode
 *
 */
//FUNCTION lxw_worksheet_set_row_opt(worksheet, row, height, format, lxw_row_col_options *options)
//RETURN CallDll( "worksheet_set_row_opt", worksheet, row, ToDouble(height), format )

/**
 * @brief Set the properties for one or more columns of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_col The zero indexed first column.
 * @param last_col  The zero indexed last column.
 * @param width     The width of the column(s).
 * @param format    A pointer to a Format instance or NULL.
 *
 * The `%worksheet_set_column()` function can be used to change the default
 * properties of a single column or a range of columns:
 *
 * @code
 *     // Width of columns B:D set to 30.
 *     worksheet_set_column(worksheet, 1, 3, 30, NULL)
 *
 * @endcode
 *
 * If `%worksheet_set_column()` is applied to a single column the value of
 * `first_col` and `last_col` should be the same:
 *
 * @code
 *     // Width of column B set to 30.
 *     worksheet_set_column(worksheet, 1, 1, 30, NULL)
 *
 * @endcode
 *
 * It is also possible, and generally clearer, to specify a column range using
 * the form of `COLS()` macro:
 *
 * @code
 *     worksheet_set_column(worksheet, 4, 4, 20, NULL)
 *     worksheet_set_column(worksheet, 5, 8, 30, NULL)
 *
 *     // Same as the examples above but clearer.
 *     worksheet_set_column(worksheet, COLS("E:E"), 20, NULL)
 *     worksheet_set_column(worksheet, COLS("F:H"), 30, NULL)
 *
 * @endcode
 *
 * The `width` parameter sets the column width in the same units used by Excel
 * which is: the number of characters in the default font. The default width
 * is 8.43 in the default font of Calibri 11. The actual relationship between
 * a string width and a column width in Excel is complex. See the
 * [following explanation of column widths](https://support.microsoft.com/en-us/kb/214123)
 * from the Microsoft support documentation for more details.
 *
 * There is no way to specify "AutoFit" for a column in the Excel file
 * format. This feature is only available at runtime from within Excel. It is
 * possible to simulate "AutoFit" in your application by tracking the maximum
 * width of the data in the column as your write it and then adjusting the
 * column width at the end.
 *
 * As usual the @ref format.h `format` parameter is optional. If you wish to
 * set the format without changing the width you can pass a default column
 * width of #LXW_DEF_COL_WIDTH = 8.43:
 *
 * @code
 *     bold = workbook_add_format(workbook)
 *     format_set_bold(bold)
 *
 *     // Set the first column to bold.
 *     worksheet_set_column(worksheet, 0, 0, LXW_DEF_COL_HEIGHT, bold)
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the column that
 * don't have a format. For example:
 *
 * @code
 *     // Column 1 has format1.
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, format1)
 *
 *     // Cell A1 in column 1 defaults to format1.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL)
 *
 *     // Cell A2 in column 1 keeps format2.
 *     worksheet_write_string(worksheet, 1, 0, "Hello", format2)
 * @endcode
 *
 * As in Excel a row format takes precedence over a default column format:
 *
 * @code
 *     // Row 1 has format1.
 *     worksheet_set_row(worksheet, 0, 15, format1)
 *
 *     // Col 1 has format2.
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, format2)
 *
 *     // Cell A1 defaults to format1, the row format.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL)
 *
 *    // Cell A2 keeps format2, the column format.
 *     worksheet_write_string(worksheet, 1, 0, "Hello", NULL)
 * @endcode
 */
FUNCTION lxw_worksheet_set_column(worksheet, first_col, last_col, width, format)
RETURN CallDll( "worksheet_set_column", worksheet, first_col, last_col, ToDouble(width), format )

 /**
  * @brief Set the properties for one or more columns of cells with options.
  *
  * @param worksheet Pointer to a lxw_worksheet instance to be updated.
  * @param first_col The zero indexed first column.
  * @param last_col  The zero indexed last column.
  * @param width     The width of the column(s).
  * @param format    A pointer to a Format instance or NULL.
  * @param options   Optional row parameters: hidden, level, collapsed.
  *
  * The `%worksheet_set_column_opt()` function  is the same as
  * `worksheet_set_column()` with an additional `options` parameter.
  *
  * The `options` parameter is a #lxw_row_col_options struct. It has the
  * following members but currently only the `hidden` property is supported:
  *
  * - `hidden`
  * - `level`
  * - `collapsed`
  *
  * The `"hidden"` option is used to hide a column. This can be used, for
  * example, to hide intermediary steps in a complicated calculation:
  *
  * @code
  *     lxw_row_col_options options = {.hidden = 1, .level = 0, .collapsed = 0};
  *
  *     worksheet_set_column_opt(worksheet, COLS("A:A"), 8.43, NULL, &options)
  * @endcode
  *
  */
//FUNCTION lxw_worksheet_set_column_opt(worksheet, first_col, last_col, width, format, lxw_row_col_options *options)
//RETURN CallDll( "worksheet_set_column_opt", worksheet, first_col, last_col, ToDouble(width), format, options )

/**
 * @brief Insert an image in a worksheet cell.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param filename  The image filename, with path if required.
 *
 * @return A #FUNCTION code.
 *
 * This function can be used to insert a image into a worksheet. The image can
 * be in PNG, JPEG or BMP format:
 *
 * @code
 *     worksheet_insert_image(worksheet, 2, 1, "logo.png")
 * @endcode
 *
 * @image html insert_image.png
 *
 * The `worksheet_insert_image_opt()` function takes additional optional
 * parameters to position and scale the image, see below.
 *
 * **Note**:
 * The scaling of a image may be affected if is crosses a row that has its
 * default height changed due to a font that is larger than the default font
 * size or that has text wrapping turned on. To aFUNCTION this you should
 * explicitly set the height of the row using `worksheet_set_row()` if it
 * crosses an inserted image.
 *
 * BMP images are only supported for backward compatibility. In general it is
 * best to aFUNCTION BMP images since they aren't compressed. If used, BMP images
 * must be 24 bit, true color, bitmaps.
 */
FUNCTION lxw_worksheet_insert_image(worksheet, row, col, filename)
RETURN CallDll( "worksheet_insert_image", worksheet, row, col, filename )
/**
 * @brief Insert an image in a worksheet cell, with options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param filename  The image filename, with path if required.
 * @param options   Optional image parameters.
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_insert_image_opt()` function is like
 * `worksheet_insert_image()` function except that it takes an optional
 * #lxw_image_options struct to scale and position the image:
 *
 * @code
 *    lxw_image_options options = {.x_offset = 30,  .y_offset = 10,
 *                                 .x_scale  = 0.5, .y_scale  = 0.5};
 *
 *    worksheet_insert_image_opt(worksheet, 2, 1, "logo.png", &options)
 *
 * @endcode
 *
 * @image html insert_image_opt.png
 *
 * @note See the notes about row scaling and BMP images in
 * `worksheet_insert_image()` above.
 */
//FUNCTION lxw_worksheet_insert_image_opt(worksheet, row, col, filename, lxw_image_options *options)
//RETURN CallDll( "worksheet_insert_image_opt", worksheet, row, col, filename, options )
/**
 * @brief Insert a chart object into a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param chart     A #lxw_chart object created via workbook_add_chart().
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_insert_chart()` can be used to insert a chart into a
 * worksheet. The chart object must be created first using the
 * `workbook_add_chart()` function and configured using the @ref chart.h
 * functions.
 *
 * @code
 *     // Create a chart object.
 *     lxw_chart *chart = workbook_add_chart(workbook, LXW_CHART_LINE)
 *
 *     // Add a data series to the chart.
 *     chart_add_series(chart, NULL, "=Sheet1!$A$1:$A$6")
 *
 *     // Insert the chart into the worksheet
 *     worksheet_insert_chart(worksheet, 0, 2, chart)
 * @endcode
 *
 * @image html chart_working.png
 *
 *
 * **Note:**
 *
 * A chart may only be inserted into a worksheet once. If several similar
 * charts are required then each one must be created separately with
 * `%worksheet_insert_chart()`.
 *
 */
FUNCTION lxw_worksheet_insert_chart(worksheet, row, col, chart)
RETURN CallDll( "worksheet_insert_chart", worksheet, row, col, chart )

/**
 * @brief Insert a chart object into a worksheet, with options.
 *
 * @param worksheet    Pointer to a lxw_worksheet instance to be updated.
 * @param row          The zero indexed row number.
 * @param col          The zero indexed column number.
 * @param chart        A #lxw_chart object created via workbook_add_chart().
 * @param user_options Optional chart parameters.
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_insert_chart_opt()` function is like
 * `worksheet_insert_chart()` function except that it takes an optional
 * #lxw_image_options struct to scale and position the image of the chart:
 *
 * @code
 *    lxw_image_options options = {.x_offset = 30,  .y_offset = 10,
 *                                 .x_scale  = 0.5, .y_scale  = 0.75};
 *
 *    worksheet_insert_chart_opt(worksheet, 0, 2, chart, &options)
 *
 * @endcode
 *
 * @image html chart_line_opt.png
 *
 * The #lxw_image_options struct is the same struct used in
 * `worksheet_insert_image_opt()` to position and scale images.
 *
 */
//FUNCTION lxw_worksheet_insert_chart_opt(worksheet, row, col, chart, lxw_image_options *user_options)
//RETURN CallDll( "worksheet_insert_chart_opt", worksheet, row, col, chart, user_options )

/**
 * @brief Merge a range of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 * @param string    String to write to the merged range.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_merge_range()` function allows cells to be merged together
 * so that they act as a single area.
 *
 * Excel generally merges and centers cells at same time. To get similar
 * behavior with libxlsxwriter you need to apply a @ref format.h "Format"
 * object with the appropriate alignment:
 *
 * @code
 *     merge_format = workbook_add_format(workbook)
 *     format_set_align(merge_format, LXW_ALIGN_CENTER)
 *
 *     worksheet_merge_range(worksheet, 1, 1, 1, 3, "Merged Range", merge_format)
 *
 * @endcode
 *
 * It is possible to apply other formatting to the merged cells as well:
 *
 * @code
 *    format_set_align   (merge_format, LXW_ALIGN_CENTER)
 *    format_set_align   (merge_format, LXW_ALIGN_VERTICAL_CENTER)
 *    format_set_border  (merge_format, LXW_BORDER_DOUBLE)
 *    format_set_bold    (merge_format)
 *    format_set_bg_color(merge_format, 0xD7E4BC)
 *
 *    worksheet_merge_range(worksheet, 2, 1, 3, 3, "Merged Range", merge_format)
 *
 * @endcode
 *
 * @image html merge.png
 *
 * The `%worksheet_merge_range()` function writes a `char*` string using
 * `worksheet_write_string()`. In order to write other data types, such as a
 * number or a formula, you can overwrite the first cell with a call to one of
 * the other write functions. The same Format should be used as was used in
 * the merged range.
 *
 * @code
 *    // First write a range with a blank string.
 *    worksheet_merge_range (worksheet, 1, 1, 1, 3, "", format)
 *
 *    // Then overwrite the first cell with a number.
 *    worksheet_write_number(worksheet, 1, 1, 123, format)
 * @endcode
 *
 * @note Merged ranges generally don’t work in libxlsxwriter when the Workbook
 * #lxw_workbook_options `constant_memory` mode is enabled.
 */
FUNCTION lxw_worksheet_merge_range(worksheet, first_row, first_col, last_row, last_col, string, format)
RETURN CallDll( "worksheet_merge_range", worksheet, first_row, first_col, last_row, last_col, string, format )

/**
 * @brief Set the autofilter area in the worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_autofilter()` function allows an autofilter to be added to
 * a worksheet.
 *
 * An autofilter is a way of adding drop down lists to the headers of a 2D
 * range of worksheet data. This allows users to filter the data based on
 * simple criteria so that some data is shown and some is hidden.
 *
 * @image html autofilter.png
 *
 * To add an autofilter to a worksheet:
 *
 * @code
 *     worksheet_autofilter(worksheet, 0, 0, 50, 3)
 *
 *     // Same as above using the RANGE() macro.
 *     worksheet_autofilter(worksheet, RANGE("A1:D51"))
 * @endcode
 *
 * Note: it isn't currently possible to apply filter conditions to the
 * autofilter.
 */
FUNCTION lxw_worksheet_autofilter(worksheet, first_row, first_col, last_row, last_col)
RETURN CallDll( "worksheet_autofilter", worksheet, first_row, first_col, last_row, last_col )

 /**
  * @brief Make a worksheet the active, i.e., visible worksheet.
  *
  * @param worksheet Pointer to a lxw_worksheet instance to be updated.
  *
  * The `%worksheet_activate()` function is used to specify which worksheet is
  * initially visible in a multi-sheet workbook:
  *
  * @code
  *     worksheet1 = workbook_add_worksheet(workbook, NULL)
  *     worksheet2 = workbook_add_worksheet(workbook, NULL)
  *     worksheet3 = workbook_add_worksheet(workbook, NULL)
  *
  *     worksheet_activate(worksheet3)
  * @endcode
  *
  * @image html worksheet_activate.png
  *
  * More than one worksheet can be selected via the `worksheet_select()`
  * function, see below, however only one worksheet can be active.
  *
  * The default active worksheet is the first worksheet.
  *
  */
FUNCTION lxw_worksheet_activate(worksheet)
RETURN CallDll( "worksheet_activate", worksheet )

 /**
  * @brief Set a worksheet tab as selected.
  *
  * @param worksheet Pointer to a lxw_worksheet instance to be updated.
  *
  * The `%worksheet_select()` function is used to indicate that a worksheet is
  * selected in a multi-sheet workbook:
  *
  * @code
  *     worksheet_activate(worksheet1)
  *     worksheet_select(worksheet2)
  *     worksheet_select(worksheet3)
  *
  * @endcode
  *
  * A selected worksheet has its tab highlighted. Selecting worksheets is a
  * way of grouping them together so that, for example, several worksheets
  * could be printed in one go. A worksheet that has been activated via the
  * `worksheet_activate()` function will also appear as selected.
  *
  */
FUNCTION lxw_worksheet_select(worksheet)
RETURN CallDll( "worksheet_select", worksheet )

/**
 * @brief Hide the current worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `%worksheet_hide()` function is used to hide a worksheet:
 *
 * @code
 *     worksheet_hide(worksheet2)
 * @endcode
 *
 * You may wish to hide a worksheet in order to aFUNCTION confusing a user with
 * intermediate data or calculations.
 *
 * @image html hide_sheet.png
 *
 * A hidden worksheet can not be activated or selected so this function is
 * mutually exclusive with the `worksheet_activate()` and `worksheet_select()`
 * functions. In addition, since the first worksheet will default to being the
 * active worksheet, you cannot hide the first worksheet without activating
 * another sheet:
 *
 * @code
 *     worksheet_activate(worksheet2)
 *     worksheet_hide(worksheet1)
 * @endcode
 */
FUNCTION lxw_worksheet_hide(worksheet)
RETURN CallDll( "worksheet_hide", worksheet )

/**
 * @brief Set current worksheet as the first visible sheet tab.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `worksheet_activate()` function determines which worksheet is initially
 * selected.  However, if there are a large number of worksheets the selected
 * worksheet may not appear on the screen. To aFUNCTION this you can select the
 * leftmost visible worksheet tab using `%worksheet_set_first_sheet()`:
 *
 * @code
 *     worksheet_set_first_sheet(worksheet19) // First visible worksheet tab.
 *     worksheet_activate(worksheet20)        // First visible worksheet.
 * @endcode
 *
 * This function is not required very often. The default value is the first
 * worksheet.
 */
FUNCTION lxw_worksheet_set_first_sheet(worksheet)
RETURN CallDll( "worksheet_set_first_sheet", worksheet )

/**
 * @brief Split and freeze a worksheet into panes.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The cell row (zero indexed).
 * @param col       The cell column (zero indexed).
 *
 * The `%worksheet_freeze_panes()` function can be used to divide a worksheet
 * into horizontal or vertical regions known as panes and to "freeze" these
 * panes so that the splitter bars are not visible.
 *
 * The parameters `row` and `col` are used to specify the location of the
 * split. It should be noted that the split is specified at the top or left of
 * a cell and that the function uses zero based indexing. Therefore to freeze
 * the first row of a worksheet it is necessary to specify the split at row 2
 * (which is 1 as the zero-based index).
 *
 * You can set one of the `row` and `col` parameters as zero if you do not
 * want either a vertical or horizontal split.
 *
 * Examples:
 *
 * @code
 *     worksheet_freeze_panes(worksheet1, 1, 0) // Freeze the first row.
 *     worksheet_freeze_panes(worksheet2, 0, 1) // Freeze the first column.
 *     worksheet_freeze_panes(worksheet3, 1, 1) // Freeze first row/column.
 *
 * @endcode
 *
 */
FUNCTION lxw_worksheet_freeze_panes(worksheet, row, col)
RETURN CallDll( "worksheet_freeze_panes", worksheet, row, col )

/**
 * @brief Split a worksheet into panes.
 *
 * @param worksheet  Pointer to a lxw_worksheet instance to be updated.
 * @param vertical   The position for the vertical split.
 * @param horizontal The position for the horizontal split.
 *
 * The `%worksheet_split_panes()` function can be used to divide a worksheet
 * into horizontal or vertical regions known as panes. This function is
 * different from the `worksheet_freeze_panes()` function in that the splits
 * between the panes will be visible to the user and each pane will have its
 * own scroll bars.
 *
 * The parameters `vertical` and `horizontal` are used to specify the vertical
 * and horizontal position of the split. The units for `vertical` and
 * `horizontal` are the same as those used by Excel to specify row height and
 * column width. However, the vertical and horizontal units are different from
 * each other. Therefore you must specify the `vertical` and `horizontal`
 * parameters in terms of the row heights and column widths that you have set
 * or the default values which are 15 for a row and 8.43 for a column.
 *
 * Examples:
 *
 * @code
 *     worksheet_split_panes(worksheet1, 15, 0)    // First row.
 *     worksheet_split_panes(worksheet2, 0,  8.43) // First column.
 *     worksheet_split_panes(worksheet3, 15, 8.43) // First row and column.
 *
 * @endcode
 *
 */
FUNCTION lxw_worksheet_split_panes(worksheet, vertical, horizontal)
RETURN CallDll( "worksheet_split_panes", worksheet, ToDouble(vertical), ToDouble(horizontal) )

/* worksheet_freeze_panes() with infrequent options. Undocumented for now. */
FUNCTION lxw_worksheet_freeze_panes_opt(worksheet, first_row, first_col, top_row, left_col, type)
RETURN CallDll( "worksheet_freeze_panes_opt", worksheet, first_row, first_col, top_row, left_col, type )

/* worksheet_split_panes() with infrequent options. Undocumented for now. */
FUNCTION lxw_worksheet_split_panes_opt(worksheet, vertical, horizontal, top_row, left_col)
RETURN CallDll( "worksheet_split_panes_opt", worksheet, ToDouble(vertical), ToDouble(horizontal), top_row, left_col )

/**
 * @brief Set the selected cell or cells in a worksheet:
 *
 * @param worksheet   Pointer to a lxw_worksheet instance to be updated.
 * @param first_row   The first row of the range. (All zero indexed.)
 * @param first_col   The first column of the range.
 * @param last_row    The last row of the range.
 * @param last_col    The last col of the range.
 *
 *
 * The `%worksheet_set_selection()` function can be used to specify which cell
 * or range of cells is selected in a worksheet: The most common requirement
 * is to select a single cell, in which case the `first_` and `last_`
 * parameters should be the same.
 *
 * The active cell within a selected range is determined by the order in which
 * `first_` and `last_` are specified.
 *
 * Examples:
 *
 * @code
 *     worksheet_set_selection(worksheet1, 3, 3, 3, 3)     // Cell D4.
 *     worksheet_set_selection(worksheet2, 3, 3, 6, 6)     // Cells D4 to G7.
 *     worksheet_set_selection(worksheet3, 6, 6, 3, 3)     // Cells G7 to D4.
 *     worksheet_set_selection(worksheet5, RANGE("D4:G7")) // Using the RANGE macro.
 *
 * @endcode
 *
 */
FUNCTION lxw_worksheet_set_selection(worksheet, first_row, first_col, last_row, last_col)
RETURN CallDll( "worksheet_set_selection", worksheet, first_row, first_col, last_row, last_col )

/**
 * @brief Set the page orientation as landscape.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to landscape:
 *
 * @code
 *     worksheet_set_landscape(worksheet)
 * @endcode
 */
FUNCTION lxw_worksheet_set_landscape(worksheet)
RETURN CallDll( "worksheet_set_landscape", worksheet )

/**
 * @brief Set the page orientation as portrait.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to portrait. The default worksheet orientation is portrait, so this
 * function isn't generally required:
 *
 * @code
 *     worksheet_set_portrait(worksheet)
 * @endcode
 */
FUNCTION lxw_worksheet_set_portrait(worksheet)
RETURN CallDll( "worksheet_set_portrait", worksheet )

/**
 * @brief Set the page layout to page view mode.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to display the worksheet in "Page View/Layout" mode:
 *
 * @code
 *     worksheet_set_page_view(worksheet)
 * @endcode
 */
FUNCTION lxw_worksheet_set_page_view(worksheet)
RETURN CallDll( "worksheet_set_page_view", worksheet )

/**
 * @brief Set the paper type for printing.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param paper_type The Excel paper format type.
 *
 * This function is used to set the paper format for the printed output of a
 * worksheet. The following paper styles are available:
 *
 *
 *   Index    | Paper format            | Paper size
 *   :------- | :---------------------- | :-------------------
 *   0        | Printer default         | Printer default
 *   1        | Letter                  | 8 1/2 x 11 in
 *   2        | Letter Small            | 8 1/2 x 11 in
 *   3        | Tabloid                 | 11 x 17 in
 *   4        | Ledger                  | 17 x 11 in
 *   5        | Legal                   | 8 1/2 x 14 in
 *   6        | Statement               | 5 1/2 x 8 1/2 in
 *   7        | Executive               | 7 1/4 x 10 1/2 in
 *   8        | A3                      | 297 x 420 mm
 *   9        | A4                      | 210 x 297 mm
 *   10       | A4 Small                | 210 x 297 mm
 *   11       | A5                      | 148 x 210 mm
 *   12       | B4                      | 250 x 354 mm
 *   13       | B5                      | 182 x 257 mm
 *   14       | Folio                   | 8 1/2 x 13 in
 *   15       | Quarto                  | 215 x 275 mm
 *   16       | ---                     | 10x14 in
 *   17       | ---                     | 11x17 in
 *   18       | Note                    | 8 1/2 x 11 in
 *   19       | Envelope 9              | 3 7/8 x 8 7/8
 *   20       | Envelope 10             | 4 1/8 x 9 1/2
 *   21       | Envelope 11             | 4 1/2 x 10 3/8
 *   22       | Envelope 12             | 4 3/4 x 11
 *   23       | Envelope 14             | 5 x 11 1/2
 *   24       | C size sheet            | ---
 *   25       | D size sheet            | ---
 *   26       | E size sheet            | ---
 *   27       | Envelope DL             | 110 x 220 mm
 *   28       | Envelope C3             | 324 x 458 mm
 *   29       | Envelope C4             | 229 x 324 mm
 *   30       | Envelope C5             | 162 x 229 mm
 *   31       | Envelope C6             | 114 x 162 mm
 *   32       | Envelope C65            | 114 x 229 mm
 *   33       | Envelope B4             | 250 x 353 mm
 *   34       | Envelope B5             | 176 x 250 mm
 *   35       | Envelope B6             | 176 x 125 mm
 *   36       | Envelope                | 110 x 230 mm
 *   37       | Monarch                 | 3.875 x 7.5 in
 *   38       | Envelope                | 3 5/8 x 6 1/2 in
 *   39       | Fanfold                 | 14 7/8 x 11 in
 *   40       | German Std Fanfold      | 8 1/2 x 12 in
 *   41       | German Legal Fanfold    | 8 1/2 x 13 in
 *
 * Note, it is likely that not all of these paper types will be available to
 * the end user since it will depend on the paper formats that the user's
 * printer supports. Therefore, it is best to stick to standard paper types:
 *
 * @code
 *     worksheet_set_paper(worksheet1, 1)  // US Letter
 *     worksheet_set_paper(worksheet2, 9)  // A4
 * @endcode
 *
 * If you do not specify a paper type the worksheet will print using the
 * printer's default paper style.
 */
FUNCTION lxw_worksheet_set_paper(worksheet, paper_type)
RETURN CallDll( "worksheet_set_paper", worksheet, paper_type )

/**
 * @brief Set the worksheet margins for the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param left    Left margin in inches.   Excel default is 0.7.
 * @param right   Right margin in inches.  Excel default is 0.7.
 * @param top     Top margin in inches.    Excel default is 0.75.
 * @param bottom  Bottom margin in inches. Excel default is 0.75.
 *
 * The `%worksheet_set_margins()` function is used to set the margins of the
 * worksheet when it is printed. The units are in inches. Specifying `-1` for
 * any parameter will give the default Excel value as shown above.
 *
 * @code
 *    worksheet_set_margins(worksheet, 1.3, 1.2, -1, -1)
 * @endcode
 *
 */
FUNCTION lxw_worksheet_set_margins(worksheet, left, right, top, bottom)
RETURN CallDll( "worksheet_set_margins", worksheet, ToDouble(left), ToDouble(right), ToDouble(top), ToDouble(bottom) )

/**
 * @brief Set the printed page header caption.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The header string.
 *
 * @return A #FUNCTION code.
 *
 * Headers and footers are generated using a string which is a combination of
 * plain text and control characters.
 *
 * The available control character are:
 *
 *
 *   | Control         | Category      | Description           |
 *   | --------------- | ------------- | --------------------- |
 *   | `&L`            | Justification | Left                  |
 *   | `&C`            |               | Center                |
 *   | `&R`            |               | Right                 |
 *   | `&P`            | Information   | Page number           |
 *   | `&N`            |               | Total number of pages |
 *   | `&D`            |               | Date                  |
 *   | `&T`            |               | Time                  |
 *   | `&F`            |               | File name             |
 *   | `&A`            |               | Worksheet name        |
 *   | `&Z`            |               | Workbook path         |
 *   | `&fontsize`     | Font          | Font size             |
 *   | `&"font,style"` |               | Font name and style   |
 *   | `&U`            |               | Single underline      |
 *   | `&E`            |               | underline      |
 *   | `&S`            |               | Strikethrough         |
 *   | `&X`            |               | Superscript           |
 *   | `&Y`            |               | Subscript             |
 *
 *
 * Text in headers and footers can be justified (aligned) to the left, center
 * and right by prefixing the text with the control characters `&L`, `&C` and
 * `&R`.
 *
 * For example (with ASCII art representation of the results):
 *
 * @code
 *     worksheet_set_header(worksheet, "&LHello")
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     | Hello                                                         |
 *     |                                                               |
 *
 *
 *     worksheet_set_header(worksheet, "&CHello")
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                          Hello                                |
 *     |                                                               |
 *
 *
 *     worksheet_set_header(worksheet, "&RHello")
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                                                         Hello |
 *     |                                                               |
 *
 *
 * @endcode
 *
 * For simple text, if you do not specify any justification the text will be
 * centered. However, you must prefix the text with `&C` if you specify a font
 * name or any other formatting:
 *
 * @code
 *     worksheet_set_header(worksheet, "Hello")
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                          Hello                                |
 *     |                                                               |
 *
 * @endcode
 *
 * You can have text in each of the justification regions:
 *
 * @code
 *     worksheet_set_header(worksheet, "&LCiao&CBello&RCielo")
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     | Ciao                     Bello                          Cielo |
 *     |                                                               |
 *
 * @endcode
 *
 * The information control characters act as variables that Excel will update
 * as the workbook or worksheet changes. Times and dates are in the users
 * default format:
 *
 * @code
 *     worksheet_set_header(worksheet, "&CPage &P of &N")
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                        Page 1 of 6                            |
 *     |                                                               |
 *
 *     worksheet_set_header(worksheet, "&CUpdated at &T")
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                    Updated at 12:30 PM                        |
 *     |                                                               |
 *
 * @endcode
 *
 * You can specify the font size of a section of the text by prefixing it with
 * the control character `&n` where `n` is the font size:
 *
 * @code
 *     worksheet_set_header(worksheet1, "&C&30Hello Big")
 *     worksheet_set_header(worksheet2, "&C&10Hello Small")
 *
 * @endcode
 *
 * You can specify the font of a section of the text by prefixing it with the
 * control sequence `&"font,style"` where `fontname` is a font name such as
 * Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
 * "Courier New" or "Times New Roman" and `style` is one of the standard
 *
 * @code
 *     worksheet_set_header(worksheet1, "&C&\"Courier New,Italic\"Hello")
 *     worksheet_set_header(worksheet2, "&C&\"Courier New,Bold Italic\"Hello")
 *     worksheet_set_header(worksheet3, "&C&\"Times New Roman,Regular\"Hello")
 *
 * @endcode
 *
 * It is possible to combine all of these features together to create
 * sophisticated headers and footers. As an aid to setting up complicated
 * headers and footers you can record a page set-up as a macro in Excel and
 * look at the format strings that VBA produces. Remember however that VBA
 * uses two quotes `""` to indicate a single quote. For the last
 * example above the equivalent VBA code looks like this:
 *
 * @code
 *     .LeftHeader = ""
 *     .CenterHeader = "&""Times New Roman,Regular""Hello"
 *     .RightHeader = ""
 *
 * @endcode
 *
 * Alternatively you can inspect the header and footer strings in an Excel
 * file by unzipping it and grepping the XML sub-files. The following shows
 * how to do that using libxml's xmllint to format the XML for clarity:
 *
 * @code
 *
 *    $ unzip myfile.xlsm -d myfile
 *    $ xmllint --format `find myfile -name "*.xml" | xargs` | egrep "Header|Footer"
 *
 *      <headerFooter scaleWithDoc="0">
 *        <oddHeader>&amp;L&amp;P</oddHeader>
 *      </headerFooter>
 *
 * @endcode
 *
 * Note that in this case you need to unescape the Html. In the above example
 * the header string would be `&L&P`.
 *
 * To include a single literal ampersand `&` in a header or footer you should
 * use a ampersand `&&`:
 *
 * @code
 *     worksheet_set_header(worksheet, "&CCuriouser && Curiouser - Attorneys at Law")
 * @endcode
 *
 * Note, the header or footer string must be less than 255 characters. Strings
 * longer than this will not be written.
 *
 */
FUNCTION lxw_worksheet_set_header(worksheet, string)
RETURN CallDll( "worksheet_set_header", worksheet, string )

/**
 * @brief Set the printed page footer caption.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The footer string.
 *
 * @return A #FUNCTION code.
 *
 * The syntax of this function is the same as worksheet_set_header().
 *
 */
FUNCTION lxw_worksheet_set_footer(worksheet, string)
RETURN CallDll( "worksheet_set_footer", worksheet, string )

/**
 * @brief Set the printed page header caption with additional options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The header string.
 * @param options   Header options.
 *
 * @return A #FUNCTION code.
 *
 * The syntax of this function is the same as worksheet_set_header() with an
 * additional parameter to specify options for the header.
 *
 * Currently, the only available option is the header margin:
 *
 * @code
 *
 *    lxw_header_footer_options header_options = { 0.2 };
 *
 *    worksheet_set_header_opt(worksheet, "Some text", &header_options)
 *
 * @endcode
 *
 */
//FUNCTION lxw_worksheet_set_header_opt(worksheet, string, lxw_header_footer_options *options)
//RETURN CallDll( "worksheet_set_header_opt", worksheet, string, options )

/**
 * @brief Set the printed page footer caption with additional options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The footer string.
 * @param options   Footer options.
 *
 * @return A #FUNCTION code.
 *
 * The syntax of this function is the same as worksheet_set_header_opt().
 *
 */
//FUNCTION lxw_worksheet_set_footer_opt(worksheet, string, lxw_header_footer_options *options)
//RETURN CallDll( "worksheet_set_footer_opt", worksheet, string, options )

/**
 * @brief Set the horizontal page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_set_h_pagebreaks()` function adds horizontal page breaks to
 * a worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Horizontal page breaks act between rows.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref and the last element of the array must be 0:
 *
 * @code
 *    breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    breaks2[] = {20, 40, 60, 80, 0};
 *
 *    worksheet_set_h_pagebreaks(worksheet1, breaks1)
 *    worksheet_set_h_pagebreaks(worksheet2, breaks2)
 * @endcode
 *
 * To create a page break between rows 20 and 21 you must specify the break at
 * row 21. However in zero index notation this is actually row 20:
 *
 * @code
 *    // Break between row 20 and 21.
 *    breaks[] = {20, 0};
 *
 *    worksheet_set_h_pagebreaks(worksheet, breaks)
 * @endcode
 *
 * There is an Excel limitation of 1023 horizontal page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
FUNCTION lxw_worksheet_set_h_pagebreaks(worksheet, breaks )
RETURN CallDll( "worksheet_set_h_pagebreaks", worksheet, breaks )

/**
 * @brief Set the vertical page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * @return A #FUNCTION code.
 *
 * The `%worksheet_set_v_pagebreaks()` function adds vertical page breaks to a
 * worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Vertical page breaks act between columns.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref and the last element of the array must be 0:
 *
 * @code
 *    breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    breaks2[] = {20, 40, 60, 80, 0};
 *
 *    worksheet_set_v_pagebreaks(worksheet1, breaks1)
 *    worksheet_set_v_pagebreaks(worksheet2, breaks2)
 * @endcode
 *
 * To create a page break between columns 20 and 21 you must specify the break
 * at column 21. However in zero index notation this is actually column 20:
 *
 * @code
 *    // Break between column 20 and 21.
 *    breaks[] = {20, 0};
 *
 *    worksheet_set_v_pagebreaks(worksheet, breaks)
 * @endcode
 *
 * There is an Excel limitation of 1023 vertical page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
FUNCTION lxw_worksheet_set_v_pagebreaks(worksheet, breaks)
RETURN CallDll( "worksheet_set_v_pagebreaks", worksheet, breaks )

/**
 * @brief Set the order in which pages are printed.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `%worksheet_print_across()` function is used to change the default
 * print direction. This is referred to by Excel as the sheet "page order":
 *
 * @code
 *     worksheet_print_across(worksheet)
 * @endcode
 *
 * The default page order is shown below for a worksheet that extends over 4
 * pages. The order is called "down then across":
 *
 *     [1] [3]
 *     [2] [4]
 *
 * However, by using the `print_across` function the print order will be
 * changed to "across then down":
 *
 *     [1] [2]
 *     [3] [4]
 *
 */
FUNCTION lxw_worksheet_print_across(worksheet)
RETURN CallDll( "worksheet_print_across", worksheet )

/**
 * @brief Set the worksheet zoom factor.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param scale     Worksheet zoom factor.
 *
 * Set the worksheet zoom factor in the range `10 <= zoom <= 400`:
 *
 * @code
 *     worksheet_set_zoom(worksheet1, 50)
 *     worksheet_set_zoom(worksheet2, 75)
 *     worksheet_set_zoom(worksheet3, 300)
 *     worksheet_set_zoom(worksheet4, 400)
 * @endcode
 *
 * The default zoom factor is 100. It isn't possible to set the zoom to
 * "Selection" because it is calculated by Excel at run-time.
 *
 * Note, `%worksheet_zoom()` does not affect the scale of the printed
 * page. For that you should use `worksheet_set_print_scale()`.
 */
FUNCTION lxw_worksheet_set_zoom(worksheet, scale)
RETURN CallDll( "worksheet_set_zoom", worksheet, scale )

/**
 * @brief Set the option to display or hide gridlines on the screen and
 *        the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param option    Gridline option.
 *
 * Display or hide screen and print gridlines using one of the values of
 * @ref lxw_gridlines.
 *
 * @code
 *    worksheet_gridlines(worksheet1, LXW_HIDE_ALL_GRIDLINES)
 *
 *    worksheet_gridlines(worksheet2, LXW_SHOW_PRINT_GRIDLINES)
 * @endcode
 *
 * The Excel default is that the screen gridlines are on  and the printed
 * worksheet is off.
 *
 */
FUNCTION lxw_worksheet_gridlines(worksheet, option)
RETURN CallDll( "worksheet_gridlines", worksheet, option )

/**
 * @brief Center the printed page horizontally.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * Center the worksheet data horizontally between the margins on the printed
 * page:
 *
 * @code
 *     worksheet_center_horizontally(worksheet)
 * @endcode
 *
 */
FUNCTION lxw_worksheet_center_horizontally(worksheet)
RETURN CallDll( "worksheet_center_horizontally", worksheet )

/**
 * @brief Center the printed page vertically.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * Center the worksheet data vertically between the margins on the printed
 * page:
 *
 * @code
 *     worksheet_center_vertically(worksheet)
 * @endcode
 *
 */
FUNCTION lxw_worksheet_center_vertically(worksheet)
RETURN CallDll( "worksheet_center_vertically", worksheet )

/**
 * @brief Set the option to print the row and column headers on the printed
 *        page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * When printing a worksheet from Excel the row and column headers (the row
 * numbers on the left and the column letters at the top) aren't printed by
 * default.
 *
 * This function sets the printer option to print these headers:
 *
 * @code
 *    worksheet_print_row_col_headers(worksheet)
 * @endcode
 *
 */
FUNCTION lxw_worksheet_print_row_col_headers(worksheet)
RETURN CallDll( "worksheet_print_row_col_headers", worksheet )

/**
 * @brief Set the number of rows to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row First row of repeat range.
 * @param last_row  Last row of repeat range.
 *
 * @return A #FUNCTION code.
 *
 * For large Excel documents it is often desirable to have the first row or
 * rows of the worksheet print out at the top of each page.
 *
 * This can be achieved by using this function. The parameters `first_row`
 * and `last_row` are zero based:
 *
 * @code
 *     worksheet_repeat_rows(worksheet, 0, 0) // Repeat the first row.
 *     worksheet_repeat_rows(worksheet, 0, 1) // Repeat the first two rows.
 * @endcode
 */
FUNCTION lxw_worksheet_repeat_rows(worksheet, first_row, last_row)
RETURN CallDll( "worksheet_repeat_rows", worksheet, first_row, last_row )

/**
 * @brief Set the number of columns to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_col First column of repeat range.
 * @param last_col  Last column of repeat range.
 *
 * @return A #FUNCTION code.
 *
 * For large Excel documents it is often desirable to have the first column or
 * columns of the worksheet print out at the left of each page.
 *
 * This can be achieved by using this function. The parameters `first_col`
 * and `last_col` are zero based:
 *
 * @code
 *     worksheet_repeat_columns(worksheet, 0, 0) // Repeat the first col.
 *     worksheet_repeat_columns(worksheet, 0, 1) // Repeat the first two cols.
 * @endcode
 */
FUNCTION lxw_worksheet_repeat_columns(worksheet, first_col, last_col)
RETURN CallDll( "worksheet_repeat_columns", worksheet, first_col, last_col )

/**
 * @brief Set the print area for a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * @return A #FUNCTION code.
 *
 * This function is used to specify the area of the worksheet that will be
 * printed. The RANGE() macro is often convenient for this.
 *
 * @code
 *     worksheet_print_area(worksheet, 0, 0, 41, 10) // A1:K42.
 *
 *     // Same as:
 *     worksheet_print_area(worksheet, RANGE("A1:K42"))
 * @endcode
 *
 * In order to set a row or column range you must specify the entire range:
 *
 * @code
 *     worksheet_print_area(worksheet, RANGE("A1:H1048576")) // Same as A:H.
 * @endcode
 */
FUNCTION lxw_worksheet_print_area(worksheet, first_row, first_col, last_row, last_col)
RETURN CallDll( "worksheet_print_area", worksheet, first_row, first_col, last_row, last_col )

/**
 * @brief Fit the printed area to a specific number of pages both vertically
 *        and horizontally.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param width     Number of pages horizontally.
 * @param height    Number of pages vertically.
 *
 * The `%worksheet_fit_to_pages()` function is used to fit the printed area to
 * a specific number of pages both vertically and horizontally. If the printed
 * area exceeds the specified number of pages it will be scaled down to
 * fit. This ensures that the printed area will always appear on the specified
 * number of pages even if the page size or margins change:
 *
 * @code
 *     worksheet_fit_to_pages(worksheet1, 1, 1) // Fit to 1x1 pages.
 *     worksheet_fit_to_pages(worksheet2, 2, 1) // Fit to 2x1 pages.
 *     worksheet_fit_to_pages(worksheet3, 1, 2) // Fit to 1x2 pages.
 * @endcode
 *
 * The print area can be defined using the `worksheet_print_area()` function
 * as described above.
 *
 * A common requirement is to fit the printed output to `n` pages wide but
 * have the height be as long as necessary. To achieve this set the `height`
 * to zero:
 *
 * @code
 *     // 1 page wide and as long as necessary.
 *     worksheet_fit_to_pages(worksheet, 1, 0)
 * @endcode
 *
 * **Note**:
 *
 * - Although it is valid to use both `%worksheet_fit_to_pages()` and
 *   `worksheet_set_print_scale()` on the same worksheet Excel only allows one
 *   of these options to be active at a time. The last function call made will
 *   set the active option.
 *
 * - The `%worksheet_fit_to_pages()` function will override any manual page
 *   breaks that are defined in the worksheet.
 *
 * - When using `%worksheet_fit_to_pages()` it may also be required to set the
 *   printer paper size using `worksheet_set_paper()` or else Excel will
 *   default to "US Letter".
 *
 */
FUNCTION lxw_worksheet_fit_to_pages(worksheet, width, height)
RETURN CallDll( "worksheet_fit_to_pages", worksheet, width, height )

/**
 * @brief Set the start page number when printing.
 *
 * @param worksheet  Pointer to a lxw_worksheet instance to be updated.
 * @param start_page Starting page number.
 *
 * The `%worksheet_set_start_page()` function is used to set the number of
 * the starting page when the worksheet is printed out:
 *
 * @code
 *     // Start print from page 2.
 *     worksheet_set_start_page(worksheet, 2)
 * @endcode
 */
FUNCTION lxw_worksheet_set_start_page(worksheet, start_page)
RETURN CallDll( "worksheet_set_start_page", worksheet, start_page )

/**
 * @brief Set the scale factor for the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param scale     Print scale of worksheet to be printed.
 *
 * This function sets the scale factor of the printed page. The Scale factor
 * must be in the range `10 <= scale <= 400`:
 *
 * @code
 *     worksheet_set_print_scale(worksheet1, 75)
 *     worksheet_set_print_scale(worksheet2, 400)
 * @endcode
 *
 * The default scale factor is 100. Note, `%worksheet_set_print_scale()` does
 * not affect the scale of the visible page in Excel. For that you should use
 * `worksheet_set_zoom()`.
 *
 * Note that although it is valid to use both `worksheet_fit_to_pages()` and
 * `%worksheet_set_print_scale()` on the same worksheet Excel only allows one
 * of these options to be active at a time. The last function call made will
 * set the active option.
 *
 */
FUNCTION lxw_worksheet_set_print_scale(worksheet, scale)
RETURN CallDll( "worksheet_set_print_scale", worksheet, scale )

/**
 * @brief Display the worksheet cells from right to left for some versions of
 *        Excel.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
  * The `%worksheet_right_to_left()` function is used to change the default
 * direction of the worksheet from left-to-right, with the `A1` cell in the
 * top left, to right-to-left, with the `A1` cell in the top right.
 *
 * @code
 *     worksheet_right_to_left(worksheet1)
 * @endcode
 *
 * This is useful when creating Arabic, Hebrew or other near or far eastern
 * worksheets that use right-to-left as the default direction.
 */
FUNCTION lxw_worksheet_right_to_left(worksheet)
RETURN CallDll( "worksheet_right_to_left", worksheet )

/**
 * @brief Hide zero values in worksheet cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `%worksheet_hide_zero()` function is used to hide any zero values that
 * appear in cells:
 *
 * @code
 *     worksheet_hide_zero(worksheet1)
 * @endcode
 */
FUNCTION lxw_worksheet_hide_zero(worksheet)
RETURN CallDll( "worksheet_hide_zero", worksheet )

/**
 * @brief Set the color of the worksheet tab.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param color     The tab color.
 *
 * The `%worksheet_set_tab_color()` function is used to change the color of the worksheet
 * tab:
 *
 * @code
 *      worksheet_set_tab_color(worksheet1, LXW_COLOR_RED)
 *      worksheet_set_tab_color(worksheet2, LXW_COLOR_GREEN)
 *      worksheet_set_tab_color(worksheet3, 0xFF9900) // Orange.
 * @endcode
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 */
FUNCTION lxw_worksheet_set_tab_color(worksheet, color)
RETURN CallDll( "worksheet_set_tab_color", worksheet, color )

/**
 * @brief Protect elements of a worksheet from modification.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param password  A worksheet password.
 * @param options   Worksheet elements to protect.
 *
 * The `%worksheet_protect()` function protects worksheet elements from modification:
 *
 * @code
 *     worksheet_protect(worksheet, "Some Password", options)
 * @endcode
 *
 * The `password` and lxw_protection pointer are both optional:
 *
 * @code
 *     worksheet_protect(worksheet1, NULL,       NULL)
 *     worksheet_protect(worksheet2, NULL,       my_options)
 *     worksheet_protect(worksheet3, "password", NULL)
 *     worksheet_protect(worksheet4, "password", my_options)
 * @endcode
 *
 * Passing a `NULL` password is the same as turning on protection without a
 * password. Passing a `NULL` password and `NULL` options, or any other
 * combination has the effect of enabling a cell's `locked` and `hidden`
 * properties if they have been set.
 *
 * A *locked* cell cannot be edited and this property is on by default for all
 * cells. A *hidden* cell will display the results of a formula but not the
 * formula itself. These properties can be set using the format_set_unlocked()
 * and format_set_hidden() format functions.
 *
 * You can specify which worksheet elements you wish to protect by passing a
 * lxw_protection pointer in the `options` argument with any or all of the
 * following members set:
 *
 *     no_select_locked_cells
 *     no_select_unlocked_cells
 *     format_cells
 *     format_columns
 *     format_rows
 *     insert_columns
 *     insert_rows
 *     insert_hyperlinks
 *     delete_columns
 *     delete_rows
 *     sort
 *     autofilter
 *     pivot_tables
 *     scenarios
 *     objects
 *
 * All parameters are off by default. Individual elements can be protected as
 * follows:
 *
 * @code
 *     lxw_protection options = {
 *         .format_cells             = 1,
 *         .insert_hyperlinks        = 1,
 *         .insert_rows              = 1,
 *         .delete_rows              = 1,
 *         .insert_columns           = 1,
 *         .delete_columns           = 1,
 *     };
 *
 *     worksheet_protect(worksheet, NULL, &options)
 *
 * @endcode
 *
 * See also the format_set_unlocked() and format_set_hidden() format functions.
 *
 * **Note:** Worksheet level passwords in Excel offer **very** weak
 * protection. They don't encrypt your data and are very easy to
 * deactivate. Full workbook encryption is not supported by `libxlsxwriter`
 * since it requires a completely different file format and would take several
 * man months to implement.
 */
//FUNCTION lxw_worksheet_protect(worksheet, password, lxw_protection *options)
FUNCTION lxw_worksheet_protect(worksheet, password, options)
RETURN CallDll( "worksheet_protect", worksheet, password, options )

/**
 * @brief Set the default row properties.
 *
 * @param worksheet        Pointer to a lxw_worksheet instance to be updated.
 * @param height           Default row height.
 * @param hide_unused_rows Hide unused cells.
 *
 * The `%worksheet_set_default_row()` function is used to set Excel default
 * row properties such as the default height and the option to hide unused
 * rows. These parameters are an optimization used by Excel to set row
 * properties without generating a very large file with an entry for each row.
 *
 * To set the default row height:
 *
 * @code
 *     worksheet_set_default_row(worksheet, 24, LXW_FALSE)
 *
 * @endcode
 *
 * To hide unused rows:
 *
 * @code
 *     worksheet_set_default_row(worksheet, 15, LXW_TRUE)
 * @endcode
 *
 * Note, in the previous case we use the default height #LXW_DEF_ROW_HEIGHT =
 * 15 so the the height remains unchanged.
 */
FUNCTION lxw_worksheet_set_default_row(worksheet, height, hide_unused_rows)
RETURN CallDll( "worksheet_set_default_row", worksheet, ToDouble(height), hide_unused_rows )


// FORMAT =============================================================================================================

/**
 * @brief Set the font used in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param font_name Cell font name.
 *
 * Specify the font used used in the cell format:
 *
 * @code
 *     format_set_font_name(format, "Avenir Black Oblique")
 * @endcode
 *
 * @image html format_set_font_name.png
 *
 * Excel can only display fonts that are installed on the system that it is
 * running on. Therefore it is generally best to use the fonts that come as
 * standard with Excel such as Calibri, Times New Roman and Courier New.
 *
 * The default font in Excel 2007, and later, is Calibri.
 */
FUNCTION lxw_format_set_font_name(format, font_name)
RETURN CallDll( "format_set_font_name", format, font_name )

/**
 * @brief Set the size of the font used in the cell.
 *
 * @param format Pointer to a Format instance.
 * @param size   The cell font size.
 *
 * Set the font size of the cell format:
 *
 * @code
 *     format_set_font_size(format, 30)
 * @endcode
 *
 * @image html format_font_size.png
 *
 * Excel adjusts the height of a row to accommodate the largest font
 * size in the row. You can also explicitly specify the height of a
 * row using the worksheet_set_row() function.
 */
FUNCTION lxw_format_set_font_size(format, size)
RETURN CallDll( "format_set_font_size", format, size )

/**
 * @brief Set the color of the font used in the cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell font color.
 *
 *
 * Set the font color:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_font_color(format, LXW_COLOR_RED)
 *
 *     worksheet_write_string(worksheet, 0, 0, "Wheelbarrow", format)
 * @endcode
 *
 * @image html format_font_color.png
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 * @note
 * The format_set_font_color() method is used to set the font color in a
 * cell. To set the color of a cell background use the format_set_bg_color()
 * and format_set_pattern() methods.
 */
FUNCTION lxw_format_set_font_color(format, color)
RETURN CallDll( "format_set_font_color", format, color )

/**
 * @brief Turn on bold for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the bold property of the font:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_bold(format)
 *
 *     worksheet_write_string(worksheet, 0, 0, "Bold Text", format)
 * @endcode
 *
 * @image html format_font_bold.png
 */
FUNCTION lxw_format_set_bold(format)
RETURN CallDll( "format_set_bold", format )

/**
 * @brief Turn on italic for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the italic property of the font:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_italic(format)
 *
 *     worksheet_write_string(worksheet, 0, 0, "Italic Text", format)
 * @endcode
 *
 * @image html format_font_italic.png
 */
FUNCTION lxw_format_set_italic(format)
RETURN CallDll( "format_set_italic", format )

/**
 * @brief Turn on underline for the format:
 *
 * @param format Pointer to a Format instance.
 * @param style Underline style.
 *
 * Set the underline property of the format:
 *
 * @code
 *     format_set_underline(format, LXW_UNDERLINE_SINGLE)
 * @endcode
 *
 * @image html format_font_underlined.png
 *
 * The available underline styles are:
 *
 * - #LXW_UNDERLINE_SINGLE
 * - #LXW_UNDERLINE_DOUBLE
 * - #LXW_UNDERLINE_SINGLE_ACCOUNTING
 * - #LXW_UNDERLINE_DOUBLE_ACCOUNTING
 *
 */
FUNCTION lxw_format_set_underline(format, style)
RETURN CallDll( "format_set_underline", format, style )

/**
 * @brief Set the strikeout property of the font.
 *
 * @param format Pointer to a Format instance.
 *
 * @image html format_font_strikeout.png
 *
 */
FUNCTION lxw_format_set_font_strikeout(format)
RETURN CallDll( "format_set_font_strikeout", format )

/**
 * @brief Set the superscript/subscript property of the font.
 *
 * @param format Pointer to a Format instance.
 * @param style  Superscript or subscript style.
 *
 * Set the superscript o subscript property of the font.
 *
 * @image html format_font_script.png
 *
 * The available script styles are:
 *
 * - #LXW_FONT_SUPERSCRIPT
 * - #LXW_FONT_SUBSCRIPT
 */
FUNCTION lxw_format_set_font_script(format, style)
RETURN CallDll( "format_set_font_script", format, style )

/**
 * @brief Set the number format for a cell.
 *
 * @param format      Pointer to a Format instance.
 * @param num_format The cell number format string.
 *
 * This method is used to define the numerical format of a number in
 * Excel. It controls whether a number is displayed as an integer, a
 * floating point number, a date, a currency value or some other user
 * defined format.
 *
 * The numerical format of a cell can be specified by using a format
 * string:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_num_format(format, "d mmm yyyy")
 * @endcode
 *
 * Format strings can control any aspect of number formatting allowed by Excel:
 *
 * @dontinclude format_num_format.c
 * @skipline set_num_format
 * @until 1209
 *
 * @image html format_set_num_format.png
 *
 * The number system used for dates is described in @ref working_with_dates.
 *
 * For more information on number formats in Excel refer to the
 * [Microsoft documentation on cell formats](http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx).
 */
FUNCTION lxw_format_set_num_format(format, num_format)
RETURN CallDll( "format_set_num_format", format, num_format )

/**
 * @brief Set the Excel built-in number format for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param index  The built-in number format index for the cell.
 *
 * This function is similar to format_set_num_format() except that it takes an
 * index to a limited number of Excel's built-in number formats instead of a
 * user defined format string:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_num_format(format, 0x0F)     // d-mmm-yy
 * @endcode
 *
 * @note
 * Unless you need to specifically access one of Excel's built-in number
 * formats the format_set_num_format() function above is a better
 * solution. The format_set_num_format_index() function is mainly included for
 * backward compatibility and completeness.
 *
 * The Excel built-in number formats as shown in the table below:
 *
 *   | Index | Index | Format String                                        |
 *   | ----- | ----- | ---------------------------------------------------- |
 *   | 0     | 0x00  | `General`                                            |
 *   | 1     | 0x01  | `0`                                                  |
 *   | 2     | 0x02  | `0.00`                                               |
 *   | 3     | 0x03  | `#,##0`                                              |
 *   | 4     | 0x04  | `#,##0.00`                                           |
 *   | 5     | 0x05  | `($#,##0_)($#,##0)`                                 |
 *   | 6     | 0x06  | `($#,##0_)[Red]($#,##0)`                            |
 *   | 7     | 0x07  | `($#,##0.00_)($#,##0.00)`                           |
 *   | 8     | 0x08  | `($#,##0.00_)[Red]($#,##0.00)`                      |
 *   | 9     | 0x09  | `0%`                                                 |
 *   | 10    | 0x0a  | `0.00%`                                              |
 *   | 11    | 0x0b  | `0.00E+00`                                           |
 *   | 12    | 0x0c  | `# ?/?`                                              |
 *   | 13    | 0x0d  | `# ??/??`                                            |
 *   | 14    | 0x0e  | `m/d/yy`                                             |
 *   | 15    | 0x0f  | `d-mmm-yy`                                           |
 *   | 16    | 0x10  | `d-mmm`                                              |
 *   | 17    | 0x11  | `mmm-yy`                                             |
 *   | 18    | 0x12  | `h:mm AM/PM`                                         |
 *   | 19    | 0x13  | `h:mm:ss AM/PM`                                      |
 *   | 20    | 0x14  | `h:mm`                                               |
 *   | 21    | 0x15  | `h:mm:ss`                                            |
 *   | 22    | 0x16  | `m/d/yy h:mm`                                        |
 *   | ...   | ...   | ...                                                  |
 *   | 37    | 0x25  | `(#,##0_)(#,##0)`                                   |
 *   | 38    | 0x26  | `(#,##0_)[Red](#,##0)`                              |
 *   | 39    | 0x27  | `(#,##0.00_)(#,##0.00)`                             |
 *   | 40    | 0x28  | `(#,##0.00_)[Red](#,##0.00)`                        |
 *   | 41    | 0x29  | `_(* #,##0_)_(* (#,##0)_(* "-"_)_(@_)`            |
 *   | 42    | 0x2a  | `_($* #,##0_)_($* (#,##0)_($* "-"_)_(@_)`         |
 *   | 43    | 0x2b  | `_(* #,##0.00_)_(* (#,##0.00)_(* "-"??_)_(@_)`    |
 *   | 44    | 0x2c  | `_($* #,##0.00_)_($* (#,##0.00)_($* "-"??_)_(@_)` |
 *   | 45    | 0x2d  | `mm:ss`                                              |
 *   | 46    | 0x2e  | `[h]:mm:ss`                                          |
 *   | 47    | 0x2f  | `mm:ss.0`                                            |
 *   | 48    | 0x30  | `##0.0E+0`                                           |
 *   | 49    | 0x31  | `@`                                                  |
 *
 * @note
 *  -  Numeric formats 23 to 36 are not documented by Microsoft and may differ
 *     in international versions. The listed date and currency formats may also
 *     vary depending on system settings.
 *  - The dollar sign in the above format appears as the defined local currency
 *    symbol.
 *  - These formats can also be set via format_set_num_format().
 */
FUNCTION lxw_format_set_num_format_index(format, index)
RETURN CallDll( "format_set_num_format_index", format, index )

/**
 * @brief Set the cell unlocked state.
 *
 * @param format Pointer to a Format instance.
 *
 * This property can be used to allow modification of a cell in a protected
 * worksheet. In Excel, cell locking is turned on by default for all
 * cells. However, it only has an effect if the worksheet has been protected
 * using the worksheet worksheet_protect() function:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_unlocked(format)
 *
 *     // Enable worksheet protection, without password or options.
 *     worksheet_protect(worksheet, NULL, NULL)
 *
 *     // This cell cannot be edited.
 *     worksheet_write_formula(worksheet, 0, 0, "=1+2", NULL)
 *
 *     // This cell can be edited.
 *     worksheet_write_formula(worksheet, 1, 0, "=1+2", format)
 * @endcode
 */
FUNCTION lxw_format_set_unlocked(format)
RETURN CallDll( "format_set_unlocked", format )

/**
 * @brief Hide formulas in a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This property is used to hide a formula while still displaying its
 * result. This is generally used to hide complex calculations from end users
 * who are only interested in the result. It only has an effect if the
 * worksheet has been protected using the worksheet worksheet_protect()
 * function:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_hidden(format)
 *
 *     // Enable worksheet protection, without password or options.
 *     worksheet_protect(worksheet, NULL, NULL)
 *
 *     // The formula in this cell isn't visible.
 *     worksheet_write_formula(worksheet, 0, 0, "=1+2", format)
 * @endcode
 */
FUNCTION lxw_format_set_hidden(format)
RETURN CallDll( "format_set_hidden", format )

/**
 * @brief Set the alignment for data in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param alignment The horizontal and or vertical alignment direction.
 *
 * This method is used to set the horizontal and vertical text alignment within a
 * cell. The following are the available horizontal alignments:
 *
 * - #LXW_ALIGN_LEFT
 * - #LXW_ALIGN_CENTER
 * - #LXW_ALIGN_RIGHT
 * - #LXW_ALIGN_FILL
 * - #LXW_ALIGN_JUSTIFY
 * - #LXW_ALIGN_CENTER_ACROSS
 * - #LXW_ALIGN_DISTRIBUTED
 *
 * The following are the available vertical alignments:
 *
 * - #LXW_ALIGN_VERTICAL_TOP
 * - #LXW_ALIGN_VERTICAL_BOTTOM
 * - #LXW_ALIGN_VERTICAL_CENTER
 * - #LXW_ALIGN_VERTICAL_JUSTIFY
 * - #LXW_ALIGN_VERTICAL_DISTRIBUTED
 *
 * As in Excel, vertical and horizontal alignments can be combined:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *
 *     format_set_align(format, LXW_ALIGN_CENTER)
 *     format_set_align(format, LXW_ALIGN_VERTICAL_CENTER)
 *
 *     worksheet_set_row(0, 30)
 *     worksheet_write_string(worksheet, 0, 0, "Some Text", format)
 * @endcode
 *
 * @image html format_font_align.png
 *
 * Text can be aligned across two or more adjacent cells using the
 * center_across property. However, for genuine merged cells it is better to
 * use the worksheet_merge_range() worksheet method.
 *
 * The vertical justify option can be used to provide automatic text wrapping
 * in a cell. The height of the cell will be adjusted to accommodate the
 * wrapped text. To specify where the text wraps use the
 * format_set_text_wrap() method.
 */
FUNCTION lxw_format_set_align(format, alignment)
RETURN CallDll( "format_set_align", format, alignment )

/**
 * @brief Wrap text in a cell.
 *
 * Turn text wrapping on for text in a cell.
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_text_wrap(format)
 *
 *     worksheet_write_string(worksheet, 0, 0, "Some long text to wrap in a cell", format)
 * @endcode
 *
 * If you wish to control where the text is wrapped you can add newline characters
 * to the string:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_text_wrap(format)
 *
 *     worksheet_write_string(worksheet, 0, 0, "It's\na bum\nwrap", format)
 * @endcode
 *
 * @image html format_font_text_wrap.png
 *
 * Excel will adjust the height of the row to accommodate the wrapped text. A
 * similar effect can be obtained without newlines using the
 * format_set_align() function with #LXW_ALIGN_VERTICAL_JUSTIFY.
 */
FUNCTION lxw_format_set_text_wrap(format)
RETURN CallDll( "format_set_text_wrap", format )

/**
 * @brief Set the rotation of the text in a cell.
 *
 * @param format Pointer to a Format instance.
 * @param angle  Rotation angle in the range -90 to 90 and 270.
 *
 * Set the rotation of the text in a cell. The rotation can be any angle in the
 * range -90 to 90 degrees:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_rotation(format, 30)
 *
 *     worksheet_write_string(worksheet, 0, 0, "This text is rotated", format)
 * @endcode
 *
 * @image html format_font_text_rotated.png
 *
 * The angle 270 is also supported. This indicates text where the letters run from
 * top to bottom.
 */
FUNCTION lxw_format_set_rotation(format, angle)
RETURN CallDll( "format_set_rotation", format, angle )

/**
 * @brief Set the cell text indentation level.
 *
 * @param format Pointer to a Format instance.
 * @param level  Indentation level.
 *
 * This method can be used to indent text in a cell. The argument, which should be
 * an integer, is taken as the level of indentation:
 *
 * @code
 *     format1 = workbook_add_format(workbook)
 *     format2 = workbook_add_format(workbook)
 *
 *     format_set_indent(format1, 1)
 *     format_set_indent(format2, 2)
 *
 *     worksheet_write_string(worksheet, 0, 0, "This text is indented 1 level",  format1)
 *     worksheet_write_string(worksheet, 1, 0, "This text is indented 2 levels", format2)
 * @endcode
 *
 * @image html text_indent.png
 *
 * @note
 * Indentation is a horizontal alignment property. It will override any other
 * horizontal properties but it can be used in conjunction with vertical
 * properties.
 */
FUNCTION lxw_format_set_indent(format, level)
RETURN CallDll( "format_set_indent", format, level )

/**
 * @brief Turn on the text "shrink to fit" for a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This method can be used to shrink text so that it fits in a cell:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *     format_set_shrink(format)
 *
 *     worksheet_write_string(worksheet, 0, 0, "Honey, I shrunk the text!", format)
 * @endcode
 */
FUNCTION lxw_format_set_shrink(format)
RETURN CallDll( "format_set_shrink", format )

/**
 * @brief Set the background fill pattern for a cell
 *
 * @param format Pointer to a Format instance.
 * @param index  Pattern index.
 *
 * Set the background pattern for a cell.
 *
 * The most common pattern is a solid fill of the background color:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *
 *     format_set_pattern (format, LXW_PATTERN_SOLID)
 *     format_set_bg_color(format, LXW_COLOR_YELLOW)
 * @endcode
 *
 * The available fill patterns are:
 *
 *    Fill Type                     | Define
 *    ----------------------------- | -----------------------------
 *    Solid                         | #LXW_PATTERN_SOLID
 *    Medium gray                   | #LXW_PATTERN_MEDIUM_GRAY
 *    Dark gray                     | #LXW_PATTERN_DARK_GRAY
 *    Light gray                    | #LXW_PATTERN_LIGHT_GRAY
 *    Dark horizontal line          | #LXW_PATTERN_DARK_HORIZONTAL
 *    Dark vertical line            | #LXW_PATTERN_DARK_VERTICAL
 *    Dark diagonal stripe          | #LXW_PATTERN_DARK_DOWN
 *    Reverse dark diagonal stripe  | #LXW_PATTERN_DARK_UP
 *    Dark grid                     | #LXW_PATTERN_DARK_GRID
 *    Dark trellis                  | #LXW_PATTERN_DARK_TRELLIS
 *    Light horizontal line         | #LXW_PATTERN_LIGHT_HORIZONTAL
 *    Light vertical line           | #LXW_PATTERN_LIGHT_VERTICAL
 *    Light diagonal stripe         | #LXW_PATTERN_LIGHT_DOWN
 *    Reverse light diagonal stripe | #LXW_PATTERN_LIGHT_UP
 *    Light grid                    | #LXW_PATTERN_LIGHT_GRID
 *    Light trellis                 | #LXW_PATTERN_LIGHT_TRELLIS
 *    12.5% gray                    | #LXW_PATTERN_GRAY_125
 *    6.25% gray                    | #LXW_PATTERN_GRAY_0625
 *
 */
FUNCTION lxw_format_set_pattern(format, index)
RETURN CallDll( "format_set_pattern", format, index )

/**
 * @brief Set the pattern background color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern background color.
 *
 * The format_set_bg_color() method can be used to set the background color of
 * a pattern. Patterns are defined via the format_set_pattern() method. If a
 * pattern hasn't been defined then a solid fill pattern is used as the
 * default.
 *
 * Here is an example of how to set up a solid fill in a cell:
 *
 * @code
 *     format = workbook_add_format(workbook)
 *
 *     format_set_pattern (format, LXW_PATTERN_SOLID)
 *     format_set_bg_color(format, LXW_COLOR_GREEN)
 *
 *     worksheet_write_string(worksheet, 0, 0, "Ray", format)
 * @endcode
 *
 * @image html formats_set_bg_color.png
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
FUNCTION lxw_format_set_bg_color(format, color)
RETURN CallDll( "format_set_bg_color", format, color )

/**
 * @brief Set the pattern foreground color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern foreground  color.
 *
 * The format_set_fg_color() method can be used to set the foreground color of
 * a pattern.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
FUNCTION lxw_format_set_fg_color(format, color)
RETURN CallDll( "format_set_fg_color", format, color )

/**
 * @brief Set the cell border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell border style:
 *
 * @code
 *     format_set_border(format, LXW_BORDER_THIN)
 * @endcode
 *
 * Individual border elements can be configured using the following functions with
 * the same parameters:
 *
 * - format_set_bottom()
 * - format_set_top()
 * - format_set_left()
 * - format_set_right()
 *
 * A cell border is comprised of a border on the bottom, top, left and right.
 * These can be set to the same value using format_set_border() or
 * individually using the relevant method calls shown above.
 *
 * The following border styles are available:
 *
 * - #LXW_BORDER_THIN
 * - #LXW_BORDER_MEDIUM
 * - #LXW_BORDER_DASHED
 * - #LXW_BORDER_DOTTED
 * - #LXW_BORDER_THICK
 * - #LXW_BORDER_DOUBLE
 * - #LXW_BORDER_HAIR
 * - #LXW_BORDER_MEDIUM_DASHED
 * - #LXW_BORDER_DASH_DOT
 * - #LXW_BORDER_MEDIUM_DASH_DOT
 * - #LXW_BORDER_DASH_DOT_DOT
 * - #LXW_BORDER_MEDIUM_DASH_DOT_DOT
 * - #LXW_BORDER_SLANT_DASH_DOT
 *
 *  The most commonly used style is the `thin` style.
 */
FUNCTION lxw_format_set_border(format, style)
RETURN CallDll( "format_set_border", format, style )

/**
 * @brief Set the cell bottom border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell bottom border style. See format_set_border() for details on the
 * border styles.
 */
FUNCTION lxw_format_set_bottom(format, style)
RETURN CallDll( "format_set_bottom", format, style )

/**
 * @brief Set the cell top border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell top border style. See format_set_border() for details on the border
 * styles.
 */
FUNCTION lxw_format_set_top(format, style)
RETURN CallDll( "format_set_top", format, style )

/**
 * @brief Set the cell left border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell left border style. See format_set_border() for details on the
 * border styles.
 */
FUNCTION lxw_format_set_left(format, style)
RETURN CallDll( "format_set_left", format, style )

/**
 * @brief Set the cell right border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell right border style. See format_set_border() for details on the
 * border styles.
 */
FUNCTION lxw_format_set_right(format, style)
RETURN CallDll( "format_set_right", format, style )

/**
 * @brief Set the color of the cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * Individual border elements can be configured using the following methods with
 * the same parameters:
 *
 * - format_set_bottom_color()
 * - format_set_top_color()
 * - format_set_left_color()
 * - format_set_right_color()
 *
 * Set the color of the cell borders. A cell border is comprised of a border
 * on the bottom, top, left and right. These can be set to the same color
 * using format_set_border_color() or individually using the relevant method
 * calls shown above.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 */
FUNCTION lxw_format_set_border_color(format, color)
RETURN CallDll( "format_set_border_color", format, color )

/**
 * @brief Set the color of the bottom cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
FUNCTION lxw_format_set_bottom_color(format, color)
RETURN CallDll( "format_set_bottom_color", format, color )

/**
 * @brief Set the color of the top cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
FUNCTION lxw_format_set_top_color(format, color)
RETURN CallDll( "format_set_top_color", format, color )

/**
 * @brief Set the color of the left cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
FUNCTION lxw_format_set_left_color(format, color)
RETURN CallDll( "format_set_left_color", format, color )

/**
 * @brief Set the color of the right cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
FUNCTION lxw_format_set_right_color(format, color)
RETURN CallDll( "format_set_right_color", format, color )

FUNCTION lxw_format_set_diag_type(format, value)
RETURN CallDll( "format_set_diag_type", format, value )

FUNCTION lxw_format_set_diag_color(format, color)
RETURN CallDll( "format_set_diag_color", format, color )

FUNCTION lxw_format_set_diag_border(format, value)
RETURN CallDll( "format_set_diag_border", format, value )

FUNCTION lxw_format_set_font_outline(format)
RETURN CallDll( "format_set_font_outline", format )

FUNCTION lxw_format_set_font_shadow(format)
RETURN CallDll( "format_set_font_shadow", format )

FUNCTION lxw_format_set_font_family(format, value)
RETURN CallDll( "format_set_font_family", format, value )

FUNCTION lxw_format_set_font_charset(format, value)
RETURN CallDll( "format_set_font_charset", format, value )

FUNCTION lxw_format_set_font_scheme(format, font_scheme)
RETURN CallDll( "format_set_font_scheme", format, font_scheme )

FUNCTION lxw_format_set_font_condense(format)
RETURN CallDll( "format_set_font_condense", format )

FUNCTION lxw_format_set_font_extend(format)
RETURN CallDll( "format_set_font_extend", format )

FUNCTION lxw_format_set_reading_order(format, value)
RETURN CallDll( "format_set_reading_order", format, value )

FUNCTION lxw_format_set_theme(format, value)
RETURN CallDll( "format_set_theme", format, value )

// UTILITY =============================================================================================================

/**
 * @brief Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *      worksheet_write_string(worksheet, CELL("A1"), "Foo", NULL);
 *
 *      //Same as:
 *      worksheet_write_string(worksheet, 0, 0,       "Foo", NULL);
 * @endcode
 *
 * @note
 *
 * This macro shouldn't be used in performance critical situations since it
 * expands to two function calls.
 */
//#define CELL(cell)  lxw_name_to_row(cell), lxw_name_to_col(cell)
//#xtranslate LXW_CELL([<cell>]) => lxw_name_to_row(<cell>), lxw_name_to_col(<cell>)


/**
 * @brief Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *     worksheet_set_column(worksheet, COLS("B:D"), 20, NULL, NULL);
 *
 *     // Same as:
 *     worksheet_set_column(worksheet, 1, 3,        20, NULL, NULL);
 * @endcode
 *
 */
//#define COLS(cols)  lxw_name_to_col(cols), lxw_name_to_col_2(cols)
//#xtranslate LXW_COLS([<cols>]) => lxw_name_to_col(<cols>), lxw_name_to_col_2(<cols>)

/**
 * @brief Convert an Excel `A1:B2` range into a `(first_row, first_col,
 *        last_row, last_col)` sequence.
 *
 * Convert an Excel `A1:B2` range into a `(first_row, first_col, last_row,
 * last_col)` sequence.
 *
 * This is a little syntactic shortcut to help with worksheet layout.
 *
 * @code
 *     worksheet_print_area(worksheet, 0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     worksheet_print_area(worksheet, RANGE("A1:K42"));
 * @endcode
 */
//#define RANGE(range) lxw_name_to_row(range), lxw_name_to_col(range), lxw_name_to_row_2(range), lxw_name_to_col_2(range)
//#xtranslate LXW_RANGE([<range>]) => lxw_name_to_row(<range>), lxw_name_to_col(<range>), lxw_name_to_row_2(<range>), lxw_name_to_col_2(<range>)

FUNCTION lxw_col_to_name(col_name, col_num, absolute)
LOCAL ret
col_name:= SPACE(4)
ret:= CallDll( "lxw_col_to_name", @col_name, col_num, absolute )
col_name:= LEFT(col_name, LEN(RTRIM(col_name))-1)
RETURN ret

FUNCTION lxw_rowcol_to_cell(cell_name, row, col)
RETURN CallDll( "lxw_rowcol_to_cell", cell_name, row, col )

FUNCTION lxw_rowcol_to_cell_abs(cell_name, row, col, abs_row, abs_col)
RETURN CallDll( "lxw_rowcol_to_cell_abs", cell_name, row, col, abs_row, abs_col )

FUNCTION lxw_rowcol_to_range(range, first_row, first_col, last_row, last_col)
RETURN CallDll( "lxw_rowcol_to_range", range, first_row, first_col, last_row, last_col )

FUNCTION lxw_rowcol_to_range_abs(range, first_row, first_col, last_row, last_col)
RETURN CallDll( "lxw_rowcol_to_range_abs", range, first_row, first_col, last_row, last_col )

FUNCTION lxw_rowcol_to_formula_abs(formula, sheetname, first_row,  first_col, last_row,  last_col)
RETURN CallDll( "lxw_rowcol_to_formula_abs", formula, sheetname, first_row,  first_col, last_row,  last_col )

FUNCTION lxw_name_to_row( row_str)
RETURN CallDll( "lxw_name_to_row", row_str )

FUNCTION lxw_name_to_col( col_str)
RETURN hb_DynCall( { "lxw_name_to_col", nHDll, hb_bitOr( HB_DYN_CTYPE_SHORT_UNSIGNED, HB_DYN_CALLCONV_SYSCALL ) }, col_str )
//RETURN CallDll( "lxw_name_to_col", col_str )

FUNCTION lxw_name_to_row_2( row_str)
RETURN CallDll( "lxw_name_to_row_2", row_str )

FUNCTION lxw_name_to_col_2( col_str)
RETURN CallDll( "lxw_name_to_col_2", col_str )
