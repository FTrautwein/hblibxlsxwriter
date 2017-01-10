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

/** Alignment values for format_set_align(). */
        /** No alignment. Cell will use Excel's default for the data type */
#define LXW_ALIGN_NONE  0

        /** Left horizontal alignment */
#define LXW_ALIGN_LEFT  1

        /** Center horizontal alignment */
#define LXW_ALIGN_CENTER  2

        /** Right horizontal alignment */
#define LXW_ALIGN_RIGHT  3
 
        /** Cell fill horizontal alignment */
#define LXW_ALIGN_FILL  4

        /** Justify horizontal alignment */
#define LXW_ALIGN_JUSTIFY  5

        /** Center Across horizontal alignment */
#define LXW_ALIGN_CENTER_ACROSS  6

        /** Left horizontal alignment */
#define LXW_ALIGN_DISTRIBUTED  7

        /** Top vertical alignment */
#define LXW_ALIGN_VERTICAL_TOP  8

        /** Bottom vertical alignment */
#define LXW_ALIGN_VERTICAL_BOTTOM  9
 
        /** Center vertical alignment */
#define LXW_ALIGN_VERTICAL_CENTER  10

        /** Justify vertical alignment */
#define LXW_ALIGN_VERTICAL_JUSTIFY  11

        /** Distributed vertical alignment */
#define LXW_ALIGN_VERTICAL_DISTRIBUTED  12


/** Predefined values for common colors. */
        /** Black */
#define LXW_COLOR_BLACK  0x1000000

        /** Blue */
#define LXW_COLOR_BLUE  0x0000FF

        /** Brown */
#define LXW_COLOR_BROWN  0x800000

        /** Cyan */
#define LXW_COLOR_CYAN  0x00FFFF

        /** Gray */
#define LXW_COLOR_GRAY  0x808080

        /** Green */
#define LXW_COLOR_GREEN  0x008000

        /** Lime */
#define LXW_COLOR_LIME  0x00FF00

        /** Magenta */
#define LXW_COLOR_MAGENTA  0xFF00FF

        /** Navy */
#define LXW_COLOR_NAVY  0x000080

        /** Orange */
#define LXW_COLOR_ORANGE  0xFF6600

        /** Pink */
#define LXW_COLOR_PINK  0xFF00FF

        /** Purple */
#define LXW_COLOR_PURPLE  0x800080

        /** Red */
#define LXW_COLOR_RED  0xFF0000

        /** Silver */
#define LXW_COLOR_SILVER  0xC0C0C0

        /** White */
#define LXW_COLOR_WHITE  0xFFFFFF

        /** Yellow */
#define LXW_COLOR_YELLOW  0xFFFF00


/** Cell border styles for use with format_set_border(). */
        /** No border */
#define LXW_BORDER_NONE  0

        /** Thin border style */
#define LXW_BORDER_THIN  1

        /** Medium border style */
#define LXW_BORDER_MEDIUM  2

        /** Dashed border style */
#define LXW_BORDER_DASHED  3

        /** Dotted border style */
#define LXW_BORDER_DOTTED  4

        /** Thick border style */
#define LXW_BORDER_THICK  5

        /** Double border style */
#define LXW_BORDER_DOUBLE  6

        /** Hair border style */
#define LXW_BORDER_HAIR  7

        /** Medium dashed border style */
#define LXW_BORDER_MEDIUM_DASHED  8

        /** Dash-dot border style */
#define LXW_BORDER_DASH_DOT  9

        /** Medium dash-dot border style */
#define LXW_BORDER_MEDIUM_DASH_DOT  10

        /** Dash-dot-dot border style */
#define LXW_BORDER_DASH_DOT_DOT  11

        /** Medium dash-dot-dot border style */
#define LXW_BORDER_MEDIUM_DASH_DOT_DOT  12

        /** Slant dash-dot border style */
#define LXW_BORDER_SLANT_DASH_DOT 13


/** Format underline values for format_set_underline(). */
        /** Single underline */
#define LXW_UNDERLINE_SINGLE  1

        /** Double underline */
#define LXW_UNDERLINE_DOUBLE  2

        /** Single accounting underline */
#define LXW_UNDERLINE_SINGLE_ACCOUNTING  3

        /** Double accounting underline */
#define LXW_UNDERLINE_DOUBLE_ACCOUNTING  4


/** Superscript and subscript values for format_set_font_script(). */
        /** Superscript font */
#define LXW_FONT_SUPERSCRIPT  1

        /** Subscript font */
#define LXW_FONT_SUBSCRIPT  2


/* lxw_format_diagonal_types */
#define LXW_DIAGONAL_BORDER_UP  1
#define LXW_DIAGONAL_BORDER_DOWN  2
#define LXW_DIAGONAL_BORDER_UP_DOWN  3


/** Pattern value for use with format_set_pattern(). */
        /** Empty pattern */
#define LXW_PATTERN_NONE   0

        /** Solid pattern */
#define LXW_PATTERN_SOLID  2

        /** Medium gray pattern */
#define LXW_PATTERN_MEDIUM_GRAY  3

        /** Dark gray pattern */
#define LXW_PATTERN_DARK_GRAY  4

        /** Light gray pattern */
#define LXW_PATTERN_LIGHT_GRAY  5

        /** Dark horizontal line pattern */
#define LXW_PATTERN_DARK_HORIZONTAL  6

        /** Dark vertical line pattern */
#define LXW_PATTERN_DARK_VERTICAL  7

        /** Dark diagonal stripe pattern */
#define LXW_PATTERN_DARK_DOWN  8

        /** Reverse dark diagonal stripe pattern */
#define LXW_PATTERN_DARK_UP  9

        /** Dark grid pattern */
#define LXW_PATTERN_DARK_GRID  10

        /** Dark trellis pattern */
#define LXW_PATTERN_DARK_TRELLIS  11

        /** Light horizontal Line pattern */
#define LXW_PATTERN_LIGHT_HORIZONTAL  12

        /** Light vertical line pattern */
#define LXW_PATTERN_LIGHT_VERTICAL  13

        /** Light diagonal stripe pattern */
#define LXW_PATTERN_LIGHT_DOWN  14

        /** Reverse light diagonal stripe pattern */
#define LXW_PATTERN_LIGHT_UP  15

        /** Light grid pattern */
#define LXW_PATTERN_LIGHT_GRID  16

        /** Light trellis pattern */
#define LXW_PATTERN_LIGHT_TRELLIS  17

        /** 12.5% gray pattern */
#define LXW_PATTERN_GRAY_125  18

        /** 6.25% gray pattern */
#define LXW_PATTERN_GRAY_0625  19

/** Gridline options using in `worksheet_gridlines()`. */
        /** Hide screen and print gridlines. */
#define LXW_HIDE_ALL_GRIDLINES   0
        /** Show screen gridlines. */
#define LXW_SHOW_SCREEN_GRIDLINES  1
        /** Show print gridlines. */
#define LXW_SHOW_PRINT_GRIDLINES  2
        /** Show screen and print gridlines. */
#define LXW_SHOW_ALL_GRIDLINES  3

#xtranslate LXW_CELL([<cell>]) => lxw_name_to_row(<cell>), lxw_name_to_col(<cell>)
#xtranslate LXW_COLS([<cols>]) => lxw_name_to_col(<cols>), lxw_name_to_col_2(<cols>)
#xtranslate LXW_RANGE([<range>]) => lxw_name_to_row(<range>), lxw_name_to_col(<range>), lxw_name_to_row_2(<range>), lxw_name_to_col_2(<range>)
