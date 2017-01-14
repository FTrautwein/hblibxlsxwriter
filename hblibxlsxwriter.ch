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

    /** False value. */
#define LXW_FALSE  0
    /** True value. */
#define LXW_TRUE   1

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


/**
 * @brief Available chart types.
 */
    /** None. */
#define LXW_CHART_NONE  0

    /** Area chart. */
#define LXW_CHART_AREA  1

    /** Area chart - stacked. */
#define LXW_CHART_AREA_STACKED  2

    /** Area chart - percentage stacked. */
#define LXW_CHART_AREA_STACKED_PERCENT  3

    /** Bar chart. */
#define LXW_CHART_BAR  4

    /** Bar chart - stacked. */
#define LXW_CHART_BAR_STACKED  5

    /** Bar chart - percentage stacked. */
#define LXW_CHART_BAR_STACKED_PERCENT  6

    /** Column chart. */
#define LXW_CHART_COLUMN  7

    /** Column chart - stacked. */
#define LXW_CHART_COLUMN_STACKED  8

    /** Column chart - percentage stacked. */
#define LXW_CHART_COLUMN_STACKED_PERCENT  9

    /** Doughnut chart. */
#define LXW_CHART_DOUGHNUT  10

    /** Line chart. */
#define LXW_CHART_LINE  11

    /** Pie chart. */
#define LXW_CHART_PIE  12

    /** Scatter chart. */
#define LXW_CHART_SCATTER  13

    /** Scatter chart - straight. */
#define LXW_CHART_SCATTER_STRAIGHT  14

    /** Scatter chart - straight with markers. */
#define LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS  15

    /** Scatter chart - smooth. */
#define LXW_CHART_SCATTER_SMOOTH  16

    /** Scatter chart - smooth with markers. */
#define LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS  17

    /** Radar chart. */
#define LXW_CHART_RADAR  18

    /** Radar chart - with markers. */
#define LXW_CHART_RADAR_WITH_MARKERS  19

    /** Radar chart - filled. */
#define LXW_CHART_RADAR_FILLED  20


/**
 * @brief Chart legend positions.
 */
    /** No chart legend. */
#define LXW_CHART_LEGEND_NONE             0

    /** Chart legend positioned at right side. */
#define LXW_CHART_LEGEND_RIGHT            1

    /** Chart legend positioned at left side. */
#define XW_CHART_LEGEND_LEFT              2

    /** Chart legend positioned at top. */
#define LXW_CHART_LEGEND_TOP              3

    /** Chart legend positioned at bottom. */
#define LXW_CHART_LEGEND_BOTTOM           4

    /** Chart legend overlaid at right side. */
#define LXW_CHART_LEGEND_OVERLAY_RIGHT    5

    /** Chart legend overlaid at left side. */
#define LXW_CHART_LEGEND_OVERLAY_LEFT     6


/**
 * @brief Chart pattern types.
 */
    /** None pattern. */
#define LXW_CHART_PATTERN_NONE                                   0

    /** 5 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_5                              1

    /** 10 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_10                             2

    /** 20 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_20                             3

    /** 25 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_25                             4

    /** 30 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_30                             5

    /** 40 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_40                             6

    /** 50 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_50                             7

    /** 60 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_60                             8

    /** 70 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_70                             9

    /** 75 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_75                            10

    /** 80 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_80                            11

    /** 90 Percent pattern. */
#define LXW_CHART_PATTERN_PERCENT_90                            12

    /** Light downward diagonal pattern. */
#define LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL               13

    /** Light upward diagonal pattern. */
#define LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL                 14

    /** Dark downward diagonal pattern. */
#define LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL                15

    /** Dark upward diagonal pattern. */
#define LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL                  16

    /** Wide downward diagonal pattern. */
#define LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL                17

    /** Wide upward diagonal pattern. */
#define LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL                  18

    /** Light vertical pattern. */
#define LXW_CHART_PATTERN_LIGHT_VERTICAL                        19

    /** Light horizontal pattern. */
#define LXW_CHART_PATTERN_LIGHT_HORIZONTAL                      20

    /** Narrow vertical pattern. */
#define LXW_CHART_PATTERN_NARROW_VERTICAL                       21

    /** Narrow horizontal pattern. */
#define LXW_CHART_PATTERN_NARROW_HORIZONTAL                     22

    /** Dark vertical pattern. */
#define LXW_CHART_PATTERN_DARK_VERTICAL                         23

    /** Dark horizontal pattern. */
#define LXW_CHART_PATTERN_DARK_HORIZONTAL                       24

    /** Dashed downward diagonal pattern. */
#define LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL              25

    /** Dashed upward diagonal pattern. */
#define LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL                26

    /** Dashed horizontal pattern. */
#define LXW_CHART_PATTERN_DASHED_HORIZONTAL                     27

    /** Dashed vertical pattern. */
#define LXW_CHART_PATTERN_DASHED_VERTICAL                       28

    /** Small confetti pattern. */
#define LXW_CHART_PATTERN_SMALL_CONFETTI                        29

    /** Large confetti pattern. */
#define LXW_CHART_PATTERN_LARGE_CONFETTI                        30

    /** Zigzag pattern. */
#define LXW_CHART_PATTERN_ZIGZAG                                31

    /** Wave pattern. */
#define LXW_CHART_PATTERN_WAVE                                  32

    /** Diagonal brick pattern. */
#define LXW_CHART_PATTERN_DIAGONAL_BRICK                        33

    /** Horizontal brick pattern. */
#define LXW_CHART_PATTERN_HORIZONTAL_BRICK                      34

    /** Weave pattern. */
#define LXW_CHART_PATTERN_WEAVE                                 35

    /** Plaid pattern. */
#define LXW_CHART_PATTERN_PLAID                                 36

    /** Divot pattern. */
#define LXW_CHART_PATTERN_DIVOT                                 37

    /** Dotted grid pattern. */
#define LXW_CHART_PATTERN_DOTTED_GRID                           38

    /** Dotted diamond pattern. */
#define LXW_CHART_PATTERN_DOTTED_DIAMOND                        39

    /** Shingle pattern. */
#define LXW_CHART_PATTERN_SHINGLE                               40

    /** Trellis pattern. */
#define LXW_CHART_PATTERN_TRELLIS                               41

    /** Sphere pattern. */
#define LXW_CHART_PATTERN_SPHERE                                42

    /** Small grid pattern. */
#define LXW_CHART_PATTERN_SMALL_GRID                            43

    /** Large grid pattern. */
#define LXW_CHART_PATTERN_LARGE_GRID                            44

    /** Small check pattern. */
#define LXW_CHART_PATTERN_SMALL_CHECK                           45

    /** Large check pattern. */
#define LXW_CHART_PATTERN_LARGE_CHECK                           46

    /** Outlined diamond pattern. */
#define LXW_CHART_PATTERN_OUTLINED_DIAMOND                      47

    /** Solid diamond pattern. */
#define LXW_CHART_PATTERN_SOLID_DIAMOND                         48
