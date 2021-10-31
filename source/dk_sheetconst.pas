unit DK_SheetConst;

{$mode ObjFPC}{$H+}

interface

uses
  {Classes, SysUtils} fpstypes, Graphics;

const
  {---COMMON PROPERTIES---}
  //Dimensions
  WIDTH_PX_RATIO = 7;
  HEIGHT_PX_RATIO = 14;
  //Alignment
  ALIGN_HOR_DEFAULT:  TsHorAlignment  = haCenter;
  ALIGN_VERT_DEFAULT: TsVertAlignment = vaCenter;
  //RowHeight
  ROW_HEIGHT_DEFAULT = 21;
  //Color
  TRANSPARENT_COLOR_INDEX = 0;

  {---SHEET PROPERTIES---}
  //Font
  FONT_NAME_DEFAULT: String = 'Arial';
  FONT_SIZE_DEFAULT: Single = 8;
  FONT_SIZE_MINIMUM: Single = 6;
  FONT_STYLE_DEFAULT: TsFontStyles = [];
  FONT_COLOR_DEFAULT: TsColor = scBlack;
  //Background
  BG_STYLE_DEFAULT: TsFillStyle = fsNoFill;
  BG_COLOR_DEFAULT: TsColor = scTransparent;
  PATTERN_COLOR_DEFAULT: TsColor = scTransparent;
  //Borders
  BORDER_STYLE_DEFAULT: TsLineStyle = lsThin;
  BORDER_COLOR_DEFAULT: TsColor = scBlack;
  //Sheet
  MAX_SHEETNAME_LENGTH = 31;

  {---GRID PROPERTIES---}
  GRID_COLOR_DEFAULT: TColor = clWindow;
  GRID_LINE_COLOR_DEFAULT: TColor = clWindowText;
  GRID_FONT_COLOR_DEFAULT: TColor = clWindowText;
  GRID_SELECTED_ROW_COLOR_DEFAULT: TColor = clHighlight;
  GRID_SELECTED_FONT_COLOR_DEFAULT: TColor = clHighlightText;
  //GRID_SELECTED_CELL_COLOR_DEFAULT - расчет через TintedColor от GRID_SELECTED_ROW_COLOR_DEFAULT
  //TintedColor(ColorToRGB(GRID_SELECTED_ROW_COLOR_DEFAULT), -0.4);
implementation

end.

