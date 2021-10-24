unit DK_SheetConst;

{$mode ObjFPC}{$H+}

interface

uses
  {Classes, SysUtils} fpstypes;

const
  //Dimensions
  WIDTH_PX_RATIO = 7;
  HEIGHT_PX_RATIO = 14;
  //Color
  TRANSPARENT_COLOR_INDEX = 0;
  //Font
  FONT_NAME_DEFAULT: String = 'Arial';
  FONT_SIZE_DEFAULT: Single = 9;
  FONT_SIZE_MINIMUM: Single = 6;
  FONT_STYLE_DEFAULT: TsFontStyles = [];
  FONT_COLOR_DEFAULT: TsColor = scBlack;
  //Alignment
  ALIGN_HOR_DEFAULT:  TsHorAlignment  = haCenter;
  ALIGN_VERT_DEFAULT: TsVertAlignment = vaCenter;
  //Background
  BG_STYLE_DEFAULT: TsFillStyle = fsNoFill;
  BG_COLOR_DEFAULT: TsColor = scTransparent;
  PATTERN_COLOR_DEFAULT: TsColor = scTransparent;
  //Borders
  BORDER_STYLE_DEFAULT: TsLineStyle = lsThin;
  BORDER_COLOR_DEFAULT: TsColor = scBlack;
  //RowHeight
  ROW_HEIGHT_DEFAULT = 21;

  MAX_SHEETNAME_LENGTH = 31;

implementation

end.

