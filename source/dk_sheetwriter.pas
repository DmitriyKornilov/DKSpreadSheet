unit DK_SheetWriter;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Grids, fpstypes, fpspreadsheet, fpspreadsheetgrid, Graphics,
  DK_Const, DK_Vector, DK_Matrix, DK_TextUtils, DK_StrUtils, DK_SheetConst,
  DK_SheetUtils;

type

  TCellBorderType = (cbtNone, cbtLeft, cbtRight, cbtTop, cbtBottom,
                     cbtOuter, cbtInner, cbtAll);

  { TSheetWriter }

  TSheetWriter = class(TObject)
  private
    FWorksheet: TsWorksheet;
    FGrid: TsWorksheetGrid;
    FRowHeights: TIntVector;
    FColWidths: TIntVector;
    FFirstCol: Integer;
    FFirstRow: Integer;
    //Color
    FBGColorMatrix: TIntMatrix;
    //Font
    FFontName: String;
    FFontSize: Single;
    FFontStyle: TsFontStyles;
    FFontColor: TsColor;
    //Alignment
    FHorAlignment: TsHorAlignment;
    FVertAlignment: TsVertAlignment;
    //Background
    FBGStyle: TsFillStyle;
    FBGColor: TsColor;
    FPatternColor: TsColor;
    //Borders
    FLeftBorderStyle: TsLineStyle;
    FRightBorderStyle: TsLineStyle;
    FTopBorderStyle: TsLineStyle;
    FBottomBorderStyle: TsLineStyle;
    FInnerBorderStyle: TsLineStyle;
    FLeftBorderColor: TsColor;
    FRightBorderColor: TsColor;
    FTopBorderColor: TsColor;
    FBottomBorderColor: TsColor;
    FInnerBorderColor: TsColor;

    function GetHasGrid: Boolean;
    function GetRowCount: Integer;
    function GetColCount: Integer;
    function GetRowHeight(const ARow: Integer): Integer;
    function ColIndex(const ACol: Integer): Integer;
    function RowIndex(const ARow: Integer): Integer;
    procedure CellIndex(var ARow, ACol: Integer);
    procedure CellIndex(var ARow1, ACol1, ARow2, ACol2: Integer);
    procedure SetFontName(const AName: String);
    procedure GetBordersNeed(const ABordersType: TCellBorderType;
                           out ALeft, ARight, ATop, ABottom, AInner: Boolean);
    procedure DrawCellBorders(const ARow, ACol: Integer;
               const ALeftNeed: Boolean; const ALeftStyle: TsLineStyle; const ALeftColor: TsColor;
               const ARightNeed: Boolean; const ARightStyle: TsLineStyle; const ARightColor: TsColor;
               const ATopNeed: Boolean; const ATopStyle: TsLineStyle; const ATopColor: TsColor;
               const ABottomNeed: Boolean; const ABottomStyle: TsLineStyle; const ABottomColor: TsColor);
    procedure DrawBorders(const ARow, ACol: Integer;
                          const ALeftNeed: Boolean = True; const ARightNeed: Boolean = True;
                          const ATopNeed: Boolean = True; const ABottomNeed: Boolean = True);
    procedure DrawBorders(const ARow1, ACol1, ARow2, ACol2: Integer;
                          const ALeftNeed: Boolean = True; const ARightNeed: Boolean = True;
                          const ATopNeed: Boolean = True; const ABottomNeed: Boolean = True;
                          const AInnerNeed: Boolean = False);
    procedure SetCellMainSettings(const ARow, ACol: Integer; const AWordWrap: Boolean);
    procedure SetCellSettings(const ARow1, ACol1, ARow2, ACol2, ARowHeight: Integer;
                              const AWordWrap: Boolean; const ABordersType: TCellBorderType);
    procedure SetWidth(const ACol, AValue: Integer);
    procedure SetHeight(const ARow, AValue: Integer);
    procedure SetLineHeight(const ARow, AHeight: Integer; const AMinValue: Integer = ROW_HEIGHT_DEFAULT);
    procedure SetNewRowHeight(const ARow, AHeight: Integer; const AMinValue: Integer = ROW_HEIGHT_DEFAULT);

    function CalcCellHeight(var AText: String; const ACol1, ACol2: Integer;
                           const AWrapToWordParts: Boolean = False;
                           const ARedStrWidth: Integer = 0): Integer;
    procedure SetGridRowCount(const ACount: Integer);
    procedure SetGridColCount(const ACount: Integer);

    procedure SetDefaultGridSettings;
  public
    constructor Create(const AColWidths: TIntVector; const AWorksheet: TsWorksheet; const AGrid: TsWorksheetGrid = nil);
    destructor  Destroy; override;
    procedure Clear;
    procedure SetDefaultCellSettings;
    procedure BeginEdit;
    procedure EndEdit;
    //Colors
    procedure AddCellBGColorIndex(const ARow, ACol, AColorIndex: Integer);
    procedure AddCellBGColorIndex(const ARow1, ACol1, ARow2, ACol2, AColorIndex: Integer);
    procedure DelCellBGColorIndex(const ARow, ACol: Integer);
    procedure ApplyBGColors(const ABGColors: TColorVector);
    procedure ClearBGColors;
    //Font
    procedure SetFontDefault;
    procedure SetFont(const AName: String; const ASize: Single; const AStyle: TsFontStyles; const AColor: TsColor);
    procedure SetFont(const AName: String; const ASize: Single; const AStyle: TFontStyles; const AColor: TColor);
    procedure SetFont(const AFont: TFont);
    //Alignment
    procedure SetAlignmentDefault;
    procedure SetAlignment(const AHorAlignment: TsHorAlignment; const AVertAlignment: TsVertAlignment);
    //Background
    procedure SetBackgroundClear;
    procedure SetBackgroundDefault;
    procedure SetBackground(const ABGStyle: TsFillStyle; const ABGColor: TsColor; const APatternColor: TsColor);
    procedure SetBackground(const ABGStyle: TsFillStyle; const ABGColor: TColor;  const APatternColor: TColor);
    procedure SetBackground(const ABGColor: TsColor);
    procedure SetBackground(const ABGColor: TColor);
    //Borders
    procedure SetBordersDefault;
    procedure SetBordersStyle(const ALeftStyle,ARightStyle,ATopStyle,ABottomStyle,AInnerStyle: TsLineStyle);
    procedure SetBordersColor(const ALeftColor,ARightColor,ATopColor,ABottomColor,AInnerColor: TsColor);
    procedure SetBordersColor(const ALeftColor,ARightColor,ATopColor,ABottomColor,AInnerColor: TColor);
    procedure SetBordersColor(const AAllColor: TsColor);
    procedure SetBordersColor(const AAllColor: TColor);
    procedure SetBorders(const AAllStyle: TsLineStyle; const AAllColor: TsColor);
    procedure SetBorders(const AAllStyle: TsLineStyle; const AAllColor: TColor);
    procedure SetBorders(const AOuterStyle: TsLineStyle; const AOuterColor: TsColor; const AInnerStyle: TsLineStyle; const AInnerColor: TsColor);
    procedure SetBorders(const AOuterStyle: TsLineStyle; const AOuterColor: TColor; const AInnerStyle: TsLineStyle; const AInnerColor: TColor);
    procedure DrawBorders(ARow, ACol: Integer; const ABordersType: TCellBorderType);
    procedure DrawBorders(ARow1, ACol1, ARow2, ACol2: Integer; const ABordersType: TCellBorderType);
    //Sizes
    procedure SetColWidth(ACol, AValue: Integer);
    procedure SetRowHeight(ARow, AValue: Integer);
    procedure SetColumns;
    procedure SetRows;
    //Frozen
    procedure SetFrozenCols(AFixColCount: Integer);
    procedure SetFrozenRows(AFixRowCount: Integer);
    //Repeated
    procedure SetRepeatedRows(ABeginRow, AEndRow: Integer);
    procedure SetRepeatedCols(ABeginCol, AEndCol: Integer);
    //Image
    procedure WriteImage(ARow, ACol: Integer; AFileName: String;
                         AOffsetX: Double= 0.0; AOffsetY: Double=0.0;
                         AScaleX: Double=1.0; AScaleY: Double=1.0);
    //WriteCellValue
    procedure WriteText(const ARow, ACol: Integer; const AValue: String;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AWordWrap: Boolean = True;
                        const AAutoHeight: Boolean = False;
                        const AWrapToWordParts: Boolean = False;
                        const ARedStrWidth: Integer = 0;
                        const ARichTextParams: TsRichTextParams = nil);
    procedure WriteText(ARow1, ACol1, ARow2, ACol2: Integer; AValue: String;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AWordWrap: Boolean = True;
                        const AAutoHeight: Boolean = False;
                        const AWrapToWordParts: Boolean = False;
                        const ARedStrWidth: Integer = 0;
                        const ARichTextParams: TsRichTextParams = nil);
    procedure WriteTextVertical(const ARow, ACol: Integer; const AValue: String;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AWordWrap: Boolean = True;
                        const AWrapToWordParts: Boolean = False;
                        const ARedStrWidth: Integer = 0;
                        const ARichTextParams: TsRichTextParams = nil);
    procedure WriteTextVertical(ARow1, ACol1, ARow2, ACol2: Integer; AValue: String;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AWordWrap: Boolean = True;
                        const AWrapToWordParts: Boolean = False;
                        const ARedStrWidth: Integer = 0;
                        const ARichTextParams: TsRichTextParams = nil);
    procedure WriteNumber(const ARow, ACol: Integer; const AValue: Double;
                          const ABordersType: TCellBorderType = cbtNone;
                          const AFormatString: String = '');
    procedure WriteNumber(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: Double;
                          const ABordersType: TCellBorderType = cbtNone;
                          const AFormatString: String = '');
    procedure WriteNumber(const ARow, ACol: Integer; const AValue: Double;
                          const ADecimals: Byte;
                          const ABordersType: TCellBorderType = cbtNone;
                          const ANumberFormat: TsNumberFormat = nfGeneral);
    procedure WriteNumber(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: Double;
                          const ADecimals: Byte;
                          const ABordersType: TCellBorderType = cbtNone;
                          const ANumberFormat: TsNumberFormat = nfGeneral);
    procedure WriteCurrency(const ARow, ACol: Integer; const AValue: Double;
                          const ABordersType: TCellBorderType = cbtNone;
                          const AFormatString: String = '');
    procedure WriteCurrency(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: Double;
                          const ABordersType: TCellBorderType = cbtNone;
                          const AFormatString: String = '');
    procedure WriteDate(const ARow, ACol: Integer; const AValue: TDate;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AFormatString: String = '');
    procedure WriteDate(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: TDate;
                        const ABordersType: TCellBorderType = cbtNone;
                        AFormatString: String = '');
    procedure WriteTime(const ARow, ACol: Integer; const AValue: TDateTime;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AFormatString: String = '');
    procedure WriteTime(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: TDateTime;
                        const ABordersType: TCellBorderType = cbtNone;
                        AFormatString: String = '');
    procedure WriteDateTime(const ARow, ACol: Integer; const AValue: TDateTime;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AFormatString: String = '');
    procedure WriteDateTime(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: TDateTime;
                        const ABordersType: TCellBorderType = cbtNone;
                        AFormatString: String = '');
    property WorkSheet: TsWorkSheet read FWorkSheet;
    property Grid: TsWorksheetGrid read FGrid;
    property HasGrid: Boolean read GetHasGrid;
    property RowCount: Integer read GetRowCount;
    property ColCount: Integer read GetColCount;
    property RowHeight[const ARow: Integer]: Integer read GetRowHeight;
  end;



implementation



{ TSheetWriter }

procedure TSheetWriter.SetDefaultCellSettings;
begin
  SetFontDefault;
  SetAlignmentDefault;
  SetBackgroundDefault;
  SetBordersDefault;
end;

procedure TSheetWriter.BeginEdit;
begin
  if HasGrid then
  begin
    FGrid.Visible:= False;
    FGrid.AutoExpand:= [aeData, aeNavigation];
  end;
  Clear;
  SetColumns;
end;

procedure TSheetWriter.EndEdit;
begin
  SetRows;
  SetColumns;
  if HasGrid then
  begin
    FGrid.AutoExpand:= [];
    FGrid.LeftCol:= 0;
    FGrid.TopRow:= 0;
    FGrid.Visible:= True;
  end;
end;

procedure TSheetWriter.DelCellBGColorIndex(const ARow, ACol: Integer);
var
  i: Integer;
begin
  for i:=0 to High(FBGColorMatrix) do
  begin
    if (FBGColorMatrix[i,0]=ARow) and (FBGColorMatrix[i,1]=ACol) then
    begin
      MDel(FBGColorMatrix, i);
      Exit;
    end
  end;
end;

procedure TSheetWriter.AddCellBGColorIndex(const ARow, ACol, AColorIndex: Integer);
var
  V: TIntVector;
begin
  V:= nil;
  VDim(V, 3);
  V[0]:= RowIndex(ARow);
  V[1]:= ColIndex(ACol);
  V[2]:= AColorIndex;
  MAppend(FBGColorMatrix, V);
end;

procedure TSheetWriter.AddCellBGColorIndex(const ARow1, ACol1, ARow2, ACol2, AColorIndex: Integer);
var
  i,j: Integer;
begin
  for i:= ARow1 to ARow2 do
    for j:= ACol1 to ACol2 do
      AddCellBGColorIndex(i,j, AColorIndex);
end;

procedure TSheetWriter.ApplyBGColors(const ABGColors: TColorVector);
var
  i, R, C: Integer;
  Cl: TsColor;
begin
  for i:= 0 to High(FBGColorMatrix) do
  begin
    R:= FBGColorMatrix[i,0];
    C:= FBGColorMatrix[i,1];
    if FBGColorMatrix[i,2] = TRANSPARENT_COLOR_INDEX then
      Cl:= scTransparent
    else
      Cl:= ColorGraphicsToSheets(ABGColors[FBGColorMatrix[i,2]]);
    FWorksheet.WriteBackground(R, C, fsSolidFill, Cl, Cl);
  end;
end;

procedure TSheetWriter.ClearBGColors;
var
  i, R, C: Integer;
begin
  for i:= 0 to High(FBGColorMatrix) do
  begin
    R:= FBGColorMatrix[i,0];
    C:= FBGColorMatrix[i,1];
    FWorksheet.WriteBackground(R, C, fsNoFill, scTransparent, scTransparent);
  end;
end;

constructor TSheetWriter.Create(const AColWidths: TIntVector;
          const AWorksheet: TsWorksheet; const AGrid: TsWorksheetGrid = nil);
begin
  inherited Create;
  FWorksheet:= AWorksheet;
  FGrid:= AGrid;
  FFirstCol:= 0;
  FFirstRow:= 0;
  FColWidths:= VCut(AColWidths);
  if HasGrid then
  begin
    FFirstCol:= 1;
    FFirstRow:= 1;
    VIns(FColWidths, 0, 1);
    VAppend(FColWidths, 0);
  end;
  Clear;
  SetDefaultGridSettings;
  SetColumns;
  SetDefaultCellSettings;
end;

destructor TSheetWriter.Destroy;
begin
  inherited Destroy;
end;

procedure TSheetWriter.Clear;
begin
  FWorksheet.Clear;
  FRowHeights:= nil;
  FBGColorMatrix:= nil;
  if HasGrid then
  begin
    VAppend(FRowHeights, 1);
    VAppend(FRowHeights, 0);
  end;
end;

function TSheetWriter.GetHasGrid: Boolean;
begin
  Result:= (FGrid<>nil);
end;

function TSheetWriter.GetColCount: Integer;
begin
  Result:= Length(FColWidths) - 2*Ord(HasGrid);
end;

function TSheetWriter.GetRowHeight(const ARow: Integer): Integer;
begin
  Result:= FRowHeights[RowIndex(ARow)];
end;

function TSheetWriter.GetRowCount: Integer;
begin
  Result:= Length(FRowHeights) - 2*Ord(HasGrid);
end;

procedure TSheetWriter.CellIndex(var ARow, ACol: Integer);
begin
  ARow:= RowIndex(ARow);
  ACol:= ColIndex(ACol);
end;

procedure TSheetWriter.CellIndex(var ARow1, ACol1, ARow2, ACol2: Integer);
begin
  CellIndex(ARow1, ACol1);
  CellIndex(ARow2, ACol2);
end;

function TSheetWriter.RowIndex(const ARow: Integer): Integer;
begin
  Result:= ARow+FFirstRow-1;
end;

function TSheetWriter.ColIndex(const ACol: Integer): Integer;
begin
  Result:= ACol+FFirstCol-1;
end;

procedure TSheetWriter.SetRowHeight(ARow, AValue: Integer);
begin
  ARow:= RowIndex(ARow);
  SetHeight(ARow, AValue);
end;

procedure TSheetWriter.SetColWidth(ACol, AValue: Integer);
begin
  ACol:= ColIndex(ACol);
  SetWidth(ACol, AValue)
end;

procedure TSheetWriter.SetWidth(const ACol, AValue: Integer);
begin
  FColWidths[ACol]:= AValue;
  if HasGrid then
    FGrid.ColWidths[ACol]:= AValue;
  FWorksheet.WriteColWidth(ACol, WidthPxToPt(AValue), suChars);
end;

procedure TSheetWriter.SetHeight(const ARow, AValue: Integer);
begin
  FWorksheet.WriteRowHeight(ARow, HeightPxToPt(AValue), suLines);
  FRowHeights[ARow]:= AValue;
  if HasGrid then
    FGrid.RowHeights[ARow]:= AValue;
end;

procedure TSheetWriter.SetLineHeight(const ARow, AHeight: Integer; const AMinValue: Integer = ROW_HEIGHT_DEFAULT);
var
  OldRowCount, i: Integer;
begin
  OldRowCount:= Length(FRowHeights);
  if OldRowCount-FFirstRow<ARow then
  begin
    for i:= OldRowCount-FFirstRow to ARow-1 do
      SetNewRowHeight(i, ROW_HEIGHT_DEFAULT);
  end;
  SetNewRowHeight(ARow, AHeight, AMinValue);
end;

procedure TSheetWriter.SetNewRowHeight(const ARow, AHeight: Integer; const AMinValue: Integer = ROW_HEIGHT_DEFAULT);
var
  OldRowCount, HValue: Integer;
begin
  HValue:= AMinValue;
  OldRowCount:= Length(FRowHeights);
  if OldRowCount-FFirstRow>ARow then
    if FRowHeights[ARow]>HValue then
      HValue:= FRowHeights[ARow];
  if AHeight>HValue then
    HValue:= AHeight;
  if OldRowCount-FFirstRow=ARow then
  begin
    if HasGrid then
    begin
      SetGridRowCount(ARow+1+FFirstRow);
      VIns(FRowHeights, ARow, HValue);
      SetHeight(ARow+1, FRowHeights[ARow+1]);
    end
    else
      VAppend(FRowHeights, HValue);
  end;
  SetHeight(ARow, HValue);
end;

procedure TSheetWriter.SetColumns;
var
  i: Integer;
begin
  if HasGrid then
    SetGridColCount(Length(FColWidths)+1);  ///!
  for i:= 0 to High(FColWidths) do
    SetWidth(i, FColWidths[i]);
  if HasGrid then
    SetGridColCount(Length(FColWidths));
end;

procedure TSheetWriter.SetRows;
var
  i: Integer;
begin
  if HasGrid then
    SetGridRowCount(Length(FRowHeights));
  for i:= 0 to High(FRowHeights) do
    SetHeight(i, FRowHeights[i]);
end;

procedure TSheetWriter.SetFrozenRows(AFixRowCount: Integer);
begin
  if HasGrid then
  begin
    if AFixRowCount>0 then
      AFixRowCount:= AFixRowCount + 1;
    if FGrid.RowCount>=AFixRowCount then
      FGrid.FrozenRows:= AFixRowCount;
  end;
  //else
    FWorkSheet.TopPaneHeight:= AFixRowCount;
end;

procedure TSheetWriter.SetRepeatedRows(ABeginRow, AEndRow: Integer);
begin
  ABeginRow:= RowIndex(ABeginRow);
  AEndRow:= RowIndex(AEndRow);
  FWorksheet.PageLayout.SetRepeatedRows(ABeginRow, AEndRow);
end;

procedure TSheetWriter.SetRepeatedCols(ABeginCol, AEndCol: Integer);
begin
  ABeginCol:= ColIndex(ABeginCol);
  AEndCol:= ColIndex(AEndCol);
  FWorksheet.PageLayout.SetRepeatedCols(ABeginCol, AEndCol);
end;

procedure TSheetWriter.SetFrozenCols(AFixColCount: Integer);
begin
  if HasGrid then
  begin
    if AFixColCount>0 then
      AFixColCount:= AFixColCount + 1;
    if FGrid.ColCount>= AFixColCount then
      FGrid.FrozenCols:= AFixColCount;  //На гридах глюки с отображением - использовать только для экспорта
  end;
  //else
    FWorkSheet.LeftPaneWidth:= AFixColCount;
end;

procedure TSheetWriter.WriteImage(ARow, ACol: Integer; AFileName: String;
      AOffsetX: Double = 0.0; AOffsetY: Double = 0.0;
      AScaleX: Double = 1.0; AScaleY: Double = 1.0);
begin
  CellIndex(ARow, ACol);
  FWorksheet.WriteImage(ARow, ACol, AFileName, AOffsetX, AOffsetY, AScaleX, AScaleY);
end;

procedure TSheetWriter.WriteText(const ARow, ACol: Integer; const AValue: String;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AWordWrap: Boolean = True;
                        const AAutoHeight: Boolean = False;
                        const AWrapToWordParts: Boolean = False;
                        const ARedStrWidth: Integer = 0;
                        const ARichTextParams: TsRichTextParams = nil);
begin
  WriteText(ARow, ACol, ARow, ACol, AValue, ABordersType, AWordWrap, AAutoHeight,
            AWrapToWordParts, ARedStrWidth, ARichTextParams);
end;

procedure TSheetWriter.WriteText(ARow1, ACol1, ARow2, ACol2: Integer; AValue: String;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AWordWrap: Boolean = True;
                        const AAutoHeight: Boolean = False;
                        const AWrapToWordParts: Boolean = False;
                        const ARedStrWidth: Integer = 0;
                        const ARichTextParams: TsRichTextParams = nil);
var
  CellHeight: Integer;
begin
  CellIndex(ARow1, ACol1, ARow2, ACol2);
  if AValue=EmptyStr then
  begin
    SetCellSettings(ARow1, ACol1, ARow2, ACol2, ROW_HEIGHT_DEFAULT, AWordWrap, ABordersType);
    FWorksheet.WriteBlank(ARow1, ACol1);
  end
  else begin
    CellHeight:= ROW_HEIGHT_DEFAULT;
    if AAutoHeight then
      CellHeight:= Round(CalcCellHeight(AValue, ACol1, ACol2, AWrapToWordParts, ARedStrWidth)/(ARow2-ARow1+1));
    SetCellSettings(ARow1, ACol1, ARow2, ACol2, CellHeight, AWordWrap, ABordersType);
    FWorksheet.WriteText(ARow1, ACol1, AValue, ARichTextParams);
  end;
end;

procedure TSheetWriter.WriteTextVertical(const ARow, ACol: Integer;
  const AValue: String; const ABordersType: TCellBorderType;
  const AWordWrap: Boolean;
  const AWrapToWordParts: Boolean; const ARedStrWidth: Integer;
  const ARichTextParams: TsRichTextParams);
begin
  WriteTextVertical(ARow, ACol, ARow, ACol, AValue, ABordersType, AWordWrap,
            AWrapToWordParts, ARedStrWidth, ARichTextParams);
end;

procedure TSheetWriter.WriteTextVertical(ARow1, ACol1, ARow2, ACol2: Integer;
  AValue: String; const ABordersType: TCellBorderType;
  const AWordWrap: Boolean;
  const AWrapToWordParts: Boolean; const ARedStrWidth: Integer;
  const ARichTextParams: TsRichTextParams);
begin
  CellIndex(ARow1, ACol1, ARow2, ACol2);
  SetCellSettings(ARow1, ACol1, ARow2, ACol2, ROW_HEIGHT_DEFAULT, AWordWrap, ABordersType);
  FWorksheet.WriteTextRotation(ARow1, ACol1, rt90DegreeCounterClockwiseRotation);
  if AValue=EmptyStr then
    FWorksheet.WriteBlank(ARow1, ACol1)
  else
    FWorksheet.WriteText(ARow1, ACol1, AValue, ARichTextParams);
end;



procedure TSheetWriter.WriteNumber(const ARow, ACol: Integer; const AValue: Double;
                            const ABordersType: TCellBorderType = cbtNone;
                            const AFormatString: String = '');
begin
  WriteNumber(ARow, ACol, ARow, ACol, AValue, ABordersType, AFormatString);
end;

procedure TSheetWriter.WriteNumber(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: Double;
                          const ABordersType: TCellBorderType = cbtNone;
                          const AFormatString: String = '');
begin
  CellIndex(ARow1, ACol1, ARow2, ACol2);
  SetCellSettings(ARow1, ACol1, ARow2, ACol2, ROW_HEIGHT_DEFAULT, False, ABordersType);
  if AFormatString<>EmptyStr then
    FWorksheet.WriteNumberFormat(ARow1, ACol1, nfCustom, AFormatString);
  FWorksheet.WriteNumber(ARow1, ACol1, AValue);
end;

procedure TSheetWriter.WriteNumber(const ARow, ACol: Integer; const AValue: Double;
                            const ADecimals: Byte;
                            const ABordersType: TCellBorderType = cbtNone;
                            const ANumberFormat: TsNumberFormat = nfGeneral);
begin
  WriteNumber(ARow, ACol, ARow, ACol, AValue, ADecimals, ABordersType, ANumberFormat);
end;

procedure TSheetWriter.WriteNumber(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: Double;
                            const ADecimals: Byte;
                            const ABordersType: TCellBorderType = cbtNone;
                            const ANumberFormat: TsNumberFormat = nfGeneral);
begin
  CellIndex(ARow1, ACol1, ARow2, ACol2);
  SetCellSettings(ARow1, ACol1, ARow2, ACol2, ROW_HEIGHT_DEFAULT, False, ABordersType);
  FWorksheet.WriteNumber(ARow1, ACol1, AValue, ANumberFormat, ADecimals);
end;

procedure TSheetWriter.WriteCurrency(const ARow, ACol: Integer; const AValue: Double;
                          const ABordersType: TCellBorderType = cbtNone;
                          const AFormatString: String = '');
begin
  WriteCurrency(ARow, ACol, ARow, ACol, AValue, ABordersType, AFormatString);
end;

procedure TSheetWriter.WriteCurrency(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: Double;
                          const ABordersType: TCellBorderType = cbtNone;
                          const AFormatString: String = '');
begin
  CellIndex(ARow1, ACol1, ARow2, ACol2);
  SetCellSettings(ARow1, ACol1, ARow2, ACol2, ROW_HEIGHT_DEFAULT, False, ABordersType);
  if AFormatString<>EmptyStr then
    FWorksheet.WriteCurrency(ARow1, ACol1, AValue, nfCustom, AFormatString)
  else
    FWorksheet.WriteCurrency(ARow1, ACol1, AValue, nfCurrency, 2, EmptyStr);
end;

procedure TSheetWriter.WriteDate(const ARow, ACol: Integer; const AValue: TDate;
                        const ABordersType: TCellBorderType = cbtNone;
                        const AFormatString: String = '');
begin
  WriteDate(ARow, ACol, ARow, ACol, AValue, ABordersType, AFormatString);
end;

procedure TSheetWriter.WriteDate(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: TDate;
                        const ABordersType: TCellBorderType = cbtNone;
                        AFormatString: String = '');
begin
  CellIndex(ARow1, ACol1, ARow2, ACol2);
  SetCellSettings(ARow1, ACol1, ARow2, ACol2, ROW_HEIGHT_DEFAULT, False, ABordersType);
  if AFormatString=EmptyStr then
    AFormatString:= 'dd.mm.yyyy';
  FWorksheet.WriteDateTime(ARow1, ACol1, AValue, AFormatString);
end;

procedure TSheetWriter.WriteTime(const ARow, ACol: Integer; const AValue: TDateTime;
  const ABordersType: TCellBorderType; const AFormatString: String);
begin
  WriteTime(ARow, ACol, ARow, ACol, AValue, ABordersType, AFormatString);
end;

procedure TSheetWriter.WriteTime(ARow1, ACol1, ARow2, ACol2: Integer; const AValue: TDateTime;
                        const ABordersType: TCellBorderType = cbtNone;
                        AFormatString: String = '');
begin
  CellIndex(ARow1, ACol1, ARow2, ACol2);
  SetCellSettings(ARow1, ACol1, ARow2, ACol2, ROW_HEIGHT_DEFAULT, False, ABordersType);
  if AFormatString=EmptyStr then
    AFormatString:= 'hh:mm:ss';
  FWorksheet.WriteDateTime(ARow1, ACol1, AValue, AFormatString);
end;

procedure TSheetWriter.WriteDateTime(const ARow, ACol: Integer; const AValue: TDateTime;
  const ABordersType: TCellBorderType; const AFormatString: String);
begin
  WriteDateTime(ARow, ACol, ARow, ACol, AValue, ABordersType, AFormatString);
end;

procedure TSheetWriter.WriteDateTime(ARow1, ACol1, ARow2, ACol2: Integer;
  const AValue: TDateTime; const ABordersType: TCellBorderType; AFormatString: String);
begin
  CellIndex(ARow1, ACol1, ARow2, ACol2);
  SetCellSettings(ARow1, ACol1, ARow2, ACol2, ROW_HEIGHT_DEFAULT, False, ABordersType);
  if AFormatString=EmptyStr then
    AFormatString:= 'dd.mm.yyyy hh:mm:ss';
  FWorksheet.WriteDateTime(ARow1, ACol1, AValue, AFormatString);
end;

function TSheetWriter.CalcCellHeight(var AText: String; const ACol1, ACol2: Integer;
                             const AWrapToWordParts: Boolean = False;
                             const ARedStrWidth: Integer = 0): Integer;
var
  CellWidth: Integer;
  BreakSymbol: String;
  Font: TFont;

  function CalcCellWidth(ACol1, ACol2: Integer): Integer;
  var
    i: Integer;
  begin
    Result:= 0;
    for i:= ACol1 to ACol2 do
      Result:= Result + FColWidths[i];
  end;

begin
  if HasGrid then
    BreakSymbol:= SYMBOL_BREAK
  else
    BreakSymbol:= ' ';
  CellWidth:= CalcCellWidth(ACol1, ACol2);
  Font:= TFont.Create;
  Font.Name:= FFontName;
  Font.Size:= Round(FFontSize);
  Font.Style:= FontStyleSheetsToGraphics(FFontStyle);
  Result:= TextToCell(AText, Font, CellWidth, ARedStrWidth, AWrapToWordParts, BreakSymbol);
  FreeAndNil(Font);
end;

procedure TSheetWriter.SetGridRowCount(const ACount: Integer);
begin
  if FGrid.FrozenRows>ACount then
    SetFrozenRows(0);
  FGrid.RowCount:= ACount;
end;

procedure TSheetWriter.SetGridColCount(const ACount: Integer);
begin
  if FGrid.FrozenCols>ACount then
    SetFrozenCols(0);
  FGrid.ColCount:= ACount;
end;

procedure TSheetWriter.SetDefaultGridSettings;
begin
  if not HasGrid then Exit;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;
end;

procedure TSheetWriter.SetCellMainSettings(const ARow, ACol: Integer; const AWordWrap: Boolean);
begin
  FWorksheet.WriteFont(ARow, ACol, FFontName, FFontSize, FFontStyle, FFontColor);
  FWorksheet.WriteHorAlignment(ARow, ACol, FHorAlignment);
  FWorksheet.WriteVertAlignment(ARow, ACol, FVertAlignment);
  FWorksheet.WriteBackground(ARow, ACol, FBGStyle, FPatternColor, FBGColor);
  FWorksheet.WriteWordwrap(ARow, ACol, AWordWrap);
end;

procedure TSheetWriter.GetBordersNeed(const ABordersType: TCellBorderType;
                           out ALeft, ARight, ATop, ABottom, AInner: Boolean);
begin
  ALeft:= False;
  ARight:= False;
  ATop:= False;
  ABottom:= False;
  AInner:= False;
  if ABordersType=cbtNone then Exit;
  if (ABordersType=cbtAll) or (ABordersType=cbtOuter) or (ABordersType=cbtInner) then
  begin
    if (ABordersType=cbtAll) or (ABordersType=cbtOuter) then
    begin
      ALeft:= True;
      ARight:= True;
      ATop:= True;
      ABottom:= True;
    end;
    if (ABordersType=cbtAll) or (ABordersType=cbtInner) then
      AInner:= True;
  end
  else begin
    if ABordersType= cbtLeft then ALeft:= True
    else if ABordersType= cbtRight then ARight:= True
    else if ABordersType= cbtTop then ATop:= True
    else if ABordersType= cbtBottom then ABottom:= True;
  end;
end;

procedure TSheetWriter.SetCellSettings(const ARow1, ACol1, ARow2, ACol2, ARowHeight: Integer;
  const AWordWrap: Boolean; const ABordersType: TCellBorderType);
var
  ALeft, ARight, ATop, ABottom, AInner: Boolean;
  i: Integer;
begin
  for i:= ARow1 to ARow2 do
    SetLineHeight(i, ARowHeight);
  GetBordersNeed(ABordersType, ALeft, ARight, ATop, ABottom, AInner);
  if (ARow1=ARow2) and (ACol1=ACol2) then
  begin
    DrawBorders(ARow1, ACol1, ALeft, ARight, ATop, ABottom);
    if HasGrid then
    begin
      if (ACol2=High(FColWidths)-1) and ARight then
        DrawBorders(ARow1, ACol2+1, True, False, False, False);
      if (ARow2=High(FRowHeights)-1) and ABottom then
        DrawBorders(ARow1+1, ACol2, False, False, True, False);
    end;
  end
  else begin
    FWorksheet.MergeCells(ARow1, ACol1, ARow2, ACol2);
    DrawBorders(ARow1, ACol1, ARow2, ACol2, ALeft, ARight, ATop, ABottom);
    if HasGrid then
    begin
      if (ACol2=High(FColWidths)-1) and ARight then
        DrawBorders(ARow1, ACol2+1, ARow2, ACol2+1, True, False, False, False);
      if (ARow2=High(FRowHeights)-1) and ABottom then
        DrawBorders(ARow1+1, ACol1, ARow1+1, ACol2, False, False, True, False);
    end;
  end;
  SetCellMainSettings(ARow1, ACol1, AWordWrap);
end;


procedure TSheetWriter.DrawCellBorders(const ARow, ACol: Integer;
               const ALeftNeed: Boolean; const ALeftStyle: TsLineStyle; const ALeftColor: TsColor;
               const ARightNeed: Boolean; const ARightStyle: TsLineStyle; const ARightColor: TsColor;
               const ATopNeed: Boolean; const ATopStyle: TsLineStyle; const ATopColor: TsColor;
               const ABottomNeed: Boolean; const ABottomStyle: TsLineStyle; const ABottomColor: TsColor);
var
  CellBorders: TsCellBorders;

  function GetBorders(const ALeftNeed, ARightNeed, ATopNeed, ABottomNeed: Boolean): TsCellBorders;
  begin
    Result:= [];
    if ALeftNeed then Result:= Result + [cbWest];
    if ARightNeed then Result:= Result + [cbEast];
    if ATopNeed then Result:= Result + [cbNorth];
    if ABottomNeed then Result:= Result + [cbSouth];
  end;

begin
  CellBorders:= GetBorders(ALeftNeed, ARightNeed, ATopNeed, ABottomNeed);
  FWorksheet.WriteBorders(ARow, ACol, CellBorders);
  if ALeftNeed then
    FWorksheet.WriteBorderStyle(ARow,ACol, cbWest, ALeftStyle, ALeftColor);
  if ARightNeed then
    FWorksheet.WriteBorderStyle(ARow,ACol, cbEast, ARightStyle, ARightColor);
  if ATopNeed then
    FWorksheet.WriteBorderStyle(ARow,ACol, cbNorth, ATopStyle, ATopColor);
  if ABottomNeed then
    FWorksheet.WriteBorderStyle(ARow,ACol, cbSouth, ABottomStyle, ABottomColor);
end;

procedure TSheetWriter.DrawBorders(ARow, ACol: Integer; const ABordersType: TCellBorderType);
var
  ALeft, ARight, ATop, ABottom, AInner: Boolean;
begin
  CellIndex(ARow, ACol);
  GetBordersNeed(ABordersType, ALeft, ARight, ATop, ABottom, AInner);
  DrawBorders(ARow, ACol, ALeft, ARight, ATop, ABottom);
end;

procedure TSheetWriter.DrawBorders(const ARow, ACol: Integer;
                          const ALeftNeed: Boolean = True; const ARightNeed: Boolean = True;
                          const ATopNeed: Boolean = True; const ABottomNeed: Boolean = True);
begin
  DrawCellBorders(ARow, ACol,
                  ALeftNeed, FLeftBorderStyle, FLeftBorderColor,
                  ARightNeed, FRightBorderStyle, FRightBorderColor,
                  ATopNeed, FTopBorderStyle, FTopBorderColor,
                  ABottomNeed, FBottomBorderStyle, FBottomBorderColor);
end;

procedure TSheetWriter.DrawBorders(ARow1, ACol1, ARow2, ACol2: Integer;
                             const ABordersType: TCellBorderType);
var
  ALeft, ARight, ATop, ABottom, AInner: Boolean;
begin
  CellIndex(ARow1, ACol1, ARow2, ACol2);
  GetBordersNeed(ABordersType, ALeft, ARight, ATop, ABottom, AInner);
  DrawBorders(ARow1, ACol1, ARow2, ACol2, ALeft, ARight, ATop, ABottom, AInner);
end;


procedure TSheetWriter.DrawBorders(const ARow1, ACol1, ARow2, ACol2: Integer;
                          const ALeftNeed: Boolean = True; const ARightNeed: Boolean = True;
                          const ATopNeed: Boolean = True; const ABottomNeed: Boolean = True;
                          const AInnerNeed: Boolean = False);
var
  i,j: Integer;
begin
  {одна ячейка }
  if (ARow1=ARow2) and (ACol1=ACol2) then
  begin
    DrawBorders(ARow1, ACol1, ALeftNeed, ARightNeed, ATopNeed, ABottomNeed);
    Exit;
  end;
  {одна строка}
  if ARow1=ARow2 then
  begin
    //внутренние ячейки
    for i:=ACol1+1 to ACol2-1 do
      DrawCellBorders(ARow1, i,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ATopNeed, FTopBorderStyle, FTopBorderColor,
                      ABottomNeed, FBottomBorderStyle, FBottomBorderColor);
    //первая и последняя ячейка строки
    DrawCellBorders(ARow1, ACol1,
                      ALeftNeed, FLeftBorderStyle, FLeftBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ATopNeed, FTopBorderStyle, FTopBorderColor,
                      ABottomNeed, FBottomBorderStyle, FBottomBorderColor);
    DrawCellBorders(ARow1, ACol2,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ARightNeed, FRightBorderStyle, FRightBorderColor,
                      ATopNeed, FTopBorderStyle, FTopBorderColor,
                      ABottomNeed, FBottomBorderStyle, FBottomBorderColor);
    Exit;
  end;
  {один столбец}
  if ACol1=ACol2 then
  begin
    //внутренние ячейки
    for i:=ARow1+1 to ARow2-1 do
      DrawCellBorders(i, ACol1,
                      ALeftNeed, FLeftBorderStyle, FLeftBorderColor,
                      ARightNeed, FRightBorderStyle, FRightBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor);
    //первая и последняя ячейка столбца
    DrawCellBorders(ARow1, ACol1,
                      ALeftNeed, FLeftBorderStyle, FLeftBorderColor,
                      ARightNeed, FRightBorderStyle, FRightBorderColor,
                      ATopNeed, FTopBorderStyle, FTopBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor);
    DrawCellBorders(ARow2, ACol1,
                      ALeftNeed, FLeftBorderStyle, FLeftBorderColor,
                      ARightNeed, FRightBorderStyle, FRightBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ABottomNeed, FBottomBorderStyle, FBottomBorderColor);
    Exit;
  end;
  {несколько строк и несколько столбцов}
  //внутренние
  for i:= ARow1+1 to ARow2-1 do
    for j:= ACol1+1 to ACol2-1 do
      DrawCellBorders(i, j,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor);
  //угловые ячейки
  DrawCellBorders(ARow1, ACol1,
                      ALeftNeed, FLeftBorderStyle, FLeftBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ATopNeed, FTopBorderStyle, FTopBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor);
  DrawCellBorders(ARow1, ACol2,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ARightNeed, FRightBorderStyle, FRightBorderColor,
                      ATopNeed, FTopBorderStyle, FTopBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor);
  DrawCellBorders(ARow2, ACol1,
                      ALeftNeed, FLeftBorderStyle, FLeftBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ABottomNeed, FBottomBorderStyle, FBottomBorderColor);
  DrawCellBorders(ARow2, ACol2,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ARightNeed, FRightBorderStyle, FRightBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ABottomNeed, FBottomBorderStyle, FBottomBorderColor);
  //граничные ячейки
  for i:= ARow1+1 to ARow2-1 do
    DrawCellBorders(i, ACol1,
                      ALeftNeed, FLeftBorderStyle, FLeftBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor);
  for i:= ARow1+1 to ARow2-1 do
    DrawCellBorders(i, ACol2,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ARightNeed, FRightBorderStyle, FRightBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor);
  for i:= ACol1+1 to ACol2-1 do
    DrawCellBorders(ARow1, i,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ATopNeed, FTopBorderStyle, FTopBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor);
  for i:= ACol1+1 to ACol2-1 do
    DrawCellBorders(ARow2, i,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      AInnerNeed, FInnerBorderStyle, FInnerBorderColor,
                      ABottomNeed, FBottomBorderStyle, FBottomBorderColor);
end;

procedure TSheetWriter.SetBordersStyle(const ALeftStyle,ARightStyle,ATopStyle,ABottomStyle,AInnerStyle: TsLineStyle);
begin
  FLeftBorderStyle:= ALeftStyle;
  FRightBorderStyle:= ARightStyle;
  FTopBorderStyle:= ATopStyle;
  FBottomBorderStyle:= ABottomStyle;
  FInnerBorderStyle:= AInnerStyle;
end;

procedure TSheetWriter.SetBordersColor(const ALeftColor,ARightColor,ATopColor,ABottomColor,AInnerColor: TColor);
begin
  SetBordersColor(ColorGraphicsToSheets(ALeftColor),
                  ColorGraphicsToSheets(ARightColor),
                  ColorGraphicsToSheets(ATopColor),
                  ColorGraphicsToSheets(ABottomColor),
                  ColorGraphicsToSheets(AInnerColor));
end;

procedure TSheetWriter.SetBordersColor(const AAllColor: TColor);
begin
  SetBordersColor(AAllColor, AAllColor, AAllColor, AAllColor, AAllColor);
end;

procedure TSheetWriter.SetBordersColor(const AAllColor: TsColor);
begin
  SetBordersColor(AAllColor, AAllColor, AAllColor, AAllColor, AAllColor);
end;

procedure TSheetWriter.SetBordersColor(const ALeftColor,ARightColor,ATopColor,ABottomColor,AInnerColor: TsColor);
begin
  FLeftBorderColor:= ALeftColor;
  FRightBorderColor:= ARightColor;
  FTopBorderColor:= ATopColor;
  FBottomBorderColor:= ABottomColor;
  FInnerBorderColor:= AInnerColor;
end;

procedure TSheetWriter.SetBorders(const AAllStyle: TsLineStyle; const AAllColor: TsColor);
begin
  SetBordersStyle(AAllStyle, AAllStyle, AAllStyle, AAllStyle, AAllStyle);
  SetBordersColor(AAllColor, AAllColor, AAllColor, AAllColor, AAllColor);
end;

procedure TSheetWriter.SetBorders(const AOuterStyle: TsLineStyle; const AOuterColor: TsColor;
                           const AInnerStyle: TsLineStyle; const AInnerColor: TsColor);
begin
  SetBordersStyle(AOuterStyle, AOuterStyle, AOuterStyle, AOuterStyle, AInnerStyle);
  SetBordersColor(AOuterColor, AOuterColor, AOuterColor, AOuterColor, AInnerColor);
end;

procedure TSheetWriter.SetBorders(const AAllStyle: TsLineStyle; const AAllColor: TColor);
begin
  SetBorders(AAllStyle, ColorGraphicsToSheets(AAllColor));
end;

procedure TSheetWriter.SetBorders(const AOuterStyle: TsLineStyle; const AOuterColor: TColor;
                           const AInnerStyle: TsLineStyle; const AInnerColor: TColor);
var
  OuterColor, InnerColor: TsColor;
begin
  OuterColor:= ColorGraphicsToSheets(AOuterColor);
  InnerColor:= ColorGraphicsToSheets(AInnerColor);
  SetBorders(AOuterStyle, OuterColor, AInnerStyle, InnerColor);
end;

procedure TSheetWriter.SetBordersDefault;
begin
  SetBorders(BORDER_STYLE_DEFAULT, BORDER_COLOR_DEFAULT);
end;

procedure TSheetWriter.SetAlignmentDefault;
begin
  SetAlignment(ALIGN_HOR_DEFAULT, ALIGN_VERT_DEFAULT);
end;

procedure TSheetWriter.SetAlignment(const AHorAlignment: TsHorAlignment;
                             const AVertAlignment: TsVertAlignment);
begin
  FHorAlignment:= AHorAlignment;
  FVertAlignment:= AVertAlignment;
end;

procedure TSheetWriter.SetBackgroundClear;
begin
   SetBackground(fsNoFill, scTransparent, scTransparent);
end;

procedure TSheetWriter.SetBackgroundDefault;
begin
  SetBackground(BG_STYLE_DEFAULT, BG_COLOR_DEFAULT, PATTERN_COLOR_DEFAULT);
end;

procedure TSheetWriter.SetBackground(const ABGStyle: TsFillStyle; const ABGColor: TsColor;
  const APatternColor: TsColor);
begin
  FBGStyle:= ABGStyle;
  FBGColor:= ABGColor;
  FPatternColor:= APatternColor;
end;

procedure TSheetWriter.SetBackground(const ABGStyle: TsFillStyle; const ABGColor: TColor;
  const APatternColor: TColor);
begin
  SetBackground(ABGStyle, ColorGraphicsToSheets(ABGColor),
                ColorGraphicsToSheets(APatternColor));
end;

procedure TSheetWriter.SetBackground(const ABGColor: TsColor);
begin
  SetBackground(fsSolidFill, ABGColor, ABGColor);
end;

procedure TSheetWriter.SetBackground(const ABGColor: TColor);
begin
  SetBackground(fsSolidFill, ABGColor, ABGColor);
end;

procedure TSheetWriter.SetFontDefault;
begin
  SetFont(FONT_NAME_DEFAULT, FONT_SIZE_DEFAULT,
          FONT_STYLE_DEFAULT, FONT_COLOR_DEFAULT);
end;

procedure TSheetWriter.SetFontName(const AName: String);
begin
  if STrim(AName)=EmptyStr then
    FFontName:= FONT_NAME_DEFAULT
  else
    FFontName:= AName;
end;

procedure TSheetWriter.SetFont(const AName: String; const ASize: Single;
                        const AStyle: TsFontStyles; const AColor: TsColor);
begin
  SetFontName(AName);
  FFontSize:= ASize;
  FFontStyle:= AStyle;
  FFontColor:= AColor;
end;

procedure TSheetWriter.SetFont(const AName: String; const ASize: Single; const AStyle: TFontStyles;
  const AColor: TColor);
begin
  SetFontName(AName);
  FFontSize:= ASize;
  FFontStyle:= FontStyleGraphicsToSheets(AStyle);
  FFontColor:= ColorGraphicsToSheets(AColor);
end;

procedure TSheetWriter.SetFont(const AFont: TFont);
begin
  SetFont(AFont.Name, AFont.Size, AFont.Style, AFont.Color);
end;

end.

