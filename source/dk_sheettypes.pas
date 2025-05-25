unit DK_SheetTypes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Graphics, fpstypes, fpspreadsheetgrid,
  DK_SheetConst, DK_SheetUtils, DK_SheetWriter, DK_SheetExporter,
  DK_Vector, DK_Color;

const
  haLeft    = fpsTypes.haLeft;
  haCenter  = fpsTypes.haCenter;
  haRight   = fpsTypes.haRight;
  haDefault = fpsTypes.haDefault;
  vaTop     = fpsTypes.vaTop;
  vaCenter  = fpsTypes.vaCenter;
  vaBottom  = fpsTypes.vaBottom;
  vaDefault = fpsTypes.vaDefault;

  cbtNone   = DK_SheetWriter.cbtNone;
  cbtLeft   = DK_SheetWriter.cbtLeft;
  cbtRight  = DK_SheetWriter.cbtRight;
  cbtTop    = DK_SheetWriter.cbtTop;
  cbtBottom = DK_SheetWriter.cbtBottom;
  cbtOuter  = DK_SheetWriter.cbtOuter;
  cbtInner  = DK_SheetWriter.cbtInner;
  cbtAll    = DK_SheetWriter.cbtAll;

type
  TCellBorderType = DK_SheetWriter.TCellBorderType;
  TSheetEvent = procedure of Object;
  TZoomEvent = procedure(const AZoomPercent: Integer) of Object;

  { TCustomSheet }

  TCustomSheet = class (TObject)
  protected
    FSelectedRows: TIntVector;
    FSelectedCols: TIntVector;
    FSelectedExtraRows: TIntVector;
    FSelectedExtraCols: TIntVector;
    function SetWidths: TIntVector; virtual; abstract;
  private
    FWriter: TSheetWriter;
    FFont: TFont;
    FColorVector: TColorVector;
    FColorIsNeed: Boolean;

    procedure SetColorIsNeed(AValue: Boolean);
    procedure SetFont(AValue: TFont);
    function GetCellColor(const ARow, ACol: Integer): TsColor;

    function GetCellSelectionIndex(const ARow, ACol: Integer): Integer;
    function GetCellSelectionExtraIndex(const ARow, ACol: Integer): Integer;

    function GetShowFrozenLine: Boolean;
    procedure SetShowFrozenLine(AValue: Boolean);
  public
    constructor Create(const AWorksheet: TsWorksheet; const AGrid: TsWorksheetGrid;
                       const AFont: TFont; const ARowHeightDefault: Integer = ROW_HEIGHT_DEFAULT);
    destructor  Destroy; override;

    procedure Clear; virtual;

    procedure Zoom(const APercents: Integer);
    procedure SetFontDefault;
    procedure Save(const ASheetName: String = 'Лист1';
                   const ADoneMessage: String = 'Выполнено!';
                   const ALandscape: Boolean = False);

    procedure BordersDraw(const ARow1: Integer = 0; const ARow2: Integer = 0;
                          const ACol1: Integer = 0; const ACol2: Integer = 0);
    procedure ColorsUpdate(const AColorVector: TColorVector);
    procedure ColorsClear;
    property  ColorIsNeed: Boolean read FColorIsNeed write SetColorIsNeed;

    procedure SelectionAddCell(const ARow, ACol: Integer);
    procedure SelectionDelCell(const ARow, ACol: Integer);
    procedure SelectionClear; virtual;

    procedure SelectionExtraAddCell(const ARow, ACol: Integer);
    procedure SelectionExtraDelCell(const ARow, ACol: Integer);
    procedure SelectionExtraClear;

    property Font: TFont read FFont write SetFont;
    property Writer: TSheetWriter read FWriter;
    property ShowFrozenLine: Boolean read GetShowFrozenLine write SetShowFrozenLine;
  end;

  {correct saving sheet with zoom<>100%}
  procedure SheetFromGridSave(const ASheet: TCustomSheet;
                   const AZoomPercent: Integer;
                   const ADrawProc: TZoomEvent;
                   const ASheetName: String = 'Лист1';
                   const ADoneMessage: String = 'Выполнено!';
                   const ALandscape: Boolean = False);

implementation

procedure SheetFromGridSave(const ASheet: TCustomSheet;
                   const AZoomPercent: Integer;
                   const ADrawProc: TZoomEvent;
                   const ASheetName: String = 'Лист1';
                   const ADoneMessage: String = 'Выполнено!';
                   const ALandscape: Boolean = False);
var
  Percent: Integer;
begin
  Percent:= AZoomPercent;
  if Percent<>100 then
    ADrawProc(100);
  ASheet.Save(ASheetName, ADoneMessage, ALandscape);
  if Percent<>100 then
    ADrawProc(Percent);
end;

{ TCustomSheet }

procedure TCustomSheet.SetFont(AValue: TFont);
begin
  if not Assigned(AValue) then Exit;
  FFont.Assign(AValue);
  if FFont.Size<FONT_SIZE_MINIMUM then
    FFont.Size:= FONT_SIZE_MINIMUM;
end;

procedure TCustomSheet.SetColorIsNeed(AValue: Boolean);
begin
  if FColorIsNeed=AValue then Exit;
  FColorIsNeed:= AValue;
  if FColorIsNeed and (not VIsNil(FColorVector)) then
    ColorsUpdate(FColorVector)
  else
    ColorsClear;
end;

function TCustomSheet.GetCellColor(const ARow, ACol: Integer): TsColor;
var
  i, ColorIndex: Integer;
begin
  Result:= scTransparent;
  if not ColorIsNeed then Exit;
  if VIsNil(FColorVector) then Exit;

  for i:= 0 to High(Writer.BGColorMatrix) do
  begin
    if (Writer.BGColorMatrix[i,0]=ARow) and (Writer.BGColorMatrix[i,1]=ACol) then
    begin
      ColorIndex:= Writer.BGColorMatrix[i,2];
      if (ColorIndex<>TRANSPARENT_COLOR_INDEX) and (ColorIndex<=High(FColorVector)) then
        Result:= ColorGraphicsToSheets(FColorVector[ColorIndex]);
    end;
  end;
end;

function TCustomSheet.GetCellSelectionIndex(const ARow, ACol: Integer): Integer;
begin
  Result:= VIndexOf(FSelectedRows, FSelectedCols, ARow, ACol);
end;

function TCustomSheet.GetCellSelectionExtraIndex(const ARow, ACol: Integer): Integer;
begin
  Result:= VIndexOf(FSelectedExtraRows, FSelectedExtraCols, ARow, ACol);
end;

function TCustomSheet.GetShowFrozenLine: Boolean;
begin
  Result:= Writer.ShowFrozenLine;
end;

procedure TCustomSheet.SetShowFrozenLine(AValue: Boolean);
begin
  Writer.ShowFrozenLine:= AValue;
end;

constructor TCustomSheet.Create(const AWorksheet: TsWorksheet; const AGrid: TsWorksheetGrid;
                                const AFont: TFont; const ARowHeightDefault: Integer = ROW_HEIGHT_DEFAULT);
begin
  FFont:= TFont.Create;
  if Assigned(AFont) then
    Font:= AFont
  else
    SetFontDefault;
  FColorIsNeed:= True;
  FWriter:= TSheetWriter.Create(SetWidths, AWorksheet, AGrid, ARowHeightDefault);
end;

destructor TCustomSheet.Destroy;
begin
  FreeAndNil(FFont);
  FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TCustomSheet.Clear;
begin
  ColorsClear;
  SelectionExtraClear;
  SelectionClear;
  FWriter.Clear;
end;

procedure TCustomSheet.Zoom(const APercents: Integer);
begin
  FWriter.SetZoom(APercents);
end;

procedure TCustomSheet.SetFontDefault;
begin
  FFont.Name:= FONT_NAME_DEFAULT;
  FFont.Size:= FONT_SIZE_DEFAULT;
  FFont.Style:= [];
  FFont.Color:= FONT_COLOR_DEFAULT;
end;

procedure TCustomSheet.Save(const ASheetName: String = 'Лист1';
                   const ADoneMessage: String = 'Выполнено!';
                   const ALandscape: Boolean = False);
var
  Exporter: TSheetExporter;
begin
  Exporter:= TSheetExporter.Create(Writer.Worksheet);
  try
    Exporter.SheetName:= ASheetName;
    if ALandscape then Exporter.PageSettings(spoLandscape);
    Exporter.Save(ADoneMessage);
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TCustomSheet.BordersDraw(const ARow1: Integer = 0; const ARow2: Integer = 0;
                                   const ACol1: Integer = 0; const ACol2: Integer = 0);
var
  i, j, R1, R2, C1, C2: Integer;
begin
  R1:= 1;
  if ARow1>R1 then R1:= ARow1;
  R2:= FWriter.RowCount;
  if (ARow2>0) and (ARow2<R2) then R2:= ARow2;

  C1:= 1;
  if ACol1>C1 then C1:= ACol1;
  C2:= FWriter.ColCount;
  if (ACol2>0) and (ACol2<C2) then C2:= ACol2;

  for i:= R1 to R2 do
    for j:= C1 to C2 do
      FWriter.DrawBorders(i, j, cbtOuter);
end;

procedure TCustomSheet.ColorsUpdate(const AColorVector: TColorVector);
begin
  FColorVector:= AColorVector;
  FWriter.ApplyBGColors(FColorVector);
end;

procedure TCustomSheet.ColorsClear;
begin
  FWriter.ClearBGColors;
end;

procedure TCustomSheet.SelectionAddCell(const ARow, ACol: Integer);
var
  CellSelectionIndex: Integer;
  Cl: TsColor;
begin
  CellSelectionIndex:= GetCellSelectionIndex(ARow, ACol);
  if CellSelectionIndex>=0 then Exit;
  VAppend(FSelectedRows, ARow);
  VAppend(FSelectedCols, ACol);
  Cl:= DefaultSelectionBGColor;
  FWriter.Worksheet.WriteBackground(ARow, ACol, fsSolidFill, Cl, Cl);
end;

procedure TCustomSheet.SelectionDelCell(const ARow, ACol: Integer);
var
  CellSelectionIndex: Integer;
  Cl: TsColor;
begin
  CellSelectionIndex:= GetCellSelectionIndex(ARow, ACol);
  if CellSelectionIndex<0 then Exit;
  VDel(FSelectedRows, CellSelectionIndex);
  VDel(FSelectedCols, CellSelectionIndex);
  Cl:= GetCellColor(ARow, ACol);
  FWriter.Worksheet.WriteBackground(ARow, ACol, fsSolidFill, Cl, Cl);
end;

procedure TCustomSheet.SelectionClear;
var
  i: Integer;
  R, C: TIntVector;
begin
  R:= VCut(FSelectedRows);
  C:= VCut(FSelectedCols);
  for i:= 0 to High(R) do
    SelectionDelCell(R[i], C[i]);
end;

procedure TCustomSheet.SelectionExtraAddCell(const ARow, ACol: Integer);
var
  CellSelectionIndex: Integer;
  Cl: TsColor;
begin
  CellSelectionIndex:= GetCellSelectionExtraIndex(ARow, ACol);
  if CellSelectionIndex>=0 then Exit;
  VAppend(FSelectedExtraRows, ARow);
  VAppend(FSelectedExtraCols, ACol);
  Cl:= DefaultSelectionBGExtraColor;
  FWriter.Worksheet.WriteBackground(ARow, ACol, fsSolidFill, Cl, Cl);
end;

procedure TCustomSheet.SelectionExtraDelCell(const ARow, ACol: Integer);
var
  CellSelectionIndex: Integer;
  Cl: TsColor;
begin
  CellSelectionIndex:= GetCellSelectionExtraIndex(ARow, ACol);
  if CellSelectionIndex<0 then Exit;
  VDel(FSelectedExtraRows, CellSelectionIndex);
  VDel(FSelectedExtraCols, CellSelectionIndex);
  CellSelectionIndex:= GetCellSelectionIndex(ARow, ACol);
  if CellSelectionIndex>=0 then
    Cl:= DefaultSelectionBGColor
  else
    Cl:= GetCellColor(ARow, ACol);
  FWriter.Worksheet.WriteBackground(ARow, ACol, fsSolidFill, Cl, Cl);
end;

procedure TCustomSheet.SelectionExtraClear;
var
  i: Integer;
  R, C: TIntVector;
begin
  R:= VCut(FSelectedExtraRows);
  C:= VCut(FSelectedExtraCols);
  for i:= 0 to High(R) do
    SelectionExtraDelCell(R[i], C[i]);
end;



end.

