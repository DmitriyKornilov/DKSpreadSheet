unit DK_SheetTypes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Graphics, Controls, lcltype, fpstypes, fpspreadsheetgrid,
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

  spoPortrait  = fpsTypes.spoPortrait;
  spoLandscape = fpsTypes.spoLandscape;

  nfGeneral       = fpsTypes.nfGeneral;
  nfFixed         = fpsTypes.nfFixed;
  nfFixedTh       = fpsTypes.nfFixedTh;
  nfExp           = fpsTypes.nfExp;
  nfPercentage    = fpsTypes.nfPercentage;
  nfFraction      = fpsTypes.nfFraction;
  nfCurrency      = fpsTypes.nfCurrency;
  nfCurrencyRed   = fpsTypes.nfCurrencyRed;
  nfShortDateTime = fpsTypes.nfShortDateTime;
  nfShortDate     = fpsTypes.nfShortDate;
  nfLongDate      = fpsTypes.nfLongDate;
  nfShortTime     = fpsTypes.nfShortTime;
  nfLongTime      = fpsTypes.nfLongTime;
  nfShortTimeAM   = fpsTypes.nfShortTimeAM;
  nfLongTimeAM    = fpsTypes.nfLongTimeAM;
  nfDayMonth      = fpsTypes.nfDayMonth;
  nfMonthYear     = fpsTypes.nfMonthYear;
  nfTimeInterval  = fpsTypes.nfTimeInterval;
  nfText          = fpsTypes.nfText;
  nfCustom        = fpsTypes.nfCustom;

  cbtNone   = DK_SheetWriter.cbtNone;
  cbtLeft   = DK_SheetWriter.cbtLeft;
  cbtRight  = DK_SheetWriter.cbtRight;
  cbtTop    = DK_SheetWriter.cbtTop;
  cbtBottom = DK_SheetWriter.cbtBottom;
  cbtOuter  = DK_SheetWriter.cbtOuter;
  cbtInner  = DK_SheetWriter.cbtInner;
  cbtAll    = DK_SheetWriter.cbtAll;

  pfNone    = DK_SheetExporter.pfNone;
  pfOnePage = DK_SheetExporter.pfOnePage;
  pfWidth   = DK_SheetExporter.pfWidth;
  pfHeight  = DK_SheetExporter.pfHeight;

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

    procedure MouseWheel(Sender: TObject; {%H-}Shift: TShiftState;
                         {%H-}WheelDelta: Integer; {%H-}MousePos: TPoint;
                         var {%H-}Handled: Boolean);
  public
    constructor Create(const AWorksheet: TsWorksheet; const AGrid: TsWorksheetGrid;
                       const AFont: TFont; const ARowHeightDefault: Integer = ROW_HEIGHT_DEFAULT);
    destructor  Destroy; override;

    procedure Clear; virtual;

    procedure Zoom(const APercents: Integer);
    procedure SetFontDefault;
    procedure Save(const ASheetName: String = 'Лист1';
                   const ADoneMessage: String = 'Выполнено!';
                   const AOrient: TsPageOrientation = spoPortrait;
                   const APageFit: TPageFit = pfWidth;
                   const AShowHeaders: Boolean = True;
                   const AShowGridLines: Boolean = True;
                   const ATopMargin: Double=10;     //mm
                   const ALeftMargin: Double=10;    //mm
                   const ARightMargin: Double=10;   //mm
                   const ABottomMargin: Double=10;  //mm
                   const AFooterMargin: Double=0;   //mm
                   const AHeaderMargin: Double=0);

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

type

  { TCustomSelectableSheet }

  TCustomSelectableSheet = class(TCustomSheet)
  private
    FAutosizeColumnNumber: Integer;
    FColumnWidthBeforeAutosize: Integer;
    FColumnWidthWithoutAutosize: Integer;

    FOnSelect: TSheetEvent;
    FCanSelect: Boolean;
    FCanUnselect: Boolean;

    FOnDblClick: TSheetEvent;
    FOnReturnKeyDown: TSheetEvent;
    FOnDelKeyDown: TSheetEvent;

    procedure SetCanSelect(const AValue: Boolean);
    procedure SetCanUnselect(const AValue: Boolean);

    procedure ChangeBounds(Sender: TObject);
    procedure MouseDown(Sender: TObject; Button: TMouseButton;
                        {%H-}Shift: TShiftState; X, Y: Integer);
    procedure KeyDown(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
    procedure DblClick(Sender: TObject);
  protected
    function IsCellSelectable(const {%H-}ARow, ACol: Integer): Boolean; virtual;
    procedure SetSelection(const ARow, ACol: Integer); virtual; abstract;
    procedure DelSelection; virtual; abstract;
    function GetIsSelected: Boolean; virtual; abstract;
    procedure SelectionMove(const ADelta: Integer); virtual; abstract;
  public
    constructor Create(const AWorksheet: TsWorksheet;
                       const AGrid: TsWorksheetGrid;
                       const AFont: TFont;
                       const ARowHeightDefault: Integer = ROW_HEIGHT_DEFAULT);

    procedure AutosizeColumnEnable(const AColNumber: Integer);
    procedure AutosizeColumnEnableLast;
    procedure AutosizeColumnDisable;
    procedure AutoSizeColumnWidth;

    procedure Select(const ARow, ACol: Integer; const ADoEvent: Boolean = True);
    procedure Unselect(const ADoEvent: Boolean = True);

    property CanSelect: Boolean read FCanSelect write SetCanSelect;
    property CanUnselect: Boolean read FCanUnselect write SetCanUnselect;
    property IsSelected: Boolean read GetIsSelected;

    property OnSelect: TSheetEvent read FOnSelect write FOnSelect;
    property OnReturnKeyDown: TSheetEvent read FOnReturnKeyDown write FOnReturnKeyDown;
    property OnDelKeyDown: TSheetEvent read FOnDelKeyDown write FOnDelKeyDown;
    property OnDblClick: TSheetEvent read FOnDblClick write FOnDblClick;
  end;

  { TSingleSelectableSheet }

  TSingleSelectableSheet = class(TCustomSelectableSheet)
  protected
    FSelectedIndex: Integer;
    FFirstRows: TIntVector;
    FLastRows: TIntVector;

    function IsCellSelectable(const ARow, ACol: Integer): Boolean; override;
    function GetIsSelected: Boolean; override;
    procedure SetSelection(const ARow, {%H-}ACol: Integer); override;
    procedure DelSelection; override;
    procedure SelectionMove(const ADelta: Integer); override;
  public
    constructor Create(const AWorksheet: TsWorksheet;
                       const AGrid: TsWorksheetGrid;
                       const AFont: TFont;
                       const ARowHeightDefault: Integer = ROW_HEIGHT_DEFAULT);

    function ReSelect(const AIDVector: TIntVector;
                      const AIDValue: Integer;
                      const AFirstRowSelectIfNotFound: Boolean = False): Boolean;

    property SelectedIndex: Integer read FSelectedIndex;
  end;

implementation

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

procedure TCustomSheet.MouseWheel(Sender: TObject; Shift: TShiftState;
                         WheelDelta: Integer; MousePos: TPoint;
                         var Handled: Boolean);
begin
  (Sender as TsWorksheetGrid).Invalidate;
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
  if Assigned(AGrid) then
    FWriter.Grid.OnMouseWheel:= @MouseWheel;
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
                   const AOrient: TsPageOrientation = spoPortrait;
                   const APageFit: TPageFit = pfWidth;
                   const AShowHeaders: Boolean = True;
                   const AShowGridLines: Boolean = True;
                   const ATopMargin: Double=10;     //mm
                   const ALeftMargin: Double=10;    //mm
                   const ARightMargin: Double=10;   //mm
                   const ABottomMargin: Double=10;  //mm
                   const AFooterMargin: Double=0;   //mm
                   const AHeaderMargin: Double=0);
var
  Exporter: TSheetExporter;
begin
  Exporter:= TSheetExporter.Create(Writer.Worksheet);
  if Writer.HasGrid then
    Writer.Grid.Visible:= False;
  try
    Exporter.SheetName:= ASheetName;
    Exporter.PageSettings(AOrient, APageFit, AShowHeaders, AShowGridLines,
                          ATopMargin, ALeftMargin, ARightMargin, ABottomMargin,
                          AFooterMargin, AHeaderMargin);
    Exporter.Save(ADoneMessage);
  finally
    FreeAndNil(Exporter);
    Writer.EndEdit(False);
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

{ TCustomSelectableSheet }

procedure TCustomSelectableSheet.MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  R, C: Integer;
begin
  if Button=mbLeft then
  begin
    if CanSelect then
    begin
      (Sender as TsWorksheetGrid).MouseToCell(X, Y, C, R);
      Select(R, C);
    end;
  end
  else if Button=mbRight then
  begin
    if CanUnselect then
      Unselect;
  end;
end;

procedure TCustomSelectableSheet.KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_UP: SelectionMove(-1);
    VK_DOWN: SelectionMove(1);
    VK_RETURN: if Assigned(FOnReturnKeyDown) then FOnReturnKeyDown;
    VK_DELETE: if Assigned(FOnDelKeyDown) then FOnDelKeyDown;
  end;
end;

procedure TCustomSelectableSheet.DblClick(Sender: TObject);
begin
  if not IsCellSelectable((Sender as TsWorksheetGrid).Row,
                          (Sender as TsWorksheetGrid).Col) then Exit;
  if Assigned(FOnDblClick) then FOnDblClick;
end;

function TCustomSelectableSheet.IsCellSelectable(const ARow, ACol: Integer): Boolean;
begin
  Result:= (ACol>=1) and (ACol<=Writer.ColCount);
end;

procedure TCustomSelectableSheet.Select(const ARow, ACol: Integer;
  const ADoEvent: Boolean);
begin
  if not IsCellSelectable(ARow, ACol) then Exit;

  if IsSelected then DelSelection;
  SetSelection(ARow, ACol);
  if ADoEvent and Assigned(FOnSelect) then FOnSelect;
end;

procedure TCustomSelectableSheet.Unselect(const ADoEvent: Boolean);
begin
  if not IsSelected then Exit;
  DelSelection;
  if ADoEvent and Assigned(FOnSelect) then FOnSelect;
end;

procedure TCustomSelectableSheet.SetCanSelect(const AValue: Boolean);
begin
  if FCanSelect=AValue then Exit;
  if not AValue then Unselect;
  FCanSelect:=AValue;
end;

procedure TCustomSelectableSheet.SetCanUnselect(const AValue: Boolean);
begin
  if FCanUnselect=AValue then Exit;
  FCanUnselect:= AValue;
end;

procedure TCustomSelectableSheet.ChangeBounds(Sender: TObject);
begin
  AutoSizeColumnWidth;
end;

constructor TCustomSelectableSheet.Create(const AWorksheet: TsWorksheet;
                       const AGrid: TsWorksheetGrid;
                       const AFont: TFont;
                       const ARowHeightDefault: Integer = ROW_HEIGHT_DEFAULT);
begin
  inherited Create(AWorksheet, AGrid, AFont, ARowHeightDefault);

  FAutosizeColumnNumber:= AUTOSIZE_NONE_COLUMN_NUMBER;
  FOnSelect:= nil;
  FCanSelect:= False;
  FCanUnselect:= True;

  if Assigned(AGrid) then
  begin
    Writer.Grid.OnMouseDown:= @MouseDown;
    Writer.Grid.OnChangeBounds:= @ChangeBounds;
    Writer.Grid.OnKeyDown:= @KeyDown;
    Writer.Grid.OnDblClick:= @DblClick;
  end;
end;

procedure TCustomSelectableSheet.AutoSizeColumnWidth;
var
  W: Integer;
begin
  if not Writer.HasGrid then Exit;
  if FAutosizeColumnNumber=AUTOSIZE_NONE_COLUMN_NUMBER then Exit;

  W:= Writer.Grid.Width -
      Writer.Grid.Scale96ToScreen(FColumnWidthWithoutAutosize+AUTOSIZE_ADDITION_WIDTH);
  if W<0 then
    W:= FColumnWidthBeforeAutosize
  else
    W:= Writer.Grid.ScaleScreenTo96(W);
  Writer.ColWidth[FAutosizeColumnNumber]:= W;
end;

procedure TCustomSelectableSheet.AutosizeColumnEnable(const AColNumber: Integer);
begin
  if not Writer.HasGrid then Exit;
  if FAutosizeColumnNumber=AColNumber then Exit;
  if ((AColNumber<1) and (AColNumber<>AUTOSIZE_LAST_COLUMN_NUMBER)) or
     (AColNumber>Writer.ColCount) then Exit;

  if AColNumber=AUTOSIZE_LAST_COLUMN_NUMBER then
    FAutosizeColumnNumber:= Writer.ColCount
  else
    FAutosizeColumnNumber:= AColNumber;
  FColumnWidthBeforeAutosize:= Writer.ColWidth[FAutosizeColumnNumber];
  FColumnWidthWithoutAutosize:= Writer.ColsWidth(1, Writer.ColCount) -
                                  FColumnWidthBeforeAutosize;

  AutoSizeColumnWidth;
end;

procedure TCustomSelectableSheet.AutosizeColumnEnableLast;
begin
  AutosizeColumnEnable(AUTOSIZE_LAST_COLUMN_NUMBER);
end;

procedure TCustomSelectableSheet.AutosizeColumnDisable;
var
  ColNum: Integer;
begin
  if not Writer.HasGrid then Exit;
  if FAutosizeColumnNumber= AUTOSIZE_NONE_COLUMN_NUMBER then Exit;

  if FAutosizeColumnNumber=AUTOSIZE_LAST_COLUMN_NUMBER then
     ColNum:= Writer.ColCount
  else
     ColNum:= FAutosizeColumnNumber;

  Writer.ColWidth[ColNum]:= FColumnWidthBeforeAutosize;
  FAutosizeColumnNumber:= AUTOSIZE_NONE_COLUMN_NUMBER;
end;

{ TSingleSelectableSheet }

function TSingleSelectableSheet.IsCellSelectable(const ARow, ACol: Integer): Boolean;
begin
  Result:= inherited IsCellSelectable(ARow, ACol) and
           (VIndexOf(FFirstRows, FLastRows, ARow)>=0);
end;

function TSingleSelectableSheet.GetIsSelected: Boolean;
begin
  Result:= FSelectedIndex>=0;
end;

procedure TSingleSelectableSheet.SetSelection(const ARow, ACol: Integer);
var
  i, j, k: Integer;
begin
  k:= VIndexOf(FFirstRows, FLastRows, ARow);
  FSelectedIndex:= k;
  for i:= FFirstRows[k] to FLastRows[k] do
    for j:= 1 to Writer.ColCount do
      SelectionAddCell(i, j);
end;

procedure TSingleSelectableSheet.DelSelection;
begin
  FSelectedIndex:= -1;
  SelectionClear;
end;

procedure TSingleSelectableSheet.SelectionMove(const ADelta: Integer);
var
  NewSelectedIndex: Integer;
begin
  if not IsSelected then Exit;
  NewSelectedIndex:= SelectedIndex + ADelta;
  if not CheckIndex(High(FFirstRows), NewSelectedIndex) then Exit;
  Select(FFirstRows[NewSelectedIndex], 1);
end;

constructor TSingleSelectableSheet.Create(const AWorksheet: TsWorksheet;
                       const AGrid: TsWorksheetGrid;
                       const AFont: TFont;
                       const ARowHeightDefault: Integer = ROW_HEIGHT_DEFAULT);
begin
  inherited Create(AWorksheet, AGrid, AFont, ARowHeightDefault);
  FSelectedIndex:= -1;
  FFirstRows:= nil;
  FLastRows:= nil;
end;

function TSingleSelectableSheet.ReSelect(const AIDVector: TIntVector;
                                         const AIDValue: Integer;
                                         const AFirstRowSelectIfNotFound: Boolean): Boolean;
var
  Index: Integer;
begin
  Result:= False;
  if VIsNil(AIDVector) or VIsNil(FFirstRows) then Exit;

  Index:= -1;
  if AIDValue<=0 then
  begin
    if AFirstRowSelectIfNotFound then
      Index:= 0;
  end
  else begin
    Index:= VIndexOf(AIDVector, AIDValue);
    if (Index<0) and AFirstRowSelectIfNotFound then
      Index:= 0;
  end;

  if Index>=0 then
  begin
    Select(FFirstRows[Index], 1);
    Result:= True;
  end;
end;

end.

