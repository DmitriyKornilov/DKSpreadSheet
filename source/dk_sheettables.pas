unit DK_SheetTables;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics, Controls, fpsTypes, fpspreadsheetgrid,
  DK_Const, DK_Vector, DK_Matrix, DK_StrUtils,  DK_SheetWriter,
  DK_SheetExporter;

const
  haLeft   = fpsTypes.haLeft;
  haCenter = fpsTypes.haCenter;
  haRight  = fpsTypes.haRight;
  vaTop    = fpsTypes.vaTop;
  vaCenter = fpsTypes.vaCenter;
  vaBottom = fpsTypes.vaBottom;

type
  TsHorAlignment  = fpsTypes.TsHorAlignment;  //(haDefault, haLeft, haCenter, haRight);
  TsVertAlignment = fpsTypes.TsVertAlignment; //(vaDefault, vaTop, vaCenter, vaBottom);

  TSheetSelectEvent = procedure of object;

  TSheetColumnType = (
    ctUndefined,
    ctOrder,
    ctInteger,
    ctString,
    ctDate,
    ctTime
  );

  { TSheetTable }

  TSheetTable = class(TObject)
  private
    function GetColumnVisible(const ACol: Integer): Boolean;
    function GetHeaderRowBegin: Integer;
    function GetHeaderRowEnd: Integer;
    function GetIsEmpty: Boolean;
    function GetIsSelected: Boolean;
    function GetValuesRowBegin: Integer;
    function GetValuesRowEnd: Integer;

    procedure SetColumnVisibles;
    procedure SetColumnVisible(const ACol: Integer; AValue: Boolean);
    procedure SetCanSelect(AValue: Boolean);
    procedure SetCanUnselect(AValue: Boolean);

    procedure MouseDown(Sender: TObject; Button: TMouseButton;
      {%H-}Shift: TShiftState; X, Y: Integer);

  protected
    FOnSelect: TSheetSelectEvent;

    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FSelectedIndex: Integer;
    FIsEmptyDraw: Boolean;

    FValuesFont: TFont;
    FHeaderFont: TFont;
    FSelectedFont: TFont;
    FRowBeforeFont: TFont;
    FRowAfterFont: TFont;

    FValuesBGColor: TColor;
    FHeaderBGColor: TColor;
    FSelectedBGColor: TColor;
    FRowBeforeBGColor: TColor;
    FRowAfterBGColor: TColor;

    FColumnWidths: TIntVector;
    FColumnNames: TStrVector;
    FColumnFormatStrings: TStrVector;
    FColumnTypes: TIntVector;
    FColumnHorAlignments, FColumnVertAlignments: TIntVector;
    FColumnValues: TStrMatrix;
    FColumnValuesBGColors: TColorVector;
    FColumnVisibles: TBoolVector;

    FHeaderFrozen: Boolean;
    FHeaderRows1, FHeaderRows2: TIntVector;
    FHeaderCols1, FHeaderCols2: TIntVector;
    FHeaderHorAlignments, FHeaderVertAlignments: TIntVector;
    FHeaderCaptions: TStrVector;
    FHeaderBGColors: TColorVector;

    FRowBeforeValue: String;
    FRowBeforeHorAlignment: TsHorAlignment;
    FRowBeforeVertAlignment: TsVertAlignment;
    FRowBeforeBorderType: TCellBorderType;

    FRowAfterValue: String;
    FRowAfterHorAlignment: TsHorAlignment;
    FRowAfterVertAlignment: TsVertAlignment;
    FRowAfterBorderType: TCellBorderType;

    FExtraFontNames: TStrVector;
    FExtraFontSizes: TIntVector;
    FExtraFontStyles: array of TFontStyles;
    FExtraFontColors: TColorVector;
    FExtraFontColumnNames: TStrVector;
    FExtraFontIfColumnNames: TStrVector;
    FExtraFontIfColumnValues: TStrVector;

    FCanSelect: Boolean;
    FCanUnselect: Boolean;

    FZoomPercents: Integer;

    procedure PrepareData;
    procedure FreezeHeader;
    procedure DrawHeader;
    procedure DrawData;
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure DrawRowBefore;
    procedure DrawRowAfter;

    function LineIndexFromRow(const ARow: Integer): Integer;
    function RowFromLineIndex(const AIndex: Integer): Integer;
    function IsLineIndexCorrect(const AIndex: Integer): Boolean;

    function IsColIndexCorrect(const AIndex: Integer): Boolean;

    function ColumnIndexByName(const AName: String): Integer;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure Zoom(const APercents: Integer);

    procedure SetExtraFont(const AColumnName, AIfColumnName, AIfColumnValue: String;
                        const AFontName: String;
                        const AFontSize: Integer;
                        const AFontStyles: TFontStyles;
                        const AFontColor: TColor);

    procedure AddToHeader(const ARow, ACol: Integer;
                          const ACaption: String;
                          const AHorAlignment: TsHorAlignment = haCenter;
                          const AVertAlignment: TsVertAlignment = vaCenter;
                          const ABGColor: TColor = clNone);
    procedure AddToHeader(const ARow1, ACol1, ARow2, ACol2: Integer;
                          const ACaption: String;
                          const AHorAlignment: TsHorAlignment = haCenter;
                          const AVertAlignment: TsVertAlignment = vaCenter;
                          const ABGColor: TColor = clNone);

    procedure AddColumn(const AName: String;
                        const AWidth: Integer = 100;
                        const AHorAlignment: TsHorAlignment = haCenter;
                        const AVertAlignment: TsVertAlignment = vaCenter;
                        const ABGColor: TColor = clNone);

    procedure SetRowBefore(const AValue: String;
                        const AHorAlignment: TsHorAlignment = haCenter;
                        const AVertAlignment: TsVertAlignment = vaCenter;
                        const ABGColor: TColor = clNone;
                        const ABorderType: TCellBorderType = cbtNone);
    procedure SetRowAfter(const AValue: String;
                        const AHorAlignment: TsHorAlignment = haCenter;
                        const AVertAlignment: TsVertAlignment = vaCenter;
                        const ABGColor: TColor = clNone;
                        const ABorderType: TCellBorderType = cbtNone);

    procedure SetColumnOrder(const AName: String);
    procedure SetColumnInteger(const AName: String; const AValues: TIntVector);
    procedure SetColumnString(const AName: String; const AValues: TStrVector);
    procedure SetColumnDate(const AName: String; const AValues: TDateVector;
                            const AFormatString: String = 'dd.mm.yyyy';
                            const ABoundaryIsEmpty: Boolean = True);
    procedure SetColumnTime(const AName: String; const AValues: TTimeVector;
                            const AFormatString: String = 'hh:nn');

    procedure SetColumnOrder(const AColIndex: Integer);
    procedure SetColumnInteger(const AColIndex: Integer; const AValues: TIntVector);
    procedure SetColumnString(const AColIndex: Integer; const AValues: TStrVector);
    procedure SetColumnDate(const AColIndex: Integer; const AValues: TDateVector;
                            const AFormatString: String = 'dd.mm.yyyy';
                            const ABoundaryIsEmpty: Boolean = True);
    procedure SetColumnTime(const AColIndex: Integer; const AValues: TTimeVector;
                            const AFormatString: String = 'hh:nn');

    procedure SetFontsName(const AName: String);
    procedure SetFontsSize(const ASize: Integer);

    property ColumnVisible[const ACol: Integer]: Boolean read GetColumnVisible write SetColumnVisible;

    property HeaderFont: TFont read FHeaderFont write FHeaderFont;
    property ValuesFont: TFont read FValuesFont write FValuesFont;
    property SelectedFont: TFont read FSelectedFont write FSelectedFont;
    property RowBeforeFont: TFont read FRowBeforeFont write FRowBeforeFont;
    property RowAfterFont: TFont read FRowAfterFont write FRowAfterFont;

    property ValuesBGColor: TColor read FValuesBGColor write FValuesBGColor;
    property HeaderBGColor: TColor read FHeaderBGColor write FHeaderBGColor;
    property SelectedBGColor: TColor read FSelectedBGColor write FSelectedBGColor;
    property RowBeforeBGColor: TColor read FRowBeforeBGColor write FRowBeforeBGColor;
    property RowAfterBGColor: TColor read FRowAfterBGColor write FRowAfterBGColor;

    procedure Save(const ADoneMessage: String);
    procedure Draw;
    property IsEmptyDraw: Boolean read FIsEmptyDraw write FIsEmptyDraw;
    property IsEmpty: Boolean read GetIsEmpty;

    procedure SelectIndex(const AIndex: Integer);
    procedure SelectRow(const ARow: Integer);
    procedure Unselect;
    property IsSelected: Boolean read GetIsSelected;
    property SelectedIndex: Integer read FSelectedIndex;

    property ValuesRowBegin: Integer read GetValuesRowBegin;
    property ValuesRowEnd: Integer read GetValuesRowEnd;

    property HeaderRowBegin: Integer read GetHeaderRowBegin;
    property HeaderRowEnd: Integer read GetHeaderRowEnd;
    property HeaderFrozen: Boolean read FHeaderFrozen write FHeaderFrozen;

    property CanSelect: Boolean read FCanSelect write SetCanSelect;
    property CanUnselect: Boolean read FCanUnselect write SetCanUnselect;

    property OnSelect: TSheetSelectEvent read FOnSelect write FOnSelect;
  end;


implementation

{ TSheetTable }

function TSheetTable.GetColumnVisible(const ACol: Integer): Boolean;
var
  Index: Integer;
begin
  Index:= ACol - 1;
  if not IsColIndexCorrect(Index) then Exit;
  Result:= FColumnVisibles[Index];
end;

function TSheetTable.GetHeaderRowBegin: Integer;
begin
  Result:= Ord(not SEmpty(FRowBeforeValue));
  if VIsNil(FHeaderRows1) then Exit;
  Result:= VMin(FHeaderRows1);
end;

function TSheetTable.GetHeaderRowEnd: Integer;
begin
  Result:= HeaderRowBegin;
  if VIsNil(FHeaderRows2) then Exit;
  Result:= VMax(FHeaderRows2);
end;

function TSheetTable.GetIsEmpty: Boolean;
begin
  Result:= MIsNil(FColumnValues) or VIsNil(FColumnValues[0]);
end;

function TSheetTable.GetIsSelected: Boolean;
begin
  Result:= FSelectedIndex>=0;
end;

function TSheetTable.GetValuesRowBegin: Integer;
begin
  Result:= HeaderRowEnd;
  if IsEmpty then Exit;
  Result:= Result + 1;
end;

function TSheetTable.GetValuesRowEnd: Integer;
begin
  Result:= ValuesRowBegin;
  if IsEmpty then Exit;
  Result:= Result + Length(FColumnValues[0]) - 1;
end;

procedure TSheetTable.MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  R,C: Integer;
begin
  if not CanSelect then Exit;

  if Button=mbLeft then
  begin
    (Sender as TsWorksheetGrid).MouseToCell(X,Y,C,R);
    SelectRow(R);
  end
  else if Button=mbRight then
    if CanUnselect then
      Unselect;
end;

procedure TSheetTable.SetColumnVisibles;
var
  i: Integer;
begin
  for i:= 0 to High(FColumnVisibles) do
  begin
    if FColumnVisibles[i] then
      FGrid.ShowCol(i+1)
    else
      FGrid.HideCol(i+1);
  end;
end;

procedure TSheetTable.SetColumnVisible(const ACol: Integer; AValue: Boolean);
var
  Index: Integer;
begin
  Index:= ACol - 1;
  if not IsColIndexCorrect(Index) then Exit;
  if FColumnVisibles[Index]=AValue then Exit;
  FColumnVisibles[Index]:= AValue;
  if AValue then
    FGrid.ShowCol(Index+1)
  else
    FGrid.HideCol(Index+1);
end;

procedure TSheetTable.SetCanSelect(AValue: Boolean);
begin
  if FCanSelect=AValue then Exit;
  if not AValue then Unselect;
  FCanSelect:= AValue;
end;

procedure TSheetTable.SetCanUnselect(AValue: Boolean);
begin
  if FCanUnselect=AValue then Exit;
  FCanUnselect:= AValue;
end;

procedure TSheetTable.DrawHeader;
var
  i: Integer;
begin
  FWriter.SetFont(FHeaderFont);
  FWriter.SetBackground(FHeaderBGColor);
  for i:= 0 to High(FHeaderCaptions) do
  begin
    FWriter.SetAlignment(TsHorAlignment(FHeaderHorAlignments[i]),
                         TsVertAlignment(FHeaderVertAlignments[i]));
    FWriter.WriteText(FHeaderRows1[i], FHeaderCols1[i],
                      FHeaderRows2[i], FHeaderCols2[i],
                      FHeaderCaptions[i], cbtNone, True, True);
  end;
  FWriter.DrawBorders(HeaderRowBegin, 1, HeaderRowEnd, FWriter.ColCount, cbtAll);
end;

procedure TSheetTable.DrawLine(const AIndex: Integer; const ASelected: Boolean);
var
  i, R, C: Integer;
  S: String;

  procedure SetBGColor(const AColumnIndex: Integer);
  begin
    if ASelected then
      FWriter.SetBackground(FSelectedBGColor)
    else if FColumnValuesBGColors[AColumnIndex]<>clNone then
      FWriter.SetBackground(FColumnValuesBGColors[AColumnIndex])
    else
      FWriter.SetBackground(FValuesBGColor);
  end;

  function ExtraFont(const AColumnIndex: Integer): Boolean;
  var
    k, Ind: Integer;
    ThisColumnName: String;
  begin
    Result:= False;

    ThisColumnName:= FColumnNames[AColumnIndex];
    //пробегаем по всем именам столбцов с доп шрифтами
    for k:=0 to High(FExtraFontColumnNames) do
    begin
      if FExtraFontColumnNames[k]<>ThisColumnName then continue;
      //индекс столбца, по которому опредяем соблюдение условия
      Ind:= ColumnIndexByName(FExtraFontIfColumnNames[k]);
      //проверяем соблюдение условия
      if FExtraFontIfColumnValues[k]=FColumnValues[Ind, AIndex] then
      begin //если есть - меняем шрфит
        FWriter.SetFont(FExtraFontNames[k], FExtraFontSizes[k],
                        FExtraFontStyles[k], FExtraFontColors[k]);
        Result:= True;
      end;
    end;
  end;

  procedure SetFont(const AColumnIndex: Integer);
  begin
    if ExtraFont(AColumnIndex) then Exit;
    if ASelected then
      FWriter.SetFont(FSelectedFont)
    else
      FWriter.SetFont(FValuesFont);
  end;

begin
  for i:= 0 to High(FColumnValues) do
  begin
    SetBGColor(i);
    SetFont(i);
    FWriter.SetAlignment(TsHorAlignment(FColumnHorAlignments[i]),
                         TsVertAlignment(FColumnVertAlignments[i]));
    R:= RowFromLineIndex(AIndex);
    C:= i + 1;
    S:= FColumnValues[i, AIndex];
    if TSheetColumnType(FColumnTypes[i])=ctOrder then
      FWriter.WriteNumber(R, C, AIndex+1, cbtOuter)
    else if SEmpty(S) then
      FWriter.WriteText(R, C, S, cbtOuter)
    else begin
      case TSheetColumnType(FColumnTypes[i]) of
        //ctUndefined,
        ctInteger: FWriter.WriteNumber(R, C, StrToInt(S), cbtOuter);
        ctString:  FWriter.WriteText(R, C, S, cbtOuter, True, True);
        ctDate:    FWriter.WriteDate(R,C, StrToDate(S), cbtOuter, FColumnFormatStrings[i]);
        ctTime:    FWriter.WriteTime(R,C, StrToTime(S), cbtOuter, FColumnFormatStrings[i]);
      end;
    end;
  end;

  if IsEmpty then Exit;
  FWriter.DrawBorders(ValuesRowBegin, 1, ValuesRowEnd, FWriter.ColCount, cbtAll);
end;

procedure TSheetTable.DrawRowBefore;
begin
  if SEmpty(FRowBeforeValue) then Exit;
  FWriter.SetFont(FRowBeforeFont);
  FWriter.SetAlignment(FRowBeforeHorAlignment, FRowBeforeVertAlignment);
  if FRowBeforeBGColor<>clNone then
    FWriter.SetBackground(FRowBeforeBGColor)
  else
    FWriter.SetBackground(FValuesBGColor);
  FWriter.WriteText(1, 1, 1, FWriter.ColCount,
                    FRowBeforeValue, FRowBeforeBorderType, True, True);
end;

procedure TSheetTable.DrawRowAfter;
var
  R: Integer;
begin
  if SEmpty(FRowAfterValue) then Exit;
  FWriter.SetFont(FRowAfterFont);
  FWriter.SetAlignment(FRowAfterHorAlignment, FRowAfterVertAlignment);
  if FRowAfterBGColor<>clNone then
    FWriter.SetBackground(FRowAfterBGColor)
  else
    FWriter.SetBackground(FValuesBGColor);
  R:= ValuesRowEnd + 1;
  FWriter.WriteText(R, 1, R, FWriter.ColCount,
                   FRowAfterValue, FRowAfterBorderType, True, True);
end;

function TSheetTable.LineIndexFromRow(const ARow: Integer): Integer;
var
  i: Integer;
begin
  Result:= -1;
  if ARow<0 then Exit;
  i:= ARow - GetHeaderRowEnd - 1;
  if IsLineIndexCorrect(i) then
    Result:= i;
end;

function TSheetTable.RowFromLineIndex(const AIndex: Integer): Integer;
begin
  Result:= GetHeaderRowEnd + 1 + AIndex;
end;

function TSheetTable.IsLineIndexCorrect(const AIndex: Integer): Boolean;
begin
  Result:= (AIndex>=0) and (AIndex<MMaxLength(FColumnValues));
end;

function TSheetTable.IsColIndexCorrect(const AIndex: Integer): Boolean;
begin
  Result:= (AIndex>=0) and (AIndex<=High(FColumnNames));
end;

function TSheetTable.ColumnIndexByName(const AName: String): Integer;
begin
  Result:= VIndexOf(FColumnNames, AName);
end;

procedure TSheetTable.PrepareData;
var
  i, MaxLength: Integer;
begin
  for i:= 0 to High(FColumnTypes) do
    if TSheetColumnType(FColumnTypes[i])=ctOrder then
      FColumnValues[i]:= nil;
  MaxLength:= MMaxLength(FColumnValues);
  for i:= 0 to High(FColumnValues) do
    if Length(FColumnValues[i])<MaxLength then
      VReDim(FColumnValues[i], MaxLength, EmptyStr);
end;

procedure TSheetTable.DrawData;
var
  i: Integer;
begin
  for i:= 0 to High(FColumnValues[0]) do
    DrawLine(i, False);
end;

constructor TSheetTable.Create(const AGrid: TsWorksheetGrid);
begin
  FGrid:= AGrid;
  FGrid.OnMouseDown:= @MouseDown;

  FHeaderFont:= TFont.Create;
  FValuesFont:= TFont.Create;
  FSelectedFont:= TFont.Create;
  FRowBeforeFont:= TFont.Create;
  FRowAfterFont:= TFont.Create;
  FHeaderFont.Assign(FGrid.Font);
  FValuesFont.Assign(FGrid.Font);
  FSelectedFont.Assign(FGrid.Font);
  FRowBeforeFont.Assign(FGrid.Font);
  FRowAfterFont.Assign(FGrid.Font);

  FValuesBGColor:= FGrid.Color;
  FHeaderBGColor:= FValuesBGColor;
  FSelectedBGColor:= DefaultSelectionBGColor;
  FRowBeforeBGColor:= FValuesBGColor;
  FRowAfterBGColor:= FValuesBGColor;

  FHeaderFrozen:= True;
  FCanSelect:= True;
  FCanUnselect:= True;
  FZoomPercents:= 100;

  FSelectedIndex:= -1;
end;

destructor TSheetTable.Destroy;
begin
  FreeAndNil(FHeaderFont);
  FreeAndNil(FValuesFont);
  FreeAndNil(FSelectedFont);
  FreeAndNil(FRowBeforeFont);
  FreeAndNil(FRowAfterFont);
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TSheetTable.Zoom(const APercents: Integer);
begin
  FZoomPercents:= APercents;
end;

procedure TSheetTable.SetExtraFont(
                        const AColumnName, AIfColumnName, AIfColumnValue: String;
                        const AFontName: String;
                        const AFontSize: Integer;
                        const AFontStyles: TFontStyles;
                        const AFontColor: TColor);
begin
  VAppend(FExtraFontNames, AFontName);
  VAppend(FExtraFontSizes, AFontSize);
  SetLength(FExtraFontStyles, Length(FExtraFontStyles)+1);
  FExtraFontStyles[High(FExtraFontStyles)]:= AFontStyles;
  VAppend(FExtraFontColors, AFontColor);
  VAppend(FExtraFontColumnNames, AColumnName);
  VAppend(FExtraFontIfColumnNames, AIfColumnName);
  VAppend(FExtraFontIfColumnValues, AIfColumnValue);
end;

procedure TSheetTable.AddToHeader(const ARow, ACol: Integer;
                          const ACaption: String;
                          const AHorAlignment: TsHorAlignment = haCenter;
                          const AVertAlignment: TsVertAlignment = vaCenter;
                          const ABGColor: TColor = clNone);
begin
  AddToHeader(ARow, ACol, ARow, ACol, ACaption, AHorAlignment, AVertAlignment, ABGColor);
end;

procedure TSheetTable.AddToHeader(const ARow1, ACol1, ARow2, ACol2: Integer;
                          const ACaption: String;
                          const AHorAlignment: TsHorAlignment = haCenter;
                          const AVertAlignment: TsVertAlignment = vaCenter;
                          const ABGColor: TColor = clNone);
begin
  VAppend(FHeaderRows1, ARow1);
  VAppend(FHeaderRows2, ARow2);
  VAppend(FHeaderCols1, ACol1);
  VAppend(FHeaderCols2, ACol2);
  VAppend(FHeaderHorAlignments, Ord(AHorAlignment));
  VAppend(FHeaderVertAlignments, Ord(AVertAlignment));
  VAppend(FHeaderCaptions, ACaption);
  VAppend(FHeaderBGColors, ABGColor);
end;

procedure TSheetTable.AddColumn(const AName: String;
                        const AWidth: Integer = 100;
                        const AHorAlignment: TsHorAlignment = haCenter;
                        const AVertAlignment: TsVertAlignment = vaCenter;
                        const ABGColor: TColor = clNone);
begin
  VAppend(FColumnFormatStrings, EmptyStr);
  VAppend(FColumnTypes, Ord(ctUndefined));
  VAppend(FColumnNames, AName);
  VAppend(FColumnWidths, AWidth);
  VAppend(FColumnHorAlignments, Ord(AHorAlignment));
  VAppend(FColumnVertAlignments, Ord(AVertAlignment));
  VAppend(FColumnValuesBGColors, ABGColor);
  VAppend(FColumnVisibles, True);
  MAppend(FColumnValues, nil);
end;

procedure TSheetTable.SetRowBefore(const AValue: String;
                        const AHorAlignment: TsHorAlignment = haCenter;
                        const AVertAlignment: TsVertAlignment = vaCenter;
                        const ABGColor: TColor = clNone;
                        const ABorderType: TCellBorderType = cbtNone);
begin
  FRowBeforeValue:= AValue;
  FRowBeforeHorAlignment:= AHorAlignment;
  FRowBeforeVertAlignment:= AVertAlignment;
  FRowBeforeBGColor:= ABGColor;
  FRowBeforeBorderType:= ABorderType;
end;

procedure TSheetTable.SetRowAfter(const AValue: String;
                        const AHorAlignment: TsHorAlignment = haCenter;
                        const AVertAlignment: TsVertAlignment = vaCenter;
                        const ABGColor: TColor = clNone;
                        const ABorderType: TCellBorderType = cbtNone);
begin
  FRowAfterValue:= AValue;
  FRowAfterHorAlignment:= AHorAlignment;
  FRowAfterVertAlignment:= AVertAlignment;
  FRowAfterBGColor:= ABGColor;
  FRowAfterBorderType:= ABorderType;
end;

procedure TSheetTable.SetColumnOrder(const AName: String);
begin
  SetColumnOrder(ColumnIndexByName(AName));
end;

procedure TSheetTable.SetColumnInteger(const AName: String; const AValues: TIntVector);
begin
  SetColumnInteger(ColumnIndexByName(AName), AValues);
end;

procedure TSheetTable.SetColumnString(const AName: String; const AValues: TStrVector);
begin
  SetColumnString(ColumnIndexByName(AName), AValues);
end;

procedure TSheetTable.SetColumnDate(const AName: String; const AValues: TDateVector;
                                    const AFormatString: String = 'dd.mm.yyyy';
                                    const ABoundaryIsEmpty: Boolean = True);
begin
  SetColumnDate(ColumnIndexByName(AName), AValues, AFormatString, ABoundaryIsEmpty);
end;

procedure TSheetTable.SetColumnTime(const AName: String; const AValues: TTimeVector;
                                    const AFormatString: String = 'hh:nn');
begin
  SetColumnTime(ColumnIndexByName(AName), AValues, AFormatString);
end;

procedure TSheetTable.SetColumnOrder(const AColIndex: Integer);
begin
  if not IsColIndexCorrect(AColIndex) then Exit;
  FColumnTypes[AColIndex]:= Ord(ctOrder);
end;

procedure TSheetTable.SetColumnInteger(const AColIndex: Integer; const AValues: TIntVector);
begin
  if not IsColIndexCorrect(AColIndex) then Exit;
  FColumnValues[AColIndex]:= VIntToStr(AValues);
  FColumnTypes[AColIndex]:= Ord(ctInteger);
end;

procedure TSheetTable.SetColumnString(const AColIndex: Integer; const AValues: TStrVector);
begin
  if not IsColIndexCorrect(AColIndex) then Exit;
  FColumnValues[AColIndex]:= VCut(AValues);
  FColumnTypes[AColIndex]:= Ord(ctString);
end;

procedure TSheetTable.SetColumnDate(const AColIndex: Integer; const AValues: TDateVector;
                                    const AFormatString: String = 'dd.mm.yyyy';
                                    const ABoundaryIsEmpty: Boolean = True);
begin
  if not IsColIndexCorrect(AColIndex) then Exit;
  FColumnValues[AColIndex]:= VFormatDateTime(AFormatString, AValues, ABoundaryIsEmpty);
  FColumnTypes[AColIndex]:= Ord(ctDate);
  FColumnFormatStrings[AColIndex]:= AFormatString;
end;

procedure TSheetTable.SetColumnTime(const AColIndex: Integer; const AValues: TTimeVector;
                                    const AFormatString: String = 'hh:nn');
begin
  if not IsColIndexCorrect(AColIndex) then Exit;
  FColumnValues[AColIndex]:= VFormatDateTime(AFormatString, AValues);
  FColumnTypes[AColIndex]:= Ord(ctTime);
  FColumnFormatStrings[AColIndex]:= AFormatString;
end;

procedure TSheetTable.SetFontsName(const AName: String);
begin
  FValuesFont.Name:= AName;
  FHeaderFont.Name:= AName;
  FSelectedFont.Name:= AName;
  FRowBeforeFont.Name:= AName;
  FRowAfterFont.Name:= AName;
end;

procedure TSheetTable.SetFontsSize(const ASize: Integer);
begin
  FValuesFont.Size:= ASize;
  FHeaderFont.Size:= ASize;
  FSelectedFont.Size:= ASize;
  FRowBeforeFont.Size:= ASize;
  FRowAfterFont.Size:= ASize;
end;

procedure TSheetTable.Save(const ADoneMessage: String);
var
  Exporter: TGridExporter;
begin
  if IsSelected then
    DrawLine(FSelectedIndex, False);
  Exporter:= TGridExporter.Create(FGrid);
  try
    Exporter.Save(ADoneMessage);
  finally
    FreeAndNil(Exporter);
  end;
  if IsSelected then
    DrawLine(FSelectedIndex, True);
end;

procedure TSheetTable.FreezeHeader;
var
  N: Integer;
begin
  if not HeaderFrozen then Exit;
  if VIsNil(FColumnValues[0]) then Exit; //no data
  N:= HeaderRowEnd;
  if N=0 then Exit;
  FWriter.SetFrozenRows(N);
end;

procedure TSheetTable.Draw;
begin
  PrepareData;
  FGrid.Clear;
  FSelectedIndex:= -1;

  if IsEmpty and (not IsEmptyDraw) then Exit;

  if Assigned(FWriter) then FreeAndNil(FWriter);
  FWriter:= TSheetWriter.Create(FColumnWidths, FGrid.Worksheet, FGrid);

  FWriter.SetZoom(FZoomPercents);
  FGrid.ZoomFactor:= FZoomPercents/100;

  FWriter.BeginEdit;

  DrawRowBefore;
  DrawHeader;
  DrawData;
  DrawRowAfter;
  FreezeHeader;

  FWriter.EndEdit;

  SetColumnVisibles;
end;

procedure TSheetTable.SelectIndex(const AIndex: Integer);

  procedure DoUnselect;
  begin
    if not IsSelected then  Exit;
    DrawLine(FSelectedIndex, False);
    FSelectedIndex:= -1;
  end;

begin
  if (AIndex<0) then //unselect only
    DoUnselect
  else begin
    if (AIndex>=0) and (AIndex<>FSelectedIndex) then
    begin
      //unselect
      DoUnselect;
      //SelectRow
      FSelectedIndex:= AIndex;
      FGrid.Row:= RowFromLineIndex(FSelectedIndex); //show
      DrawLine(FSelectedIndex, True);
    end;
  end;

  if Assigned(FOnSelect) then FOnSelect;
end;

procedure TSheetTable.SelectRow(const ARow: Integer);
var
  NewSelectedIndex: Integer;
begin
  NewSelectedIndex:= LineIndexFromRow(ARow);
  SelectIndex(NewSelectedIndex);
end;

procedure TSheetTable.Unselect;
begin
  SelectIndex(-1);
end;

end.

