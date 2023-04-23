unit DK_SheetTables;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics, Controls, fpsTypes,
  fpspreadsheetgrid, DK_Const, DK_Vector, DK_Matrix, DK_SheetWriter;

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
    function GetHeaderRowBegin: Integer;
    function GetHeaderRowEnd: Integer;
    function GetIsSelected: Boolean;
    procedure MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  protected
    FOnSelect: TSheetSelectEvent;

    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FSelectedIndex: Integer;

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

    procedure DrawHeader;
    procedure FreezeHeader;

    procedure PrepareData;
    procedure DrawData;
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure DrawRowAddition(const AFont: TFont; const AValue: String;
                        const AHorAlignment: TsHorAlignment;
                        const AVertAlignment: TsVertAlignment;
                        const ABGColor: TColor;
                        const ABorderType: TCellBorderType);

    function LineIndexFromRow(const ARow: Integer): Integer;
    function RowFromLineIndex(const AIndex: Integer): Integer;
    function IsLineIndexCorrect(const AIndex: Integer): Boolean;

    function IsColIndexCorrect(const AIndex: Integer): Boolean;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;

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


    procedure Draw;
    procedure Select(const ARow: Integer);
    procedure Unselect;
    property IsSelected: Boolean read GetIsSelected;
    property SelectedIndex: Integer read FSelectedIndex;

    property HeaderRowBegin: Integer read GetHeaderRowBegin;
    property HeaderRowEnd: Integer read GetHeaderRowEnd;
    property HeaderFrozen: Boolean read FHeaderFrozen write FHeaderFrozen;

    property OnSelect: TSheetSelectEvent read FOnSelect write FOnSelect;
  end;


implementation

{ TSheetTable }

function TSheetTable.GetHeaderRowBegin: Integer;
begin
  Result:= 0;
  if VIsNil(FHeaderRows1) then Exit;
  Result:= VMin(FHeaderRows1);
end;

function TSheetTable.GetHeaderRowEnd: Integer;
begin
  Result:= 0;
  if VIsNil(FHeaderRows2) then Exit;
  Result:= VMin(FHeaderRows2);
end;

function TSheetTable.GetIsSelected: Boolean;
begin
  Result:= FSelectedIndex>=0;
end;

procedure TSheetTable.MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  R,C: Integer;
begin
  if Button=mbLeft then
  begin
    (Sender as TsWorksheetGrid).MouseToCell(X,Y,C,R);
    Select(R);
  end
  else if Button=mbRight then
    Unselect;
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
begin
  if ASelected then
  begin
    FWriter.SetFont(FSelectedFont);
    FWriter.SetBackground(FSelectedBGColor);
  end
  else begin
    FWriter.SetFont(FValuesFont);
  end;

  for i:= 0 to High(FColumnValues) do
  begin
    if not ASelected then
    begin
      if FColumnValuesBGColors[i]<>clNone then
        FWriter.SetBackground(FColumnValuesBGColors[i])
      else
        FWriter.SetBackground(FValuesBGColor);
    end;
    FWriter.SetAlignment(TsHorAlignment(FColumnHorAlignments[i]),
                         TsVertAlignment(FColumnVertAlignments[i]));
    R:= RowFromLineIndex(AIndex);
    C:= i + 1;
    S:= FColumnValues[i, AIndex];
    if TSheetColumnType(FColumnTypes[i])=ctOrder then
      FWriter.WriteNumber(R, C, AIndex+1, cbtOuter)
    else if S=EmptyStr then
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
end;

procedure TSheetTable.DrawRowAddition(const AFont: TFont; const AValue: String;
                        const AHorAlignment: TsHorAlignment;
                        const AVertAlignment: TsVertAlignment;
                        const ABGColor: TColor;
                        const ABorderType: TCellBorderType);
begin
  if AValue=EmptyStr then Exit;
  FWriter.SetFont(AFont);
  FWriter.SetAlignment(AHorAlignment, AVertAlignment);
  if ABGColor<>clNone then
    FWriter.SetBackground(ABGColor)
  else
    FWriter.SetBackground(FValuesBGColor);
  FWriter.WriteText(1, 1, 1, FWriter.ColCount, AValue, ABorderType, True, True);
end;

function TSheetTable.LineIndexFromRow(const ARow: Integer): Integer;
var
  i: Integer;
begin
  Result:= -1;
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

procedure TSheetTable.PrepareData;
var
  i, MaxLength: Integer;
begin
  MaxLength:= MMaxLength(FColumnValues);
  for i:= 0 to High(FColumnValues) do
    if Length(FColumnValues[i])<MaxLength then
      VReDim(FColumnValues[i], MaxLength, EmptyStr);
end;

procedure TSheetTable.DrawData;
var
  i: Integer;
begin
  PrepareData;
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
var
  ColIndex: Integer;
begin
  ColIndex:= VIndexOf(FColumnNames, AName);
  SetColumnOrder(ColIndex);
end;

procedure TSheetTable.SetColumnInteger(const AName: String; const AValues: TIntVector);
var
  ColIndex: Integer;
begin
  ColIndex:= VIndexOf(FColumnNames, AName);
  SetColumnInteger(ColIndex, AValues);
end;

procedure TSheetTable.SetColumnString(const AName: String; const AValues: TStrVector);
var
  ColIndex: Integer;
begin
  ColIndex:= VIndexOf(FColumnNames, AName);
  SetColumnString(ColIndex, AValues);
end;

procedure TSheetTable.SetColumnDate(const AName: String; const AValues: TDateVector;
                                    const AFormatString: String = 'dd.mm.yyyy';
                                    const ABoundaryIsEmpty: Boolean = True);
var
  ColIndex: Integer;
begin
  ColIndex:= VIndexOf(FColumnNames, AName);
  SetColumnDate(ColIndex, AValues, AFormatString, ABoundaryIsEmpty);
end;

procedure TSheetTable.SetColumnTime(const AName: String; const AValues: TTimeVector;
                                    const AFormatString: String = 'hh:nn');
var
  ColIndex: Integer;
begin
  ColIndex:= VIndexOf(FColumnNames, AName);
  SetColumnTime(ColIndex, AValues, AFormatString);
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
  FSelectedIndex:= -1;

  if Assigned(FWriter) then FreeAndNil(FWriter);
  FWriter:= TSheetWriter.Create(FColumnWidths, FGrid.Worksheet, FGrid);
  FWriter.BeginEdit;

  DrawRowAddition(FRowBeforeFont, FRowBeforeValue,
                  FRowBeforeHorAlignment, FRowBeforeVertAlignment,
                  FRowBeforeBGColor, FRowBeforeBorderType);
  DrawHeader;
  DrawData;
  DrawRowAddition(FRowAfterFont, FRowAfterValue,
                  FRowAfterHorAlignment, FRowAfterVertAlignment,
                  FRowAfterBGColor, FRowAfterBorderType);
  FreezeHeader;

  FWriter.EndEdit;
end;

procedure TSheetTable.Select(const ARow: Integer);
var
  NewSelectedIndex: Integer;

  procedure DoUnselect;
  begin
    if IsSelected then
    begin
      DrawLine(FSelectedIndex, False);
      FSelectedIndex:= -1;
    end;
  end;

begin
  if ARow<0 then
    DoUnselect
  else begin
    NewSelectedIndex:= LineIndexFromRow(ARow);
    if (NewSelectedIndex>=0) and (NewSelectedIndex<>FSelectedIndex) then
    begin
      //unselect
      DoUnselect;
      //select
      FSelectedIndex:= NewSelectedIndex;
      DrawLine(FSelectedIndex, True);
    end;
  end;

  if Assigned(FOnSelect) then FOnSelect;
end;

procedure TSheetTable.Unselect;
begin
  Select(-1);
end;

end.

