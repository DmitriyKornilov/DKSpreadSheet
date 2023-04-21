unit DK_SheetTables;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics, Controls, fpsTypes,
  fpspreadsheetgrid, DK_Const, DK_Vector, DK_Matrix, DK_SheetWriter;

type
  TsHorAlignment = fpsTypes.TsHorAlignment; //(haDefault, haLeft, haCenter, haRight);
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
    procedure SetHeaderBGColor(AValue: TColor);
    procedure SetHeaderFont(AValue: TFont);
    procedure SetSelectedBGColor(AValue: TColor);
    procedure SetSelectedFont(AValue: TFont);
    procedure SetValuesBGColor(AValue: TColor);
    procedure SetValuesFont(AValue: TFont);
    procedure MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  protected
    FOnSelect: TSheetSelectEvent;

    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FSelectedIndex: Integer;

    FHeaderFont: TFont;
    FValuesFont: TFont;
    FSelectedFont: TFont;

    FValuesBGColor: TColor;
    FHeaderBGColor: TColor;
    FSelectedBGColor: TColor;

    FColumnWidths: TIntVector;
    FColumnNames: TStrVector;
    FColumnFormatStrings: TStrVector;
    FColumnTypes: TIntVector;
    FColumnHorAlignments, FColumnVertAlignments: TIntVector;
    FColumnValues: TStrMatrix;
    FColumnValuesBGColors: TColorVector;

    FHeaderRows1, FHeaderRows2: TIntVector;
    FHeaderCols1, FHeaderCols2: TIntVector;
    FHeaderHorAlignments, FHeaderVertAlignments: TIntVector;
    FHeaderCaptions: TStrVector;
    FHeaderBGColors: TColorVector;

    procedure DrawHeader;

    procedure PrepareData;
    procedure DrawData;
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);

    function LineIndexFromRow(const ARow: Integer): Integer;
    function RowFromLineIndex(const AIndex: Integer): Integer;
    function IsLineIndexCorrect(const AIndex: Integer): Boolean;

    procedure Clear;

    function IsColIndexCorrect(const AIndex: Integer): Boolean;

  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;

    procedure AddToHeader(const ARow1, ACol1, ARow2, ACol2: Integer;
                          const ACaption: String;
                          const AHorAlignment: TsHorAlignment = haCenter;
                          const AVertAlignment: TsVertAlignment = vaCenter;
                          const AHeaderBGColor: TColor = clNone);

    procedure AddColumn(const AName: String;
                        const AWidth: Integer = 100;
                        const AHorAlignment: TsHorAlignment = haCenter;
                        const AVertAlignment: TsVertAlignment = vaCenter;
                        const AColumnValuesBGColor: TColor = clNone);

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

    property HeaderFont: TFont read FHeaderFont write SetHeaderFont;
    property ValuesFont: TFont read FValuesFont write SetValuesFont;
    property SelectedFont: TFont read FSelectedFont write SetSelectedFont;

    property ValuesBGColor: TColor read FValuesBGColor write SetValuesBGColor;
    property HeaderBGColor: TColor read FHeaderBGColor write SetHeaderBGColor;
    property SelectedBGColor: TColor read FSelectedBGColor write SetSelectedBGColor;

    procedure Draw;
    procedure Select(const ARow: Integer);
    procedure Unselect;
    property IsSelected: Boolean read GetIsSelected;
    property SelectedIndex: Integer read FSelectedIndex;

    property HeaderRowBegin: Integer read GetHeaderRowBegin;
    property HeaderRowEnd: Integer read GetHeaderRowEnd;

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

procedure TSheetTable.SetHeaderBGColor(AValue: TColor);
begin
  if FHeaderBGColor=AValue then Exit;
  FHeaderBGColor:=AValue;
  //Refresh
end;

procedure TSheetTable.SetHeaderFont(AValue: TFont);
begin
  if FHeaderFont=AValue then Exit;
  FHeaderFont:=AValue;
  //FGrid.Refresh;
end;

procedure TSheetTable.SetSelectedBGColor(AValue: TColor);
begin
  if FSelectedBGColor=AValue then Exit;
  FSelectedBGColor:=AValue;
  //Refresh
end;

procedure TSheetTable.SetSelectedFont(AValue: TFont);
begin
  if FSelectedFont=AValue then Exit;
  FSelectedFont:=AValue;
  //Refresh
end;

procedure TSheetTable.SetValuesBGColor(AValue: TColor);
begin
  if FValuesBGColor=AValue then Exit;
  FValuesBGColor:=AValue;
  //Refresh
end;

procedure TSheetTable.SetValuesFont(AValue: TFont);
begin
  if FValuesFont=AValue then Exit;
  FValuesFont:=AValue;
  //Refresh
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

procedure TSheetTable.Clear;
begin


  FSelectedIndex:= -1;
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
  FHeaderFont.Assign(FGrid.Font);
  FValuesFont.Assign(FGrid.Font);
  FSelectedFont.Assign(FGrid.Font);

  FValuesBGColor:= FGrid.Color;
  FHeaderBGColor:= FValuesBGColor;
  FSelectedBGColor:= DefaultSelectionBGColor;
end;

destructor TSheetTable.Destroy;
begin
  FreeAndNil(FHeaderFont);
  FreeAndNil(FValuesFont);
  FreeAndNil(FSelectedFont);
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TSheetTable.AddToHeader(const ARow1, ACol1, ARow2, ACol2: Integer;
                          const ACaption: String;
                          const AHorAlignment: TsHorAlignment = haCenter;
                          const AVertAlignment: TsVertAlignment = vaCenter;
                          const AHeaderBGColor: TColor = clNone);
begin
  VAppend(FHeaderRows1, ARow1);
  VAppend(FHeaderRows2, ARow2);
  VAppend(FHeaderCols1, ACol1);
  VAppend(FHeaderCols2, ACol2);
  VAppend(FHeaderHorAlignments, Ord(AHorAlignment));
  VAppend(FHeaderVertAlignments, Ord(AVertAlignment));
  VAppend(FHeaderCaptions, ACaption);
  VAppend(FHeaderBGColors, AHeaderBGColor);
end;



procedure TSheetTable.AddColumn(const AName: String;
                        const AWidth: Integer = 100;
                        const AHorAlignment: TsHorAlignment = haCenter;
                        const AVertAlignment: TsVertAlignment = vaCenter;
                        const AColumnValuesBGColor: TColor = clNone);
begin
  VAppend(FColumnFormatStrings, EmptyStr);
  VAppend(FColumnTypes, Ord(ctUndefined));
  VAppend(FColumnNames, AName);
  VAppend(FColumnWidths, AWidth);
  VAppend(FColumnHorAlignments, Ord(AHorAlignment));
  VAppend(FColumnVertAlignments, Ord(AVertAlignment));
  VAppend(FColumnValuesBGColors, AColumnValuesBGColor);
  MAppend(FColumnValues, nil);
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

procedure TSheetTable.Draw;
begin
  Clear;

  if Assigned(FWriter) then FreeAndNil(FWriter);
  FWriter:= TSheetWriter.Create(FColumnWidths, FGrid.Worksheet, FGrid);
  FWriter.BeginEdit;

  DrawHeader;
  DrawData;

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

