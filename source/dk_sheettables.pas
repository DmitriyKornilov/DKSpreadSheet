unit DK_SheetTables;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics, Controls,
  fpspreadsheetgrid, DK_Vector, DK_SheetWriter;

type

  { TSheetTable }

  TSheetTable = class(TObject)
  private
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
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FSelectedIndex: Integer;

    FHeaderFont: TFont;
    FValuesFont: TFont;
    FSelectedFont: TFont;

    FValuesBGColor: TColor;
    FHeaderBGColor: TColor;
    FSelectedBGColor: TColor;

    FLineRowBegins, FLineRowEnds: TIntVector;

    function GetColumnWidths: TIntVector; virtual; abstract; //write in child all
    procedure DrawHeader(var ARow: Integer); virtual;        //write in child with inherited first
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean); virtual; //write in child with inherited first

    function IndexFromRow(const ARow: Integer): Integer;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;

    property HeaderFont: TFont read FHeaderFont write SetHeaderFont;
    property ValuesFont: TFont read FValuesFont write SetValuesFont;
    property SelectedFont: TFont read FSelectedFont write SetSelectedFont;

    property ValuesBGColor: TColor read FValuesBGColor write SetValuesBGColor;
    property HeaderBGColor: TColor read FHeaderBGColor write SetHeaderBGColor;
    property SelectedBGColor: TColor read FSelectedBGColor write SetSelectedBGColor;

    procedure Draw; virtual;  //write in child with inherited first
    procedure Select(const ARow: Integer);
    procedure Unselect;
    property IsSelected: Boolean read GetIsSelected;
    property SelectedIndex: Integer read FSelectedIndex;
  end;


implementation

{ TSheetTable }

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

procedure TSheetTable.DrawHeader(var ARow: Integer);
begin
  FWriter.SetFont(FHeaderFont);
  FWriter.SetBackground(FHeaderBGColor);
end;

procedure TSheetTable.DrawLine(const AIndex: Integer;
  const ASelected: Boolean);
begin
  if ASelected then
  begin
    FWriter.SetFont(FSelectedFont);
    FWriter.SetBackground(FSelectedBGColor);
  end
  else begin
    FWriter.SetFont(FValuesFont);
    FWriter.SetBackground(FValuesBGColor);
  end;
end;

function TSheetTable.IndexFromRow(const ARow: Integer): Integer;
var
  i: Integer;
begin
  Result:= -1;
  if VIsNil(FLineRowBegins) then Exit;
  for i:= 0 to High(FLineRowBegins) do
  begin
    if (ARow>=FLineRowBegins[i]) and (ARow<=FLineRowEnds[i]) then
    begin
      Result:= i;
      break;
    end;
  end;
end;

constructor TSheetTable.Create(const AGrid: TsWorksheetGrid);
begin
  FGrid:= AGrid;
  FGrid.OnMouseDown:= @MouseDown;
  FWriter:= TSheetWriter.Create(GetColumnWidths, AGrid.Worksheet, AGrid);

  FSelectedIndex:= -1;

  FHeaderFont:= TFont.Create;
  FValuesFont:= TFont.Create;
  FSelectedFont:= TFont.Create;
  FHeaderFont.Assign(FGrid.Font);
  FValuesFont.Assign(FGrid.Font);
  FSelectedFont.Assign(FGrid.Font);

  FValuesBGColor:= clWindow;
  FHeaderBGColor:= FValuesBGColor;
  FSelectedBGColor:= clHighlight;
end;

destructor TSheetTable.Destroy;
begin
  FreeAndNil(FHeaderFont);
  FreeAndNil(FValuesFont);
  FreeAndNil(FSelectedFont);
  FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TSheetTable.Draw;
begin
  FLineRowBegins:= nil;
  FLineRowEnds:= nil;
  FSelectedIndex:= -1;
end;

procedure TSheetTable.Select(const ARow: Integer);
begin
  Unselect;
  FSelectedIndex:= IndexFromRow(ARow);
  if FSelectedIndex>=0 then
    DrawLine(FSelectedIndex, True);
end;

procedure TSheetTable.Unselect;
begin
  if not IsSelected then Exit;
  DrawLine(FSelectedIndex, False);
  FSelectedIndex:= -1;
end;

end.

