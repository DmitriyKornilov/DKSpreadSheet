unit DK_SheetUtils;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, {SysUtils,} fpspreadsheetgrid, Graphics;

  //use in OnDrawCell TsWorksheetGrid to draw borders for frozen cells
  procedure DrawFrozenBorders(const AGrid: TsWorksheetGrid;
                            const ACol, ARow: Integer;
                            const AFrozenCol1, AFrozenRow1, AFrozenCol2, AFrozenRow2: Integer;
                            const ARect: TRect;
                            const ALineColor: TColor = clBlack);

implementation

procedure DrawFrozenBorders(const AGrid: TsWorksheetGrid;
                            const ACol, ARow: Integer;
                            const AFrozenCol1, AFrozenRow1, AFrozenCol2, AFrozenRow2: Integer;
                            const ARect: TRect;
                            const ALineColor: TColor = clBlack);
begin
  if (ACol=AFrozenCol1-1)and
     ((ARow>=AFrozenRow1) and (ARow<=AFrozenRow2)) then
  begin
    AGrid.Canvas.Pen.Color:= ALineColor;
    AGrid.Canvas.MoveTo(ARect.Right-1, ARect.Top);
    AGrid.Canvas.LineTo(ARect.Right-1, ARect.Bottom);
  end;
  if ((ACol>=AFrozenCol1) and (ACol<=AFrozenCol2)) and
     ((ARow>=AFrozenRow1) and (ARow<=AFrozenRow2)) then
  begin
    AGrid.Canvas.Pen.Color:= ALineColor;
    AGrid.Canvas.MoveTo(ARect.Left-1, ARect.Top);
    AGrid.Canvas.LineTo(ARect.Right-1, ARect.Top);
    AGrid.Canvas.LineTo(ARect.Right-1, ARect.Bottom);
    AGrid.Canvas.LineTo(ARect.Left-1, ARect.Bottom);
    AGrid.Canvas.LineTo(ARect.Left-1, ARect.Top);
  end;
end;


end.

