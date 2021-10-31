unit DK_SheetUtils;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, {SysUtils,} fpspreadsheetgrid, fpstypes, Graphics, DK_SheetConst;

  //use in OnDrawCell TsWorksheetGrid to draw borders for frozen cells
  procedure DrawFrozenBorders(const AGrid: TsWorksheetGrid;
                            const ACol, ARow: Integer;
                            const AFrozenCol1, AFrozenRow1, AFrozenCol2, AFrozenRow2: Integer;
                            const ARect: TRect;
                            const ALineColor: TColor = clBlack);
  function FontStyleSheetsToGraphics(const AFontStyle: TsFontStyles): TFontStyles;
  function FontStyleGraphicsToSheets(const AFontStyle: TFontStyles): TsFontStyles;
  function ColorGraphicsToSheets(const AColor: TColor): TsColor;
  function ColorSheetsToGraphics(const AColor: TsColor): TColor;
  function WidthPxToPt(const AValuePx: Integer): Single;
  function HeightPxToPt(const AValuePx: Integer): Single;

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



function WidthPxToPt(const AValuePx: Integer): Single;
begin
  WidthPxToPt:= AValuePx/WIDTH_PX_RATIO;
end;

function HeightPxToPt(const AValuePx: Integer): Single;
begin
  HeightPxToPt:= AValuePx/HEIGHT_PX_RATIO;
end;

function FontStyleSheetsToGraphics(const AFontStyle: TsFontStyles): TFontStyles;
begin
  Result:= [];
  if fssBold in AFontStyle then Result:= Result + [fsBold];
  if fssItalic in AFontStyle then Result:= Result + [fsItalic];
  if fssStrikeOut in AFontStyle then Result:= Result + [fsStrikeOut];
  if fssUnderline in AFontStyle then Result:= Result + [fsUnderline];
end;

function FontStyleGraphicsToSheets(const AFontStyle: TFontStyles): TsFontStyles;
begin
  Result:= [];
  if fsBold in AFontStyle then Result:= Result + [fssBold];
  if fsItalic in AFontStyle then Result:= Result + [fssItalic];
  if fsStrikeOut in AFontStyle then Result:= Result + [fssStrikeOut];
  if fsUnderline in AFontStyle then Result:= Result + [fssUnderline];
end;

function ColorGraphicsToSheets(const AColor: TColor): TsColor;
begin
  if AColor=clNone then
    Result:= scTransparent
  else
    Result:= ColorToRGB(AColor);
end;

function ColorSheetsToGraphics(const AColor: TsColor): TColor;
begin
  if (AColor=scTransparent) or (AColor=scNotDefined) then
    Result:= clNone
  else
    Result:= TColor(AColor);
end;


end.

