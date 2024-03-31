unit DK_SheetTypes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Graphics, fpstypes, fpspreadsheetgrid,
  DK_SheetConst, DK_SheetWriter, DK_SheetExporter, DK_Vector;

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

  { TCustomSheet }

  TCustomSheet = class (TObject)
  protected
    function SetWidths: TIntVector; virtual; abstract;
  private
    FWriter: TSheetWriter;
    FFont: TFont;
    procedure SetFont(AValue: TFont);
  public
    constructor Create(const AWorksheet: TsWorksheet; const AGrid: TsWorksheetGrid = nil);
    destructor  Destroy; override;
    procedure Zoom(const APercents: Integer);
    procedure SetFontDefault;
    procedure Save(const ASheetName: String = 'Лист1';
                   const ADoneMessage: String = 'Выполнено!';
                   const ALandscape: Boolean = False);
    property Font: TFont read FFont write SetFont;
    property Writer: TSheetWriter read FWriter;
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

constructor TCustomSheet.Create(const AWorksheet: TsWorksheet; const AGrid: TsWorksheetGrid = nil);
begin
  FFont:= TFont.Create;
  SetFontDefault;
  FWriter:= TSheetWriter.Create(SetWidths, AWorksheet, AGrid);
end;

destructor TCustomSheet.Destroy;
begin
  FreeAndNil(FFont);
  FreeAndNil(FWriter);
  inherited Destroy;
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

end.

