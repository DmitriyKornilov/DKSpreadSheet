unit DK_SheetExporter;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DK_Dialogs, DK_StrUtils, DK_SheetConst, Dialogs,
  fpstypes, fpspreadsheet, fpspreadsheetgrid, fpsallformats;

const
  spoPortrait  = fpsTypes.spoPortrait;
  spoLandscape = fpsTypes.spoLandscape;

type
  TsPageOrientation = fpsTypes.TsPageOrientation;
  TsWorksheet       = fpspreadsheet.TsWorksheet;
  TsWorksheetGrid   = fpspreadsheetgrid.TsWorksheetGrid;

  TPageFit = (pfOnePage, //всё на одной странице
              pfWidth,   //заполнить по ширине 1 страницы (по высоте - сколько нужно, чтобы вместить все)
              pfHeight); //заполнить по высоте 1 страницы (по ширине - сколько нужно, чтобы вместить все)

  { TGridExporter }

  TGridExporter = class (TObject)
  private
      FWorkbook: TsWorkbook;
      procedure SetSheetName(const AName: String);
      function GetSheetName: String;
  public
      constructor Create(const AGrid: TsWorksheetGrid);
      destructor  Destroy; override;

      //Direct save to FileName
      procedure SaveToXLSX(const ADoneMessage: String = '';
                           const AFileName: String = '';
                           const AOverwriteExistingFile: Boolean = False);
      procedure SaveToODS(const ADoneMessage: String = '';
                           const AFileName: String = '';
                           const AOverwriteExistingFile: Boolean = False);
      //Save with SaveDialog
      procedure Save(const ADoneMessage: String = '';
                     const ADefaultFileName: String = '';
                     const AOverwriteExistingFile: Boolean = False);

      procedure PageSettings(const AOrient: TsPageOrientation = spoPortrait;
                             const APageFit: TPageFit = pfWidth;
                             const AShowHeaders: Boolean = True;
                             const AShowGridLines: Boolean = True;
                             const ATopMargin: Double=10;     //mm
                             const ALeftMargin: Double=10;    //mm
                             const ARightMargin: Double=10;   //mm
                             const ABottomMargin: Double=10;  //mm
                             const AFooterMargin: Double=0;   //mm
                             const AHeaderMargin: Double=0);  //mm
      property SheetName: String read GetSheetName write SetSheetName;
  end;

  { TSheetExporter }

  TSheetExporter = class (TObject)
    private
      FWorkbook: TsWorkbook;
    public
      constructor Create;
      destructor  Destroy; override;
      function AddWorksheet(const AName: String): TsWorksheet;

      //Direct save to FileName
      procedure SaveToXLSX(const ADoneMessage: String = '';
                           const AFileName: String = '';
                           const AOverwriteExistingFile: Boolean = False);
      procedure SaveToODS(const ADoneMessage: String = '';
                           const AFileName: String = '';
                           const AOverwriteExistingFile: Boolean = False);
      //Save with SaveDialog
      procedure Save(const ADoneMessage: String = '';
                     const ADefaultFileName: String = '';
                     const AOverwriteExistingFile: Boolean = False);

      procedure PageSettings(const AOrient: TsPageOrientation = spoPortrait;
                             const APageFit: TPageFit = pfWidth;
                             const AShowHeaders: Boolean = True;
                             const AShowGridLines: Boolean = True;
                             const ATopMargin: Double=10;     //mm
                             const ALeftMargin: Double=10;    //mm
                             const ARightMargin: Double=10;   //mm
                             const ABottomMargin: Double=10;  //mm
                             const AFooterMargin: Double=0;   //mm
                             const AHeaderMargin: Double=0);  //mm

    end;

  function TextToSheetName(const AText: String): String;
  function TextToFileName(const AText: String): String;

implementation

function TextToSheetName(const AText: String): String;
const
  INVALID_CHARS: array [0..6] of String = ('[', ']', ':', '*', '?', '/', '\');
var
  i: Integer;
begin
  Result:= SCopy(AText, 1, MAX_SHEETNAME_LENGTH);
  for i:=0 to High(INVALID_CHARS) do
    StringReplace(Result, INVALID_CHARS[i], '-', [rfReplaceAll]);
end;

function TextToFileName(const AText: String): String;
const
  INVALID_CHARS: array [0..8] of String = ('"', '*', '|', '\', ':', '<', '>', '?', '/');
var
  i, k: Integer;
begin
  Result:= AText;
  //for i:=0 to High(INVALID_CHARS) do
  //  StringReplace(Result, INVALID_CHARS[i], '', [rfReplaceAll]);
  for i:=0 to High(INVALID_CHARS) do
  begin
    k:= SPos(Result, INVALID_CHARS[i]);
    while k>0 do
    begin
      Result:= SDel(Result, k, k);
      k:= SPos(Result, INVALID_CHARS[i]);
    end;
  end;
end;

function OpenSaveDialog(var AFileName: String; out AFormat: TsSpreadsheetFormat): Boolean;
var
  SD: TSaveDialog;
  FileExtention: String;
begin
  Result:= False;
  SD:= TSaveDialog.Create(nil);
  SD.FileName:= AFileName;
  SD.Filter:= 'Электронная таблица (*.xlsx)|*.xlsx|Электронная таблица (*.ods)|*.ods';
  SD.Title:= 'Сохранить как';
  Result:= SD.Execute;
  if Result then
  begin
    case SD.FilterIndex of
    1: begin
         AFormat:= sfOOXML;
         FileExtention:= 'xlsx';
       end;
    2: begin
         AFormat:= sfOpenDocument;
         FileExtention:= 'ods';
       end;
    end;
    AFileName:= SFileName(SD.FileName, FileExtention);
  end;
  FreeAndNil(SD);
end;

procedure SaveToFormat(const AWorkbook: TsWorkbook;
                       const AFormat: TsSpreadsheetFormat;
                       const ADoneMessage: String = '';
                       const AFileName: String = '';
                       const AOverwriteExistingFile: Boolean = False);
var
  FileName: String;
begin
  FileName:= AFileName;
  if not AOverwriteExistingFile then
    if FileExists(FileName) then
      if not Confirm('Файл "' + FileName +
                     '" уже существует! Перезаписать файл?') then Exit;
  AWorkbook.WriteToFile(FileName, AFormat, True);
  if ADoneMessage<>EmptyStr then ShowInfo(ADoneMessage);
end;

procedure SaveWithDialog(const AWorkbook: TsWorkbook;
                         const ADoneMessage: String;
                         const ADefaultFileName: String;
                         const AOverwriteExistingFile: Boolean = False);
var
  FileFormat: TsSpreadsheetFormat;
  FileName: String;
begin
  FileName:=  TextToFileName(ADefaultFileName);
  if not OpenSaveDialog(FileName, FileFormat) then Exit;
  SaveToFormat(AWorkbook, FileFormat, ADoneMessage,
               FileName, AOverwriteExistingFile);
end;

procedure SheetPageSettings(const ASheet: TsWorksheet;
                       const AOrient: TsPageOrientation = spoPortrait;
                       const APageFit: TPageFit = pfWidth;
                       const AShowHeaders: Boolean = True;
                       const AShowGridLines: Boolean = True;
                       const ATopMargin: Double=10;     //mm
                       const ALeftMargin: Double=10;    //mm
                       const ARightMargin: Double=10;   //mm
                       const ABottomMargin: Double=10;  //mm
                       const AFooterMargin: Double=0;   //mm
                       const AHeaderMargin: Double=0);  //mm
begin
  ASheet.Options:= [];
  if AShowHeaders then
    ASheet.Options:= ASheet.Options + [soShowHeaders];
  if AShowGridLines then
    ASheet.Options:= ASheet.Options + [soShowGridLines];
  ASheet.PageLayout.Orientation:= AOrient;
  ASheet.PageLayout.TopMargin:= ATopMargin;
  ASheet.PageLayout.LeftMargin:= ALeftMargin;
  ASheet.PageLayout.RightMargin:= ARightMargin;
  ASheet.PageLayout.BottomMargin:= ABottomMargin;
  ASheet.PageLayout.HeaderMargin:= AHeaderMargin;
  ASheet.PageLayout.FooterMargin:= AFooterMargin;
  ASheet.PageLayout.Options:=
    ASheet.PageLayout.Options + [poFitPages, poHorCentered];
  case APageFit of
    pfOnePage: begin
                 ASheet.PageLayout.FitWidthToPages:= 1;
                 ASheet.PageLayout.FitHeightToPages:= 1;
               end;
    pfWidth  : begin
                 ASheet.PageLayout.FitWidthToPages:= 1;
                 ASheet.PageLayout.FitHeightToPages:= 0;
               end;
    pfHeight : begin
                 ASheet.PageLayout.FitWidthToPages:= 0;
                 ASheet.PageLayout.FitHeightToPages:= 1;
               end;
  end;
end;

{ TGridExporter }

constructor TGridExporter.Create(const AGrid: TsWorksheetGrid);
begin
  FWorkbook:= AGrid.Workbook;
  PageSettings;
end;

destructor TGridExporter.Destroy;
begin
  inherited Destroy;
end;

procedure TGridExporter.SetSheetName(const AName: String);
begin
  FWorkbook.ActiveWorksheet.Name:= TextToSheetName(AName);
end;

function TGridExporter.GetSheetName: String;
begin
  Result:= FWorkbook.ActiveWorksheet.Name;
end;

procedure TGridExporter.SaveToXLSX(const ADoneMessage: String;
  const AFileName: String; const AOverwriteExistingFile: Boolean);
begin
  SaveToFormat(FWorkbook, sfOOXML, ADoneMessage,
               AFileName, AOverwriteExistingFile);
end;

procedure TGridExporter.SaveToODS(const ADoneMessage: String;
  const AFileName: String; const AOverwriteExistingFile: Boolean);
begin
  SaveToFormat(FWorkbook, sfOpenDocument, ADoneMessage,
               AFileName, AOverwriteExistingFile);
end;

procedure TGridExporter.Save(const ADoneMessage: String;
                              const ADefaultFileName: String;
                              const AOverwriteExistingFile: Boolean = False);
begin
  SaveWithDialog(FWorkbook, ADoneMessage,
                 ADefaultFileName, AOverwriteExistingFile);
end;

procedure TGridExporter.PageSettings(
                             const AOrient: TsPageOrientation = spoPortrait;
                             const APageFit: TPageFit = pfWidth;
                             const AShowHeaders: Boolean = True;
                             const AShowGridLines: Boolean = True;
                             const ATopMargin: Double=10;     //mm
                             const ALeftMargin: Double=10;    //mm
                             const ARightMargin: Double=10;   //mm
                             const ABottomMargin: Double=10;  //mm
                             const AFooterMargin: Double=0;   //mm
                             const AHeaderMargin: Double=0);  //mm
begin
  SheetPageSettings(FWorkBook.ActiveWorksheet, AOrient, APageFit,
              AShowHeaders, AShowGridLines,
              ATopMargin, ALeftMargin, ARightMargin, ABottomMargin,
              AFooterMargin, AHeaderMargin);

end;

{ TSheetExporter }

constructor TSheetExporter.Create;
begin
  inherited Create;
  FWorkbook:= TsWorkbook.Create;
end;

destructor TSheetExporter.Destroy;
begin
  FreeAndNil(FWorkbook);
  inherited Destroy;
end;

function TSheetExporter.AddWorksheet(const AName: String): TsWorksheet;
begin
  Result:= FWorkbook.AddWorksheet(TextToSheetName(AName));
  FWorkbook.ActiveWorksheet:= Result;
  PageSettings;
end;

procedure TSheetExporter.SaveToXLSX(const ADoneMessage: String;
  const AFileName: String; const AOverwriteExistingFile: Boolean);
begin
  SaveToFormat(FWorkbook, sfOOXML, ADoneMessage,
               AFileName, AOverwriteExistingFile);
end;

procedure TSheetExporter.SaveToODS(const ADoneMessage: String;
  const AFileName: String; const AOverwriteExistingFile: Boolean);
begin
  SaveToFormat(FWorkbook, sfOpenDocument, ADoneMessage,
               AFileName, AOverwriteExistingFile);
end;

procedure TSheetExporter.Save(const ADoneMessage: String;
                              const ADefaultFileName: String;
                              const AOverwriteExistingFile: Boolean = False);
begin
  SaveWithDialog(FWorkbook, ADoneMessage,
                 ADefaultFileName, AOverwriteExistingFile);
end;

procedure TSheetExporter.PageSettings(
                             const AOrient: TsPageOrientation = spoPortrait;
                             const APageFit: TPageFit = pfWidth;
                             const AShowHeaders: Boolean = True;
                             const AShowGridLines: Boolean = True;
                             const ATopMargin: Double=10;     //mm
                             const ALeftMargin: Double=10;    //mm
                             const ARightMargin: Double=10;   //mm
                             const ABottomMargin: Double=10;  //mm
                             const AFooterMargin: Double=0;   //mm
                             const AHeaderMargin: Double=0);  //mm
begin
  SheetPageSettings(FWorkBook.ActiveWorksheet, AOrient, APageFit,
              AShowHeaders, AShowGridLines,
              ATopMargin, ALeftMargin, ARightMargin, ABottomMargin,
              AFooterMargin, AHeaderMargin);

end;

end.

