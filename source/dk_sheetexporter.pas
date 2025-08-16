unit DK_SheetExporter;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Dialogs,
  fpstypes, fpspreadsheet, fpspreadsheetgrid, fpsallformats,
  DK_Dialogs, DK_StrUtils, DK_SheetConst, DK_SheetExportFolderForm;


const
  spoPortrait  = fpsTypes.spoPortrait;
  spoLandscape = fpsTypes.spoLandscape;

type
  TsPageOrientation = fpsTypes.TsPageOrientation;
  TsWorksheet       = fpspreadsheet.TsWorksheet;
  TsWorksheetGrid   = fpspreadsheetgrid.TsWorksheetGrid;

  TPageFit = (pfNone,
              pfOnePage, //всё на одной странице
              pfWidth,   //заполнить по ширине 1 страницы (по высоте - сколько нужно, чтобы вместить все)
              pfHeight); //заполнить по высоте 1 страницы (по ширине - сколько нужно, чтобы вместить все)

  { TCustomExporter }

  TCustomExporter = class (TObject)
  protected
    FWorkbook: TsWorkbook;
  public
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

    procedure PageMargins(const ATopMargin: Double=10;     //mm
                           const ALeftMargin: Double=10;    //mm
                           const ARightMargin: Double=10;   //mm
                           const ABottomMargin: Double=10;  //mm
                           const AFooterMargin: Double=0;   //mm
                           const AHeaderMargin: Double=0);  //mm

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

  { TSingleExporter }

  TSingleExporter = class (TCustomExporter)
  private
    procedure SetSheetName(const ASheetName: String);
    function GetSheetName: String;
  public
    property SheetName: String read GetSheetName write SetSheetName;
  end;

  { TGridExporter }

  TGridExporter = class (TSingleExporter)
  public
    constructor Create(const AGrid: TsWorksheetGrid);
  end;

  { TSheetExporter }

  TSheetExporter = class (TSingleExporter)
  public
    constructor Create(const ASheet: TsWorksheet);
  end;

  { TSheetsExporter }

  TSheetsExporter = class (TCustomExporter)
  public
    constructor Create;
    destructor  Destroy; override;
    function AddWorksheet(const ASheetName: String): TsWorksheet;
  end;

  { TBooksExporter }

  TBooksExporter = class (TObject)
  private
    FWorkbook: TsWorkbook;
    FFolderName: String;
    FFileExtension: String;
    FFormat: TsSpreadsheetFormat;
    FCanExport: Boolean;
  public
    constructor Create;
    destructor  Destroy; override;
    function BeginExport: Boolean;
    procedure EndExport(const ADoneMessage: String = '');
    procedure Save(const AFileName: String);
    function AddWorksheet(const ASheetName: String;
                          const AInNewWorkbook: Boolean = True): TsWorksheet;
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
  if ADoneMessage<>EmptyStr then Inform(ADoneMessage);
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

procedure SheetMargins(const ASheet: TsWorksheet;
                       const ATopMargin: Double=10;     //mm
                       const ALeftMargin: Double=10;    //mm
                       const ARightMargin: Double=10;   //mm
                       const ABottomMargin: Double=10;  //mm
                       const AFooterMargin: Double=0;   //mm
                       const AHeaderMargin: Double=0);
begin
  ASheet.PageLayout.TopMargin:= ATopMargin;
  ASheet.PageLayout.LeftMargin:= ALeftMargin;
  ASheet.PageLayout.RightMargin:= ARightMargin;
  ASheet.PageLayout.BottomMargin:= ABottomMargin;
  ASheet.PageLayout.HeaderMargin:= AHeaderMargin;
  ASheet.PageLayout.FooterMargin:= AFooterMargin;
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

  SheetMargins(ASheet, ATopMargin, ALeftMargin, ARightMargin, ABottomMargin,
               AHeaderMargin, AFooterMargin);

  if APageFit=pfNone then
  begin
    ASheet.PageLayout.Options:=
      ASheet.PageLayout.Options - [poFitPages, poHorCentered];
    ASheet.PageLayout.FitWidthToPages:= 0;
    ASheet.PageLayout.FitHeightToPages:= 0;
    Exit;
  end;

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

{ TCustomExporter }

procedure TCustomExporter.SaveToXLSX(const ADoneMessage: String = '';
                         const AFileName: String = '';
                         const AOverwriteExistingFile: Boolean = False);
begin
  if FWorkbook.GetWorksheetCount>0 then
    FWorkbook.ActiveWorksheet:= FWorkbook.GetFirstWorksheet;
  SaveToFormat(FWorkbook, sfOOXML, ADoneMessage, AFileName, AOverwriteExistingFile);
end;

procedure TCustomExporter.SaveToODS(const ADoneMessage: String = '';
                         const AFileName: String = '';
                         const AOverwriteExistingFile: Boolean = False);
begin
  if FWorkbook.GetWorksheetCount>0 then
    FWorkbook.ActiveWorksheet:= FWorkbook.GetFirstWorksheet;
  SaveToFormat(FWorkbook, sfOpenDocument, ADoneMessage, AFileName, AOverwriteExistingFile);
end;

procedure TCustomExporter.Save(const ADoneMessage: String = '';
                   const ADefaultFileName: String = '';
                   const AOverwriteExistingFile: Boolean = False);
begin
  if FWorkbook.GetWorksheetCount>0 then
    FWorkbook.ActiveWorksheet:= FWorkbook.GetFirstWorksheet;
  SaveWithDialog(FWorkbook, ADoneMessage, ADefaultFileName, AOverwriteExistingFile);
end;

procedure TCustomExporter.PageMargins(const ATopMargin: Double=10;     //mm
                           const ALeftMargin: Double=10;    //mm
                           const ARightMargin: Double=10;   //mm
                           const ABottomMargin: Double=10;  //mm
                           const AFooterMargin: Double=0;   //mm
                           const AHeaderMargin: Double=0);  //mm
begin
  SheetMargins(FWorkBook.ActiveWorksheet, ATopMargin, ALeftMargin, ARightMargin,
               ABottomMargin, AHeaderMargin, AFooterMargin);
end;

procedure TCustomExporter.PageSettings(const AOrient: TsPageOrientation = spoPortrait;
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

{ TSingleExporter }

procedure TSingleExporter.SetSheetName(const ASheetName: String);
begin
  FWorkbook.ActiveWorksheet.Name:= TextToSheetName(ASheetName);
end;

function TSingleExporter.GetSheetName: String;
begin
  Result:= FWorkbook.ActiveWorksheet.Name;
end;

{ TGridExporter }

constructor TGridExporter.Create(const AGrid: TsWorksheetGrid);
begin
  FWorkbook:= AGrid.Workbook;
  PageSettings;
end;

{ TSheetExporter }

constructor TSheetExporter.Create(const ASheet: TsWorksheet);
begin
  FWorkbook:= ASheet.Workbook;
  PageSettings;
end;

{ TSheetsExporter }

constructor TSheetsExporter.Create;
begin
  inherited Create;
  FWorkbook:= TsWorkbook.Create;
end;

destructor TSheetsExporter.Destroy;
begin
  FreeAndNil(FWorkbook);
  inherited Destroy;
end;

function TSheetsExporter.AddWorksheet(const ASheetName: String): TsWorksheet;
begin
  Result:= FWorkbook.AddWorksheet(TextToSheetName(ASheetName));
  FWorkbook.ActiveWorksheet:= Result;
  PageSettings;
end;

{ TBooksExporter }

constructor TBooksExporter.Create;
begin
  FCanExport:= False;
end;

destructor TBooksExporter.Destroy;
begin
  if Assigned(FWorkbook) then FreeAndNil(FWorkbook);
  inherited Destroy;
end;

function TBooksExporter.BeginExport: Boolean;
var
  FileType: Byte;
begin
  FFolderName:= EmptyStr;
  FileType:= 0;
  Result:= SheetExportFolderFormOpen(FFolderName, FileType);
  if not Result then Exit;
  case FileType of
  0:
    begin
      FFormat:= sfOOXML;
      FFileExtension:= '.xlsx';
    end;
  1:
    begin
      FFormat:= sfOpenDocument;
      FFileExtension:= '.ods';
    end;
  end;
  FCanExport:= True;
end;

procedure TBooksExporter.EndExport(const ADoneMessage: String = '');
begin
  if ADoneMessage<>EmptyStr then
    Inform(ADoneMessage);
end;

procedure TBooksExporter.Save(const AFileName: String);
var
  FullFileName: String;
begin
  if not FCanExport then Exit;
  if FWorkbook.GetWorksheetCount>0 then
    FWorkbook.ActiveWorksheet:= FWorkbook.GetFirstWorksheet;
  FullFileName:= FFolderName + DirectorySeparator + AFileName + FFileExtension;
  SaveToFormat(FWorkbook, FFormat, EmptyStr, FullFileName, True);
end;

function TBooksExporter.AddWorksheet(const ASheetName: String;
                                     const AInNewWorkbook: Boolean = True): TsWorksheet;
begin
  if AInNewWorkbook then
  begin
    if Assigned(FWorkbook) then
      FreeAndNil(FWorkbook);
  end;
  if not Assigned(FWorkbook) then
    FWorkbook:= TsWorkbook.Create;
  Result:= FWorkbook.AddWorksheet(TextToSheetName(ASheetName));
  FWorkbook.ActiveWorksheet:= Result;
  PageSettings;
end;

procedure TBooksExporter.PageSettings(const AOrient: TsPageOrientation = spoPortrait;
                           const APageFit: TPageFit = pfWidth;
                           const AShowHeaders: Boolean = True;
                           const AShowGridLines: Boolean = True;
                           const ATopMargin: Double=10;     //mm
                           const ALeftMargin: Double=10;    //mm
                           const ARightMargin: Double=10;   //mm
                           const ABottomMargin: Double=10;  //mm
                           const AFooterMargin: Double=0;   //mm
                           const AHeaderMargin: Double=0);
begin
  SheetPageSettings(FWorkBook.ActiveWorksheet, AOrient, APageFit,
              AShowHeaders, AShowGridLines,
              ATopMargin, ALeftMargin, ARightMargin, ABottomMargin,
              AFooterMargin, AHeaderMargin);
end;

end.

