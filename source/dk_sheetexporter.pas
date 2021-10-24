unit DK_SheetExporter;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DK_Dialogs, DK_StrUtils, DK_SheetConst, Dialogs,
  fpstypes, fpspreadsheet;

type

  TPageFit = (pfOnePage, //всё на одной странице
              pfWidth,   //заполнить по ширине 1 страницы (по высоте - сколько нужно, чтобы вместить все)
              pfHeight); //заполнить по высоте 1 страницы (по ширине - сколько нужно, чтобы вместить все)

  { TSheetExporter }

  TSheetExporter = class (TObject)
    private
      FWorkbook: TsWorkbook;
      procedure SaveToFormat(const AFormat: TsSpreadsheetFormat;
                           const ADoneMessage: String = '';
                           const AFileName: String = '';
                           const AOverwriteExistingFile: Boolean = False);
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

function OpenSaveDialog(var AFileName: String; out AFormat: TsSpreadsheetFormat): Boolean;
var
  SD: TSaveDialog;
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
    1: AFormat:= sfOOXML;
    2: AFormat:= sfOpenDocument;
    end;
    AFileName:= SD.FileName;
  end;
  FreeAndNil(SD);
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
  PageSettings;
end;

procedure TSheetExporter.SaveToFormat(const AFormat: TsSpreadsheetFormat;
                           const ADoneMessage: String = '';
                           const AFileName: String = '';
                           const AOverwriteExistingFile: Boolean = False);
begin
  if not AOverwriteExistingFile then
    if FileExists(AFileName) then
      if not Confirm('Файл "' + AFileName +
                     '" уже существует! Перезаписать файл?') then Exit;
  FWorkbook.WriteToFile(AFileName, AFormat, True);
  if ADoneMessage<>EmptyStr then ShowInfo(ADoneMessage);
end;

procedure TSheetExporter.SaveToXLSX(const ADoneMessage: String;
  const AFileName: String; const AOverwriteExistingFile: Boolean);
begin
  SaveToFormat(sfOOXML, ADoneMessage, AFileName, AOverwriteExistingFile);
end;

procedure TSheetExporter.SaveToODS(const ADoneMessage: String;
  const AFileName: String; const AOverwriteExistingFile: Boolean);
begin
  SaveToFormat(sfOpenDocument, ADoneMessage, AFileName, AOverwriteExistingFile);
end;

procedure TSheetExporter.Save(const ADoneMessage: String;
                              const ADefaultFileName: String;
                              const AOverwriteExistingFile: Boolean = False);
var
  FileFormat: TsSpreadsheetFormat;
  FileName: String;
begin
  FileName:=  ADefaultFileName;
  if not OpenSaveDialog(FileName, FileFormat) then Exit;
  SaveToFormat(FileFormat, ADoneMessage, FileName, AOverwriteExistingFile);
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
var
  WSheet: TsWorksheet;
begin
  WSheet:= FWorkBook.ActiveWorksheet;
  WSheet.Options:= [];
  if AShowHeaders then
    WSheet.Options:= WSheet.Options + [soShowHeaders];
  if AShowGridLines then
    WSheet.Options:= WSheet.Options + [soShowGridLines];
  WSheet.PageLayout.Orientation:= AOrient;
  WSheet.PageLayout.TopMargin:= ATopMargin;
  WSheet.PageLayout.LeftMargin:= ALeftMargin;
  WSheet.PageLayout.RightMargin:= ARightMargin;
  WSheet.PageLayout.BottomMargin:= ABottomMargin;
  WSheet.PageLayout.HeaderMargin:= AHeaderMargin;
  WSheet.PageLayout.FooterMargin:= AFooterMargin;
  WSheet.PageLayout.Options:=
    WSheet.PageLayout.Options + [poFitPages, poHorCentered];
  case APageFit of
    pfOnePage: begin
                 WSheet.PageLayout.FitWidthToPages:= 1;
                 WSheet.PageLayout.FitHeightToPages:= 1;
               end;
    pfWidth  : begin
                 WSheet.PageLayout.FitWidthToPages:= 1;
                 WSheet.PageLayout.FitHeightToPages:= 0;
               end;
    pfHeight : begin
                 WSheet.PageLayout.FitWidthToPages:= 0;
                 WSheet.PageLayout.FitHeightToPages:= 1;
               end;
  end;
end;

end.

