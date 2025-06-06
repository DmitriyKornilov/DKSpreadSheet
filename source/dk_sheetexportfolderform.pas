unit DK_SheetExportFolderForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Buttons,
  ExtCtrls, DK_StrUtils, DK_CtrlUtils, DK_Dialogs;

type

  { TSheetExportFolderForm }

  TSheetExportFolderForm = class(TForm)
    ButtonPanel: TPanel;
    ButtonPanelBevel: TBevel;
    CancelButton: TSpeedButton;
    FileTypeComboBox: TComboBox;
    FolderEdit: TEdit;
    FolderLabel: TLabel;
    FileTypeLabel: TLabel;
    FolderButton: TSpeedButton;
    PX24: TImageList;
    PX30: TImageList;
    PX36: TImageList;
    PX42: TImageList;
    SaveButton: TSpeedButton;
    FolderDialog: TSelectDirectoryDialog;
    procedure CancelButtonClick(Sender: TObject);
    procedure FolderButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
  private

  public

  end;

var
  SheetExportFolderForm: TSheetExportFolderForm;

  function SheetExportFolderFormOpen(var AFolderName: String;
                                     var AFileType: Byte //0-XLSX, 1-ODS
                                     ): Boolean;

implementation

function SheetExportFolderFormOpen(var AFolderName: String; var AFileType: Byte): Boolean;
var
  FolderForm: TSheetExportFolderForm;
begin
  FolderForm:= TSheetExportFolderForm.Create(nil);
  try
    FolderForm.FolderEdit.Text:= AFolderName;
    FolderForm.FileTypeComboBox.ItemIndex:= AFileType;
    Result:= FolderForm.ShowModal=mrOK;
    AFolderName:= FolderForm.FolderEdit.Text;
    AFileType:= FolderForm.FileTypeComboBox.ItemIndex;
  finally
    FreeAndNil(FolderForm)
  end;
end;

{$R *.lfm}

{ TSheetExportFolderForm }

procedure TSheetExportFolderForm.CancelButtonClick(Sender: TObject);
begin
  ModalResult:= mrCancel;
end;

procedure TSheetExportFolderForm.FolderButtonClick(Sender: TObject);
begin
  if not FolderDialog.Execute then Exit;
  FolderEdit.Text:= FolderDialog.Filename;
end;

procedure TSheetExportFolderForm.FormShow(Sender: TObject);
var
  Images: TImageList;
begin
  Images:= ChooseImageListForScreenPPI(PX24, PX30, PX36, PX42);
  FolderButton.Images:= Images;
  SaveButton.Images:= Images;
  CancelButton.Images:= Images;
end;

procedure TSheetExportFolderForm.SaveButtonClick(Sender: TObject);
var
  S: String;
begin
  S:= STrim(FolderEdit.Text);

  if SEmpty(S) then
  begin
    Inform('Не указана папка!');
    Exit;
  end;

  if not DirectoryExists(S) then
  begin
    if Confirm('Папки "' + S + '" не существует. Создать новую папку?') then
    begin
      if not CreateDir(S) then
      begin
        Inform('Не удалось создать папку "'+ S + '"!');
        Exit;
      end;
    end
    else Exit;
  end;

  FolderEdit.Text:= S;
  ModalResult:= mrOK;
end;

end.

