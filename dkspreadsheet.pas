{ This file was automatically created by Lazarus. Do not edit!
  This source is only used to compile and install the package.
 }

unit DKSpreadSheet;

{$warn 5023 off : no warning about unused units}
interface

uses
  DK_SheetConst, DK_SheetExporter, DK_SheetWriter, DK_SheetUtils, 
  DK_SheetTables, LazarusPackageIntf;

implementation

procedure Register;
begin
end;

initialization
  RegisterPackage('DKSpreadSheet', @Register);
end.
