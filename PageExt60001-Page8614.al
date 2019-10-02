pageextension 60001 "Config. Package Card Ext." extends "Config. Package Card"
{
    actions
    {
        // Add changes to page actions here
        addafter(ImportFromExcel)
        {
            action(ExportBigDataExcel)
            {
                ApplicationArea = "#Basic", "#Suite";
                Promoted = true;
                PromotedCategory = Category5;
                CaptionML = ENU = 'Export BLOB to Excel', RUS = 'Экспорт BLOB в Excel';
                ToolTipML = ENU = 'Export the BLOB data in the package to Excel.', RUS = 'Экспорт BLOB данных из пакета в Excel.';
                Image = ExportToExcel;

                trigger OnAction()
                begin
                    ExportBigDataExcel;
                end;
            }
            action(ImportBigDataExcel)
            {
                ApplicationArea = "#Basic", "#Suite";
                Promoted = true;
                PromotedCategory = Category5;
                CaptionML = ENU = 'Import BLOB from Excel', RUS = 'Импорт BLOB из Excel';
                ToolTipML = ENU = 'Begin the migration of legacy BLOB data.', RUS = 'Начало миграции унаследованных BLOB данных.';
                Image = ImportExcel;

                trigger OnAction()
                begin
                    ImportBigDataExcel;
                end;

            }
        }
    }

    var
        ConfigExcelExchangeExt: Codeunit "Config. Excel Exchange Ext.";
        ConfirmManagement: Codeunit "Confirm Management";
        Text004: TextConst ENU = 'Export package %1 with %2 tables?', RUS = 'Экспортировать пакет %1, в котором содержится таблиц: %2?';
        // SingleTableSelectedQst: TextConst ENU = 'One table has been selected. Do you want to continue?', RUS = 'Выбрана одна таблица. Продолжить?';

    procedure ImportBigDataExcel()
    begin
        ConfigExcelExchangeExt.ImportExcelFromSelectedPackage(Code);
    end;

    procedure ExportBigDataExcel()
    var
        ConfigPackageTable: Record "Config. Package Table";
    begin
        TESTFIELD(Code);

        ConfigPackageTable.SETRANGE("Package Code", Code);
        IF ConfirmManagement.ConfirmProcess(STRSUBSTNO(Text004, Code, ConfigPackageTable.COUNT), TRUE) THEN
            ConfigExcelExchangeExt.ExportExcelFromTables(ConfigPackageTable);
    end;
}