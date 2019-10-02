pageextension 60000 "Config. Package Subform Ext." extends "Config. Package Subform"
{
    actions
    {
        // Add changes to page actions here
        addafter(ImportFromExcel)
        {
            action(ExportBigDataExcel)
            {
                ApplicationArea = "#Basic", "#Suite";
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
        ConfigPackageTable: Record "Config. Package Table";
        MultipleTablesSelectedQst: TextConst ENU = '%1 tables have been selected. Do you want to continue?', RUS = 'Выбрано таблиц: %1. Продолжить?';
        SingleTableSelectedQst: TextConst ENU = 'One table has been selected. Do you want to continue?', RUS = 'Выбрана одна таблица. Продолжить?';

    procedure ImportBigDataExcel()
    var
        ConfigPackageTable: Record "Config. Package Table";
        ConfigExcelExchangeExt: Codeunit "Config. Excel Exchange Ext.";
    begin
        CurrPage.SETSELECTIONFILTER(ConfigPackageTable);
        ConfigExcelExchangeExt.SetSelectedTables(ConfigPackageTable);
        ConfigExcelExchangeExt.ImportExcelFromSelectedPackage("Package Code");
    end;

    procedure ExportBigDataExcel()
    var
        ConfigExcelExchangeExt: Codeunit "Config. Excel Exchange Ext.";
    begin
        CurrPage.SETSELECTIONFILTER(ConfigPackageTable);
        IF CONFIRM(SelectionConfirmMessage, TRUE) THEN
            ConfigExcelExchangeExt.ExportExcelFromTables(ConfigPackageTable);
    end;

    LOCAL procedure SelectionConfirmMessage(): Text
    begin
        IF ConfigPackageTable.COUNT <> 1 THEN
            EXIT(STRSUBSTNO(MultipleTablesSelectedQst, ConfigPackageTable.COUNT));

        EXIT(SingleTableSelectedQst);
    end;
}