codeunit 60001 "Config. Excel Exchange Ext."
{

    trigger OnRun()
    begin
    end;

    var
        SelectedConfigPackage: Record 8623;
        SelectedTable: Record 2000000026 temporary;
        ConfigXMLExchange: Codeunit "Config. XML Exchange Ext.";
        FileMgt: Codeunit 419;
        ConfigProgressBar: Codeunit 8615;
        ConfigValidateMgt: Codeunit 8617;
        CannotCreateXmlSchemaErr: Label 'Could not create XML Schema.';
        CreatingExcelMsg: Label 'Creating Excel worksheet';
        OpenXMLManagement: Codeunit 6223;
        TypeHelper: Codeunit 10;
        WrkbkReader: DotNet WorkbookReader;
        WrkbkWriter: DotNet WorkbookWriter;
        WrkShtWriter: DotNet WorksheetWriter;
        Worksheet: DotNet Worksheet;
        Workbook: DotNet Workbook;
        WorkBookPart: DotNet WorkbookPart;
        CreateWrkBkFailedErr: Label 'Could not create the Excel workbook.';
        WrkShtHelper: DotNet WorksheetHelper;
        DataSet: DotNet DataSet;
        DataTable: DotNet DataTable;
        DataColumn: DotNet DataColumn;
        StringBld: DotNet StringBuilder;
        id: BigInteger;
        HideDialog: Boolean;
        CommentVmlShapeXmlTxt: Label '<v:shape id="%1" type="#_x0000_t202" style=''position:absolute;  margin-left:59.25pt;margin-top:1.5pt;width:96pt;height:55.5pt;z-index:1;  visibility:hidden'' fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style=''mso-direction-alt:auto''><div style=''text-align:left''/></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>%2</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>%3</x:Row><x:Column>%4</x:Column></x:ClientData></v:shape>', Locked = true;
        VmlDrawingXmlTxt: Label '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"  path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>', Locked = true;
        EndXmlTokenTxt: Label '</xml>', Locked = true;
        VmlShapeAnchorTxt: Label '%1,15,%2,10,%3,31,8,9', Locked = true;
        FileExtensionFilterTok: Label 'Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*';
        ExcelFileNameTok: Label '*%1.xlsx', Comment = '%1 = String generated from current datetime to make sure file names are unique ';
        ExcelFileExtensionTok: Label '.xlsx';
        InvalidDataInSheetMsg: Label 'Data in sheet ''%1'' could not be imported, because the sheet has an unexpected format.', Comment = '%1=excel sheet name';
        ImportFromExcelMsg: Label 'Import from Excel';
        FileOnServer: Boolean;

    [Scope('Personalization')]
    procedure ExportExcelFromConfig(var ConfigLine: Record 8622): Text
    var
        ConfigPackageTable: Record 8613;
        ConfigMgt: Codeunit 8616;
        FileName: Text;
        "Filter": Text;
    begin
        ConfigLine.FINDFIRST;
        ConfigPackageTable.SETRANGE("Package Code", ConfigLine."Package Code");
        Filter := ConfigMgt.MakeTableFilter(ConfigLine, TRUE);
        IF Filter <> '' THEN
            ConfigPackageTable.SETFILTER("Table ID", Filter);

        ConfigPackageTable.SETRANGE("Dimensions as Columns", TRUE);
        IF ConfigPackageTable.FINDSET THEN
            REPEAT
                IF NOT (ConfigPackageTable.DimensionPackageDataExist OR (ConfigPackageTable.DimensionFieldsCount > 0)) THEN
                    ConfigPackageTable.InitDimensionFields;
            UNTIL ConfigPackageTable.NEXT = 0;
        ConfigPackageTable.SETRANGE("Dimensions as Columns");
        ExportExcel(FileName, ConfigPackageTable, TRUE, FALSE);
        EXIT(FileName);
    end;

    [Scope('Personalization')]
    procedure ExportExcelFromPackage(ConfigPackage: Record 8623): Boolean
    var
        ConfigPackageTable: Record 8613;
        FileName: Text;
    begin
        ConfigPackageTable.SETRANGE("Package Code", ConfigPackage.Code);
        EXIT(ExportExcel(FileName, ConfigPackageTable, FALSE, FALSE));
    end;

    [Scope('Personalization')]
    procedure ExportExcelFromTables(var ConfigPackageTable: Record 8613): Boolean
    var
        FileName: Text;
    begin
        EXIT(ExportExcel(FileName, ConfigPackageTable, FALSE, FALSE));
    end;

    [Scope('Personalization')]
    procedure ExportExcelTemplateFromTables(var ConfigPackageTable: Record 8613): Boolean
    var
        FileName: Text;
    begin
        EXIT(ExportExcel(FileName, ConfigPackageTable, FALSE, TRUE));
    end;

    [Scope('Personalization')]
    procedure ExportExcel(var FileName: Text; var ConfigPackageTable: Record 8613; ExportFromWksht: Boolean; SkipData: Boolean): Boolean
    var
        TempBlob: Record 99008535;
        VmlDrawingPart: DotNet VmlDrawingPart;
        TableDefinitionPart: DotNet TableDefinitionPart;
        TableParts: DotNet TableParts;
        TablePart: DotNet TablePart;
        SingleXMLCells: DotNet SingleXmlCells;
        XmlTextWriter: DotNet XmlTextWriters;
        FileMode: DotNet FileMode;
        Encoding: DotNet Encoding;
        Caption: Text;
        RootElementName: Text;
        TempSetupDataFileName: Text;
        TempSchemaFileName: Text;
        DataTableCounter: Integer;
    begin
        OnBeforeExportExcel(ConfigPackageTable);

        TempSchemaFileName := CreateSchemaFile(ConfigPackageTable, RootElementName);
        TempSetupDataFileName := BuildDataSetForPackageTable(ExportFromWksht, ConfigPackageTable);

        CreateBook(TempBlob);
        WrkShtHelper := WrkShtHelper.WorksheetHelper(WrkbkWriter.FirstWorksheet.Worksheet);
        OpenXMLManagement.ImportSchema(WrkbkWriter, TempSchemaFileName, 1, RootElementName);
        OpenXMLManagement.CreateSchemaConnection(WrkbkWriter, TempSetupDataFileName);

        DataTableCounter := 1;

        IF NOT HideDialog THEN
            ConfigProgressBar.Init(ConfigPackageTable.COUNT, 1, CreatingExcelMsg);

        DataTable := DataSet.Tables.Item(1);

        IF ConfigPackageTable.FINDSET THEN
            REPEAT
                IF ISNULL(StringBld) THEN BEGIN
                    StringBld := StringBld.StringBuilder;
                    StringBld.Append(VmlDrawingXmlTxt);
                END;

                ConfigPackageTable.CALCFIELDS("Table Caption");
                IF NOT HideDialog THEN
                    ConfigProgressBar.Update(ConfigPackageTable."Table Caption");

                // Initialize WorkSheetWriter
                Caption := DELCHR(ConfigPackageTable."Table Caption", '=', '/');
                IF id < 1 THEN BEGIN
                    WrkShtWriter := WrkbkWriter.FirstWorksheet;
                    WrkShtWriter.Name := Caption;
                END ELSE
                    WrkShtWriter := WrkbkWriter.AddWorksheet(GetExcelWorksheetName(Caption, FORMAT(ConfigPackageTable."Table ID")));
                Worksheet := WrkShtWriter.Worksheet;

                // Add and initialize SingleCellTable part
                WrkShtWriter.AddSingleCellTablePart;
                SingleXMLCells := SingleXMLCells.SingleXmlCells;
                Worksheet.WorksheetPart.SingleCellTablePart.SingleXmlCells := SingleXMLCells;
                id += 3;

                OpenXMLManagement.AddAndInitializeCommentsPart(WrkShtWriter, VmlDrawingPart);
                AddPackageAndTableInformation(ConfigPackageTable, SingleXMLCells);
                AddAndInitializeTableDefinitionPart(ConfigPackageTable, ExportFromWksht, DataTableCounter, TableDefinitionPart, SkipData);
                IF NOT SkipData THEN
                    OpenXMLManagement.CopyDataToExcelTable(WrkShtWriter, DataTable);

                DataTableCounter += 2;
                TableParts := WrkShtWriter.CreateTableParts(1);
                WrkShtHelper.AppendElementToOpenXmlElement(Worksheet, TableParts);
                TablePart := WrkShtWriter.CreateTablePart(Worksheet.WorksheetPart.GetIdOfPart(TableDefinitionPart));
                WrkShtHelper.AppendElementToOpenXmlElement(TableParts, TablePart);

                StringBld.Append(EndXmlTokenTxt);

                XmlTextWriter := XmlTextWriter.XmlTextWriter(VmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8);
                XmlTextWriter.WriteRaw(StringBld.ToString);
                XmlTextWriter.Flush;
                XmlTextWriter.Close;

                CLEAR(StringBld);

            UNTIL ConfigPackageTable.NEXT = 0;

        FILE.ERASE(TempSchemaFileName);
        FILE.ERASE(TempSetupDataFileName);

        OpenXMLManagement.CleanMapInfo(WrkbkWriter.Workbook.WorkbookPart.CustomXmlMappingsPart.MapInfo);
        WrkbkWriter.Workbook.Save;
        WrkbkWriter.Close;
        ClearOpenXmlVariables;

        IF NOT HideDialog THEN
            ConfigProgressBar.Close;

        IF FileName = '' THEN
            FileName :=
              STRSUBSTNO(ExcelFileNameTok, FORMAT(CURRENTDATETIME, 0, '<Day,2>_<Month,2>_<Year4>_<Hours24>_<Minutes,2>_<Seconds,2>'));

        IF NOT FileOnServer THEN
            FileName := FileMgt.BLOBExport(TempBlob, FileName, NOT HideDialog)
        ELSE
            FileMgt.BLOBExportToServerFile(TempBlob, FileName);

        EXIT(FileName <> '');
    end;

    [Scope('Personalization')]
    procedure ImportExcelFromConfig(ConfigLine: Record 8622)
    var
        ConfigPackage: Record 8623;
        TempBlob: Record 99008535;
    begin
        ConfigLine.TESTFIELD("Line Type", ConfigLine."Line Type"::Table);
        ConfigLine.TESTFIELD("Table ID");
        IF ConfigPackage.GET(ConfigLine."Package Code") AND IsFileImportedToBLOB(TempBlob) THEN
            ImportExcel(TempBlob);
    end;

    [Scope('Personalization')]
    procedure ImportExcelFromPackage(): Boolean
    var
        TempBlob: Record 99008535;
    begin
        IF IsFileImportedToBLOB(TempBlob) THEN
            EXIT(ImportExcel(TempBlob));
        EXIT(FALSE)
    end;

    [Scope('Personalization')]
    procedure ImportExcelFromSelectedPackage(PackageCode: Code[20]): Boolean
    var
        TempBlob: Record 99008535;
    begin
        SelectedConfigPackage.GET(PackageCode);
        IF IsFileImportedToBLOB(TempBlob) THEN
            EXIT(ImportExcel(TempBlob));
        EXIT(FALSE)
    end;

    procedure SetSelectedTables(var ConfigPackageTable: Record 8613)
    begin
        IF ConfigPackageTable.FINDSET THEN
            REPEAT
                SelectedTable.Number := ConfigPackageTable."Table ID";
                IF SelectedTable.INSERT THEN;
            UNTIL ConfigPackageTable.NEXT = 0;
    end;

    local procedure IsTableSelected(TableId: Integer): Boolean
    begin
        IF SelectedTable.ISEMPTY THEN
            EXIT(TRUE);
        EXIT(SelectedTable.GET(TableId));
    end;

    local procedure IsWorksheetSelected(var TempConfigPackageTable: Record 8613 temporary; WrksheetId: Integer): Boolean
    begin
        TempConfigPackageTable.RESET;
        TempConfigPackageTable.SETRANGE("Processing Order", WrksheetId);
        EXIT(NOT TempConfigPackageTable.ISEMPTY);
    end;

    local procedure IsImportFromExcelConfirmed(var TempConfigPackageTable: Record 8613 temporary): Boolean
    var
        ConfigPackageImportPreview: Page 8617;
    begin
        IF GUIALLOWED AND NOT HideDialog THEN
            IF ReadPackageTableKeysFromExcel(TempConfigPackageTable) THEN BEGIN
                ConfigPackageImportPreview.SetData(SelectedConfigPackage.Code, TempConfigPackageTable);
                ConfigPackageImportPreview.RUNMODAL;
                EXIT(ConfigPackageImportPreview.IsImportConfirmed);
            END;
        EXIT(TRUE);
    end;

    [TryFunction]
    local procedure ReadPackageTableKeysFromExcel(var TempConfigPackageTable: Record 8613 temporary)
    var
        WrkShtReader: DotNet WorksheetReader;
        Enumerator: DotNet IEnumerator;
        CellData: DotNet CellData;
        Window: Dialog;
        WrkSheetId: Integer;
        SheetCount: Integer;
    begin
        Window.OPEN(ImportFromExcelMsg);
        WrkSheetId := WrkbkReader.FirstSheetId;
        SheetCount := WrkbkReader.Workbook.Sheets.ChildElements.Count + WrkSheetId;
        REPEAT
            WrkShtReader := WrkbkReader.GetWorksheetById(FORMAT(WrkSheetId));
            Enumerator := WrkShtReader.GetEnumerator;
            WHILE NextCellInRow(Enumerator, CellData, 1) DO
                FillImportPreviewBuffer(TempConfigPackageTable, WrkSheetId, CellData.ColumnNumber, CellData.Value);
            WrkSheetId += 1;
        UNTIL WrkSheetId >= SheetCount;
        SelectedTable.DELETEALL;
        IF TempConfigPackageTable.FINDFIRST THEN;
        Window.CLOSE;
    end;

    local procedure FillImportPreviewBuffer(var TempConfigPackageTable: Record 8613 temporary; WrkSheetId: Integer; ColumnNo: Integer; Value: Text)
    var
        ConfigPackage: Record 8623;
        ConfigPackageTable: Record 8613;
    begin
        WITH TempConfigPackageTable DO
            CASE ColumnNo OF
                1: // first column contains Package Code
                    "Package Code" := COPYSTR(Value, 1, MAXSTRLEN("Package Code"));
                3: // third column contains Table ID
                    BEGIN
                        "Processing Order" := WrkSheetId;
                        EVALUATE("Table ID", Value);
                        "Delayed Insert" := NOT ConfigPackage.GET("Package Code"); // New Package flag value
                        Validated := NOT ConfigPackageTable.GET("Package Code", "Table ID"); // New Table flag value
                        IF IsTableSelected("Table ID") THEN
                            INSERT;
                    END;
            END;
    end;

    local procedure NextCellInRow(Enumerator: DotNet IEnumerator; CellData: DotNet CellData; RowNo: Integer): Boolean
    begin
        IF Enumerator.MoveNext THEN BEGIN
            CellData := Enumerator.Current;
            EXIT(CellData.RowNumber = RowNo);
        END;
    end;

    [Scope('Personalization')]
    procedure ImportExcel(var TempBlob: Record 99008535) Imported: Boolean
    var
        TempConfigPackageTable: Record 8613 temporary;
        WorkBookPart: DotNet WorkbookPart;
        InStream: InStream;
        XMLSchemaDataFile: Text;
        WrkSheetId: Integer;
        DataColumnTableId: Integer;
        SheetCount: Integer;
    begin
        TempBlob.Blob.CREATEINSTREAM(InStream);
        WrkbkReader := WrkbkReader.Open(InStream);
        IF NOT IsImportFromExcelConfirmed(TempConfigPackageTable) THEN BEGIN
            CLEAR(WrkbkReader);
            EXIT(FALSE);
        END;
        WorkBookPart := WrkbkReader.Workbook.WorkbookPart;
        XMLSchemaDataFile := OpenXMLManagement.ExtractXMLSchema(WorkBookPart);

        WrkSheetId := WrkbkReader.FirstSheetId;
        SheetCount := WorkBookPart.Workbook.Sheets.ChildElements.Count + WrkSheetId;
        DataSet := DataSet.DataSet;
        DataSet.ReadXmlSchema(XMLSchemaDataFile);

        WrkSheetId := WrkbkReader.FirstSheetId;
        DataColumnTableId := 0;
        REPEAT
            IF IsWorksheetSelected(TempConfigPackageTable, WrkSheetId) THEN
                ReadWorksheetData(WrkSheetId, DataColumnTableId);
            WrkSheetId += 1;
            DataColumnTableId += 2;
        UNTIL WrkSheetId >= SheetCount;

        TempBlob.INIT;
        TempBlob.Blob.CREATEINSTREAM(InStream);
        DataSet.WriteXml(InStream);
        ConfigXMLExchange.SetExcelMode(TRUE);
        IF ConfigXMLExchange.ImportPackageXMLFromStream(InStream) THEN
            Imported := TRUE;

        EXIT(Imported);
    end;

    local procedure ReadWorksheetData(WrkSheetId: Integer; DataColumnTableId: Integer)
    var
        TempXMLBuffer: Record 1235 temporary;
        CellData: DotNet CellData;
        DataRow: DotNet DataRow;
        DataRow2: DotNet DataRow;
        Enumerator: DotNet IEnumerator;
        Type: DotNet Type;
        WrkShtReader: DotNet WorksheetReader;
        SheetHeaderRead: Boolean;
        ColumnCount: Integer;
        TotalColumnCount: Integer;
        RowIn: Integer;
        CurrentRowIndex: Integer;
        RowChanged: Boolean;
        FirstDataRow: Integer;
        CellValueText: Text;
    begin
        WrkShtReader := WrkbkReader.GetWorksheetById(FORMAT(WrkSheetId));
        IF InitColumnMapping(WrkShtReader, TempXMLBuffer) THEN BEGIN
            Enumerator := WrkShtReader.GetEnumerator;
            IF GetDataTable(DataColumnTableId) THEN BEGIN
                DataColumn := DataTable.Columns.Item(1);
                DataColumn.DataType := Type.GetType('System.String');
                DataTable.BeginLoadData;
                DataRow := DataTable.NewRow;
                SheetHeaderRead := FALSE;
                DataColumn := DataTable.Columns.Item(1);
                RowIn := 1;
                ColumnCount := 0;
                TotalColumnCount := 0;
                CurrentRowIndex := 1;
                FirstDataRow := 4;
                WHILE Enumerator.MoveNext DO BEGIN
                    CellData := Enumerator.Current;
                    CellValueText := CellData.Value;
                    RowChanged := CurrentRowIndex <> CellData.RowNumber;
                    IF NOT SheetHeaderRead THEN BEGIN // Read config and table information
                        IF (CellData.RowNumber = 1) AND (CellData.ColumnNumber = 1) THEN
                            DataRow.Item(1, CellValueText);
                        IF (CellData.RowNumber = 1) AND (CellData.ColumnNumber = 3) THEN BEGIN
                            DataColumn := DataTable.Columns.Item(0);
                            DataRow.Item(0, CellValueText);
                            DataTable.Rows.Add(DataRow);
                            DataColumn := DataTable.Columns.Item(2);
                            DataColumn.AllowDBNull(TRUE);
                            DataTable := DataSet.Tables.Item(DataColumnTableId + 1);
                            ColumnCount := 0;
                            TotalColumnCount := DataTable.Columns.Count - 1;
                            REPEAT
                                DataColumn := DataTable.Columns.Item(ColumnCount);
                                DataColumn.DataType := Type.GetType('System.String');
                                ColumnCount += 1;
                            UNTIL ColumnCount = TotalColumnCount;
                            ColumnCount := 0;
                            DataRow2 := DataTable.NewRow;
                            DataRow2.SetParentRow(DataRow);
                            SheetHeaderRead := TRUE;
                        END;
                    END ELSE BEGIN // Read data rows
                        IF (RowIn = 1) AND (CellData.RowNumber >= FirstDataRow) AND (CellData.ColumnNumber = 1) THEN BEGIN
                            TotalColumnCount := ColumnCount;
                            ColumnCount := 0;
                            RowIn += 1;
                            FirstDataRow := CellData.RowNumber;
                        END;

                        IF RowChanged AND (CellData.RowNumber > FirstDataRow) AND (RowIn <> 1) THEN BEGIN
                            DataTable.Rows.Add(DataRow2);
                            DataTable.EndLoadData;
                            DataRow2 := DataTable.NewRow;
                            DataRow2.SetParentRow(DataRow);
                            RowIn += 1;
                            ColumnCount := 0;
                        END;

                        IF RowIn <> 1 THEN
                            IF TempXMLBuffer.GET(CellData.ColumnNumber) THEN BEGIN
                                DataColumn := DataTable.Columns.Item(TempXMLBuffer."Parent Entry No.");
                                DataColumn.AllowDBNull(TRUE);
                                DataRow2.Item(TempXMLBuffer."Parent Entry No.", CellValueText);
                            END;

                        ColumnCount := CellData.ColumnNumber + 1;
                    END;
                    CurrentRowIndex := CellData.RowNumber;
                END;
                // Add the last row
                DataTable.Rows.Add(DataRow2);
                DataTable.EndLoadData;
            END ELSE
                MESSAGE(InvalidDataInSheetMsg, WrkShtReader.Name);
        END;
    end;

    [Scope('Personalization')]
    procedure ClearOpenXmlVariables()
    begin
        CLEAR(WrkbkReader);
        CLEAR(WrkbkWriter);
        CLEAR(WrkShtWriter);
        CLEAR(Workbook);
        CLEAR(WorkBookPart);
        CLEAR(WrkShtHelper);
    end;

    [Scope('Personalization')]
    procedure CreateBook(var TempBlob: Record 99008535)
    var
        InStream: InStream;
    begin
        TempBlob.Blob.CREATEINSTREAM(InStream);
        WrkbkWriter := WrkbkWriter.Create(InStream);
        IF ISNULL(WrkbkWriter) THEN
            ERROR(CreateWrkBkFailedErr);

        Workbook := WrkbkWriter.Workbook;
        WorkBookPart := Workbook.WorkbookPart;
    end;

    [Scope('Personalization')]
    procedure SetHideDialog(NewHideDialog: Boolean)
    begin
        HideDialog := NewHideDialog;
    end;

    local procedure CreateSchemaFile(var ConfigPackageTable: Record 8613; var RootElementName: Text): Text
    var
        ConfigDataSchema: XMLport 8610;
        OStream: OutStream;
        TempSchemaFile: File;
        TempSchemaFileName: Text;
    begin
        TempSchemaFile.CREATETEMPFILE;
        TempSchemaFileName := TempSchemaFile.NAME + '.xsd';
        TempSchemaFile.CLOSE;
        TempSchemaFile.CREATE(TempSchemaFileName);
        TempSchemaFile.CREATEOUTSTREAM(OStream);
        RootElementName := ConfigDataSchema.GetRootElementName;
        ConfigDataSchema.SETDESTINATION(OStream);
        ConfigDataSchema.SETTABLEVIEW(ConfigPackageTable);
        IF NOT ConfigDataSchema.EXPORT THEN
            ERROR(CannotCreateXmlSchemaErr);
        TempSchemaFile.CLOSE;
        EXIT(TempSchemaFileName);
    end;

    local procedure CreateXMLPackage(TempSetupDataFileName: Text; ExportFromWksht: Boolean; var ConfigPackageTable: Record 8613): Text
    begin
        CLEAR(ConfigXMLExchange);
        ConfigXMLExchange.SetExcelMode(TRUE);
        ConfigXMLExchange.SetCalledFromCode(TRUE);
        ConfigXMLExchange.SetPrefixMode(TRUE);
        ConfigXMLExchange.SetExportFromWksht(ExportFromWksht);
        ConfigXMLExchange.ExportPackageXML(ConfigPackageTable, TempSetupDataFileName);
        ConfigXMLExchange.SetExcelMode(FALSE);
        EXIT(TempSetupDataFileName);
    end;

    local procedure CreateTableColumnNames(var ConfigPackageField: Record 8616; var ConfigPackageTable: Record 8613; TableColumns: DotNet TableColumns)
    var
        "Field": Record 2000000041;
        Dimension: Record 348;
        XmlColumnProperties: DotNet XmlColumnProperties;
        TableColumn: DotNet TableColumn;
        WrkShtWriter2: DotNet WorksheetWriter;
        RecRef: RecordRef;
        FieldRef: FieldRef;
        TableColumnName: Text;
        ColumnID: Integer;
    begin
        RecRef.OPEN(ConfigPackageTable."Table ID");
        ConfigPackageField.SETCURRENTKEY("Package Code", "Table ID", "Processing Order");
        IF ConfigPackageField.FINDSET THEN BEGIN
            ColumnID := 1;
            REPEAT
                IF TypeHelper.GetField(ConfigPackageField."Table ID", ConfigPackageField."Field ID", Field) OR
                   ConfigPackageField.Dimension
                THEN BEGIN
                    IF ConfigPackageField.Dimension THEN
                        TableColumnName := ConfigPackageField."Field Caption" + ' ' + STRSUBSTNO('(%1)', Dimension.TABLECAPTION)
                    ELSE
                        TableColumnName := ConfigPackageField."Field Caption";
                    XmlColumnProperties := WrkShtWriter2.CreateXmlColumnProperties(
                        1,
                        '/DataList/' + (ConfigXMLExchange.GetElementName(ConfigPackageTable."Table Caption") + 'List') +
                        '/' + ConfigXMLExchange.GetElementName(ConfigPackageTable."Table Caption") +
                        '/' + ConfigXMLExchange.GetElementName(ConfigPackageField."Field Caption"),
                        WrkShtWriter.XmlDataType2XmlDataValues(
                          ConfigXMLExchange.GetXSDType(ConfigPackageTable."Table ID", ConfigPackageField."Field ID")));
                    TableColumn := WrkShtWriter.CreateTableColumn(
                        ColumnID,
                        TableColumnName,
                        ConfigXMLExchange.GetElementName(ConfigPackageField."Field Caption"));
                    WrkShtHelper.AppendElementToOpenXmlElement(TableColumn, XmlColumnProperties);
                    WrkShtHelper.AppendElementToOpenXmlElement(TableColumns, TableColumn);
                    WrkShtWriter.SetCellValueText(
                      3, OpenXMLManagement.GetXLColumnID(ColumnID), TableColumnName, WrkShtWriter.DefaultCellDecorator);
                    IF NOT ConfigPackageField.Dimension THEN BEGIN
                        FieldRef := RecRef.FIELD(ConfigPackageField."Field ID");
                        OpenXMLManagement.SetCellComment(
                          WrkShtWriter, OpenXMLManagement.GetXLColumnID(ColumnID) + '3', ConfigValidateMgt.AddComment(FieldRef));
                        CreateCommentVmlShapeXml(ColumnID, 3);
                    END;
                END;
                ColumnID += 1;
            UNTIL ConfigPackageField.NEXT = 0;
        END;
        RecRef.CLOSE;
    end;

    local procedure CreateCommentVmlShapeXml(ColId: Integer; RowId: Integer)
    var
        Guid: Guid;
        Anchor: Text;
        CommentShape: Text;
    begin
        Guid := CREATEGUID;

        Anchor := CreateCommentVmlAnchor(ColId, RowId);

        CommentShape := STRSUBSTNO(CommentVmlShapeXmlTxt, Guid, Anchor, RowId - 1, ColId - 1);

        StringBld.Append(CommentShape);
    end;

    local procedure CreateCommentVmlAnchor(ColId: Integer; RowId: Integer): Text
    begin
        EXIT(STRSUBSTNO(VmlShapeAnchorTxt, ColId, RowId - 2, ColId + 2));
    end;

    local procedure AddPackageAndTableInformation(var ConfigPackageTable: Record 8613; var SingleXMLCells: DotNet SingleXmlCells)
    var
        SingleXMLCell: DotNet SingleXmlCell;
        RecRef: RecordRef;
        TableCaptionString: Text;
    begin
        // Add package name
        SingleXMLCell := WrkShtWriter.AddSingleXmlCell(id);
        WrkShtHelper.AppendElementToOpenXmlElement(SingleXMLCells, SingleXMLCell);
        OpenXMLManagement.AddSingleXMLCellProperties(SingleXMLCell, 'A1', '/DataList/' +
          (ConfigXMLExchange.GetElementName(ConfigPackageTable."Table Caption") + 'List') + '/' +
          ConfigXMLExchange.GetElementName(ConfigPackageTable.FIELDNAME("Package Code")), 1, 1);
        WrkShtWriter.SetCellValueText(1, 'A', ConfigPackageTable."Package Code", WrkShtWriter.DefaultCellDecorator);

        // Add Table name
        RecRef.OPEN(ConfigPackageTable."Table ID");
        TableCaptionString := RecRef.CAPTION;
        RecRef.CLOSE;
        WrkShtWriter.SetCellValueText(1, 'B', TableCaptionString, WrkShtWriter.DefaultCellDecorator);

        // Add Table id
        id += 1;
        SingleXMLCell := WrkShtWriter.AddSingleXmlCell(id);
        WrkShtHelper.AppendElementToOpenXmlElement(SingleXMLCells, SingleXMLCell);

        OpenXMLManagement.AddSingleXMLCellProperties(SingleXMLCell, 'C1', '/DataList/' +
          (ConfigXMLExchange.GetElementName(ConfigPackageTable."Table Caption") + 'List') + '/' +
          ConfigXMLExchange.GetElementName(ConfigPackageTable.FIELDNAME("Table ID")), 1, 1);
        WrkShtWriter.SetCellValueText(1, 'C', FORMAT(ConfigPackageTable."Table ID"), WrkShtWriter.DefaultCellDecorator);
    end;

    local procedure BuildDataSetForPackageTable(ExportFromWksht: Boolean; var ConfigPackageTable: Record 8613): Text
    var
        TempSetupDataFileName: Text;
    begin
        TempSetupDataFileName := CreateXMLPackage(FileMgt.ServerTempFileName(''), ExportFromWksht, ConfigPackageTable);
        DataSet := DataSet.DataSet;
        DataSet.ReadXml(TempSetupDataFileName);
        EXIT(TempSetupDataFileName);
    end;

    local procedure AddAndInitializeTableDefinitionPart(var ConfigPackageTable: Record 8613; ExportFromWksht: Boolean; DataTableCounter: Integer; var TableDefinitionPart: DotNet TableDefinitionPart; SkipData: Boolean)
    var
        ConfigPackageField: Record 8616;
        TableColumns: DotNet TableColumns;
        "Table": DotNet Table;
        BooleanValue: DotNet BooleanValue;
        StringValue: DotNet StringValue;
        RowsCount: Integer;
    begin
        TableDefinitionPart := WrkShtWriter.CreateTableDefinitionPart;
        ConfigPackageField.RESET;
        ConfigPackageField.SETRANGE("Package Code", ConfigPackageTable."Package Code");
        ConfigPackageField.SETRANGE("Table ID", ConfigPackageTable."Table ID");
        ConfigPackageField.SETRANGE("Include Field", TRUE);
        IF NOT ExportFromWksht THEN
            ConfigPackageField.SETRANGE(Dimension, FALSE);

        DataTable := DataSet.Tables.Item(DataTableCounter);

        id += 1;
        IF SkipData THEN
            RowsCount := 1
        ELSE
            RowsCount := DataTable.Rows.Count;
        Table := WrkShtWriter.CreateTable(id);
        Table.TotalsRowShown := BooleanValue.BooleanValue(FALSE);
        Table.Reference :=
          StringValue.StringValue(
            'A3:' + OpenXMLManagement.GetXLColumnID(ConfigPackageField.COUNT) + FORMAT(RowsCount + 3));
        Table.Name := StringValue.StringValue('Table' + FORMAT(id));
        Table.DisplayName := StringValue.StringValue('Table' + FORMAT(id));
        OpenXMLManagement.AppendAutoFilter(Table);
        TableColumns := WrkShtWriter.CreateTableColumns(ConfigPackageField.COUNT);

        CreateTableColumnNames(ConfigPackageField, ConfigPackageTable, TableColumns);
        WrkShtHelper.AppendElementToOpenXmlElement(Table, TableColumns);
        OpenXMLManagement.AppendTableStyleInfo(Table);
        TableDefinitionPart.Table := Table;
    end;

    [TryFunction]
    local procedure GetDataTable(TableId: Integer)
    begin
        DataTable := DataSet.Tables.Item(TableId);
    end;

    local procedure InitColumnMapping(WrkShtReader: DotNet WorksheetReader; var TempXMLBuffer: Record 1235 temporary): Boolean
    var
        "Table": DotNet Table;
        TableColumn: DotNet TableColumn;
        Enumerable: DotNet IEnumerable;
        Enumerator: DotNet IEnumerator;
        XmlColumnProperties: DotNet XmlColumnProperties;
        TableStartColumnIndex: Integer;
        Index: Integer;
    begin
        TempXMLBuffer.DELETEALL;
        IF NOT OpenXMLManagement.FindTableDefinition(WrkShtReader, Table) THEN
            EXIT(FALSE);

        TableStartColumnIndex := GetTableStartColumnIndex(Table);
        Index := 0;
        Enumerable := Table.TableColumns;
        Enumerator := Enumerable.GetEnumerator;
        WHILE Enumerator.MoveNext DO BEGIN
            TableColumn := Enumerator.Current;
            XmlColumnProperties := TableColumn.XmlColumnProperties;
            IF NOT ISNULL(XmlColumnProperties) THEN BEGIN
                // identifies column to xsd mapping.
                IF NOT ISNULL(XmlColumnProperties.XPath) THEN
                    InsertXMLBuffer(Index + TableStartColumnIndex, TempXMLBuffer);
            END;
            Index += 1;
        END;

        // RowCount > 2 means sheet has datarow(s)
        EXIT(WrkShtReader.RowCount > 2);
    end;

    local procedure GetTableStartColumnIndex("Table": DotNet Table): Integer
    var
        String: DotNet String;
        Index: Integer;
        Length: Integer;
        ColumnIndex: Integer;
    begin
        // <x:table id="5" ... ref="A3:E6" ...>
        // table.Reference = "A3:E6" (A3 - top left table corner, E6 - bottom right corner)
        // we convert "A" - to column index
        String := Table.Reference.Value;
        Length := String.IndexOf(':');
        String := DELCHR(String.Substring(0, Length), '=', '0123456789');
        Length := String.Length - 1;
        FOR Index := 0 TO Length DO
            ColumnIndex += (String.Chars(Index) - 64) + Index * 26;
        EXIT(ColumnIndex);
    end;

    local procedure InsertXMLBuffer(ColumnIndex: Integer; var TempXMLBuffer: Record 1235 temporary)
    begin
        TempXMLBuffer.INIT;
        TempXMLBuffer."Entry No." := ColumnIndex; // column index in table definition
        TempXMLBuffer."Parent Entry No." := TempXMLBuffer.COUNT; // column index in dataset
        TempXMLBuffer.INSERT;
    end;

    [Scope('Personalization')]
    procedure SetFileOnServer(NewFileOnServer: Boolean)
    begin
        FileOnServer := NewFileOnServer;
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeExportExcel(var ConfigPackageTable: Record 8613)
    begin
    end;

    local procedure GetExcelWorksheetName(Caption: Text; TableID: Text): Text
    var
        WorksheetNameMaxLen: Integer;
    begin
        // maximum Worksheet Name length in Excel
        WorksheetNameMaxLen := 31;
        IF STRLEN(Caption) > WorksheetNameMaxLen THEN
            Caption := COPYSTR(TableID + ' ' + Caption, 1, WorksheetNameMaxLen);
        EXIT(Caption);
    end;

    local procedure IsFileImportedToBLOB(var TempBlob: Record 99008535): Boolean
    var
        IsHandled: Boolean;
    begin
        OnImportExcelFile(TempBlob, IsHandled);
        IF IsHandled THEN
            EXIT(TRUE);
        EXIT(FileMgt.BLOBImportWithFilter(TempBlob, ImportFromExcelMsg, '', FileExtensionFilterTok, ExcelFileExtensionTok) <> '');
    end;

    [IntegrationEvent(false, false)]
    local procedure OnImportExcelFile(var TempBlob: Record 99008535; var IsHandled: Boolean)
    begin
    end;
}

