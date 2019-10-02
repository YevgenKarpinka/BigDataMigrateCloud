codeunit 60000 "Config. XML Exchange Ext."
{

    trigger OnRun()
    begin
    end;

    var
        FileManagement: Codeunit 419;
        ConfigPackageMgt: Codeunit 8611;
        ConfigProgressBar: Codeunit 8615;
        ConfigValidateMgt: Codeunit 8617;
        ConfigMgt: Codeunit 8616;
        ConfigPckgCompressionMgt: Codeunit 8619;
        TypeHelper: Codeunit 10;
        XMLDOMMgt: Codeunit 6224;
        ErrorTypeEnum: Option General,TableRelation;
        Advanced: Boolean;
        CalledFromCode: Boolean;
        PackageAllreadyContainsDataQst: Label 'Package %1 already contains data that will be overwritten by the import. Do you want to continue?', Comment = '%1 - Package name';
        TableContainsRecordsQst: Label 'Table %1 in the package %2 contains %3 records that will be overwritten by the import. Do you want to continue?', Comment = '%1=The ID of the table being imported. %2=The Config Package Code. %3=The number of records in the config package.';
        MissingInExcelFileErr: Label '%1 is missing in the Excel file.', Comment = '%1=The Package Code field caption.';
        ExportPackageTxt: Label 'Exporting package';
        ImportPackageTxt: Label 'Importing package';
        PackageFileNameTxt: Label 'Package%1.rapidstart', Locked = true;
        DownloadTxt: Label 'Download';
        ImportFileTxt: Label 'Import File';
        FileDialogFilterTxt: Label 'RapidStart file (*.rapidstart)|*.rapidstart|All Files (*.*)|*.*', Comment = 'Only translate ''RapidStart Files'' {Split=r"[\|\(]\*\.[^ |)]*[|) ]?"}';
        ExcelMode: Boolean;
        HideDialog: Boolean;
        DataListTxt: Label 'DataList', Locked = true;
        TableDoesNotExistErr: Label 'An error occurred while importing the %1 table. The table does not exist in the database.';
        WrongFileTypeErr: Label 'The specified file could not be imported because it is not a valid RapidStart package file.';
        ExportFromWksht: Boolean;
        RecordProgressTxt: Label 'Import %1 records', Comment = '%1=The name of the table being imported.';
        AddPrefixMode: Boolean;
        WorkingFolder: Text;
        PackageCodesMustMatchErr: Label 'The package code specified on the configuration package must be the same as the package name in the imported package.';

    local procedure AddXMLComment(var PackageXML: DotNet XmlDocument; var Node: DotNet XmlNode; Comment: Text[250])
    var
        CommentNode: DotNet XmlNode;
    begin
        CommentNode := PackageXML.CreateComment(Comment);
        Node.AppendChild(CommentNode);
    end;

    local procedure AddTableAttributes(ConfigPackageTable: Record 8613; var PackageXML: DotNet XmlDocument; var TableNode: DotNet XmlNode)
    var
        FieldNode: DotNet XmlNode;
    begin
        WITH ConfigPackageTable DO BEGIN
            IF "Page ID" > 0 THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME("Page ID")));
                FieldNode.InnerText := FORMAT("Page ID");
                TableNode.AppendChild(FieldNode);
            END;
            IF "Package Processing Order" > 0 THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME("Package Processing Order")));
                FieldNode.InnerText := FORMAT("Package Processing Order");
                TableNode.AppendChild(FieldNode);
            END;
            IF "Processing Order" > 0 THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME("Processing Order")));
                FieldNode.InnerText := FORMAT("Processing Order");
                TableNode.AppendChild(FieldNode);
            END;
            IF "Data Template" <> '' THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME("Data Template")));
                FieldNode.InnerText := FORMAT("Data Template");
                TableNode.AppendChild(FieldNode);
            END;
            IF Comments <> '' THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME(Comments)));
                FieldNode.InnerText := FORMAT(Comments);
                TableNode.AppendChild(FieldNode);
            END;
            IF "Created by User ID" <> '' THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME("Created by User ID")));
                FieldNode.InnerText := FORMAT("Created by User ID");
                TableNode.AppendChild(FieldNode);
            END;
            IF "Skip Table Triggers" THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME("Skip Table Triggers")));
                FieldNode.InnerText := '1';
                TableNode.AppendChild(FieldNode);
            END;
            IF "Parent Table ID" > 0 THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME("Parent Table ID")));
                FieldNode.InnerText := FORMAT("Parent Table ID");
                TableNode.AppendChild(FieldNode);
            END;
            IF "Delete Recs Before Processing" THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME("Delete Recs Before Processing")));
                FieldNode.InnerText := '1';
                TableNode.AppendChild(FieldNode);
            END;
            IF "Dimensions as Columns" THEN BEGIN
                FieldNode := PackageXML.CreateElement(GetElementName(FIELDNAME("Dimensions as Columns")));
                FieldNode.InnerText := '1';
                TableNode.AppendChild(FieldNode);
            END;
        END;
    end;

    local procedure AddFieldAttributes(ConfigPackageField: Record 8616; var FieldNode: DotNet XmlNode)
    begin
        IF ConfigPackageField."Primary Key" THEN
            XMLDOMMgt.AddAttribute(FieldNode, GetElementName(ConfigPackageField.FIELDNAME("Primary Key")), '1');
        IF ConfigPackageField."Validate Field" THEN
            XMLDOMMgt.AddAttribute(FieldNode, GetElementName(ConfigPackageField.FIELDNAME("Validate Field")), '1');
        IF ConfigPackageField."Create Missing Codes" THEN
            XMLDOMMgt.AddAttribute(FieldNode, GetElementName(ConfigPackageField.FIELDNAME("Create Missing Codes")), '1');
        IF ConfigPackageField."Processing Order" <> 0 THEN
            XMLDOMMgt.AddAttribute(
              FieldNode, GetElementName(ConfigPackageField.FIELDNAME("Processing Order")), FORMAT(ConfigPackageField."Processing Order"));
    end;

    local procedure AddDimensionFields(var ConfigPackageField: Record 8616; var RecRef: RecordRef; var PackageXML: DotNet XmlDocument; var RecordNode: DotNet XmlNode; var FieldNode: DotNet XmlNode; ExportValue: Boolean)
    var
        DimCode: Code[20];
    begin
        ConfigPackageField.SETRANGE(Dimension, TRUE);
        IF ConfigPackageField.FINDSET THEN
            REPEAT
                FieldNode :=
                  PackageXML.CreateElement(
                    GetElementName(ConfigValidateMgt.CheckName(ConfigPackageField."Field Name")));
                IF ExportValue THEN BEGIN
                    DimCode := COPYSTR(ConfigPackageField."Field Name", 1, 20);
                    FieldNode.InnerText := GetDimValueFromTable(RecRef, DimCode);
                    RecordNode.AppendChild(FieldNode);
                END ELSE BEGIN
                    FieldNode.InnerText := '';
                    RecordNode.AppendChild(FieldNode);
                END;
            UNTIL ConfigPackageField.NEXT = 0;
    end;

    [Scope('Personalization')]
    procedure ApplyPackageFilter(ConfigPackageTable: Record 8613; var RecRef: RecordRef)
    var
        ConfigPackageFilter: Record 8626;
        FieldRef: FieldRef;
    begin
        ConfigPackageFilter.SETRANGE("Package Code", ConfigPackageTable."Package Code");
        ConfigPackageFilter.SETRANGE("Table ID", ConfigPackageTable."Table ID");
        ConfigPackageFilter.SETRANGE("Processing Rule No.", 0);
        IF ConfigPackageFilter.FINDSET THEN
            REPEAT
                IF ConfigPackageFilter."Field Filter" <> '' THEN BEGIN
                    FieldRef := RecRef.FIELD(ConfigPackageFilter."Field ID");
                    FieldRef.SETFILTER(STRSUBSTNO('%1', ConfigPackageFilter."Field Filter"));
                END;
            UNTIL ConfigPackageFilter.NEXT = 0;
    end;

    local procedure CreateRecordNodes(var PackageXML: DotNet XmlDocument; ConfigPackageTable: Record 8613)
    var
        "Field": Record 2000000041;
        ConfigPackageField: Record 8616;
        ConfigPackage: Record 8623;
        DocumentElement: DotNet XmlNode;
        FieldNode: DotNet XmlNode;
        RecordNode: DotNet XmlNode;
        TableNode: DotNet XmlNode;
        TableIDNode: DotNet XmlNode;
        PackageCodeNode: DotNet XmlNode;
        RecRef: RecordRef;
        FieldRef: FieldRef;
        ExportMetadata: Boolean;
    begin
        ConfigPackageTable.TESTFIELD("Package Code");
        ConfigPackageTable.TESTFIELD("Table ID");
        ConfigPackage.GET(ConfigPackageTable."Package Code");
        ExcludeRemovedFields(ConfigPackageTable);
        DocumentElement := PackageXML.DocumentElement;
        TableNode := PackageXML.CreateElement(GetElementName(ConfigPackageTable."Table Name" + 'List'));
        DocumentElement.AppendChild(TableNode);

        TableIDNode := PackageXML.CreateElement(GetElementName(ConfigPackageTable.FIELDNAME("Table ID")));
        TableIDNode.InnerText := FORMAT(ConfigPackageTable."Table ID");
        TableNode.AppendChild(TableIDNode);

        IF ExcelMode THEN BEGIN
            PackageCodeNode := PackageXML.CreateElement(GetElementName(ConfigPackageTable.FIELDNAME("Package Code")));
            PackageCodeNode.InnerText := FORMAT(ConfigPackageTable."Package Code");
            TableNode.AppendChild(PackageCodeNode);
        END ELSE
            AddTableAttributes(ConfigPackageTable, PackageXML, TableNode);

        ExportMetadata := TRUE;
        RecRef.OPEN(ConfigPackageTable."Table ID");
        ApplyPackageFilter(ConfigPackageTable, RecRef);
        IF RecRef.FINDSET THEN
            REPEAT
                RecordNode := PackageXML.CreateElement(GetTableElementName(ConfigPackageTable."Table Name"));
                TableNode.AppendChild(RecordNode);

                ConfigPackageField.SETRANGE("Package Code", ConfigPackageTable."Package Code");
                ConfigPackageField.SETRANGE("Table ID", ConfigPackageTable."Table ID");
                ConfigPackageField.SETRANGE("Include Field", TRUE);
                ConfigPackageField.SETRANGE(Dimension, FALSE);
                ConfigPackageField.SETCURRENTKEY("Package Code", "Table ID", "Processing Order");
                IF ConfigPackageField.FINDSET THEN
                    REPEAT
                        FieldRef := RecRef.FIELD(ConfigPackageField."Field ID");
                        IF TypeHelper.GetField(RecRef.NUMBER, FieldRef.NUMBER, Field) THEN BEGIN
                            FieldNode :=
                              PackageXML.CreateElement(GetFieldElementName(ConfigValidateMgt.CheckName(FieldRef.NAME)));
                            FieldNode.InnerText := FormatFieldValue(FieldRef, ConfigPackage);
                            IF Advanced AND ConfigPackageField."Localize Field" THEN
                                AddXMLComment(PackageXML, FieldNode, '_locComment_text="{MaxLength=' + FORMAT(Field.Len) + '}"');
                            RecordNode.AppendChild(FieldNode); // must be after AddXMLComment and before AddAttribute.
                            IF NOT ExcelMode AND ExportMetadata THEN
                                AddFieldAttributes(ConfigPackageField, FieldNode);
                            IF Advanced THEN
                                IF ConfigPackageField."Localize Field" THEN
                                    XMLDOMMgt.AddAttribute(FieldNode, '_loc', 'locData')
                                ELSE
                                    XMLDOMMgt.AddAttribute(FieldNode, '_loc', 'locNone');
                        END;
                    UNTIL ConfigPackageField.NEXT = 0;

                IF ConfigPackageTable."Dimensions as Columns" AND ExcelMode AND ExportFromWksht THEN
                    AddDimensionFields(ConfigPackageField, RecRef, PackageXML, RecordNode, FieldNode, TRUE);
                ExportMetadata := FALSE;
            UNTIL RecRef.NEXT = 0
        ELSE BEGIN
            RecordNode := PackageXML.CreateElement(GetTableElementName(ConfigPackageTable."Table Name"));
            TableNode.AppendChild(RecordNode);

            ConfigPackageField.SETRANGE("Package Code", ConfigPackageTable."Package Code");
            ConfigPackageField.SETRANGE("Table ID", ConfigPackageTable."Table ID");
            ConfigPackageField.SETRANGE("Include Field", TRUE);
            ConfigPackageField.SETRANGE(Dimension, FALSE);
            IF ConfigPackageField.FINDSET THEN
                REPEAT
                    FieldRef := RecRef.FIELD(ConfigPackageField."Field ID");
                    FieldNode :=
                      PackageXML.CreateElement(GetFieldElementName(ConfigValidateMgt.CheckName(FieldRef.NAME)));
                    FieldNode.InnerText := '';
                    RecordNode.AppendChild(FieldNode);
                    IF NOT ExcelMode THEN
                        AddFieldAttributes(ConfigPackageField, FieldNode);
                UNTIL ConfigPackageField.NEXT = 0;

            IF ConfigPackageTable."Dimensions as Columns" AND ExcelMode AND ExportFromWksht THEN
                AddDimensionFields(ConfigPackageField, RecRef, PackageXML, RecordNode, FieldNode, FALSE);
        END;
    end;

    [Scope('Internal')]
    procedure ExportPackage(ConfigPackage: Record 8623)
    var
        ConfigPackageTable: Record 8613;
    begin
        WITH ConfigPackage DO BEGIN
            TESTFIELD(Code);
            TESTFIELD("Package Name");
            ConfigPackageTable.SETRANGE("Package Code", Code);
            ExportPackageXML(ConfigPackageTable, '');
        END;
    end;

    [Scope('Internal')]
    procedure ExportPackageXML(var ConfigPackageTable: Record 8613; XMLDataFile: Text): Boolean
    var
        ConfigPackage: Record 8623;
        PackageXML: DotNet XmlDocument;
        FileFilter: Text;
        ToFile: Text[50];
        CompressedFileName: Text;
    begin
        ConfigPackageTable.FINDFIRST;
        ConfigPackage.GET(ConfigPackageTable."Package Code");
        ConfigPackage.TESTFIELD(Code);
        ConfigPackage.TESTFIELD("Package Name");
        IF NOT ConfigPackage."Exclude Config. Tables" AND NOT ExcelMode THEN
            ConfigPackageMgt.AddConfigTables(ConfigPackage.Code);

        IF NOT CalledFromCode THEN
            XMLDataFile := FileManagement.ServerTempFileName('');
        FileFilter := GetFileDialogFilter;
        IF ToFile = '' THEN
            ToFile := STRSUBSTNO(PackageFileNameTxt, ConfigPackage.Code);

        SetWorkingFolder(FileManagement.GetDirectoryName(XMLDataFile));
        PackageXML := PackageXML.XmlDocument;
        ExportPackageXMLDocument(PackageXML, ConfigPackageTable, ConfigPackage, Advanced);

        PackageXML.Save(XMLDataFile);

        IF NOT CalledFromCode THEN BEGIN
            CompressedFileName := FileManagement.ServerTempFileName('');
            ConfigPckgCompressionMgt.ServersideCompress(XMLDataFile, CompressedFileName);

            FileManagement.DownloadHandler(CompressedFileName, DownloadTxt, '', FileFilter, ToFile);
        END;

        EXIT(TRUE);
    end;

    [Scope('Internal')]
    procedure ExportPackageXMLDocument(var PackageXML: DotNet XmlDocument; var ConfigPackageTable: Record 8613; ConfigPackage: Record 8623; Advanced: Boolean)
    var
        DocumentElement: DotNet XmlElement;
        LocXML: Text[1024];
    begin
        ConfigPackage.TESTFIELD(Code);
        ConfigPackage.TESTFIELD("Package Name");

        IF Advanced THEN
            LocXML := '<_locDefinition><_locDefault _loc="locNone"/></_locDefinition>';
        XMLDOMMgt.LoadXMLDocumentFromText(
          STRSUBSTNO(
            '<?xml version="1.0" encoding="UTF-16" standalone="yes"?><%1>%2</%1>',
            GetPackageTag,
            LocXML),
          PackageXML);

        CleanUpConfigPackageData(ConfigPackage);

        IF NOT ExcelMode THEN BEGIN
            InitializeMediaTempFolder;
            DocumentElement := PackageXML.DocumentElement;
            IF ConfigPackage."Exclude Config. Tables" THEN
                XMLDOMMgt.AddAttribute(DocumentElement, GetElementName(ConfigPackage.FIELDNAME("Exclude Config. Tables")), '1');
            IF ConfigPackage."Processing Order" > 0 THEN
                XMLDOMMgt.AddAttribute(
                  DocumentElement, GetElementName(ConfigPackage.FIELDNAME("Processing Order")), FORMAT(ConfigPackage."Processing Order"));
            IF ConfigPackage."Language ID" > 0 THEN
                XMLDOMMgt.AddAttribute(
                  DocumentElement, GetElementName(ConfigPackage.FIELDNAME("Language ID")), FORMAT(ConfigPackage."Language ID"));
            XMLDOMMgt.AddAttribute(
              DocumentElement, GetElementName(ConfigPackage.FIELDNAME("Product Version")), ConfigPackage."Product Version");
            XMLDOMMgt.AddAttribute(DocumentElement, GetElementName(ConfigPackage.FIELDNAME("Package Name")), ConfigPackage."Package Name");
            XMLDOMMgt.AddAttribute(DocumentElement, GetElementName(ConfigPackage.FIELDNAME(Code)), ConfigPackage.Code);
        END;

        IF NOT HideDialog THEN
            ConfigProgressBar.Init(ConfigPackageTable.COUNT, 1, ExportPackageTxt);
        ConfigPackageTable.SETAUTOCALCFIELDS("Table Name");
        IF ConfigPackageTable.FINDSET THEN
            REPEAT
                IF NOT HideDialog THEN
                    ConfigProgressBar.Update(ConfigPackageTable."Table Name");

                ExportConfigTableToXML(ConfigPackageTable, PackageXML);
            UNTIL ConfigPackageTable.NEXT = 0;

        IF NOT ExcelMode THEN BEGIN
            UpdateConfigPackageMediaSet(ConfigPackage);
            ExportConfigPackageMediaSetToXML(PackageXML, ConfigPackage);
        END;

        IF NOT HideDialog THEN
            ConfigProgressBar.Close;
    end;

    local procedure ExportConfigTableToXML(var ConfigPackageTable: Record 8613; var PackageXML: DotNet XmlDocument)
    begin
        CreateRecordNodes(PackageXML, ConfigPackageTable);
        ConfigPackageTable."Exported Date and Time" := CREATEDATETIME(TODAY, TIME);
        ConfigPackageTable.MODIFY;
    end;

    [Scope('Internal')]
    procedure ImportPackageXMLFromClient(): Boolean
    var
        ServerFileName: Text;
        DecompressedFileName: Text;
    begin
        ServerFileName := FileManagement.ServerTempFileName('.xml');
        IF UploadXMLPackage(ServerFileName) THEN BEGIN
            DecompressedFileName := DecompressPackage(ServerFileName);

            EXIT(ImportPackageXML(DecompressedFileName));
        END;

        EXIT(FALSE);
    end;

    [Scope('Internal')]
    procedure ImportPackageXML(XMLDataFile: Text): Boolean
    var
        PackageXML: DotNet XmlDocument;
    begin
        XMLDOMMgt.LoadXMLDocumentFromFile(XMLDataFile, PackageXML);

        EXIT(ImportPackageXMLDocument(PackageXML, ''));
    end;

    [Scope('Personalization')]
    procedure ImportPackageXMLFromStream(InStream: InStream): Boolean
    var
        PackageXML: DotNet XmlDocument;
    begin
        XMLDOMMgt.LoadXMLDocumentFromInStream(InStream, PackageXML);

        EXIT(ImportPackageXMLDocument(PackageXML, ''));
    end;

    [Scope('Personalization')]
    procedure ImportPackageXMLWithCodeFromStream(InStream: InStream; PackageCode: Code[20]): Boolean
    var
        PackageXML: DotNet XmlDocument;
    begin
        XMLDOMMgt.LoadXMLDocumentFromInStream(InStream, PackageXML);
        IF PackageCode <> '' THEN BEGIN
            IF PackageCode <> GetPackageCode(PackageXML) THEN
                ERROR(PackageCodesMustMatchErr);
        END;

        EXIT(ImportPackageXMLDocument(PackageXML, PackageCode));
    end;

    [Scope('Internal')]
    procedure ImportPackageXMLDocument(PackageXML: DotNet XmlDocument; PackageCode: Code[20]): Boolean
    var
        ConfigPackage: Record 8623;
        ConfigPackageRecord: Record 8614;
        ConfigPackageData: Record 8615;
        TempBlob: Record 99008535;
        ParallelSessionManagement: Codeunit 490;
        DocumentElement: DotNet XmlElement;
        TableNodes: DotNet XmlNodeList;
        TableNode: DotNet XmlNode;
        Value: Text;
        TableID: Integer;
        NodeCount: Integer;
        Confirmed: Boolean;
        NoOfChildNodes: Integer;
    begin
        DocumentElement := PackageXML.DocumentElement;

        IF NOT ExcelMode THEN BEGIN
            IF PackageCode = '' THEN BEGIN
                PackageCode := GetPackageCode(PackageXML);
                IF ConfigPackage.GET(PackageCode) THEN BEGIN
                    ConfigPackage.CALCFIELDS("No. of Records");
                    Confirmed := TRUE;
                    IF NOT HideDialog THEN
                        IF ConfigPackage."No. of Records" > 0 THEN
                            IF NOT CONFIRM(PackageAllreadyContainsDataQst, TRUE, PackageCode) THEN
                                Confirmed := FALSE;
                    IF NOT Confirmed THEN
                        EXIT(FALSE);
                    ConfigPackage.DELETE(TRUE);
                    COMMIT;
                END;
                ConfigPackage.INIT;
                ConfigPackage.Code := PackageCode;
                ConfigPackage.INSERT;
            END ELSE
                ConfigPackage.GET(PackageCode);

            ConfigPackage."Package Name" :=
              COPYSTR(
                GetAttribute(GetElementName(ConfigPackage.FIELDNAME("Package Name")), DocumentElement), 1,
                MAXSTRLEN(ConfigPackage."Package Name"));
            Value := GetAttribute(GetElementName(ConfigPackage.FIELDNAME("Language ID")), DocumentElement);
            IF Value <> '' THEN
                EVALUATE(ConfigPackage."Language ID", Value);
            ConfigPackage."Product Version" :=
              COPYSTR(
                GetAttribute(GetElementName(ConfigPackage.FIELDNAME("Product Version")), DocumentElement), 1,
                MAXSTRLEN(ConfigPackage."Product Version"));
            Value := GetAttribute(GetElementName(ConfigPackage.FIELDNAME("Processing Order")), DocumentElement);
            IF Value <> '' THEN
                EVALUATE(ConfigPackage."Processing Order", Value);
            Value := GetAttribute(GetElementName(ConfigPackage.FIELDNAME("Exclude Config. Tables")), DocumentElement);
            IF Value <> '' THEN
                EVALUATE(ConfigPackage."Exclude Config. Tables", Value);
            ConfigPackage.MODIFY;
        END;
        COMMIT; // to enable background processes to reference the ConfigPackage

        TableNodes := DocumentElement.ChildNodes;
        IF NOT HideDialog THEN
            ConfigProgressBar.Init(TableNodes.Count, 1, ImportPackageTxt);
        FOR NodeCount := 0 TO (TableNodes.Count - 1) DO BEGIN
            TableNode := TableNodes.Item(NodeCount);
            IF EVALUATE(TableID, FORMAT(TableNode.FirstChild.InnerText)) THEN BEGIN
                NoOfChildNodes := TableNode.ChildNodes.Count;
                IF (NoOfChildNodes < 5) OR ExcelMode THEN
                    ImportTableFromXMLNode(TableNode, PackageCode)
                ELSE BEGIN
                    // Send to background
                    TempBlob.WriteAsText('<doc>' + TableNode.OuterXml + '</doc>', TEXTENCODING::UTF8);
                    ParallelSessionManagement.NewSessionRunCodeunitWithBlob(
                      CODEUNIT::"Config. Import Table in Backgr", PackageCode, 0, TempBlob);
                END;
                IF ExcelMode THEN
                    CASE TRUE OF // Dimensions
                        ConfigMgt.IsDefaultDimTable(TableID):
                            BEGIN
                                ConfigPackageRecord.SETRANGE("Package Code", PackageCode);
                                ConfigPackageRecord.SETRANGE("Table ID", TableID);
                                IF ConfigPackageRecord.FINDSET THEN
                                    REPEAT
                                        ConfigPackageData.GET(
                                          ConfigPackageRecord."Package Code", ConfigPackageRecord."Table ID", ConfigPackageRecord."No.", 1);
                                        ConfigPackageMgt.UpdateDefaultDimValues(ConfigPackageRecord, COPYSTR(ConfigPackageData.Value, 1, 20));
                                    UNTIL ConfigPackageRecord.NEXT = 0;
                            END;
                        ConfigMgt.IsDimSetIDTable(TableID):
                            BEGIN
                                ConfigPackageRecord.SETRANGE("Package Code", PackageCode);
                                ConfigPackageRecord.SETRANGE("Table ID", TableID);
                                IF ConfigPackageRecord.FINDSET THEN
                                    REPEAT
                                        ConfigPackageMgt.HandlePackageDataDimSetIDForRecord(ConfigPackageRecord);
                                    UNTIL ConfigPackageRecord.NEXT = 0;
                            END;
                    END;
            END;
        END;
        IF NOT HideDialog THEN
            ConfigProgressBar.Close;
        IF NOT ExcelMode THEN
            ParallelSessionManagement.WaitForAllToFinish(0);

        ConfigPackageMgt.UpdateConfigLinePackageData(ConfigPackage.Code);

        // autoapply configuration lines
        ConfigPackageMgt.ApplyConfigTables(ConfigPackage);

        EXIT(TRUE);
    end;

    procedure ImportTableFromXMLNode(var TableNode: DotNet XmlNode;

    var
        PackageCode: Code[20])
    var
        ConfigPackageRecord: Record 8614;
        ConfigPackageTable: Record 8613;
        TableID: Integer;
    begin
        IF EVALUATE(TableID, FORMAT(TableNode.FirstChild.InnerText)) THEN BEGIN
            FillPackageMetadataFromXML(PackageCode, TableID, TableNode);
            IF NOT TableObjectExists(TableID) THEN BEGIN
                ConfigPackageMgt.InsertPackageTableWithoutValidation(ConfigPackageTable, PackageCode, TableID);
                ConfigPackageMgt.InitPackageRecord(ConfigPackageRecord, PackageCode, TableID);
                ConfigPackageMgt.RecordError(ConfigPackageRecord, 0, COPYSTR(STRSUBSTNO(TableDoesNotExistErr, TableID), 1, 250));
            END ELSE
                IF PackageDataExistsInXML(PackageCode, TableID, TableNode) THEN
                    FillPackageDataFromXML(PackageCode, TableID, TableNode);
        END;
    end;

    local procedure PackageDataExistsInXML(PackageCode: Code[20]; TableID: Integer; var TableNode: DotNet XmlNode): Boolean
    var
        ConfigPackageTable: Record 8613;
        ConfigPackageField: Record 8616;
        RecRef: RecordRef;
        RecordNodes: DotNet XmlNodeList;
        RecordNode: DotNet XmlNode;
        I: Integer;
    begin
        IF NOT ConfigPackageTable.GET(PackageCode, TableID) THEN
            EXIT(FALSE);

        ConfigPackageTable.CALCFIELDS("Table Name");
        RecordNodes := TableNode.SelectNodes(GetElementName(ConfigPackageTable."Table Name"));

        IF RecordNodes.Count = 0 THEN
            EXIT(FALSE);

        FOR I := 0 TO RecordNodes.Count - 1 DO BEGIN
            RecordNode := RecordNodes.Item(I);
            IF RecordNode.HasChildNodes THEN BEGIN
                RecRef.OPEN(ConfigPackageTable."Table ID");
                ConfigPackageField.SETRANGE("Package Code", ConfigPackageTable."Package Code");
                ConfigPackageField.SETRANGE("Table ID", ConfigPackageTable."Table ID");
                IF ConfigPackageField.FINDSET THEN
                    REPEAT
                        IF ConfigPackageField."Include Field" AND FieldNodeExists(RecordNode, GetElementName(ConfigPackageField."Field Name")) THEN
                            IF GetNodeValue(RecordNode, GetElementName(ConfigPackageField."Field Name")) <> '' THEN
                                EXIT(TRUE);
                    UNTIL ConfigPackageField.NEXT = 0;
                RecRef.CLOSE;
            END;
        END;

        EXIT(FALSE);
    end;

    local procedure FillPackageMetadataFromXML(var PackageCode: Code[20]; TableID: Integer; var TableNode: DotNet XmlNode)
    var
        ConfigPackage: Record 8623;
        ConfigPackageTable: Record 8613;
        ConfigPackageField: Record 8616;
        "Field": Record 2000000041;
        ConfigMgt: Codeunit 8616;
        RecordNodes: DotNet XmlNodeList;
        RecordNode: DotNet XmlNode;
        FieldNode: DotNet XmlNode;
        Value: Text;
    begin
        IF (TableID > 0) AND (NOT ConfigPackageTable.GET(PackageCode, TableID)) THEN BEGIN
            IF NOT ExcelMode THEN BEGIN
                ConfigPackageTable.INIT;
                ConfigPackageTable."Package Code" := PackageCode;
                ConfigPackageTable."Table ID" := TableID;
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Page ID")));
                IF Value <> '' THEN
                    EVALUATE(ConfigPackageTable."Page ID", Value);
                IF ConfigPackageTable."Page ID" = 0 THEN
                    ConfigPackageTable."Page ID" := ConfigMgt.FindPage(TableID);
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Package Processing Order")));
                IF Value <> '' THEN
                    EVALUATE(ConfigPackageTable."Package Processing Order", Value);
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Processing Order")));
                IF Value <> '' THEN
                    EVALUATE(ConfigPackageTable."Processing Order", Value);
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Dimensions as Columns")));
                IF Value <> '' THEN
                    EVALUATE(ConfigPackageTable."Dimensions as Columns", Value);
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Skip Table Triggers")));
                IF Value <> '' THEN
                    EVALUATE(ConfigPackageTable."Skip Table Triggers", Value);
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Parent Table ID")));
                IF Value <> '' THEN
                    EVALUATE(ConfigPackageTable."Parent Table ID", Value);
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Delete Recs Before Processing")));
                IF Value <> '' THEN
                    EVALUATE(ConfigPackageTable."Delete Recs Before Processing", Value);
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Created by User ID")));
                IF Value <> '' THEN
                    EVALUATE(ConfigPackageTable."Created by User ID", COPYSTR(Value, 1, 50));
                ConfigPackageTable."Data Template" :=
                  COPYSTR(
                    GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Data Template"))), 1,
                    MAXSTRLEN(ConfigPackageTable."Data Template"));
                ConfigPackageTable.Comments :=
                  COPYSTR(
                    GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME(Comments))),
                    1, MAXSTRLEN(ConfigPackageTable.Comments));
                ConfigPackageTable."Imported Date and Time" := CREATEDATETIME(TODAY, TIME);
                ConfigPackageTable."Imported by User ID" := USERID;
                ConfigPackageTable.INSERT(TRUE);
                ConfigPackageField.SETRANGE("Package Code", ConfigPackageTable."Package Code");
                ConfigPackageField.SETRANGE("Table ID", ConfigPackageTable."Table ID");
                ConfigPackageMgt.SelectAllPackageFields(ConfigPackageField, FALSE);
            END ELSE BEGIN // Excel import
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Package Code")));
                IF Value <> '' THEN
                    ConfigPackageTable."Package Code" := COPYSTR(Value, 1, MAXSTRLEN(ConfigPackageTable."Package Code"))
                ELSE
                    ERROR(MissingInExcelFileErr, ConfigPackageTable.FIELDCAPTION("Package Code"));
                Value := GetNodeValue(TableNode, GetElementName(ConfigPackageTable.FIELDNAME("Table ID")));
                IF Value <> '' THEN
                    EVALUATE(ConfigPackageTable."Table ID", Value)
                ELSE
                    ERROR(MissingInExcelFileErr, ConfigPackageTable.FIELDCAPTION("Table ID"));
                IF NOT ConfigPackageTable.FIND THEN BEGIN
                    IF NOT ConfigPackage.GET(ConfigPackageTable."Package Code") THEN BEGIN
                        ConfigPackage.INIT;
                        ConfigPackage.VALIDATE(Code, ConfigPackageTable."Package Code");
                        ConfigPackage.INSERT(TRUE);
                    END;
                    ConfigPackageTable.INIT;
                    ConfigPackageTable.INSERT(TRUE);
                END;
                PackageCode := ConfigPackageTable."Package Code";
            END;

            ConfigPackageTable.CALCFIELDS("Table Name");
            IF ConfigPackageTable."Table Name" <> '' THEN BEGIN
                RecordNodes := TableNode.SelectNodes(GetElementName(ConfigPackageTable."Table Name"));
                IF RecordNodes.Count > 0 THEN BEGIN
                    RecordNode := RecordNodes.Item(0);
                    IF RecordNode.HasChildNodes THEN BEGIN
                        ConfigPackageMgt.SetFieldFilter(Field, TableID, 0);
                        IF Field.FINDSET THEN
                            REPEAT
                                IF FieldNodeExists(RecordNode, GetElementName(Field.FieldName)) THEN BEGIN
                                    ConfigPackageField.GET(PackageCode, TableID, Field."No.");
                                    ConfigPackageField."Primary Key" := ConfigValidateMgt.IsKeyField(TableID, Field."No.");
                                    ConfigPackageField."Include Field" := TRUE;
                                    FieldNode := RecordNode.SelectSingleNode(GetElementName(Field.FieldName));
                                    IF NOT ISNULL(FieldNode) AND NOT ExcelMode THEN BEGIN
                                        Value := GetAttribute(GetElementName(ConfigPackageField.FIELDNAME("Primary Key")), FieldNode);
                                        ConfigPackageField."Primary Key" := Value = '1';
                                        Value := GetAttribute(GetElementName(ConfigPackageField.FIELDNAME("Validate Field")), FieldNode);
                                        ConfigPackageField."Validate Field" := (Value = '1') AND
                                          NOT ConfigPackageMgt.ValidateException(TableID, Field."No.");
                                        Value := GetAttribute(GetElementName(ConfigPackageField.FIELDNAME("Create Missing Codes")), FieldNode);
                                        ConfigPackageField."Create Missing Codes" := (Value = '1') AND
                                          NOT ConfigPackageMgt.ValidateException(TableID, Field."No.");
                                        Value := GetAttribute(GetElementName(ConfigPackageField.FIELDNAME("Processing Order")), FieldNode);
                                        IF Value <> '' THEN
                                            EVALUATE(ConfigPackageField."Processing Order", Value);
                                    END;
                                    ConfigPackageField.MODIFY;
                                END;
                            UNTIL Field.NEXT = 0;
                    END;
                END;
            END;
        END;
    end;

    local procedure FillPackageDataFromXML(PackageCode: Code[20]; TableID: Integer; var TableNode: DotNet XmlNode)
    var
        ConfigPackageTable: Record 8613;
        ConfigPackageData: Record 8615;
        ConfigPackageRecord: Record 8614;
        ConfigPackageField: Record 8616;
        TempConfigPackageField: Record 8616 temporary;
        ConfigProgressBarRecord: Codeunit 8615;
        RecRef: RecordRef;
        FieldRef: FieldRef;
        RecordNodes: DotNet XmlNodeList;
        RecordNode: DotNet XmlNode;
        NodeCount: Integer;
        RecordCount: Integer;
        StepCount: Integer;
        ErrorText: Text[250];
    begin
        IF ConfigPackageTable.GET(PackageCode, TableID) THEN BEGIN
            ExcludeRemovedFields(ConfigPackageTable);
            IF ExcelMode THEN BEGIN
                ConfigPackageTable.CALCFIELDS("No. of Package Records");
                IF ConfigPackageTable."No. of Package Records" > 0 THEN
                    IF CONFIRM(TableContainsRecordsQst, TRUE, TableID, PackageCode, ConfigPackageTable."No. of Package Records") THEN
                        ConfigPackageTable.DeletePackageData
                    ELSE
                        EXIT;
            END;
            ConfigPackageTable.CALCFIELDS("Table Name");
            IF NOT HideDialog THEN
                ConfigProgressBar.Update(ConfigPackageTable."Table Name");
            RecordNodes := TableNode.SelectNodes(GetElementName(ConfigPackageTable."Table Name"));
            RecordCount := RecordNodes.Count;

            IF NOT HideDialog AND (RecordCount > 1000) THEN BEGIN
                StepCount := ROUND(RecordCount / 100, 1);
                ConfigProgressBarRecord.Init(RecordCount, StepCount,
                  STRSUBSTNO(RecordProgressTxt, ConfigPackageTable."Table Name"));
            END;

            ConfigPackageField.SETRANGE("Package Code", ConfigPackageTable."Package Code");
            ConfigPackageField.SETRANGE("Table ID", ConfigPackageTable."Table ID");
            ConfigPackageField.SETRANGE("Include Field", TRUE);
            IF ConfigPackageField.FINDSET THEN
                REPEAT
                    TempConfigPackageField := ConfigPackageField;
                    TempConfigPackageField.INSERT;
                UNTIL ConfigPackageField.NEXT = 0;

            FOR NodeCount := 0 TO RecordCount - 1 DO BEGIN
                RecordNode := RecordNodes.Item(NodeCount);
                IF RecordNode.HasChildNodes THEN BEGIN
                    ConfigPackageMgt.InitPackageRecord(ConfigPackageRecord, PackageCode, ConfigPackageTable."Table ID");

                    RecRef.CLOSE;
                    RecRef.OPEN(ConfigPackageTable."Table ID");
                    IF TempConfigPackageField.FINDSET THEN
                        REPEAT
                            ConfigPackageData.INIT;
                            ConfigPackageData."Package Code" := TempConfigPackageField."Package Code";
                            ConfigPackageData."Table ID" := TempConfigPackageField."Table ID";
                            ConfigPackageData."Field ID" := TempConfigPackageField."Field ID";
                            ConfigPackageData."No." := ConfigPackageRecord."No.";
                            IF FieldNodeExists(RecordNode, GetElementName(TempConfigPackageField."Field Name")) OR
                               TempConfigPackageField.Dimension
                            THEN
                                GetConfigPackageDataValue(ConfigPackageData, RecordNode, GetElementName(TempConfigPackageField."Field Name"));
                            ConfigPackageData.INSERT;

                            IF NOT TempConfigPackageField.Dimension THEN BEGIN
                                FieldRef := RecRef.FIELD(ConfigPackageData."Field ID");
                                IF ConfigPackageData.Value <> '' THEN BEGIN
                                    ErrorText := ConfigValidateMgt.EvaluateValue(FieldRef, ConfigPackageData.Value, NOT ExcelMode);
                                    IF ErrorText <> '' THEN
                                        ConfigPackageMgt.FieldError(ConfigPackageData, ErrorText, ErrorTypeEnum::General)
                                    ELSE
                                        ConfigPackageData.Value := FORMAT(FieldRef.VALUE);

                                    ConfigPackageData.MODIFY;
                                END;
                            END;
                        UNTIL TempConfigPackageField.NEXT = 0;
                    ConfigPackageTable."Imported Date and Time" := CURRENTDATETIME;
                    ConfigPackageTable."Imported by User ID" := USERID;
                    ConfigPackageTable.MODIFY;
                    IF NOT HideDialog AND (RecordCount > 1000) THEN
                        ConfigProgressBarRecord.Update(
                          STRSUBSTNO('Records: %1 of %2', ConfigPackageRecord."No.", RecordCount));
                END;
            END;
            IF NOT HideDialog AND (RecordCount > 1000) THEN
                ConfigProgressBarRecord.Close;
        END;
    end;

    local procedure ExcludeRemovedFields(ConfigPackageTable: Record 8613)
    var
        "Field": Record 2000000041;
        ConfigPackageField: Record 8616;
    begin
        Field.SETRANGE(TableNo, ConfigPackageTable."Table ID");
        Field.SETRANGE(ObsoleteState, Field.ObsoleteState::Removed);
        IF Field.FINDSET THEN
            REPEAT
                IF ConfigPackageField.GET(ConfigPackageTable."Package Code", Field.TableNo, Field."No.") THEN BEGIN
                    ConfigPackageField.VALIDATE("Include Field", FALSE);
                    ConfigPackageField.MODIFY;
                END;
            UNTIL Field.NEXT = 0;
    end;

    local procedure FieldNodeExists(var RecordNode: DotNet XmlNode; FieldNodeName: Text[250]): Boolean
    var
        FieldNode: DotNet XmlNode;
    begin
        FieldNode := RecordNode.SelectSingleNode(FieldNodeName);

        IF NOT ISNULL(FieldNode) THEN
            EXIT(TRUE);
    end;

    local procedure FormatFieldValue(var FieldRef: FieldRef; ConfigPackage: Record 8623) InnerText: Text
    var
        TypeHelper: Codeunit 10;
        Date: Date;
    begin
        IF NOT (((FORMAT(FieldRef.TYPE) = 'Integer') OR (FORMAT(FieldRef.TYPE) = 'BLOB')) AND
                (FieldRef.RELATION <> 0) AND (FORMAT(FieldRef.VALUE) = '0'))
        THEN
            InnerText := FORMAT(FieldRef.VALUE, 0, ConfigValidateMgt.XMLFormat);

        IF NOT ExcelMode THEN BEGIN
            IF (FORMAT(FieldRef.TYPE) = 'Boolean') OR (FORMAT(FieldRef.TYPE) = 'Option') THEN
                InnerText := FORMAT(FieldRef.VALUE, 0, 2);
            IF (FORMAT(FieldRef.TYPE) = 'DateFormula') AND (FORMAT(FieldRef.VALUE) <> '') THEN
                InnerText := '<' + FORMAT(FieldRef.VALUE, 0, ConfigValidateMgt.XMLFormat) + '>';
            IF FORMAT(FieldRef.TYPE) = 'BLOB' THEN
                InnerText := ConvertBLOBToBase64String(FieldRef);
            IF FORMAT(FieldRef.TYPE) = 'MediaSet' THEN
                InnerText := ExportMediaSet(FieldRef);
            IF FORMAT(FieldRef.TYPE) = 'Media' THEN
                InnerText := ExportMedia(FieldRef, ConfigPackage);
        END ELSE BEGIN
            IF FORMAT(FieldRef.TYPE) = 'Option' THEN
                InnerText := FORMAT(FieldRef.VALUE);
            IF (FORMAT(FieldRef.TYPE) = 'Date') AND (ConfigPackage."Language ID" <> 0) AND (InnerText <> '') THEN BEGIN
                EVALUATE(Date, FORMAT(FieldRef.VALUE));
                InnerText := TypeHelper.FormatDate(Date, ConfigPackage."Language ID");
            END;
            // >>
            IF FORMAT(FieldRef.TYPE) = 'BLOB' THEN BEGIN
                InnerText := ConvertBLOBToString(FieldRef);
            END;
            // <<
        END;

        EXIT(InnerText);
    end;

    procedure GetAttribute(AttributeName: Text[1024]; var XMLNode: DotNet XmlNode): Text[1000]
    var
        XMLAttributes: DotNet XmlNamedNodeMap;
        XMLAttributeNode: DotNet XmlNode;
    begin
        XMLAttributes := XMLNode.Attributes;
        XMLAttributeNode := XMLAttributes.GetNamedItem(AttributeName);
        IF ISNULL(XMLAttributeNode) THEN
            EXIT('');
        EXIT(FORMAT(XMLAttributeNode.InnerText));
    end;

    local procedure GetDimValueFromTable(var RecRef: RecordRef; DimCode: Code[20]): Code[20]
    var
        DimSetEntry: Record 480;
        DefaultDim: Record 352;
        ConfigMgt: Codeunit 8616;
        FieldRef: FieldRef;
        DimSetID: Integer;
        MasterNo: Code[20];
    begin
        IF RecRef.FIELDEXIST(480) THEN BEGIN // Dimension Set ID
            FieldRef := RecRef.FIELD(480);
            DimSetID := FieldRef.VALUE;
            IF DimSetID > 0 THEN BEGIN
                DimSetEntry.SETRANGE("Dimension Set ID", DimSetID);
                DimSetEntry.SETRANGE("Dimension Code", DimCode);
                IF DimSetEntry.FINDFIRST THEN
                    EXIT(DimSetEntry."Dimension Value Code");
            END;
        END ELSE
            IF ConfigMgt.IsDefaultDimTable(RecRef.NUMBER) THEN BEGIN // Default Dimensions
                FieldRef := RecRef.FIELD(1);
                DefaultDim.SETRANGE("Table ID", RecRef.NUMBER);
                MasterNo := FORMAT(FieldRef.VALUE);
                DefaultDim.SETRANGE("No.", MasterNo);
                DefaultDim.SETRANGE("Dimension Code", DimCode);
                IF DefaultDim.FINDFIRST THEN
                    EXIT(DefaultDim."Dimension Value Code");
            END;
    end;

    [Scope('Personalization')]
    procedure GetElementName(NameIn: Text[250]): Text[250]
    var
        XMLDOMManagement: Codeunit 6224;
    begin
        OnBeforeGetElementName(NameIn);

        IF NOT XMLDOMManagement.IsValidXMLNameStartCharacter(NameIn[1]) THEN
            NameIn := '_' + NameIn;
        NameIn := COPYSTR(XMLDOMManagement.ReplaceXMLInvalidCharacters(NameIn, ' '), 1, MAXSTRLEN(NameIn));
        NameIn := DELCHR(NameIn, '=', '?''`');
        NameIn := CONVERTSTR(NameIn, '<>,./\+&()%:', '            ');
        NameIn := CONVERTSTR(NameIn, '-', '_');
        NameIn := DELCHR(NameIn, '=', ' ');
        EXIT(NameIn);
    end;

    [Scope('Personalization')]
    procedure GetFieldElementName(NameIn: Text[250]): Text[250]
    begin
        IF AddPrefixMode THEN
            NameIn := COPYSTR('Field_' + NameIn, 1, MAXSTRLEN(NameIn));

        EXIT(GetElementName(NameIn));
    end;

    [Scope('Personalization')]
    procedure GetTableElementName(NameIn: Text[250]): Text[250]
    begin
        IF AddPrefixMode THEN
            NameIn := COPYSTR('Table_' + NameIn, 1, MAXSTRLEN(NameIn));

        EXIT(GetElementName(NameIn));
    end;

    local procedure GetNodeValue(var RecordNode: DotNet XmlNode; FieldNodeName: Text[250]): Text
    var
        FieldNode: DotNet XmlNode;
    begin
        FieldNode := RecordNode.SelectSingleNode(FieldNodeName);
        IF NOT ISNULL(FieldNode) THEN
            EXIT(FieldNode.InnerText);
    end;

    local procedure GetPackageTag(): Text
    begin
        EXIT(DataListTxt);
    end;

    [Scope('Internal')]
    procedure GetPackageCode(PackageXML: DotNet XmlDocument): Code[20]
    var
        ConfigPackage: Record 8623;
        DocumentElement: DotNet XmlElement;
    begin
        DocumentElement := PackageXML.DocumentElement;
        EXIT(COPYSTR(GetAttribute(GetElementName(ConfigPackage.FIELDNAME(Code)), DocumentElement), 1, MAXSTRLEN(ConfigPackage.Code)));
    end;

    local procedure InitializeMediaTempFolder()
    var
        MediaFolder: Text;
    begin
        IF ExcelMode THEN
            EXIT;

        IF WorkingFolder = '' THEN
            EXIT;

        MediaFolder := GetCurrentMediaFolderPath;
        IF FileManagement.ServerDirectoryExists(MediaFolder) THEN
            FileManagement.ServerRemoveDirectory(MediaFolder, TRUE);

        FileManagement.ServerCreateDirectory(MediaFolder);
    end;

    local procedure GetCurrentMediaFolderPath(): Text
    begin
        EXIT(FileManagement.CombinePath(WorkingFolder, GetMediaFolderName));
    end;

    [Scope('Internal')]
    procedure GetMediaFolder(var MediaFolderPath: Text; SourcePath: Text): Boolean
    var
        SourceDirectory: Text;
    begin
        IF FileManagement.ServerFileExists(SourcePath) THEN
            SourceDirectory := FileManagement.GetDirectoryName(SourcePath)
        ELSE
            IF FileManagement.ServerDirectoryExists(SourcePath) THEN
                SourceDirectory := SourcePath;

        IF SourceDirectory = '' THEN
            EXIT(FALSE);

        MediaFolderPath := FileManagement.CombinePath(SourceDirectory, GetMediaFolderName);
        EXIT(FileManagement.ServerDirectoryExists(MediaFolderPath));
    end;

    [Scope('Personalization')]
    procedure GetMediaFolderName(): Text
    begin
        EXIT('Media');
    end;

    [Scope('Personalization')]
    procedure GetXSDType(TableID: Integer; FieldID: Integer): Text[30]
    var
        "Field": Record 2000000041;
    begin
        IF Field.GET(TableID, FieldID) THEN
            CASE Field.Type OF
                Field.Type::Integer:
                    EXIT('xsd:integer');
                Field.Type::Date:
                    EXIT('xsd:date');
                Field.Type::Time:
                    EXIT('xsd:time');
                Field.Type::Boolean:
                    EXIT('xsd:boolean');
                Field.Type::DateTime:
                    EXIT('xsd:dateTime');
                ELSE
                    EXIT('xsd:string');
            END;

        EXIT('xsd:string');
    end;

    [Scope('Personalization')]
    procedure SetAdvanced(NewAdvanced: Boolean)
    begin
        Advanced := NewAdvanced;
    end;

    [Scope('Personalization')]
    procedure SetCalledFromCode(NewCalledFromCode: Boolean)
    begin
        CalledFromCode := NewCalledFromCode;
    end;

    local procedure SetWorkingFolder(NewWorkingFolder: Text)
    begin
        WorkingFolder := NewWorkingFolder;
    end;

    [Scope('Personalization')]
    procedure SetExcelMode(NewExcelMode: Boolean)
    begin
        ExcelMode := NewExcelMode;
    end;

    [Scope('Personalization')]
    procedure SetHideDialog(NewHideDialog: Boolean)
    begin
        HideDialog := NewHideDialog;
    end;

    [Scope('Personalization')]
    procedure SetExportFromWksht(NewExportFromWksht: Boolean)
    begin
        ExportFromWksht := NewExportFromWksht;
    end;

    [Scope('Personalization')]
    procedure SetPrefixMode(PrefixMode: Boolean)
    begin
        AddPrefixMode := PrefixMode;
    end;

    [Scope('Personalization')]
    procedure TableObjectExists(TableId: Integer): Boolean
    var
        TableMetadata: Record 2000000136;
    begin
        EXIT(TableMetadata.GET(TableId) AND (TableMetadata.ObsoleteState <> TableMetadata.ObsoleteState::Removed));
    end;

    [Scope('Internal')]
    procedure DecompressPackage(ServerFileName: Text) DecompressedFileName: Text
    begin
        DecompressedFileName := FileManagement.ServerTempFileName('');
        IF NOT ConfigPckgCompressionMgt.ServersideDecompress(ServerFileName, DecompressedFileName) THEN
            ERROR(WrongFileTypeErr);
    end;

    [Scope('Personalization')]
    procedure DecompressPackageToBlob(var TempBlob: Record 99008535; var TempBlobUncompressed: Record 99008535)
    var
        InStream: InStream;
        OutStream: OutStream;
        CompressionMode: DotNet CompressionMode;
        CompressedStream: DotNet GZipStream;
    begin
        TempBlob.Blob.CREATEINSTREAM(InStream);
        CompressedStream := CompressedStream.GZipStream(InStream, CompressionMode.Decompress); // Decompress the stream
        TempBlobUncompressed.Blob.CREATEOUTSTREAM(OutStream);  // Creates outstream to enable you to write data to the blob.
        COPYSTREAM(OutStream, CompressedStream); // Copy contents from the CompressedStream to the OutStream, this populates the blob with the decompressed file.
    end;

    local procedure UploadXMLPackage(ServerFileName: Text): Boolean
    begin
        EXIT(UPLOAD(ImportFileTxt, '', GetFileDialogFilter, '', ServerFileName));
    end;

    [Scope('Personalization')]
    procedure GetFileDialogFilter(): Text
    begin
        EXIT(FileDialogFilterTxt);
    end;

    local procedure ConvertBLOBToBase64String(var FieldRef: FieldRef): Text
    var
        TempBlob: Record 99008535;
    begin
        FieldRef.CALCFIELD;
        TempBlob.Blob := FieldRef.VALUE;
        EXIT(TempBlob.ToBase64String);
    end;

    local procedure ExportMediaSet(var FieldRef: FieldRef): Text
    var
        TempConfigMediaBuffer: Record 8630 temporary;
        FilesExported: Integer;
        ItemPrefixPath: Text;
        MediaFolder: Text;
    begin
        IF ExcelMode THEN
            EXIT;

        IF NOT GetMediaFolder(MediaFolder, WorkingFolder) THEN
            EXIT('');

        TempConfigMediaBuffer.INIT;
        TempConfigMediaBuffer."Media Set" := FieldRef.VALUE;
        TempConfigMediaBuffer.INSERT;
        IF TempConfigMediaBuffer."Media Set".COUNT = 0 THEN
            EXIT;

        ItemPrefixPath := MediaFolder + '\' + FORMAT(TempConfigMediaBuffer."Media Set");
        FilesExported := TempConfigMediaBuffer."Media Set".EXPORTFILE(ItemPrefixPath);
        IF FilesExported <= 0 THEN
            EXIT('');

        EXIT(FORMAT(FieldRef.VALUE));
    end;

    local procedure ExportMedia(var FieldRef: FieldRef; ConfigPackage: Record 8623): Text
    var
        ConfigMediaBuffer: Record 8630;
        TempConfigMediaBuffer: Record 8630 temporary;
        MediaOutStream: OutStream;
        MediaIDGuidText: Text;
        BlankGuid: Guid;
    begin
        IF ExcelMode THEN
            EXIT;

        MediaIDGuidText := FORMAT(FieldRef.VALUE);
        IF (MediaIDGuidText = '') OR (MediaIDGuidText = FORMAT(BlankGuid)) THEN
            EXIT;

        ConfigMediaBuffer.INIT;
        ConfigMediaBuffer."Package Code" := ConfigPackage.Code;
        ConfigMediaBuffer."Media ID" := MediaIDGuidText;
        ConfigMediaBuffer."No." := ConfigMediaBuffer.GetNextNo;
        ConfigMediaBuffer.INSERT;

        ConfigMediaBuffer."Media Blob".CREATEOUTSTREAM(MediaOutStream);

        TempConfigMediaBuffer.INIT;
        TempConfigMediaBuffer.Media := FieldRef.VALUE;
        TempConfigMediaBuffer.INSERT;
        TempConfigMediaBuffer.Media.EXPORTSTREAM(MediaOutStream);

        ConfigMediaBuffer.MODIFY;

        EXIT(MediaIDGuidText);
    end;

    local procedure GetConfigPackageDataValue(var ConfigPackageData: Record 8615; var RecordNode: DotNet XmlNode; FieldNodeName: Text[250])
    var
        TempBlob: Record 99008535;
    begin
        // >>
        //IF ConfigPackageMgt.IsBLOBField(ConfigPackageData."Table ID",ConfigPackageData."Field ID") AND NOT ExcelMode THEN BEGIN
        IF ConfigPackageMgt.IsBLOBField(ConfigPackageData."Table ID", ConfigPackageData."Field ID") THEN BEGIN
            IF ExcelMode THEN
                TempBlob.WriteAsText(GetNodeValue(RecordNode, FieldNodeName), TEXTENCODING::UTF8)
            ELSE
                // <<
                TempBlob.FromBase64String(GetNodeValue(RecordNode, FieldNodeName));
            ConfigPackageData."BLOB Value" := TempBlob.Blob;
        END ELSE
            ConfigPackageData.Value := COPYSTR(GetNodeValue(RecordNode, FieldNodeName), 1, MAXSTRLEN(ConfigPackageData.Value));
    end;

    local procedure UpdateConfigPackageMediaSet(ConfigPackage: Record 8623)
    var
        TempNameValueBuffer: Record 823 temporary;
        FileManagement: Codeunit 419;
        MediaFolder: Text;
    begin
        IF NOT GetMediaFolder(MediaFolder, WorkingFolder) THEN
            EXIT;

        FileManagement.GetServerDirectoryFilesList(TempNameValueBuffer, MediaFolder);
        IF NOT TempNameValueBuffer.FINDSET THEN
            EXIT;

        REPEAT
            ImportMediaSetFromFile(ConfigPackage, TempNameValueBuffer.Name);
        UNTIL TempNameValueBuffer.NEXT = 0;

        FileManagement.ServerRemoveDirectory(MediaFolder, TRUE);
    end;

    local procedure ExportConfigPackageMediaSetToXML(var PackageXML: DotNet XmlDocument; ConfigPackage: Record 8623)
    var
        ConfigMediaBuffer: Record 8630;
        ConfigPackageTable: Record 8613;
        ConfigPackageManagement: Codeunit 8611;
    begin
        ConfigMediaBuffer.SETRANGE("Package Code", ConfigPackage.Code);
        IF ConfigMediaBuffer.ISEMPTY THEN
            EXIT;

        ConfigPackageManagement.InsertPackageTable(ConfigPackageTable, ConfigPackage.Code, DATABASE::"Config. Media Buffer");
        ConfigPackageTable.CALCFIELDS("Table Name");
        ExportConfigTableToXML(ConfigPackageTable, PackageXML);
    end;

    local procedure ImportMediaSetFromFile(ConfigPackage: Record 8623; FileName: Text)
    var
        TempBlob: Record 99008535 temporary;
        ConfigMediaBuffer: Record 8630;
        FileManagement: Codeunit 419;
        DummyGuid: Guid;
    begin
        ConfigMediaBuffer.INIT;
        FileManagement.BLOBImportFromServerFile(TempBlob, FileName);
        ConfigMediaBuffer."Media Blob" := TempBlob.Blob;
        ConfigMediaBuffer."Package Code" := ConfigPackage.Code;
        ConfigMediaBuffer."Media Set ID" := COPYSTR(FileManagement.GetFileNameWithoutExtension(FileName), 1, STRLEN(FORMAT(DummyGuid)));
        ConfigMediaBuffer."No." := ConfigMediaBuffer.GetNextNo;
        ConfigMediaBuffer.INSERT;
    end;

    local procedure CleanUpConfigPackageData(ConfigPackage: Record 8623)
    var
        ConfigMediaBuffer: Record 8630;
    begin
        ConfigMediaBuffer.SETRANGE("Package Code", ConfigPackage.Code);
        ConfigMediaBuffer.DELETEALL;
    end;

    LOCAL procedure ConvertBLOBToString(VAR FieldRef: FieldRef): Text
    var
        TempBlob: Record TempBlob;
        CR: Text[1];
    begin
        // >>
        CR[1] := 10;
        FieldRef.CALCFIELD;
        TempBlob.Blob := FieldRef.VALUE;
        EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
        // <<
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeGetElementName(var NameIn: Text[250])
    begin
    end;
}

