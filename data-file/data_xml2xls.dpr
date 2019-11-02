{$IFDEF FPC}
 {$MODE Delphi}
{$ELSE}
 {$APPTYPE CONSOLE}
{$ENDIF}

uses
   SysUtils, windows, xml_work, xls_work, data_structure;

var
  XMLin : TXML;
  XLS : TXLS;
  XMLout : text;

procedure ErrorExit(ErrorMsg : widestring);
 begin
   writeln;
   SetConsoleOutputCP(CP_UTF8);
   writeln(ansistring(ErrorMsg));
   halt(1);
 end; { ErrorExit }

//////////////////////////////////////////////////////
// Структура

procedure ReadTablesHeaders(CreateSheets : boolean); forward;

procedure CreateTablesStructure(CreateSheets : boolean);
 begin
   CreateTables();
   ReadTablesHeaders(CreateSheets);
 end; { CreateTablesStructure }

//////////////////////////////////////////////////////
// Запись XLS на основе XML

//procedure WriteCell(TabNum : TTableNum; Row,Col : integer; Value : TXMLString);
procedure WriteCellText(XLSSheetID : TXLSSheetID; Row,Col : integer; Value : TXMLString);
 {uses: XLS, Tables, MaxTable}
 begin
   XLS.WriteCellData(XLSSheetID,Row,Col,Value,'String');
 end; { WriteCell }

procedure WriteCell(XLSSheetID : TXLSSheetID; Row,Col : integer; Value : TXMLString; var Typ : TXMLString);
 {uses: XLS, Tables, MaxTable}
 begin
   XLS.WriteCellData(XLSSheetID,Row,Col,Value,Typ);
 end; { WriteCell }

procedure WriteTableToStructure(TableNum,ParentID : integer);
 {uses: XLS, Tables}
 begin
   with Tables[TableNum] do begin
     if AllXMLDataFilled then exit;
     if RowInStructure = 0 then begin
       inc(Tables[0].Height);
       RowInStructure := Tables[0].Height;
       inc(Tables[0].Height);
     end;
     WriteCellText(Tables[0].XLSSheetID,RowInStructure,1,'table_'+inttostr(ParentID)+'_'+inttostr(TableNum));
     WriteCellText(Tables[0].XLSSheetID,RowInStructure,2,XMLtoXMLText(TableKeyPath));
     WriteCellText(Tables[0].XLSSheetID,RowInStructure,3,XMLtoXMLText(XLSSheetName));
     WriteCellText(Tables[0].XLSSheetID,RowInStructure+1,1,'XML data');
     WriteCellText(Tables[0].XLSSheetID,RowInStructure+1,2,XMLtoXMLText(TableKeyValue));
     WriteCellText(Tables[0].XLSSheetID,RowInStructure+1,3,XMLtoXMLText(RowsKeyValue));
     if (TableKeyPath <> '') and (XLSSheetName<>'') and (TableKeyValue<>'') and (RowsKeyValue<>'')
      then AllXMLDataFilled := true;
   end;
 end; { WriteTableToStructure }

procedure WriteTextToStructure(XMLText : TXMLString);
 {uses: XLS, Tables}
 begin
   inc(Tables[0].Height);
   WriteCellText(Tables[0].XLSSheetID,Tables[0].Height,1,'text');
   WriteCellText(Tables[0].XLSSheetID,Tables[0].Height,2,XMLtoXMLText(XMLText));
 end; { WriteTextToStructure }

procedure WriteTablesHeaders();
 {uses: XLS, Tables, MaxTable}
 var TN,FN : integer;
 begin
//сделано при чтении таблиц из DATA_BIN:
//WriteCellText(Tables[0].XLSSheetID,Tables[0].Height,1,'table_'+inttostr(ParentID)+'_'+inttostr(TableNum));
//WriteCellText(Tables[0].XLSSheetID,Tables[0].Height,2,XML.Key_Name);
//WriteCellText(Tables[0].XLSSheetID,Tables[0].Height,3,XMLtoXMLText(Tables[TableNum].XLSSheetName));
   for TN:=0 to MaxTable do with Tables[TN] do begin
     for FN:=1 to Width do begin
       //для таблицы 0 поля не нужны, вывод срезан через RowInStructure < 0
       WriteCellText(Tables[0].XLSSheetID,RowInStructure  ,3+FN,Fields[FN].KeyName);
       WriteCellText(Tables[0].XLSSheetID,RowInStructure+1,3+FN,XMLtoXMLText(Fields[FN].KeyValue));
       WriteCellText(XLSSheetID,1,Fields[FN].Col,Fields[FN].KeyName);
       WriteCellText(XLSSheetID,2,Fields[FN].Col,Fields[FN].ColName);
     end;
     XLS.TruncateRow(Tables[0].XLSSheetID,RowInStructure  ,3+Width);
     XLS.TruncateRow(Tables[0].XLSSheetID,RowInStructure+1,3+Width);
     XLS.TruncateTable(Tables[TN].XLSSheetID,Tables[TN].Height);
   end;
 end; { WriteTablesHeaders }

//////////////////////////////////////////////////////
// расширение чтения XML

procedure ErrorExit_InXML(ErrorMsg : widestring);
 {uses: XMLin}
 begin
   ErrorExit(
        'Error in XML: '+ErrorMsg+#13#10+
        'Last open path: '+XMLtoWide(XMLin.Key_Path)+#13#10+
        'Last text value: '+XMLtoWide(XMLin.Text_Value)+#13#10+
        'Last key value: '+XMLtoWide(XMLin.Key_Value)
   );
 end; { ErrorExit_InXML }

procedure ErrorExit_InXLS(SheetID,Row,Col : integer; ErrorMsg : widestring);
 {uses: Tables}
 begin
   ErrorExit(
        'Error in XLS: '+ErrorMsg+#13#10+
        'Table: "'+XMLtoWide(Tables[SheetID].TableKeyName)+'"'#13#10+
        'Sheet: "'+XMLtoWide(Tables[SheetID].XLSSheetName)+'"'#13#10+
        'Row: '+inttostr(Row)+' Col: '+inttostr(Col)
   );
 end; { ErrorExit_InXLS }

procedure ExtXML_ReadKey();
 {uses: XMLin}
 begin
   if not XMLin.ReadKey() then ErrorExit_InXML('unexpected end of file');
 end; { ExtXML_ReadKey }

procedure ExtXML_ReadKeyOpening(KeyName : TXMLString);
 {uses: XMLin}
 begin
   if not XMLin.ReadNeededKey(KeyName) or (XMLin.KeyIsClosed <> false)
    then ErrorExit_InXML('key "'+XMLtoWide(KeyName)+'" expected');
 end; { ExtXML_ReadKeyOpening }

procedure ExtXML_ReadKeyClosing(KeyName : TXMLString);
 {uses: XMLin}
 begin
   if not XMLin.ReadNeededKey(KeyName) or (XMLin.KeyIsClosed <> true)
     then ErrorExit_InXML('closing key "'+XMLtoWide(KeyName)+'" expected');
 end; { ExtXML_ReadKeyClosing }

function ExtXML_ReadSimpleKey(KeyName : TXMLString) : TXMLString;
 {uses: XMLin}
 begin
   ExtXML_ReadKeyOpening(KeyName);
   ExtXML_ReadKeyClosing(KeyName);
   result := XMLin.Text_Value;
 end; { ExtXML_ReadSimpleKey }

function ExtXML_ReadIntegerKey(KeyName : TXMLString) : longint;
 {uses: XMLin}
 var E : integer;
 begin
   val(ExtXML_ReadSimpleKey(KeyName), result, E );
   if E <> 0 then ErrorExit_InXML('key "'+XMLtoWide(KeyName)+'" is not integer');
 end; { ExtXML_ReadIntegerKey }

//////////////////////////////////////////////////////
// чтение таблиц из XML

procedure ReadTablesHeaders(CreateSheets : boolean);
 {uses: XLS, Tables}
 var
   RN,TN,CN,FN : integer;
   TablePath,FieldName,ColName : TXMLString;
 begin
   Tables[0].XLSSheetID := XLS.GetXLSSheetID(Tables[0].XLSSheetName,CreateSheets);
   if XLS.SheetHeight(Tables[0].XLSSheetID) >=2 then begin
     //прочитать лист "Структура" (Tables[0]):
     for RN:=1 to XLS.SheetHeight(Tables[0].XLSSheetID)
      do if copy(XLS.ReadCellData(Tables[0].XLSSheetID,RN,1),1,5) = 'table' then begin
        //полное имя таблицы
        TablePath := XMLTextToXML(XLS.ReadCellData(Tables[0].XLSSheetID,RN,2));
        TN:=FindTableNum(TablePath);
        if TN = 0 then continue;
        with Tables[TN] do begin
          RowInStructure := RN;
          //имя листа для таблицы
          XLSSheetName := XMLTextToXML(XLS.ReadCellData(Tables[0].XLSSheetID,RN,3));
          XLSSheetID := XLS.GetXLSSheetID(XLSSheetName,false);
          if XLSSheetID < 0 then ErrorExit_InXLS(0,RN,3,'bad sheet name');
          //ключи XML
          TableKeyValue := XMLTextToXML(XLS.ReadCellData(Tables[0].XLSSheetID,RN+1,2));
          RowsKeyValue := XMLTextToXML(XLS.ReadCellData(Tables[0].XLSSheetID,RN+1,3));
          //состав и порядок полей таблицы
          for CN:=4 to XLS.RowWidth(Tables[0].XLSSheetID,RN) do begin
            FieldName := XLS.ReadCellData(Tables[0].XLSSheetID,RN,CN);
            if FieldName = '' then ErrorExit_InXLS(0,RN,CN,'bad field name');
            inc(Width);
            Fields[Width].KeyName := FieldName;
            Fields[Width].KeyValue := XMLTextToXML(XLS.ReadCellData(Tables[0].XLSSheetID,RN+1,CN));
          end;
          //состав и порядок полей на листе
          for CN:=1 to XLS.RowWidth(XLSSheetID,2) do begin
            FieldName := XLS.ReadCellData(XLSSheetID,1,CN);
            if FieldName = '' then continue;
            ColName := XLS.ReadCellData(XLSSheetID,2,CN);
            FN := FindFieldNum(Fields,Width,Tables[TN],FieldName,'',ColName,-1);
            if FN = 0 then ErrorExit_InXLS(TN,1,CN,'bad field name');
            Fields[FN].ColName := ColName;
            Fields[FN].Col := CN;
          end;
        end;
     end;
   end;
   //для отсутствующих листов
   for TN:=1 to MaxTable do with Tables[TN] do begin
     XLSSheetID := XLS.GetXLSSheetID(XLSSheetName,CreateSheets);
   end;
   //только для выгрузки - отметить поля-вложенные_таблицы
   for TN:=1 to MaxTable do with Tables[TN] do begin
     for FN:=1 to Width do with Fields[FN] do begin
       TableNum := FindTableNum(TableKeyPath+'>ITEM>'+KeyName);
     end;
   end;

 end; { ReadTablesHeaders }

procedure ReadTable(ParentID : integer); forward;

function ParentIDtoStr(ParentID : integer) : TXMLString;
 begin
   result:='['+inttostr(ParentID)+']';
 end; { ParentIDtoStr }

procedure ReadFields(var Table : TTable; ParentID : integer);
 {uses: XMLin, Tables(заголовки)}
 var
   Row,FN,PredFieldNum : integer;
   V : TXMLString;
 function FindFieldNum_WithCheck(FieldName,FieldKeyValue,ColumnName : TXMLString; CreateAfterField : longint) : TFieldNum;
  begin
    result:=FindFieldNum(Table.Fields,Table.Width,Table,FieldName,FieldKeyValue,ColumnName,CreateAfterField);
    if (result=0) and (CreateAfterField>=0)
     then ErrorExit_InXML('too many fields in table '+XMLtoWide(Table.TableKeyPath));
  end; { FindFieldNum_WithCheck }
 begin with Table do begin
   Row:=Height; PredFieldNum:=0;
   if ParentID <> 0 then begin
     FN:=FindFieldNum_WithCheck('PARENTID','ParentID','ParentID',0);
     PredFieldNum:=FN;
     WriteCell(XLSSheetID,Row,Fields[FN].Col,ParentIDtoStr(ParentID),Fields[FN].Typ);
   end;
   repeat
     ExtXML_ReadKey();
     if          (XMLin.Key_Name = 'ITEM') and (XMLin.KeyIsClosed = true) then begin
       V:=trim(XMLin.Text_Value);
       if V<>'' then begin
         FN:=FindFieldNum_WithCheck('[VALUE]','','[value]',PredFieldNum);
         WriteCell(XLSSheetID,Row,Fields[FN].Col,V,Fields[FN].Typ);
       end;
       break;
     end else begin
       FN:=FindFieldNum_WithCheck(XMLin.Key_Name,XMLin.Key_Value,copy(XMLin.Key_Value,2,Length(XMLin.Key_Name)),PredFieldNum);
       PredFieldNum:=FN;
       if ( pos('<'+XMLin.Key_Name+'>', FieldTable_KeyNames) > 0) and (XMLin.KeyIsClosed = false) then begin
         WriteCell(XLSSheetID,Row,Fields[FN].Col,ParentIDtoStr(Row),Fields[FN].Typ);
         ReadTable(Row);
       end else begin
         ExtXML_ReadKeyClosing(XMLin.Key_Name);
         WriteCell(XLSSheetID,Row,Fields[FN].Col,XMLin.Text_Value,Fields[FN].Typ);
       end;
     end;
   until false;
 end end; { ReadFields }

procedure ReadTable(ParentID : integer);
 {uses: XMLin, Tables}
 var
   TableNum : TTableNum;
 begin
   TableNum := FindTableNum(XMLin.Key_Path);
   if TableNum = 0 then ErrorExit_InXML('unknown table');
   with Tables[TableNum] do begin
     TableKeyValue := XMLin.Key_Value;
     Count := ExtXML_ReadIntegerKey('count');
     ItemVersion := ExtXML_ReadIntegerKey('item_version');
     repeat
       ExtXML_ReadKey();
       if          (XMLin.Key_Name = 'ITEM') and (XMLin.KeyIsClosed = false) then begin
         if RowsKeyValue = '' then RowsKeyValue := XMLin.Key_Value;
         if Height = 0 then begin
           FieldsKeyValue := XMLin.Key_Value;
         end;
         inc(Height);
         ReadFields(Tables[TableNum],ParentID);
       end else if (XMLin.Key_Name = TableKeyName) and (XMLin.KeyIsClosed = true) then begin
         break;
       end else begin
         ErrorExit_InXML('bad key in table');
       end;
     until false;
   end;
   WriteTableToStructure(TableNum,ParentID);
 end; { ReadTable }

procedure Write_Data_to_XLS();
var TN : integer;
begin

  CreateTablesStructure(true);
  for TN:=0 to MaxTable do with Tables[TN] do begin
    RowInStructure := 0; // порядок задастся входным файлом
  end;

  if not XMLin.FindKey('data_bin',false,false) then ErrorExit_InXML('Main key "data_bin" not found');

  WriteTextToStructure(trim(XMLin.Text_Value+XMLin.Key_Value));
  repeat
    ExtXML_ReadKey;
    if XMLin.KeyIsClosed = false
     then begin
       ReadTable(0);
     end else begin
       WriteTextToStructure(trim(XMLin.Text_Value+XMLin.Key_Value));
    end;
  until XMLin.Key_Path = '';
  WriteTablesHeaders();

  XLS.SaveFile();

end; { WriteData_to_XLS }

procedure WriteTable(TableNum : TTableNum; ParentIDStr,KeyPrefixString : TXMLString);
 {uses XLS, Tables }
 var RowsSelection : TRowsSelection;
     Row,FN : integer;
     CellValue,RowKeyPrefixString,ItemKeyPrefixString : TXMLString;
 begin with Tables[TableNum] do begin
   writeln(XMLout,KeyPrefixString+TableKeyValue);
   if TableKeyIsSimple = false then begin
     TableKeyValue := DeleteXMLAllAttributes(TableKeyValue);
     TableKeyIsSimple := true;
   end;
   if ParentIDStr = ''
    then RowsSelection := XLS.SelectRows(XLSSheetID,1,ParentIDStr,false)
    else RowsSelection := XLS.SelectRows(XLSSheetID,1,ParentIDStr,true);
   Count:=Length(RowsSelection);
   RowKeyPrefixString:=KeyPrefixString+#9;
   writeln(XMLout,RowKeyPrefixString+'<count>'+inttostr(Count)+'</count>');
   writeln(XMLout,RowKeyPrefixString+'<item_version>'+inttostr(ItemVersion)+'</item_version>');
   ItemKeyPrefixString:=RowKeyPrefixString+#9;
   for Row:=1 to Count do begin
     if ParentIDStr <> '' then FN:=2 else FN:=1;
     if FN=Width then begin
       CellValue:=XLS.ReadCellData(XLSSheetID,RowsSelection[Row-1],Fields[FN].Col);
       writeln(XMLout,RowKeyPrefixString+'<item>'+XMLtoXMLText(CellValue)+'</item>');
       continue;
     end;

     writeln(XMLout,RowKeyPrefixString+RowsKeyValue);
     if RowsKeyIsSimple = false then begin
       RowsKeyValue := DeleteXMLAllAttributes(RowsKeyValue);
       RowsKeyIsSimple := true;
     end;

     for FN:=FN to Width do with Fields[FN] do begin
       if ParentField > 0
        then if XLS.ReadCellData(XLSSheetID,RowsSelection[Row-1],Fields[ParentField].Col) <> NeedParentValue
              then continue;

       CellValue:=XLS.ReadCellData(XLSSheetID,RowsSelection[Row-1],Col);
       if TableNum <> 0 then begin
         if CellValue <> ''
          then WriteTable(TableNum,CellValue,ItemKeyPrefixString);
       end else begin
         write  (XMLout,ItemKeyPrefixString+KeyValue);
         write  (XMLout,CellValue);
         writeln(XMLout,CloseXMLKey(KeyValue));
       end;
     end;
     writeln(XMLout,RowKeyPrefixString+'</item>');

   end;
   writeln(XMLout,KeyPrefixString+CloseXMLKey(TableKeyValue));
 end end; { WriteTable }

procedure Write_XLS_to_Data();
var
    TN,StructureHeight : integer;
    StructureCmd,StructureData : TXMLString;
begin {$I-}

  CreateTablesStructure(false);
  for TN:=0 to MaxTable do with Tables[TN] do begin
    if XLSSheetID < 0
     then ErrorExit_InXLS(TN,0,0,'Table/sheet not found');
  end;

  StructureHeight := XLS.SheetHeight(Tables[0].XLSSheetID);
  while Tables[0].Height < StructureHeight do begin
    inc(Tables[0].Height);
    StructureCmd := XLS.ReadCellData(Tables[0].XLSSheetID,Tables[0].Height,1);
    StructureData := XMLTextToXML(XLS.ReadCellData(Tables[0].XLSSheetID,Tables[0].Height,2));
    if          StructureCmd = '' then begin
      //не надо
    end else if StructureCmd = 'text' then begin
      writeln(XMLout,StructureData);
    end else if copy(StructureCmd,1,7) = 'table_0' then begin
      TN := FindTableNum(StructureData);
      if TN = 0
       then ErrorExit_InXLS(0,Tables[0].Height,2,'Bad table path "'+XMLtoWide(StructureData)+'"');
      WriteTable(TN,'',#9);
    end else if copy(StructureCmd,1,5) = 'table' then begin
      //не надо - вложенная таблица
    end else if StructureCmd = 'XML data' then begin
      //не надо - вспомогательные данные таблицы
    end else begin
      ErrorExit_InXLS(0,Tables[0].Height,1,'Bad structure command "'+StructureCmd+'"')
    end;
  end;
  flush(XMLout);

end; { Write_XLS_to_Data }

procedure ErrorExit_InParameters(ErrorMsg : widestring);
 begin
   ErrorExit(
        'Export-import data.bin (allods2 file) between <XML format> to/from <Excel XML 2003 format>'#13#10#13#10+
        'Error in parameters: '+ErrorMsg+#13#10+
        '"data.xml file name" "excel.xml file name" - for export data to excel(EXISTING file updating)'+#13#10+
        '"excel.xml file name" "data.xml file name" /IMPORT - for import excel to data(NEW file creating)'+#13#10
   );
 end; { ErrorExit_InParameters }

procedure CheckXMLformat(ParameterNumber,XMLKeyNumber : integer; XMLKeyValue : TXMLString);
 var XML : TXML;
     Msg : widestring;
 begin
   if ParameterNumber = 0 then exit;
   XML:=TXML.Create();
   if XML = nil then ErrorExit('Internal error: XML create');
   Msg:=paramstr(ParameterNumber);
   if not XML.Open_Read(Msg)
    then ErrorExit(XML.LastErrorMsg);
   Msg:='Error in XML file "'+Msg+'"'#13#10;
   while XMLKeyNumber > 0 do begin
     if not XML.ReadKey()
      then ErrorExit(Msg+XML.LastErrorMsg);
     dec(XMLKeyNumber);
   end;
   if XML.Key_Value <> XMLKeyValue
    then ErrorExit(Msg+': need key '+XMLtoWide(XMLKeyValue));
   XML.Destroy();
 end; { CheckXMLformat }

procedure RecognizeParameters();
{uses: XMLin, XMLout, XLS }
var
    DoImport : boolean;
    Pdata,Pexcel,FileError : integer;
begin{$I-}

  if paramcount < 2 then ErrorExit_InParameters('');
  DoImport := ( AnsiUpperCase(paramstr(3)) = '/IMPORT' );

  if DoImport then begin
    Pexcel:=1; Pdata:=0;
  end else begin
    Pdata:=1; Pexcel:=2;
  end;

//<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>              //2
//<!DOCTYPE boost_serialization>                                        //2
//<boost_serialization signature="serialization::archive" version="5">  //1
//<data_bin class_id="0" tracking_level="0" version="0">                //1
  CheckXMLformat(Pdata,6,'<data_bin class_id="0" tracking_level="0" version="0">');

//<?xml version="1.0"?>                                                 //2
//<?mso-application progid="Excel.Sheet"?>                              //2
//<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"        //1
  CheckXMLformat(Pexcel,3,'<?mso-application progid="Excel.Sheet"?>');

  XLS:=TXLS.Create();
  if XLS = nil then ErrorExit('Internal error: XLS create');
  _UndefinedString := '';

  if DoImport then begin

    if not XLS.Open_Read(paramstr(1))
     then ErrorExit(XLS.LastErrorMsg);

    assign(XMLout,paramstr(2)); rewrite(XMLout);
    FileError:=ioresult;
    if FileError<>0 then ErrorExit('Error create file "'+paramstr(2)+'": '+inttostr(FileError));

    Write_XLS_to_Data();

    close(XMLout);

  end else begin

    XMLin:=TXML.Create();
    if XMLin = nil then ErrorExit('Internal error: XML create');

    if not XMLin.Open_Read(paramstr(1))
     then ErrorExit_InXML(XMLin.LastErrorMsg);

    if not XLS.Open_Write(paramstr(2))
     then ErrorExit(XLS.LastErrorMsg);

    Write_Data_to_XLS();

    XMLin.Destroy();

  end;

  XLS.Destroy();

end; { RecognizeParameters }

procedure Otladka;
 begin
  {   XMLin.ReadKey();
   writeln(XMLin.Key_Value);
   writeln(ansistring(XMLin.Key_Value));
   writeln(ansistring(XMLtoWide(XMLin.Key_Value)));
   writeln(AnsiUpperCase(XMLin.Key_Value));
   writeln(AnsiUpperCase(XMLtoWide(XMLin.Key_Value)));
   exit;{}
 end;

begin SetConsoleOutputCP(CP_UTF8);

  RecognizeParameters();

end.
