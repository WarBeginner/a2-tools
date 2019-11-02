unit data_structure;

interface

uses xml_work, xls_work;

const
  MaxTable = 18; MaxField = 67;
type
  TTableNum = byte;
  TFieldNum = byte;
  TField = record
    KeyName,
    KeyValue,
    ColName,
    Typ,
    NeedParentValue
                : TXMLString;
    Col : longint;
    ParentField : longint;
    TableNum : TTableNum;
  end;
  TFields = array [1..MaxField] of TField;
  TTable = record
    TableKeyPath,
    TableKeyValue,
    TableKeyName,
    RowsKeyValue,
    FieldTable_KeyNames,
    FieldsKeyValue,
    XLSSheetName
                  : TXMLString;
    RowInStructure : longint;
    Count : longint;
    ItemVersion : longword;
    Height,Width : longint;
    XLSSheetID : TXLSSheetID;
    Fields : TFields;
    AllXMLDataFilled : boolean;
    TableKeyIsSimple,RowsKeyIsSimple : boolean;
  end;
var
  Tables : array [0..MaxTable] of TTable;

procedure CreateTables();

procedure FillField(var Table : TTable; FieldNum : TFieldNum);

function FindTableNum(TablePath : TXMLString) : TTableNum;

function FindFieldNum(var Fields : TFields; var Width : longint; var Table : TTable; FieldName,FieldKeyValue,ColumnName : TXMLString; CreateAfterField : longint) : TFieldNum;


implementation

procedure CreateTablesData();
 begin
   //'>BOOST_SERIALIZATION>DATA_BIN'
   with Tables[0] do begin
     TableKeyName := '';
     TableKeyPath := '';
     XLSSheetName := WideToXML('Структура');
   end;
   with Tables[1] do begin
     TableKeyName := 'CLASSES';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>CLASSES';
     XLSSheetName := 'classes';
   end;
   with Tables[2] do begin
     TableKeyName := 'MATERIALS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>MATERIALS';
     XLSSheetName := 'materials';
   end;
   with Tables[3] do begin
     TableKeyName := 'MODIFIERS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>MODIFIERS';
     XLSSheetName := 'modifiers';
   end;
   with Tables[4] do begin
     TableKeyName := 'ARMOURS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>ARMOURS';
     XLSSheetName := 'armours';
     FieldTable_KeyNames := '<SHAPESALLOWED>';
   end;
   with Tables[5] do begin
     TableKeyName := 'SHAPESALLOWED';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>ARMOURS>ITEM>SHAPESALLOWED';
     XLSSheetName := 'shapesAllowed_armours';
   end;
   with Tables[6] do begin
     TableKeyName := 'SHIELDS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>SHIELDS';
     XLSSheetName := 'shields';
     FieldTable_KeyNames := '<SHAPESALLOWED>';
   end;
   with Tables[7] do begin
     TableKeyName := 'SHAPESALLOWED';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>SHIELDS>ITEM>SHAPESALLOWED';
     XLSSheetName := 'shapesAllowed_shields';
   end;
   with Tables[8] do begin
     TableKeyName := 'WEAPONS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>WEAPONS';
     XLSSheetName := 'weapons';
     FieldTable_KeyNames := '<SHAPESALLOWED>';
   end;
   with Tables[9] do begin
     TableKeyName := 'SHAPESALLOWED';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>WEAPONS>ITEM>SHAPESALLOWED';
     XLSSheetName := 'shapesAllowed_weapons';
   end;
   with Tables[10] do begin
     TableKeyName := 'MAGICITEMS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>MAGICITEMS';
     XLSSheetName := 'magicItems';
   end;
   with Tables[11] do begin
     TableKeyName := 'MONSTERS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>MONSTERS';
     XLSSheetName := 'monsters';
     FieldTable_KeyNames := '<KNOWNSPELLS>';
   end;
   with Tables[12] do begin
     TableKeyName := 'KNOWNSPELLS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>MONSTERS>ITEM>KNOWNSPELLS';
     XLSSheetName := 'knownSpells_monsters';
   end;
   with Tables[13] do begin
     TableKeyName := 'HUMANS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>HUMANS';
     XLSSheetName := 'humans';
     FieldTable_KeyNames := '<KNOWNSPELLS>';
   end;
   with Tables[14] do begin
     TableKeyName := 'KNOWNSPELLS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>HUMANS>ITEM>KNOWNSPELLS';
     XLSSheetName := 'knownSpells_humans';
   end;
   with Tables[15] do begin
     TableKeyName := 'BUILDINGS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>BUILDINGS';
     XLSSheetName := 'buildings';
     FieldTable_KeyNames := '<CANNOTPASS>, <TILES>';
   end;
   with Tables[16] do begin
     TableKeyName := 'CANNOTPASS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>BUILDINGS>ITEM>CANNOTPASS';
     XLSSheetName := 'canNotPass';
   end;
   with Tables[17] do begin
     TableKeyName := 'TILES';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>BUILDINGS>ITEM>TILES';
     XLSSheetName := 'tiles';
   end;
   with Tables[18] do begin
     TableKeyName := 'SPELLS';
     TableKeyPath := '>BOOST_SERIALIZATION>DATA_BIN>SPELLS';
     XLSSheetName := 'spells';
   end;

 end; { CreateTablesData }

procedure CreateFieldsData();
 {uses: Tables}
 begin

   with Tables[0] do begin
     RowInStructure := -1;
     Width := 4;
     with Fields[1] do begin
       KeyName := WideToXML('Тип');
       ColName := WideToXML('Расширение');
       Col     := 1;
     end;
     with Fields[2] do begin
       KeyName := WideToXML('Значение');
       ColName := WideToXML('Ключ таблицы');
       Typ := 'String';
       Col     := 2;
     end;
     with Fields[3] do begin
       KeyName := WideToXML('Имя листа');
       ColName := WideToXML('Ключ строки');
       Typ := 'String';
       Col     := 3;
     end;
     with Fields[4] do begin
       KeyName := WideToXML('Поля');
       ColName := WideToXML('Ключи полей');
       Typ := 'String';
       Col     := 4;
     end;
   end;

 end; { CreateFieldsData }

procedure CreateTables();
 {uses: Tables}
var TN : TTableNum;
begin
  CreateTablesData();
  for TN:=0 to MaxTable do with Tables[TN] do begin
    Height := 2;
    Width := 0;
    AllXMLDataFilled:=false;
  end;
  CreateFieldsData();
end; { CreateTables }

//////////////////////////////////
//procedure FillField(var Table : TTable; FieldNum : TFieldNum);
//Typ
//

procedure FillField_Typ(var Table : TTable; FieldNum : TFieldNum);
begin with Table.Fields[FieldNum] do begin
  if   (Typ = 'String')
    or (Typ = 'Number')
   then exit;
  if  (KeyName = '[VALUE]')
   or (KeyName = 'NAME')
   or (KeyName = 'EFFECT')
   or (KeyName = 'SPELL1')
   or (KeyName = 'SPELL2')
   or (KeyName = 'SPELL3')
   or (KeyName = 'EQUIPITEM1')
   or (KeyName = 'EQUIPITEM2')
   or (KeyName = 'EQUIPITEM3')
   or (KeyName = 'EQUIPITEM4')
   or (KeyName = 'EQUIPITEM5')
   or (KeyName = 'EQUIPITEM6')
   or (KeyName = 'EQUIPITEM7')
   or (KeyName = 'EQUIPITEM8')
   or (KeyName = 'EQUIPITEM9')
   or (KeyName = 'EQUIPITEM10')
   or (KeyName = 'PARENTID')
   or (KeyName = 'SHAPESALLOWED')
   or (KeyName = 'KNOWNSPELLS')
   or (KeyName = 'CANNOTPASS')
   or (KeyName = 'TILES')
   or (Table.TableKeyName = 'KNOWNSPELLS')
    and (   (KeyName = 'FIRST')
         or (KeyName = 'SECOND')
        )
   or (Table.TableKeyName = 'SHAPESALLOWED')
    and (   (KeyName = 'ITEM')
         or (KeyName = 'FIRST')
         or (KeyName = 'SECOND')
        )
   then Typ := 'String'
   else Typ := 'Number';
end end; { FillField_Typ }

procedure FillField_ParentField(var Table : TTable; FieldNum : TFieldNum);
begin with Table.Fields[FieldNum] do begin
  ParentField:=0;
  NeedParentValue:='';
  if ( (Table.TableKeyName = 'MONSTERS')
    or (Table.TableKeyName = 'HUMANS')
    or (Table.TableKeyName = 'SPELLS')
      )
   and (KeyName <> 'EMPTY')
   then begin
          ParentField:=FindFieldNum(Table.Fields,Table.Width,Table,'EMPTY','','',-1);
          NeedParentValue :='0';
        end
  else if (KeyName = 'KNOWNSPELLS')
   then begin
          ParentField:=FindFieldNum(Table.Fields,Table.Width,Table,'KNOWNSPELLSAREEMPTY','','',-1);
          NeedParentValue :='0';
        end;
end end; { FillField_ParentField }

procedure FillField(var Table : TTable; FieldNum : TFieldNum);
begin
  FillField_Typ(Table,FieldNum);
  FillField_ParentField(Table,FieldNum);
end; { FillField }

function FindTableNum(TablePath : TXMLString) : TTableNum;
 {uses: Tables, MaxTable}
 begin
   Tables[0].TableKeyPath := TablePath;
   result:=MaxTable; while Tables[result].TableKeyPath <> TablePath do dec(result);
   Tables[0].TableKeyPath := '';
 end; { FindTableNum }

function FindFieldNum(var Fields : TFields; var Width : longint; var Table : TTable; FieldName,FieldKeyValue,ColumnName : TXMLString; CreateAfterField : longint) : TFieldNum;
 {uses: MaxField}
 var FN : TFieldNum;
 begin
   for FN:=1 to Width do if Fields[FN].KeyName = FieldName then begin
     result:=FN;
     if Fields[FN].Typ = ''
      then FillField(Table,FN);
     exit;
   end;
   result:=0;
   if (CreateAfterField < 0) or (Width = MaxField)
    then exit;
   inc(Width);
   inc(CreateAfterField);
   if CreateAfterField < Width then begin
     {FN:=Width;
     while FN>CreateAfterField do begin
       Fields[FN]:=Fields[FN-1];
       dec(FN);
     end;{}
     (**)(**)(**)
     move(Fields[CreateAfterField],Fields[CreateAfterField+1],sizeof(TField)*(Width-CreateAfterField));
     //это нужно, т.к. образуются копии ссылок (TXTMLString=ansistring),
     //  о которых дельфа ничего не знает:
     FillChar(Fields[CreateAfterField],sizeof(TField),0);
   end;
   with Fields[CreateAfterField] do begin
     KeyName := FieldName;
     KeyValue := FieldKeyValue;
     ColName := ColumnName;
     Col := Width;
     Typ := '';
   end;
   FillField(Table,CreateAfterField);
   result:=CreateAfterField;
 end; { FindFieldNum }

end.
