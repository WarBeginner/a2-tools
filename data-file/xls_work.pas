{$IFDEF FPC}
 {$MODE Delphi}
{$ELSE}
 {$APPTYPE CONSOLE}
{$ENDIF}

unit xls_work;

interface

//uses sysutils;
uses xml_work;

const
   MaxRow = $10000; MaxCol = $10000;
   _UndefinedString : TXMLString = '';

type

   TRowsSelection = array of longint;

   TXLSSheetID = integer;
   TXMLData = record
     Text_Value : TXMLString;
     Key_Value : TXMLString;
     Header_Value : TXMLString;
     Footer_Value : TXMLString;
   end;
   TCell = record
     XMLData : TXMLData; //<Data>
     Value : TXMLString;
   end;
   TCells = array of TCell;
   TRow = record
     XMLData : TXMLData; //<Row>
     Cells : TCells;
   end;
   TRows = array of TRow;
   TSheet = record
     Name : TXMLString;
     XMLData : TXMLData; //<Table>
     Rows : TRows;
   end;
   TXLS = class
     Private
       XMLData : TXMLData; //<Worksheet>
       Sheets : array of TSheet;
       _ErrorMsg : widestring;
       function OpenFile(XML : TXML) : boolean;
     Public
       FileName : TXMLString;
       constructor Create();
       destructor Destroy();
       function GetXLSSheetID(SheetName : TXMLString; Create : boolean) : TXLSSheetID;
       function SheetHeight(SheetID : TXLSSheetID) : integer;
       function RowWidth(SheetID: TXLSSheetID; Row : integer): integer;
       function Open_Read(FileName : widestring) : boolean;
       function Open_Write(FileName : widestring) : boolean;
       function SaveFile() : boolean;
       property LastErrorMsg : widestring read _ErrorMsg;
       function WriteCellData(SheetID : TXLSSheetID; Row, Col : integer; CellValue,Typ : TXMLString) : boolean;
       function ReadCellData(SheetID : TXLSSheetID; Row, Col : integer) : TXMLString;
       function SelectRows(SheetID : TXLSSheetID; Col : integer; CellValue : TXMLString; SelectEqual : boolean) : TRowsSelection;
       function TruncateRow(SheetID: TXLSSheetID; Row, Col : integer) : boolean;
       function TruncateTable(SheetID : TXLSSheetID; Row : integer) : boolean;

   end;

implementation

uses sysutils;

////////////////////////////////////////////////
// вспомогашки

function GetDefaultTextValue(TextValue : TXMLString) : TXMLString;
 var I1,I2 : integer;
 begin
   I1:=Length(TextValue);
   if I1=0 then begin result:=''; exit end;
   while (I1 > 0) and (TextValue[I1] > ' ') do dec(I1);
   inc(I1);
   I2:=I1;
   while (I2<=Length(TextValue)) and (TextValue[I2] <= ' ') do inc(I2);
   result:=#13#10+copy(TextValue,I1,I2-I1);
 end; { GetDefaultTextValue }

/////////////////////////////////////////////////

function TXLS.SelectRows(SheetID : TXLSSheetID; Col : integer; CellValue : TXMLString; SelectEqual : boolean) : TRowsSelection;
 var Row,L : integer;
     Empty : boolean;
 begin with Sheets[SheetID] do begin
   SelectEqual := not SelectEqual;
   result:=nil; L:=0;
   for Row:=3 to Length(Rows) do with Rows[Row-1] do begin
     Empty := (Col>Length(Cells));
     if  ((Empty and (CellValue='')) xor SelectEqual)
      or ((not Empty and (Cells[Col-1].Value=CellValue)) xor SelectEqual)
      then begin
       inc(L);
       SetLength(result,L);
       result[L-1]:=Row;
     end;
   end;
 end end; { TXLS.SelectRows }

////////////////////////////////////////////////


constructor TXLS.Create();
begin
  (**)(**)(**)
end;

destructor TXLS.Destroy;
var S : integer;
begin
  for S:=0 to High(Sheets)
   do self.TruncateTable(S,0);
  Sheets:=nil;
end;

/////////////////////////////////////////////////////////////////////
// Чтение

procedure AddRows(var Rows : TRows; Sheet : TSheet; N : integer; SetLastRowStrings : boolean);
var RN,RMax : integer;
begin
  if N<=0 then exit;
  RN:=High(Rows); RMax:=RN+N;
  SetLength(Rows,RMax+1);
  if SetLastRowStrings = false then dec(RMax);
  while RN < RMax do begin
    inc(RN);
    with Rows[RN] do begin
      Cells:=nil;
      if RN = 0
       then XMLData.Text_Value := GetDefaultTextValue(Sheet.XMLData.Text_Value)+' '
       else XMLData.Text_Value := GetDefaultTextValue(Rows[RN-1].XMLData.Text_Value);
      XMLData.Key_Value := '';
      XMLData.Header_Value := '';
      XMLData.Footer_Value := XMLData.Text_Value+'</Row>';
    end;
  end;
end; { AddRows }

procedure AddCells(var Cells : TCells; Row : TRow; N : integer; SetLastCellStrings : boolean);
var CN,CMax : integer;
begin
  if N<=0 then exit;
  CN:=High(Cells); CMax:=CN+N;
  SetLength(Cells,CMax+1);
  if SetLastCellStrings = false then dec(CMax);
  while CN < CMax do begin
    inc(CN);
    with Cells[CN] do begin
      if CN = 0
       then XMLData.Text_Value := GetDefaultTextValue(Row.XMLData.Text_Value)+' '
       else XMLData.Text_Value := GetDefaultTextValue(Cells[CN-1].XMLData.Text_Value);
      XMLData.Key_Value := '';
      XMLData.Header_Value := '';
      XMLData.Footer_Value := '</Data></Cell>';
      Cells[CN].Value := '';
    end;
  end;
end; { AddCells }

function TXLS.OpenFile(XML : TXML) : boolean;
label StartReadRows;
var
    SN,RN : integer;
    RIndex : TXMLString;
    T : TXMLString;
    E : integer;
 procedure ReadCells(var Cells : TCells; var Row : TRow; var XML : TXML);
  var CN : integer;
      CIndex : TXMLString;
      E : integer;
  begin
    Cells:=nil; CN:=0;
    while XML.FindKey('Cell',true,true) do begin
      CIndex := GetXMLAttribute(XML.Key_Value,'ss:Index');
      if CIndex = ''
       then inc(CN)
       else begin
         val(CIndex,CN,E);
         if (E<>0) or (CN <= Length(Cells)) then begin
           _ErrorMsg := 'Bad cell number "'+XMLtoWide(CIndex);//+'" in rows worksheet "'+Name+'"';
           exit;
         end;
      end;
      AddCells(Cells,Row,CN-Length(Cells),false);
      with Cells[CN-1] do begin
        XMLData.Text_Value := XML.Text_Value;
        XMLData.Header_Value := '';
        XMLData.Key_Value := XML.Key_Value;
        XMLData.Footer_Value := '';
        if XML.FindKey('Data',true,true) then begin
          XMLData.Header_Value := XMLData.Header_Value + XML.Text_Value+XML.Key_Value;
          XML.FindKey('Data',false,false);
          Value := XML.Text_Value;
          XMLData.Footer_Value := XMLData.Footer_Value + XML.Key_Value;
        end else begin
          XMLData.Footer_Value := XMLData.Footer_Value + XML.Text_Value+XML.Key_Value;
        end;
        if (XML.Key_Name <> 'CELL') or (XML.KeyIsClosed = false) then begin
          XML.FindKey('Cell',false,false);
          XMLData.Footer_Value := XMLData.Footer_Value + XML.Text_Value+XML.Key_Value;
        end;
      end;
    end;
  end; { ReadCells }
begin { TXLS.Open }
   result:=false; _ErrorMsg := '';
   self.XMLData.Text_Value := '';
   self.XMLData.Key_Value := '';
   self.XMLData.Header_Value := '';
   self.XMLData.Footer_Value := '';
   while XML.FindKey('Worksheet',false,false) do begin
     SN:=Length(Sheets);
     inc(SN);
     SetLength(Sheets,SN);
     with Sheets[SN-1] do begin
       Name := GetXMLAttribute(XML.Key_Value,'ss:Name');
       if Name = ''
        then begin _ErrorMsg := 'Worksheet #'+inttostr(SN)+' is unnamed: '+XMLtoWide(XML.Text_Value); exit end;
       XMLData.Text_Value := XML.Text_Value;
       XMLData.Key_Value := XML.Key_Value;
       XMLData.Header_Value := '';
       XMLData.Footer_Value := '';

       if not XML.FindKey('Table',true,true)
        then begin _ErrorMsg := 'Key <Table> in worksheet "'+XMLtoWide(Name)+'" not found'; exit end;
       XMLData.Header_Value := XMLData.Header_Value + XML.Text_Value + XML.Key_Value;

       if XML.FindKey('Row',true,true)
        then XMLData.Header_Value := XMLData.Header_Value + XML.Text_Value;
       T:='';
       Rows:=nil; RN:=0;
       goto StartReadRows;
       repeat
         XML.FindKey('Row',true,true);
         T:=XML.Text_Value;
      StartReadRows:
         if (XML.Key_Name = 'ROW') and (XML.KeyIsClosed = false) then begin
           RIndex := GetXMLAttribute(XML.Key_Value,'ss:Index');
           if RIndex = ''
            then inc(RN)
            else begin
              val(RIndex,RN,E);
              if (E<>0) or (RN <= Length(Rows)) then begin
                _ErrorMsg := 'Bad row number "'+XMLtoWide(RIndex)+'" in rows worksheet "'+XMLtoWide(Name)+'"';
                exit;
              end;
           end;
           AddRows(Rows,Sheets[SN-1],RN-Length(Rows),false);
           Rows[RN-1].XMLData.Text_Value := T;
           Rows[RN-1].XMLData.Key_Value := XML.Key_Value;
           Rows[RN-1].XMLData.Header_Value := '';
           ReadCells(Rows[RN-1].Cells,Rows[RN-1],XML);
           Rows[RN-1].XMLData.Footer_Value := XML.Text_Value+XML.Key_Value;
         end else if (XML.Key_Name = 'TABLE') and (XML.KeyIsClosed = true) then begin
           break;
         end else begin
           _ErrorMsg := 'Unknown key "'+XMLtoWide(XML.Key_Name)+'" in rows worksheet "'+XMLtoWide(Name)+'"';
           exit;
         end;
       until false;
       XMLData.Footer_Value := XMLData.Footer_Value + XML.Text_Value+XML.Key_Value;
       XML.FindKey('Worksheet',false,false);
       XMLData.Footer_Value := XMLData.Footer_Value + XML.Text_Value+XML.Key_Value;
     end;
   end;
   repeat
     XMLData.Footer_Value := XMLData.Footer_Value + XML.Text_Value+XML.Key_Value;
   until not XML.ReadKey();
   XMLData.Footer_Value := XMLData.Footer_Value + XML.Text_Value+XML.Key_Value;
   if Length(Sheets) = 0 then begin _ErrorMsg := 'No XLS sheets found'; exit end;
   result:=true;
end; { TXLS.OpenFile }

function TXLS.Open_Read(FileName: widestring): boolean;
var XML : TXML;
begin
   result:=false; _ErrorMsg := '';
   self.FileName := FileName;
   XML:=TXML.Create();
   if XML = nil
    then begin _ErrorMsg := 'Internal error: XML create'; exit end;
   if not XML.Open_Read(FileName)
    then _ErrorMsg := XML.LastErrorMsg
    else result := self.OpenFile(XML);
   XML.Destroy();
end; { TXLS.Open_Read }

function TXLS.Open_Write(FileName: widestring): boolean;
var XML : TXML;
begin
   result:=false; _ErrorMsg := '';
   self.FileName := FileName;
   XML:=TXML.Create();
   if XML = nil
    then begin _ErrorMsg := 'Internal error: XML create'; exit end;
   if not XML.Open_Write(FileName)
    then _ErrorMsg := XML.LastErrorMsg
    else result := self.OpenFile(XML);
   XML.Destroy();
end; { TXLS.Open_Write }

function TXLS.GetXLSSheetID(SheetName: TXMLString; Create : boolean): TXLSSheetID;
var
  UpCaseName : widestring;
  SN : integer;
begin{$B-}
  for result:=0 to High(Sheets) do if Sheets[result].Name = SheetName then exit;
  UpCaseName:=AnsiUpperCase(XMLtoWide(SheetName));
  for result:=0 to High(Sheets) do if AnsiUpperCase(XMLtoWide(Sheets[result].Name)) = UpCaseName then exit;
  if Create = false then begin
    result:=-1; exit;
  end;
  SN:=Length(Sheets);
  SetLength(Sheets,SN+1);
  with Sheets[SN] do begin
    Name := SheetName;
    Rows:=nil;
    if SN = 0
     then XMLData.Text_Value := #13#10
     else XMLData.Text_Value := GetDefaultTextValue(Sheets[SN-1].XMLData.Text_Value);
    XMLData.Key_Value := '<Worksheet ss:Name="'+Name+'">';
    XMLData.Header_Value := XMLData.Text_Value+' <Table>';
    XMLData.Footer_Value := XMLData.Text_Value+' </Table>'+XMLData.Text_Value+'</Worksheet>';
  end;
  result:=SN;
end; { TXLS.GetXLSSheetID }

function TXLS.SheetHeight(SheetID: TXLSSheetID): integer;
begin
  with Sheets[SheetID] do begin
    result:=Length(Rows);
  end;
end; { TXLS.SheetHeight }

function TXLS.RowWidth(SheetID: TXLSSheetID; Row : integer): integer;
begin
  result:=-1;
  with Sheets[SheetID] do begin
    if Row > Length(Rows) then exit;
    with Rows[Row-1] do begin
      result:=Length(Cells);
    end;
  end;
end; { TXLS.RowWidth }

function TXLS.ReadCellData(SheetID: TXLSSheetID; Row, Col: integer): TXMLString;
begin
  result:=_UndefinedString;
  with Sheets[SheetID] do begin
    if Row > Length(Rows) then exit;
    with Rows[Row-1] do begin
      if Col > Length(Cells) then exit;
      with Cells[Col-1] do begin
        result:=Value;
      end;
    end;
  end;
end; { TXLS.ReadCellData }

///////////////////////////////////////////////////////////////////////////////////
// Изменение

procedure OpenKeyToSetValue(var XMLData : TXMLData);
 var L,KL : integer;
 begin with XMLData do begin
   if Footer_Value<>'' then exit;

   L:=Length(Text_Value);
   KL:=1; while (KL<=L) and (Text_Value[KL]<=' ') do inc(KL);
   Footer_Value :=copy(Text_Value,1,KL-1);

   L:=Length(Key_Value);
   if L<4 then exit; //<a/>
   if Key_Value[L-1] = '/' then Key_Value:=copy(Key_Value,1,L-2)+Key_Value[L];

   KL:=2; while (KL<L) and (Key_Value[KL]>' ') do inc(KL);
   Footer_Value :=Footer_Value + '</'+copy(Key_Value,2,KL-2)+'>';

 end end; { OpenKeyToSetValue }

function TXLS.WriteCellData(SheetID : TXLSSheetID; Row, Col: integer; CellValue,Typ : TXMLString): boolean;
begin
  result:=false;
  if (Row<=0) or (Col<=0) or (Row>=MaxRow) or (Col>=MaxCol) then begin
    _ErrorMsg := 'Bad cell addr: '+inttostr(Row)+':'+inttostr(Col);
    exit;
  end;
  with Sheets[SheetID] do begin
    AddRows(Rows,Sheets[SheetID],Row-Length(Rows),true);
    with Rows[Row-1] do begin
      if XMLData.Key_Value=''
       then XMLData.Key_Value := '<Row>';
      OpenKeyToSetValue(XMLData);
      AddCells(Cells,Rows[Row-1],Col-Length(Cells),true);
      with Cells[Col-1] do begin
        if XMLData.Key_Value=''
         then XMLData.Key_Value := '<Cell>';
        OpenKeyToSetValue(XMLData);
        XMLData.Header_Value := '<Data ss:Type="'+Typ+'">';
        Value := CellValue;
      end;
    end;
  end;
  result:=true;
end; { TXLS.WriteCellData }


function TXLS.TruncateRow(SheetID: TXLSSheetID; Row, Col : integer): boolean;
begin
  result:=false;
  if (Row<=0) or (Col<0) or (Row>=MaxRow) or (Col>=MaxCol) then begin
    _ErrorMsg := 'Bad cell addr: '+inttostr(Row)+':'+inttostr(Col);
    exit;
  end;
  with Sheets[SheetID]
   do if Row <= Length(Rows) then with Rows[Row-1] do begin
    if Col = 0
     then Cells:=nil
     else if Col > Length(Cells)
           then SetLength(Cells,Col);
  end;
  result:=true;
end; { TXLS.TruncateRow }

function TXLS.TruncateTable(SheetID : TXLSSheetID; Row : integer) : boolean;
var R : integer;
begin
  result:=false;
  if (Row<0) or (Row>=MaxRow) then begin
    _ErrorMsg := 'Bad row num: '+inttostr(Row);
    exit;
  end;
  with Sheets[SheetID] do begin
    R:=High(Rows);
    while R>=Row do with Rows[R] do begin
      Cells:=nil;
      dec(R);
    end;
    if Row = 0
     then Rows:=nil
     else if Row > Length(Rows)
           then SetLength(Rows,Row);
    result:=true;
  end
end; { TXLS.TruncateTable }


function TXLS.SaveFile(): boolean;
var F : text;
    E,SN,RN,CN,WritedRN,WritedCN : integer;
begin{$I-}
  result:=false;
  assign(F,FileName); rewrite(F);
  E:=ioresult();
  if E<>0 then begin
    _ErrorMsg := 'Error create file "'+FileName+'": '+inttostr(E);
    exit;
  end;
  write(F,XMLData.Text_Value);
  write(F,XMLData.Header_Value);
  for SN:=0 to High(Sheets) do with Sheets[SN] do begin

    (**)//заплатки
    XMLData.Header_Value := DeleteXMLAttribute(XMLData.Header_Value,'ss:ExpandedRowCount');
    XMLData.Header_Value := DeleteXMLAttribute(XMLData.Header_Value,'ss:ExpandedColumnCount');

    write(F,XMLData.Text_Value);
    write(F,XMLData.Key_Value);
    write(F,XMLData.Header_Value);
    WritedRN:=0;
    for RN:=0 to High(Rows) do with Rows[RN] do begin
      if XMLData.Key_Value = ''
       then continue;
      write(F,XMLData.Text_Value);
      //
      if WritedRN = RN
       then XMLData.Key_Value := DeleteXMLAttribute(XMLData.Key_Value,'ss:Index')
       else XMLData.Key_Value := SetXMLAttribute(XMLData.Key_Value,'ss:Index',inttostr(RN+1));
      WritedRN := RN+1;
      write(F,XMLData.Key_Value);
      //
      write(F,XMLData.Header_Value);
      WritedCN:=0;
      for CN:=0 to High(Cells) do with Cells[CN] do begin
        if XMLData.Key_Value = ''
         then continue;
        write(F,XMLData.Text_Value);
        //
        if WritedCN = CN
         then XMLData.Key_Value := DeleteXMLAttribute(XMLData.Key_Value,'ss:Index')
         else XMLData.Key_Value := SetXMLAttribute(XMLData.Key_Value,'ss:Index',inttostr(CN+1));
        WritedCN := CN+1;
        write(F,XMLData.Key_Value);
        //
        write(F,XMLData.Header_Value);
        write(F,Value);
        write(F,XMLData.Footer_Value);
      end;
      write(F,XMLData.Footer_Value);
    end;
    write(F,XMLData.Footer_Value);
  end;
  write(F,XMLData.Footer_Value);
  close(F);
end; { TXLS.SaveFile }

end.

Структура XML-файла начинается с корневого элемента Workbook, обозначающего книгу. Перейдем к его дочерним элементам.

    DocumentProperties (Свойства документа). В этом элементе не содержится никаких данных из таблицы, поэтому этот элемент мы рассматривать не будем.
    ExcelWorkbook (Книга Excel). Также в этом элементе не содержится важной для нас информации, поэтому также пропускаем.
    Styles (Стили). Содержит форматирование таблицы. Расмотрим этот элемент вкратце. Будем считать что нам важнее сами данные, чем их форматирование.
    Worksheet (Лист). Данных элементов может быть несколько, в зависимости от количества листов в книге. Отобрать данные элементы можно с помощью XML-метода getElementsByTagName.

WorkSheet

Обязательные параметры:

    ss:Name (Название листа).

Необязательные параметры:

    ss:Protected (Информация о защите листа)
    ss:RightToLeft (Направление текста)

Дочерние элементы:

    Table (Таблица). Данные
    WorksheetOptions (Настройки). Не содержит данных. рассматриваться не будет.

Table

Обязательные параметры: нет

Необязательные параметры:

    ss:DefaultColumnWidth (Ширина столбцов по умолчанию). Указывается в pt (1pt = 4/3px)
    ss:DefaultRowHeight (Высота строк по умолчанию). Указывается в pt (1pt = 4/3px)
    ss:ExpandedColumnCount (Общее число столбцов в этой таблице)
    ss:ExpandedRowCount (Общее число строк в этой таблице)
    ss:LeftCell (Начало таблицы слева)
    ss:StyleID (Стиль таблицы). Ссылается на элемент Styles (подробнее ниже)
    ss:TopCell (Начала таблицы сверху)

Дочерние элементы:

    Column (Столбцы)
    Row (Строки)

Column

Обязательные параметры: нет

Необязательные параметры:

    c:Caption (Заголовок)
    ss:AutoFitWidth (Автоматическая ширина столбца). Истина если содержит значение 1.
    ss:Hidden (Признак скрытия столбца)
    ss:Index (Индекс столбца)
    ss:Span (Количество столбцов с одинаковым форматированием)
    ss:StyleID (Стиль столбца)
    ss:Width (Ширина столбца). Указывается в pt (1pt = 4/3px)

Остановимся подробнее на параметрах ss:Index и ss:Span. Например, имеется 5 столбцов:

    Ширина 100pt
    Ширина 20pt
    Ширина 20pt
    Ширина 20pt
    Ширина 50pt

В XML-файле столбцы должны быть описаны следующим образом:

<column ss:width="100"/>
<column ss:span="2" ss:width="20"/>
<column ss:index="5" ss:width="50"/>

Row

Обязательные параметры: нет

Необязательные параметры:

    c:Caption (Заголовок)
    ss:AutoFitWidth (Автоматическая высота строки). Истина если содержит значение 1.
    ss:Height (Высота строки). Указывается в pt (1pt = 4/3px)
    ss:Hidden (Признак скрытия строки)
    ss:Index (Индекс строки)
    ss:Span (Количество строк с одинаковым форматированием)
    ss:StyleID (Стиль строки)

Дочерние элементы:

    Cell (Ячейка)

Cell

Обязательные параметры:

    ss:Type (Тип ячейки). Возможные значения:

        Number (Числовой)
        DateTime (Дата и время)
        Boolean (Логический)
        String (Строковый)
        Error (Ошибка). Возможные значения:
            #NULL!
            #DIV/0!
            #VALUE!
            #REF!
            #NAME?
            #NUM!
            #N/A
            #CIRC!

Необязательные параметры: нет

Дочерние элементы:

    B (Жирным)
    Font (Шрифт)
    I (Курсив)
    S (Зачеркнутый)
    Span (Форматированный)
    Sub (Верхний регистр)
    Sup (Нижний регистр)
    U (Подчеркивание)

Значение элемента: Значение ячейки