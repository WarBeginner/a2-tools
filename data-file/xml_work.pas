{$IFDEF FPC}
 {$MODE Delphi}
{$ELSE}
 {$APPTYPE CONSOLE}
{$ENDIF}

unit xml_work;

interface

//uses sysutils;
const XML_BufSize = 1024;
type
   TXMLString = ansistring;
   TXMLBuf = array [0..XML_BufSize-1] of byte;
   TChar = string[15];
   TKeyAttr = record
      NamePos,NameLen : longint;
      ValuePos,ValueLen : longint;
   end;
   TKeyAttributes = array of TKeyAttr;
   TXML = class
     Private
       _F : File;
       _FileIsOpen : boolean;
       _ErrorMsg : widestring;
       _DataBuf  : TXMLBuf;
       _DataPos,_BufSize : longword;
       _Char : TChar;
       _KeyIsClosed : boolean;
       _Key_Path  : TXMLString; //список вложенности всех КЛЮЧЕЙ от верхнего до текущего, перед каждым ключем разделитель ">"
       _Key_Name  : TXMLString; //имя КЛЮЧА (от "<" до непечатного, пробела или ">")
       _Key_Value  : TXMLString; //полное содержимое ключа (от "<" до ">" включительно)
       _Text_Value : TXMLString; //текст между этим ключем и предыдущим
       _KeyAttributes : TKeyAttributes; // указатели на атрибуты в _Key_Value
       function ReadBuf() : boolean; //обеспечивает чтением непустой буфер
       function ReadChar(InText : boolean) : boolean; //читает один символ (сворачивает особые)
       function ReadChar_UTF8 : TChar;
       function OpenFile(FileName : widestring) : boolean;
     Public
       constructor Create();
       destructor Destroy();
       function Open_Read(FileName : widestring) : boolean; // открывает файл
       function Open_Write(FileName : widestring) : boolean;
       function ReadKey() : boolean; // читает очередной ключ (выделяет места атрибутов)
       function ReadNeededKey(KeyName : TXMLString) : boolean; // читает очередной ключ и сравнивает имя
       function FindKey(KeyName : TXMLString; PathMustNoUp,PathMustEqual : boolean) : boolean; // ищет очередной ключ
       function FindPath(KeyPath : TXMLString) : boolean; // ищет очередное совпадение списка вложенности ключей

       property LastErrorMsg : widestring read _ErrorMsg; // последняя ошибка
       property Key_Path : TXMLString read _Key_Path;
       property Key_Name : TXMLString read _Key_Name;
       property KeyIsClosed : boolean read _KeyIsClosed;
       property Key_Value : TXMLString read _Key_Value;
       property Text_Value : TXMLString read _Text_Value;
       //property Key_Attr[Name : TXMLString] : ansistring read GetKeyAttr write SetKeyAttr;
     end;

function XMLtoWide(XMLString : TXMLString) : widestring;
function WideToXML(WString : widestring) : TXMLString;

function XMLtoXMLText(XMLString : TXMLString) : TXMLString;
function XMLTextToXML(XMLText : TXMLString) : TXMLString;

procedure FindXMLAttr(KeyValue,AttrName: TXMLString; var NamePos,ValueStartPos,ValueOverPos : integer);
function GetXMLAttribute(KeyValue,AttrName : TXMLString) : TXMLString;
function SetXMLAttribute(KeyValue,AttrName,AttrValue : TXMLString) : TXMLString;
function DeleteXMLAttribute(KeyValue,AttrName : TXMLString) : TXMLString;
function DeleteXMLAllAttributes(KeyValue : TXMLString) : TXMLString;
function CloseXMLKey(KeyValue : TXMLString) : TXMLString;


implementation

uses
  sysutils;

//////////////////////////////////////////////////////////////
// Функции XML

function hexchar(B : byte) : widechar;
 begin
   if B<$A
    then result:=widechar(B+ord('0'))
    else result:=widechar(B+(ord('A')-10))
 end; { hexchar }

function hexstr(N : longword) : widestring;
 begin
   if N>$10 then result:=hexstr(N shr 4) else result:='';
   result:=result+hexchar(byte(N));
 end; { hexstr }

function XMLtoWide(XMLString : TXMLString) : widestring;
 var I : integer;
     B : byte;
     N : longword;
     R : widestring;
 begin
   I:=1; R:='';
   while I<=Length(XMLString) do begin
     B:=ord(XMLString[I]); inc(I); N:=B;
     if B >= $C0 then begin
       N:=(N shl 6 or (ord(XMLString[I]) and $3F)) and $7FF;
       inc(I);
     end;
     if B >= $E0 then begin
       N:=(N shl 6 or (ord(XMLString[I]) and $3F)) and $FFFF;
       inc(I);
     end;
     if B >= $F0 then begin
       N:=(N shl 6 or (ord(XMLString[I]) and $3F)) and $1FFFFF;
       inc(I);
       R:=R+'#0'+hexstr(N)+'h';
     end else begin
       R:=R+widechar(N);
     end;
   end;
   result:=R;
 end; { XMLtoWide }

function WideToXML(WString : widestring) : TXMLString;
 var I : integer;
     B : byte;
     N : longword;
     R,C : TXMLString;
 begin
   R:='';
   for I:=1 to Length(WString) do begin
     N:=ord(WString[I]) and $1FFFFF;
     if N<$80 then begin
       R:=R+chr(lo(N)); continue;
     end;
     B:=0; C:='';
     if N > $FFFF then begin
       B:=B shr 1 or $C0;
       C:=chr(lo(N) and $3F or $80)+C;
       N:=N shr 6;
     end;
     if N > $1FFF then begin
       B:=B shr 1 or $C0;
       C:=chr(lo(N) and $3F or $80)+C;
       N:=N shr 6;
     end;
     if N > $7F then begin
       B:=B shr 1 or $C0;
       C:=chr(lo(N) and $3F or $80)+C;
       N:=N shr 6;
     end;
     R:=R+(chr(B or lo(N))+C);
   end;
   result:=R;
 end; { WideToXML }

function XMLtoXMLText(XMLString : TXMLString) : TXMLString;
 var I : integer;
 begin
   result:='';
   for I:=1 to Length(XMLString) do case XMLString[I] of
     #9   : result:=result+'&#9;';
     #13  : result:=result+'&#13;';
     #10  : result:=result+'&#10;';
     '&'  : result:=result+'&amp;';
     '''' : result:=result+'&apos;';
     '<'  : result:=result+'&lt;';
     '>'  : result:=result+'&gt;';
     '"'  : result:=result+'&quot;';
     else result:=result+XMLString[I];
   end;
 end; { XMLtoXMLText }

function XMLTextToXML(XMLText : TXMLString) : TXMLString;
 var I,Code : integer;
     State : byte;
     _Char : char;
 begin
   result:='';
   State:=0;
   for I:=1 to Length(XMLText) do begin
     _Char := XMLText[I];
     case State of
     0  : //ничего не было
          if _Char <> '&'
             then begin result:=result+_Char; continue; end
             else State:=1;
     1  : //было &
          if      (_Char = 'q') then State:=11
          else if (_Char = 'l') then State:=21
          else if (_Char = 'g') then State:=31
          else if (_Char = 'a') then State:=101
          else if (_Char = '#') then begin State:=40; Code:=0; end
          else begin result:=result+'&q'+_Char; State := 0; end; // ошибка?
     11 ://было &q    собираем &quot;
          if      (_Char = 'u') then State:=12
          else begin result:=result+'&q'+_Char; State := 0; end; // ошибка?
     12 ://было &qu   собираем  &quot;
          if      (_Char = 'o') then State:=13
          else begin result:=result+'&qu'+_Char; State := 0; end; // ошибка?
     13 ://было &quo  собираем  &quot;
          if      (_Char = 't') then begin result:=result+'"'; State:=255 end
          else begin result:=result+'&quo'+_Char; State := 0; end; // ошибка?
     21 ://было &l    собираем  &lt;
          if      (_Char = 't') then begin result:=result+'<'; State:=255 end
          else begin result:=result+'&l'+_Char; State := 0; end; // ошибка?
     31 ://было &g    собираем  &gt;
          if      (_Char = 't') then begin result:=result+'>'; State:=255 end
          else begin result:=result+'&g'+_Char; State := 0; end; // ошибка?
     40 ://было &#    собираем  &#число;
          if      (_Char >= '0') and (_Char<='9') and (Code<($FF div 10)) then Code:=Code*10+ord(_Char)-ord('0')
          else if (_Char = ';') then begin result:=result+Chr(Code); State:=0 end
          else begin result:=result+'&#'+inttostr(Code)+_Char; State := 0; end; // ошибка?
     41 ://было &#    собираем  &#число;
          if      (_Char = ';') then begin result:=result+'>'; State:=255 end
          else begin result:=result+'&#'+_Char; State := 0; end; // ошибка?
     101://было &a
          if      (_Char = 'p') then State:=112
          else if (_Char = 'm') then State:=122
          else begin result:=result+'&a'+_Char; State := 0; end; // ошибка?
     112://было &ap   собираем  &apos;
          if      (_Char = 'o') then State:=113
          else begin result:=result+'&ap'+_Char; State := 0; end; // ошибка?
     113://было &apo  собираем  &apos;
          if      (_Char = 's') then begin result:=result+''''; State:=255 end
          else begin result:=result+'&apo'+_Char; State := 0; end; // ошибка?
     122://было &am   собираем  &amp;
          if      (_Char = 'p') then begin result:=result+'&'; State:=255 end
          else begin result:=result+'&am'+_Char; State := 0; end; // ошибка?
     255://ждем  ;
          if      (_Char = ';') then State := 0 // все путем
          else begin result:=result+_Char; State := 0; end; // ошибка?
     end;
   end;
 end; { XMLTextToXML }

procedure FindXMLAttr(KeyValue,AttrName: TXMLString; var NamePos,ValueStartPos,ValueOverPos : integer);
 var I,L,MaxL,MaxI,State : integer;
     _Char : char;
 begin
   //KeyValue:=AnsiUpperCase(KeyValue);
   //AttrName:=AnsiUpperCase(AttrName);
   NamePos := 0; ValueStartPos := 0; ValueOverPos := 0;
   MaxL:=Length(AttrName)+1;
   L:=1; // место в имени для сравнения; -1 - не нашлось; 0 - нашлось
   MaxI := Length(KeyValue);
   I:=0;
   State := -1;
   while I < MaxI do begin
     I:=I+1;
     _Char:=KeyValue[I];
     if (_Char = '>') and (State <> 4) then begin
       State:=-1;
       NamePos := 0; ValueStartPos := 0; ValueOverPos := 0;
     end;
     case State of
       -1: begin//ждем "<" - начало ключа
             if _Char='<' then State:=0;
           end;
       0 : begin//ждем пробела - конца имени ключа
             if _Char<=' ' then State:=1;
           end;
       1 : begin//ждем начало имени
             if (_Char>' ') and (_Char <> '/') and (_Char <> '>') then begin
               State:=2;
               NamePos := I;
               if AttrName[1] = _Char then L:=2 else L:=-1;
             end;
           end;
       2 : begin//ждем конец имени
             if (_Char='=') or (_Char<=' ') then begin
               State:=3;
               if L=MaxL then L:=0 else L:=-1;
             end else begin
               if (L>0) and (AttrName[L] = _Char) then inc(L) else L:=-1;
             end;
           end;
       3 : begin//ждем начало значения (")
             if (_Char='"') then begin
               State:=4;
               if L = 0 then ValueStartPos := I+1;
             end;
           end;
       4 : begin//ждем конец значения (")
             if _Char='"' then begin
               State:=1;
               if L=0 then begin
                 ValueOverPos := I;
                 exit;
               end;
             end;
           end;
     end;
   end;
   if L<0 then NamePos := 0;
 end; { FindXMLAttr }

function GetXMLAttribute(KeyValue,AttrName : TXMLString) : TXMLString;
 var NamePos,ValueStartPos,ValueOverPos : integer;
 begin
   FindXMLAttr(KeyValue,AttrName,NamePos,ValueStartPos,ValueOverPos);
   if (NamePos<=0) or (ValueStartPos<=0) or (ValueOverPos<=0) then result:=''
    else result:=copy(KeyValue,ValueStartPos,ValueOverPos-ValueStartPos);
 end; { GetXMLAttribute }

function SetXMLAttribute(KeyValue,AttrName,AttrValue : TXMLString) : TXMLString;
 var NamePos,ValueStartPos,ValueOverPos : integer;
 begin
   FindXMLAttr(KeyValue,AttrName,NamePos,ValueStartPos,ValueOverPos);
   if (NamePos<=0) or (ValueStartPos<=0) or (ValueOverPos<=0) then begin
     ValueStartPos := Length(KeyValue); ValueOverPos := ValueStartPos;
     if KeyValue[ValueStartPos-1] = '/' then dec(ValueStartPos);
     AttrValue := ' '+AttrName+'="'+AttrValue+'"';
   end;
   result:=copy(KeyValue,1,ValueStartPos-1)
            + AttrValue
            + copy(KeyValue,ValueOverPos,Length(KeyValue)-ValueOverPos+1);
 end; { SetXMLAttribute }

function DeleteXMLAttribute(KeyValue,AttrName : TXMLString) : TXMLString;
 var NamePos,ValueStartPos,ValueOverPos,L : integer;
 begin
   FindXMLAttr(KeyValue,AttrName,NamePos,ValueStartPos,ValueOverPos);
   if (NamePos<=0) or (ValueStartPos<=0) or (ValueOverPos<=0) then result:=KeyValue
    else begin
      L:=Length(KeyValue);
      repeat
        inc(ValueOverPos);
      until (ValueOverPos>L) or (KeyValue[ValueOverPos]>' ');
      result:=copy(KeyValue,1,NamePos-1)
              + copy(KeyValue,ValueOverPos,Length(KeyValue)-ValueOverPos+1);
   end;
 end; { DeleteXMLAttribute }

function DeleteXMLAllAttributes(KeyValue : TXMLString) : TXMLString;
 var Lin,L : integer;
     C : char;
 begin
   result:=KeyValue;
   Lin:=2; L:=Length(KeyValue);
   while Lin < L do begin
     C:=KeyValue[Lin];
     if C<=' ' then break;
     inc(Lin);
   end;
   if Lin < L then begin
     if KeyValue[L-1]='/' then begin
       result[Lin]:='/'; inc(Lin);
     end;
     result[Lin]:='>';
     SetLength(result,Lin);
   end;
 end; { DeleteXMLAllAttributes }

function CloseXMLKey(KeyValue : TXMLString) : TXMLString;
 var Lin,Lout,L : integer;
     C : char;
 begin
   result:=KeyValue;
   if KeyValue[2]='/' then exit;
   Lin:=2; L:=Length(KeyValue);
   result[2]:='/'; SetLength(result,L+1); Lout:=3;
   while Lin < L do begin
     C:=KeyValue[Lin];
     if C<=' ' then break;
     result[Lout]:=C; inc(Lout);
     inc(Lin);
   end;
   result[Lout]:='>';
   SetLength(result,Lout);
 end; { CloseXMLKey }

//////////////////////////////////////////////////////////////
// Объект XML

destructor TXML.Destroy;
begin
  if _FileIsOpen then close(_F);
end;

constructor TXML.Create();
begin
  _DataPos := 0; _BufSize := 0; _FileIsOpen := false;
end;

function TXML.OpenFile(FileName : widestring) : boolean;
var E : integer;
begin {$I-}
  result:=false;
  assign(_F,FileName);
  reset(_F,1);
  E:=ioresult;
  if E <> 0 then begin
    _ErrorMsg := 'Error '+inttostr(E)+' opening file "'+FileName+'"';
    exit;
  end;
  _FileIsOpen := true;
  _Key_Path := '';
  result:=true;
end; { TXML.OpenFile }

function TXML.Open_Read(FileName : widestring) : boolean;
var FM : byte;
begin
  FM:=FileMode;
  FileMode := 0;
  result:=self.OpenFile(FileName);
  FileMode:=FM;
end; { TXML.Open_Read }

function TXML.Open_Write(FileName : widestring) : boolean;
var FM : byte;
begin
  FM:=FileMode;
  FileMode := 2;
  result:=self.OpenFile(FileName);
  FileMode:=FM;
end; { TXML.Open_Write }

function TXML.ReadBuf() : boolean;
begin
  result:=false;
  if _FileIsOpen = false then begin
    _ErrorMsg := 'Error: file not open';
    exit;
  end;
  while _DataPos >= _BufSize do begin
    if eof(_F) then exit;
    blockread(_F,_DataBuf,XML_BufSize,_BufSize);
    _DataPos:=0;
  end;
  result:=(_DataPos < _BufSize);
end; { TXML.ReadBuf }

function TXML.ReadChar_UTF8 : TChar;
var B : byte;
    C : TChar;
begin
  result:=''; // bad char
  if not ReadBuf() then exit;
  B:=_DataBuf[_DataPos]; inc(_DataPos);
  C:=char(B);
  if B >= $C0 then begin
    if not ReadBuf() then exit;
    C:=C+char(_DataBuf[_DataPos]); inc(_DataPos);
  end;
  if B >= $E0 then begin
    if not ReadBuf() then exit;
    C:=C+char(_DataBuf[_DataPos]); inc(_DataPos);
  end;
  if B >= $F0 then begin
    if not ReadBuf() then exit;
    C:=C+char(_DataBuf[_DataPos]); inc(_DataPos);
  end;
  result:=C;
end; { TXML.ReadChar_UTF8 }

function TXML.ReadChar(InText : boolean) : boolean;
{var
   WC : widechar;
   State : byte;}
begin
  result:=false;
  //repeat
    if not ReadBuf() then exit;
    _Char := ReadChar_UTF8();
  //until (_Char>=' ') or InText;
{  if _Char <> '&' then begin result:=true; exit; end;
  State:=0;
  repeat
    if not ReadBuf() then exit;
    WC := ReadChar_UTF8();
//было &
    if      (State = 0) and (WC = 'q') then State:=11
    else if (State = 0) and (WC = 'l') then State:=21
    else if (State = 0) and (WC = 'g') then State:=31
    else if (State = 0) and (WC = 'a') then State:=101
    else if State = 0 then exit
//"   &quot;
    else if (State = 11) and (WC = 'u') then State:=12
    else if (State = 11) then exit
    else if (State = 12) and (WC = 'o') then State:=13
    else if (State = 12) then exit
    else if (State = 13) and (WC = 't') then begin _Char := '"'; State:=255 end
    else if (State = 13) then exit
//<   &lt;
    else if (State = 21) and (WC = 't') then begin _Char := '<'; State:=255 end
    else if (State = 21) then exit
//>   &gt;
    else if (State = 31) and (WC = 't') then begin _Char := '>'; State:=255 end
    else if (State = 31) then exit
//было &a
    else if (State = 101) and (WC = 'p') then State:=112
    else if (State = 101) and (WC = 'm') then State:=122
    else if (State = 101) then exit
//'   &apos;
    else if (State = 112) and (WC = 'o') then State:=113
    else if (State = 112) then exit
    else if (State = 113) and (WC = 's') then begin _Char := ''''; State:=255 end
    else if (State = 113) then exit
//&   &amp;
    else if (State = 122) and (WC = 'p') then begin _Char := '&'; State:=255 end
    else if (State = 122) then exit
//ждем ;
    else if (State = 255) and (WC = ';') then break
    else if (State = 255) then exit

  until false;}
  result:=true;
end; { TXML.ReadChar }

function TXML.ReadKey() : boolean;
var L,State,AN : integer;
begin
  result:=false;
  _Text_Value := '';

  if    (copy(_Key_Value,1,2) = '<?')
     or (copy(_Key_Value,1,2) = '<!')
     or (copy(_Key_Value,Length(_Key_Value)-1,2) = '/>')
    then begin

    //_Key_Name := _Key_Name;
    _Key_Value := '';
    _KeyIsClosed := true;

  end else begin

    _Key_Value := '';
    _Key_Name := '';
    repeat
      if not ReadChar(true) then exit;
      _Text_Value := _Text_Value + _Char;
    until _Char = '<';
    SetLength(_Text_Value, Length(_Text_Value) - 1);

    _Key_Value := '<';
    repeat
      if not ReadChar(false) then exit;
      _Key_Name := _Key_Name + UpperCase(_Char);
      _Key_Value := _Key_Value + _Char;
    until (_Char = '>') or (_Char <= ' ');
    L:=Length(_Key_Name)-1;
    SetLength(_Key_Name, L);
    if L = 0 then exit;
    if _Key_Name[1] = '/' then begin
      delete(_Key_Name,1,1);
      dec(L);
      if L = 0 then exit;
    end;
    if _Key_Name[L] = '/' then begin
      delete(_Key_Name,L,1);
      dec(L);
      if L = 0 then exit;
    end;
    _KeyAttributes := nil;
    State:=1;
    while _Char <> '>' do begin
      if not ReadChar(false) then exit;
      _Key_Value := _Key_Value + _Char;
      case State of
       0 : begin//ждем пробела - конца имени ключа
             if _Char<=' ' then State:=1;
           end;
       1 : begin//ждем начало имени
             if (_Char>' ') and (_Char <> '/') and (_Char <> '>') then begin State:=2;
               AN:=Length(_KeyAttributes);
               SetLength(_KeyAttributes,AN+1);
               _KeyAttributes[AN].NamePos:=Length(_Key_Value);
             end;
           end;
       2 : begin//ждем конец имени
             if (_Char='=') or (_Char<=' ') then begin State:=3;
               _KeyAttributes[AN].NameLen:=Length(_Key_Value)-_KeyAttributes[AN].NamePos;
             end;
           end;
       3 : begin//ждем начало значения (")
             if (_Char='"') then begin State:=4;
               _KeyAttributes[AN].ValuePos:=Length(_Key_Value)+1;
             end;
           end;
       4 : begin//ждем конец значения (")
             if _Char='"' then begin State:=1;
               _KeyAttributes[AN].ValueLen:=Length(_Key_Value)-_KeyAttributes[AN].ValuePos;
             end;
           end;
      end;
    end;
    _KeyIsClosed := (_Key_Value[2] = '/');

  end;

  if _KeyIsClosed then begin
     L:= Length(_Key_Path) - Length(_Key_Name);
     if (_Key_Path[L] <> '>') or (copy(_Key_Path,L+1,Length(_Key_Name)) <> _Key_Name) then exit;
     SetLength(_Key_Path, L-1);
  end else begin
     _Key_Path := _Key_Path + '>' + _Key_Name;
  end;

  result:=true;
end; { TXML.ReadKey }

function TXML.ReadNeededKey(KeyName: TXMLString): boolean;
begin
  result:=false;
  if not ReadKey() then exit;
  if uppercase(KeyName) <> _Key_Name then exit;
  result:=true;
end; { TXML.ReadNeededKey }

function TXML.FindKey(KeyName: TXMLString; PathMustNoUp,PathMustEqual : boolean): boolean;
label StartReadKey,BadEndReadKey;
var
  TV : TXMLString;
  NeedPath : TXMLString;
  NeedPathLen : longint;
begin{$B-}
  result:=false;
  KeyName := UpperCase(KeyName);
  TV:=''; NeedPathLen:=Length(_Key_Path); NeedPath:=_Key_Path+'>'+KeyName;
  goto StartReadKey;
  repeat
    if PathMustNoUp and (Length(_Key_Path)<NeedPathLen) then goto BadEndReadKey;
    TV:=TV+self._Text_Value+self._Key_Value;
  StartReadKey:
    if not ReadKey() then goto BadEndReadKey;
  until (_Key_Name = KeyName) and ( (PathMustEqual=false) or (_Key_Path=NeedPath) );
  result:=true;
BadEndReadKey:
  _Text_Value :=TV+_Text_Value;
end;

function TXML.FindPath(KeyPath: TXMLString): boolean;
begin
  result:=false;
  repeat
    if not ReadKey() then exit;
  until _Key_Path = KeyPath;
  result:=(NOT _KeyIsClosed);
end;

end.
