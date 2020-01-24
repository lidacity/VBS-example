Option Explicit

' https://yandex.ru/dev/disk/poligon/#!//v1/disk/resources/CreateResource
const Authorization = "OAuth xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx_xxxxxxxx"


const stateRoot = 0
const stateNameQuoted = 1
const stateNameFinished = 2
const stateValue = 3
const stateValueQuoted = 4
const stateValueQuotedEscaped = 5
const stateValueUnquoted = 6
const stateValueUnquotedEscaped = 7

function JSONToXML(json)
 dim dom, xmlElem, i, ch, state, name, value
 set dom = CreateObject("Microsoft.XMLDOM")
 state = stateRoot
 for i = 1 to Len(json)
  ch = Mid(json, i, 1)
  Select Case state
  Case stateRoot
   Select Case ch
    Case "["
     If dom.documentElement is Nothing Then
      set xmlElem = dom.CreateElement("ARRAY")
      set dom.documentElement = xmlElem
     Else
      set xmlElem = XMLCreateChild(xmlElem, "ARRAY")
     End If
    Case "{"
     If dom.documentElement is Nothing Then
      set xmlElem = dom.CreateElement("OBJECT")
      set dom.documentElement = xmlElem
     Else
      set xmlElem = XMLCreateChild(xmlElem, "OBJECT")
     End If
    Case """"
     state = stateNameQuoted 
     name = ""
    Case "}"
     set xmlElem = xmlElem.parentNode
    Case "]"
     set xmlElem = xmlElem.parentNode
   End Select
  Case stateNameQuoted 
   Select Case ch
    Case """"
     state = stateNameFinished
    Case Else
     name = name + ch
   End Select
  Case stateNameFinished
   Select Case ch
    Case ":"
     value = ""
     State = stateValue
    End Select
  Case stateValue
   Select Case ch
    Case """"
     State = stateValueQuoted
    Case "{"
     Set xmlElem = XMLCreateChild(xmlElem, "OBJECT")
     State = stateRoot
    Case "["
     Set xmlElem = XMLCreateChild(xmlElem, "ARRAY")
     State = stateRoot
    Case " "
    Case Chr(9)
    Case vbCr
    Case vbLF
    Case Else
     value = ch
     State = stateValueUnquoted
   End Select
  Case stateValueQuoted
   Select Case ch
    Case """"
     xmlElem.setAttribute name, value
     state = stateRoot
    Case "\"
     state = stateValueQuotedEscaped
    Case Else
     value = value + ch
   End Select
  Case stateValueQuotedEscaped ' @@TODO: Handle escape sequences
   value = value + ch
   state = stateValueQuoted
  Case stateValueUnquoted
   Select Case ch
    Case "}"
     xmlElem.setAttribute name, value
     Set xmlElem = xmlElem.parentNode
     state = stateRoot
    Case "]"
     xmlElem.setAttribute name, value
     Set xmlElem = xmlElem.parentNode
     state = stateRoot
    Case ","
     xmlElem.setAttribute name, value
     state = stateRoot
    Case "\"
     state = stateValueUnquotedEscaped
    Case Else
     value = value + ch
   End Select
  Case stateValueUnquotedEscaped ' @@TODO: Handle escape sequences
   value = value + ch
   state = stateValueUnquoted
  End Select
 Next
 Set JSONToXML = dom
end function


function XMLCreateChild(xmlParent, tagName)
 Dim xmlChild
 if xmlParent is Nothing then
  set XMLCreateChild = Nothing
  Exit Function
 end if
 if xmlParent.ownerDocument is Nothing then
  set XMLCreateChild = Nothing
  Exit Function
 end if
 set xmlChild = xmlParent.ownerDocument.createElement(tagName)
 xmlParent.appendChild xmlChild
 set XMLCreateChild = xmlChild
end function


function GetValue(Json, Name)
 dim XML, Attr
 GetValue = ""
 set XML = JSONToXML(Json).documentElement
 if not XML.attributes is Nothing then
  for each Attr in XML.attributes
   if Attr.nodeName = Name then 
    GetValue = Attr.nodeValue
   end if
  next
 end if
end function


'src dst: "d" 866, "w" 1251, "u" UTF
function ConvertString(str, src, dst)
 src = lcase(src)
 dst = lcase(dst)
 dim Fsrc, Fdst, ArrFdos, ArrFwin, ArrFutf, d, Simv, n
 ArrFdos = split("128;129;130;131;132;133;134;135;136;137;138;139;140;141;142;143;144;145;146;147;148;149;150;151;152;153;154;155;156;157;158;159;160;161;162;163;164;165;166;167;168;169;170;171;172;173;174;175;224;225;226;227;228;229;230;231;232;233;234;235;236;237;238;239;240;241",";")
 ArrFwin = split("192;193;194;195;196;197;198;199;200;201;202;203;204;205;206;207;208;209;210;211;212;213;214;215;216;217;218;219;220;221;222;223;224;225;226;227;228;229;230;231;232;233;234;235;236;237;238;239;240;241;242;243;244;245;246;247;248;249;250;251;252;253;254;255;168;184",";")
 ArrFutf = split("208:144;208:145;208:146;208:147;208:148;208:149;208:150;208:151;208:152;208:153;208:154;208:155;208:156;208:157;208:158;208:159;208:160;208:161;208:162;208:163;208:164;208:165;208:166;208:167;208:168;208:169;208:170;208:171;208:172;208:173;208:174;208:175;208:176;208:177;208:178;208:179;208:180;208:181;208:182;208:183;208:184;208:185;208:186;208:187;208:188;208:189;208:190;208:191;209:128;209:129;209:130;209:131;209:132;209:133;209:134;209:135;209:136;209:137;209:138;209:139;209:140;209:141;209:142;209:143;208:129;209:145",";")
 if (src = "w" and dst = "w") or (src = "d" and dst = "d") or (src = "u" and dst = "u") then
  ConvertString = str
  exit function
 end if
 if src = "w" then
   Fsrc = ArrFwin
  elseif lcase(src) = "d" then
   Fsrc = ArrFdos
  elseif lcase(src) = "u" then
   Fsrc = ArrFutf
  else
   ConvertString = "Err: The variable src isn't true"
   exit function
 end if
 if dst = "w" then
   Fdst = ArrFwin
  elseif dst = "d" then
   Fdst = ArrFdos
  elseif dst = "u" then
   Fdst = ArrFutf
  else
   ConvertString = "Err: The variable dst isn't true"
   exit function
 end if
 set d = CreateObject("Scripting.Dictionary") 
 for n=0 to ubound(Fsrc)
  d.Add Fsrc(n), Fdst(n)
 next
 if (src = "w" and dst = "d") or (src = "d" and dst = "w") then
  for n = 1 to len(str)
   if d.item(cStr(asc(mid(str,n,1)))) <> "" then
    Simv = Simv & chr(d.item(cStr(asc(mid(str,n,1)))))
   else
    Simv = Simv & mid(str,n,1)
   end if
  next
 elseif src = "u" then
  for n = 1 to len(str)
   if asc(mid(str,n,1)) = 208 or asc(mid(str,n,1)) = 209 then
    Simv = Simv & chr(d.Item(cStr(asc(left(mid(str,n,2),1)) & ":" & asc(right(mid(str,n,2),1)))))
    n = n + 1
   else
    Simv = Simv & mid(str,n,1)
   end if
  next
 elseif dst = "u" then
  for n = 1 to len(str)
   if d.item(cStr(asc(mid(str,n,1)))) <> "" Then
    Simv = Simv & chr(left(d.item(cStr(asc(mid(str,n,1)))),3)) & chr(right(d.item(cStr(asc(mid(str,n,1)))),3)) 
   else
    Simv = Simv & mid(str,n,1)
   end if
  next
 end if
 set d = Nothing
 ConvertString = Simv
end function


Function UrlEncode(Data)
 Dim CharPosition, s, CharCode
 For CharPosition = 1 To Len(Data)
  s = Mid(Data, CharPosition,1)
  CharCode = Asc(s)
  if (charcode > 90 and charcode < 95) or charcode = 96 or (charcode > 122) then
   UrlEncode = UrlEncode & "%" & right("0" & hex(CharCode),2)
  else
   if charcode = 32 Then
    UrlEncode = UrlEncode & "+"
   else
    UrlEncode = UrlEncode & s
   end if
  end if
 Next
End Function


function ReadBinary(FileName)
 const adTypeBinary = 1
 dim BinaryStream
 set BinaryStream = CreateObject("ADODB.Stream")
 BinaryStream.Type = adTypeBinary
 BinaryStream.Open
 BinaryStream.LoadFromFile FileName
 ReadBinary = BinaryStream.Read
 BinaryStream.Close
 set BinaryStream = Nothing
end function


' -=-=-=-


sub UploadFileYaDisk(Logging, File)

 dim FSO, Size, FileName, PathName
 set FSO = CreateObject("Scripting.FileSystemObject")
 Size = FSO.GetFile(File).Size
 PathName = FSO.GetAbsolutePathName(File)
 FileName = FSO.GetFileName(File)
 set FSO = Nothing

 Logging.WriteLine Now & " INFO: " & "Upload to Yandex-Disk " & FileName
 dim YaDisk
 set YaDisk = CreateObject("Microsoft.XMLHTTP")

 ' атрымаць спасылак дл€ зал≥Ґк≥ файла на €ƒыск
 dim URL, Dest
 Dest = UrlEncode(ConvertString("Reports/" & FileName, "w", "u"))
 URL = "https://cloud-api.yandex.net/v1/disk/resources/upload?path=" & Dest & "&overwrite=true"
 Logging.WriteLine Now & " REQUEST: GET " & URL
 YaDisk.Open "GET", URL, False
 YaDisk.SetRequestHeader "Authorization", Authorization
 YaDisk.SetRequestHeader "Content-Type", "application/json"
 YaDisk.SetRequestHeader "Cache-Control", "no-cache"
 YaDisk.Send
 Logging.WriteLine Now & " RESPONSE: " & YaDisk.Status & " " & YaDisk.ResponseText

 ' зал≥ць файл на €ƒыск
 if YaDisk.Status = 200 then
  URL = GetValue(YaDisk.ResponseText, "href")
  Logging.WriteLine Now & " REQUEST: PUT " & URL
  YaDisk.Open "PUT", URL, False
  YaDisk.SetRequestHeader "Authorization", Authorization
  YaDisk.SetRequestHeader "Content-Type", "application/json" ' "application/x-www-form-urlencoded"
  YaDisk.SetRequestHeader "Cache-Control", "no-cache"
  YaDisk.SetRequestHeader "Content-Length", Size
  YaDisk.Send ReadBinary(PathName)
  Logging.WriteLine Now & " RESPONSE: " & YaDisk.Status & " OK"
 end if

 set YaDisk = Nothing

end sub


dim FSO, Logging
const ForAppending = 8

set FSO = CreateObject("Scripting.FileSystemObject")
set Logging = FSO.OpenTextFile("send.log", ForAppending, True)
Logging.WriteLine Now & " INFO: " & "Start"

Dim File
For Each File In FSO.GetFolder("D:\Reports\").Files
 If FSO.GetExtensionName(File) = "xlsx" Then
  UploadFileYaDisk Logging, File
 End If
Next

Logging.WriteLine Now & " INFO: " & "Done"
Logging.Close
set Logging = Nothing
set FSO = Nothing
