Option Explicit


const ForReading = 1, ForAppending = 8
const API_ID = "xxxxxxxxx:xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
const ChatID = "xxxxxxxxx"


function SendMessage(Logging, ID, Chat, Text)
 set Telegram = CreateObject("Microsoft.XMLHTTP")
 URL = "https://api.telegram.org/bot" & API_ID & "/sendMessage?" & "chat_id=" & ChatID & "&text=" & Text & "&parse_mode=Markdown"
 Logging.WriteLine Now & " REQUEST: " & URL
 Telegram.open "POST", URL, false
 Telegram.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
 Telegram.send
 SendMessage = Telegram.responseText
 set Telegram = Nothing
 Logging.WriteLine Now & " RESPONSE: " & SendMessage
end function


function SendPhoto(Logging, ID, Chat, FileName, Caption)
 dim URL
 URL = "https://api.telegram.org/bot" & ID & "/sendPhoto?" & "chat_id=" & Chat
 Logging.WriteLine Now & " REQUEST: " & URL
 SendPhoto = Upload(URL, FileName, "photo", "chat_id=" & Chat & "|filename=" & FileName & "|caption=" & Caption)
 Logging.WriteLine Now & " RESPONSE: " & SendPhoto
end function


'*bold text*
'_italic text_
'[text](URL)
'`inline fixed-width code`
'```pre-formatted fixed-width code block```
function SendDocument(Logging, ID, Chat, FileName, Caption)
 dim URL
 URL = "https://api.telegram.org/bot" & ID & "/sendDocument?" & "chat_id=" & Chat
 Logging.WriteLine Now & " REQUEST: " & URL
 SendDocument = Upload(URL, FileName, "document", "chat_id=" & Chat & "|filename=" & FileName & "|caption=" & Caption)
 Logging.WriteLine Now & " RESPONSE: " & SendDocument
end function


' src dst: "d" 866, "w" 1251, "u" UTF8
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



function Upload(strUploadUrl, strFilePath, strFileField, strDataPairs)
const MULTIPART_BOUNDARY = "---------------------------111111111----"
dim ado, rs
dim lngCount
dim bytFormData, bytFormStart, bytFormEnd, bytFile
dim strFormStart, strFormEnd, strDataPair
dim web
const adLongVarBinary = 205
 '
 set ado = CreateObject("ADODB.Stream")
 ado.Type = 1
 ado.Open
 ado.LoadFromFile strFilePath
 bytFile = ado.Read
 ado.Close
 set ado = Nothing
 '
 strFormEnd = vbCrLf & "--" & MULTIPART_BOUNDARY & "--" & vbCrLf
 '
 strFormStart = ""
 for each strDataPair In Split(strDataPairs, "|")
  strFormStart = strFormStart & "--" & MULTIPART_BOUNDARY & vbCrLf
  strFormStart = strFormStart & "Content-Disposition: form-data; "
  strFormStart = strFormStart & "name=""" & Split(strDataPair, "=")(0) & """"
  strFormStart = strFormStart & vbCrLf & vbCrLf
  strFormStart = strFormStart & Split(strDataPair, "=")(1)
  strFormStart = strFormStart & vbCrLf
 next
 '
 strFormStart = strFormStart & "--" & MULTIPART_BOUNDARY & vbCrLf
 strFormStart = strFormStart & "Content-Disposition: form-data; "
 strFormStart = strFormStart & "name=""" & strFileField & """; "
 strFormStart = strFormStart & "filename=""" & Mid(strFilePath, InStrRev(strFilePath, "\") + 1) & """"
 strFormStart = strFormStart & vbCrLf
 strFormStart = strFormStart & "Content-Type: application/upload"
 strFormStart = strFormStart & vbCrLf & vbCrLf
 '
 set rs = CreateObject("ADODB.Recordset")
 rs.Fields.Append "FormData", adLongVarBinary, Len(strFormStart) + LenB(bytFile) + Len(strFormEnd)
 rs.Open
 rs.AddNew
 '
 for lngCount = 1 to Len(strFormStart)
  bytFormStart = bytFormStart & ChrB(Asc(Mid(strFormStart, lngCount, 1)))
 next
 rs("FormData").AppendChunk bytFormStart & ChrB(0)
 bytFormStart = rs("formData").GetChunk(Len(strFormStart))
 rs("FormData") = ""
 '
 for lngCount = 1 to Len(strFormEnd)
  bytFormEnd = bytFormEnd & ChrB(Asc(Mid(strFormEnd, lngCount, 1)))
 next
 rs("FormData").AppendChunk bytFormEnd & ChrB(0)
 bytFormEnd = rs("formData").GetChunk(Len(strFormEnd))
 rs("FormData") = ""
 '
 rs("FormData").AppendChunk bytFormStart
 rs("FormData").AppendChunk bytFile
 rs("FormData").AppendChunk bytFormEnd
 bytFormData = rs("FormData")
 rs.Close
 set rs = Nothing
 '
 set web = CreateObject("WinHttp.WinHttpRequest.5.1")
 web.SetTimeouts 0, 0, 0, 0
 web.Open "POST", strUploadUrl, False
 web.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & MULTIPART_BOUNDARY
 web.Send bytFormData
 Upload = web.responseText
 set web = Nothing
end function


sub Telegram(PathName)

 dim FSO, Logging, FileName
 set FSO = CreateObject("Scripting.FileSystemObject")
 set Logging = FSO.OpenTextFile("telegram.log", ForAppending, True)

 for each FileName in FSO.GetFolder(PathName).Files
  if right(FileName, 4) = ".png" then
   'Text = ConvertString("проверка *связи*!", "w", "u")
   'SendMessage Logging, API_IP, ChatID, Text
   '
   'FileName = "test.png"
   'Caption = "hello"
   'SendPhoto Logging, API_ID, ChatID, FileName, Caption
   'Caption = ConvertString("Ку Ку", "w", "u")
   'SendDocument Logging, API_ID, ChatID, FileName, Caption
   '
   dim Caption
   Caption = FSO.GetFileName(FileName)
   Caption = replace(Caption, "." & FSO.GetExtensionName(Caption), "")
   Caption = replace(Caption, "_", ":")
   SendDocument Logging, API_ID, ChatID, FileName, Caption
  end if
 next

 Logging.Close
 set Logging = Nothing
 set FSO = Nothing
end sub


Telegram("D:\Reports\ScreenShot\")
