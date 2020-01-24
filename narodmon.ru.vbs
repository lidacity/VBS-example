Option Explicit


const ForReading = 1, ForAppending = 8
const DEVICE_MAC = "xx:xx:xx:xx:xx:xx"


sub NarodMon(PathName)


 dim FSO, Logging, FileName, File
 set FSO = CreateObject("Scripting.FileSystemObject")
 set Logging = FSO.OpenTextFile("narodmon.log", ForAppending, True)

 dim T, P

 FileName = PathName & "T.txt"
 if FSO.FileExists(FileName) then
  set File = FSO.OpenTextFile(FileName, ForReading)
  T = File.Readline
  File.Close
  set File = Nothing
  FSO.DeleteFile FileName
 else
  T = ""
 end if

 if T <> "" then

  dim Send
  Send = "ID=" & DEVICE_MAC
  if T <> "" then
   Send = Send & "&T=" & T
  end if
  Logging.WriteLine Now & " REQUEST: " & Send

  On Error Resume Next
  dim NarodMon
  set NarodMon = CreateObject("Microsoft.XMLHTTP")
  NarodMon.Open "POST", "http://narodmon.ru/post.php", False
  NarodMon.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  NarodMon.SetRequestHeader "Content-Length", len(Send)
  NarodMon.Send Send
  if NarodMon.Status = 200 then
   Logging.WriteLine Now & " RESPONSE: " & NarodMon.ResponseText
  else
   Logging.WriteLine Now & " ERRCODE: " & NarodMon.Status
  end if
  set NarodMon = Nothing
  if Err.Number <> 0 then
   Logging.WriteLine Now & " ERROR: " & Err.Number & " " & replace(Err.Description, vbCrLf, "")
  end if
  On Error GoTo 0

 else
  Logging.WriteLine Now & " ERROR: " & "Value-file not found!"
 end if

 Logging.Close
 set Logging = Nothing
 set FSO = Nothing
end sub


NarodMon "D:\Reports\temp\"
