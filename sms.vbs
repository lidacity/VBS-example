Option Explicit


const ForReading = 1, ForAppending = 8


'https://s3.amazonaws.com/etherarp.net/hilink.sh
sub SendSMS(FileName)

 const E303_ip = "http://192.168.1.1/"
 const HTTPPOST_SEND = "api/sms/send-sms"
 const Phone = "+7911xxxxxxx"

 dim FSO, Logging
 set FSO = CreateObject("Scripting.FileSystemObject")
 if FSO.FileExists(FileName) then

  set Logging = FSO.OpenTextFile("sms.log", ForAppending, True)
  Logging.WriteLine Now & " INFO: " & FileName
 
  dim Content, Length, DateTime
 
  dim File
  set File = FSO.OpenTextFile(FileName, ForReading)
  Content = File.ReadAll
  File.Close
  set File = Nothing
  FSO.DeleteFile FileName
  Length = len(Content)
 
  DateTime = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
 
  dim XML, SMS
  XML = "<request><Index>-1</Index><Phones><Phone>" & Phone & "</Phone></Phones><Sca></Sca><Content>" & Content & "</Content><Length>" & Length & "</Length><Reserved>1</Reserved><Date>" & DateTime & "</Date></request>"
  Logging.WriteLine Now & " REQUEST: " & Replace(XML, vbCrLf, "\n")
 
  set SMS = CreateObject("Microsoft.XMLHTTP")
  SMS.Open "POST", E303_ip & HTTPPOST_SEND, False
  SMS.SetRequestHeader "Content-Type", "application/xml"
  SMS.Send XML
 
  if SMS.Status = 200 then
   Logging.WriteLine Now & " RESPONSE: " & Replace(SMS.ResponseText, vbCrLf, "")
  else
   Logging.WriteLine Now & " ERRCODE: " & SMS.Status
  end if
  set SMS = Nothing
 
  Logging.Close
  set Logging = Nothing

 end if
 set FSO = Nothing

end sub



dim Item
for each Item in array("sms1.txt", "sms2.txt")
 SendSMS "D:\Reports\temp\" & Item
next
