Option Explicit

const DestFileName = "D:\Reports\WiP.xlsx"
const MailTo = "Получатель <mail@example.com>"
const MailFrom = "Отправитель <xxxxx@yandex.ru>"
const MailSubject = "Показания на "
const SMTP = "smtp.yandex.ru"
const UserName = "xxxxx@yandex.ru"
const Password = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"


sub SendMail(D, FileName)
 dim FSO, File, Text
 set FSO = CreateObject("Scripting.FileSystemObject")
 set File = FSO.OpenTextFile("sendmail.txt", ForReading)
 Text = File.ReadAll
 File.Close
 set File = Nothing
 set FSO = Nothing
 '
 dim Mail
 const Setting = "http://schemas.microsoft.com/cdo/configuration/"
 set Mail = CreateObject("CDO.Message")
 Mail.To = MailTo
 Mail.From = MailFrom
 Mail.Subject = MailSubject & D
 Mail.TextBody = Text
 'Mail.HtmlBody = Text
 Mail.AddAttachment FileName
 Mail.TextBodyPart.Charset = "utf-8"
 Mail.Configuration.Fields.Item(Setting & "sendusing") = 2
 Mail.Configuration.Fields.Item(Setting & "smtpserver") = SMTP
 Mail.Configuration.Fields.Item(Setting & "smtpauthenticate") = 1
 Mail.Configuration.Fields.Item(Setting & "sendusername") = UserName
 Mail.Configuration.Fields.Item(Setting & "sendpassword") = Password
 Mail.Configuration.Fields.Item(Setting & "smtpserverport") = 465
 Mail.Configuration.Fields.Item(Setting & "smtpusessl") = True
 Mail.Configuration.Fields.Item(Setting & "smtpconnectiontimeout") = 10
 Mail.Configuration.Fields.Update
 Mail.Send
 set Mail = Nothing
end sub


function GetDate(D)
 GetDate = Year(D) & "-" & right("0" & Month(D), 2) & "-" & right("0" & Day(D), 2)
end function



dim FSO, Logging
const ForReading = 1, ForAppending = 8

set FSO = CreateObject("Scripting.FileSystemObject")
set Logging = FSO.OpenTextFile("sendmail.log", ForAppending, True)

Dim Period
Period = GetDate(Now)
SendMail Period, DestFileName
FSO.DeleteFile DestFileName
Logging.WriteLine Now & " INFO: Send mail on date " & Period

SaveLastDate()

Logging.Close
set Logging = Nothing
set FSO = Nothing
