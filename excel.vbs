Option Explicit

Const PathName = "D:\Reports\"
Const FileName = "WiP.xlsx"
Const SheetConsumptionRate = "Расходная норма"

Const ForAppending = 8


Private Sub Main()
 Dim Tick
 Tick = Timer
 SaveReport()
 Log Tick
End Sub


Sub SaveReport()

 Dim FSO, Logging
 Set FSO = CreateObject("Scripting.FileSystemObject")
 set Logging = FSO.OpenTextFile("wip.log", ForAppending, True)
 Logging.WriteLine Now & " INFO: " & "Start"

 dim Excel, Source, Dest
 dim SheetName
 set Excel = CreateObject("Excel.Application")
 'Excel.Visible = True
 Excel.DisplayAlerts = False
 Excel.SheetsInNewWorkbook = 1

 SheetName = GetDate(Now)

 dim FilePath
 FilePath = PathName & Year(Now) & " " & FileName
 set Source = Excel.WorkBooks.Open(FilePath)

 FilePath = "Z:\" & Year(Now) & "_" & FileName
 if ExistFile(FilePath) then
  set Dest = Excel.WorkBooks.Open(FilePath)
 else
  set Dest = Excel.WorkBooks.Add
  Dest.SaveAs FilePath
 end if

 if not SheetExists(Dest, SheetConsumptionRate) then
  Source.Sheets(SheetConsumptionRate).Copy Dest.Sheets(1)
 end if

 if SheetExists(Dest, SheetName) then
  Dest.Sheets(SheetName).Delete
 end if

 Source.Sheets(SheetName).Copy Dest.Sheets(1)
 Dest.Sheets(1).Calculate
 Dest.Sheets(1).Activate

 Dest.Save
 Dest.Close
 Set Dest = Nothing
 Source.Close
 set Source = Nothing
 Excel.Quit
 Set Excel = Nothing

 Logging.WriteLine Now & " INFO: " & "Done"
 Logging.Close
 Set Logging = Nothing
 Set FSO = Nothing

End Sub


function GetDate(D)
 GetDate = Year(D) & "-" & right("0" & Month(D), 2) & "-" & right("0" & Day(D), 2)
end function


function SheetExists(WorkBook, SheetName)
 On Error Resume Next
 SheetExists = (LCase(WorkBook.Sheets(SheetName).Name) = LCase(SheetName))
 On Error Goto 0
end function


Function ExistFile(FileName)
 Dim FSO
 Set FSO = CreateObject("Scripting.FileSystemObject")
 ExistFile = FSO.FileExists(FileName)
 Set FSO = Nothing
End Function


Main()