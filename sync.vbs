Option Explicit


const ForReading = 1, ForWriting = 2, ForAppending = 8

class FileStruct
 public Size
 public DateTime
end class


function CreateFolder(Path)
 dim Line
 Line = Now & vbTab & "create destination directory " & Path & " .. "
 On Error Resume Next
 dim FSO
 set FSO = CreateObject("Scripting.FileSystemObject")
 FSO.CreateFolder Path
 set FSO = Nothing
 On Error Goto 0
 if Err.Number = 0 Then
  CreateFolder = Line & "ok"
 else
  CreateFolder = Line & "ERROR : " & Err.Description
  Err.Clear
 end if
end Function


function DeleteFolder(Path)
 dim Line
 Line = Now & vbTab & "delete directory " & Path & " .. "
 On Error Resume Next
 dim FSO
 set FSO = CreateObject("Scripting.FileSystemObject")
 FSO.DeleteFolder Path, True
 set FSO = Nothing
 On Error Goto 0
 if Err.Number = 0 Then
  DeleteFolder = Line & "ok"
 else
  DeleteFolder = Line & "ERROR : " & Err.Description
  Err.Clear
 end if
end Function


function CopyFile(Source, Dest)
 dim Line 
 Line = Now & vbTab & "copy file " & Source & " to " & Dest & " .. "
 On Error Resume Next
 dim FSO
 set FSO = CreateObject("Scripting.FileSystemObject")
 FSO.CopyFile Source, Dest, True
 set FSO = Nothing
 On Error Goto 0
 if Err.Number = 0 Then
  CopyFile = Line & "ok"
 else
  CopyFile = Line & "ERROR : " & Err.Description
  Err.Clear
 end if
end Function


function DeleteFile(Path)
 dim Line
 Line = Now & vbTab & "delete file " & Path & " .. "
 On Error Resume Next
 dim FSO
 set FSO = CreateObject("Scripting.FileSystemObject")
 FSO.DeleteFile Path, true
 set FSO = Nothing
 On Error Goto 0
 if Err.Number = 0 Then
  DeleteFile = Line & "ok"
 else
  DeleteFile = Line & "ERROR : " & Err.Description
  Err.Clear
 end if
end Function


function CopyFolder(Source, Dest)
 dim Line
 Line = Now & vbTab & "copy directory " & Source & " to " & Dest & " .. "
 On Error Resume Next
 dim FSO
 set FSO = CreateObject("Scripting.FileSystemObject")
 FSO.CopyFolder Source, Dest, True
 set FSO = Nothing
 On Error Goto 0
 if Err.Number = 0 Then
  CopyFolder = Line & "ok"
 else
  CopyFolder Line & "ERROR : " & Err.Description
  Err.Clear
 end if
end Function


function IsExist(Name, Names)
 IsExist = InStr(Names, ";" & Name & ";") <> 0
end function


sub SyncFolder(Logging, Source, Dest, Exclude)
 On Error Resume Next
 if FSO.FolderExists(Source) then

  if not FSO.FolderExists(Dest) then
   Logging.WriteLine CreateFolder(Dest)
  end if

  dim Dict
  set Dict = CreateObject("Scripting.Dictionary")
  dim File, Folder, Key
  dim Node

  dim Folder2
  set Folder2 = FSO.GetFolder(Dest)
  for each File in Folder2.Files
   set Node = new FileStruct
   Node.Size = File.Size
   Node.DateTime = File.DateLastModified
   Dict.Add File.Name, Node
  next
  for each Folder in Folder2.SubFolders
   set Node = new FileStruct
   Node.Size = -1
   Node.DateTime = 0
   Dict.Add Folder.Name, Node
  next

  dim Folder1
  set Folder1 = FSO.GetFolder(Source)
  for each File in Folder1.Files
   if Dict.Exists(File.Name) then
    if Dict.Item(File.Name).Size = -1 then
     Logging.WriteLine DeleteFolder(Dest & File.Name)
    end if
    if (File.Size <> Dict.Item(File.Name).Size) or (File.DateLastModified <> Dict.Item(File.Name).DateTime) then
     if not IsExist(File.Name, Exclude) then
      Logging.WriteLine CopyFile(Source & File.Name, Dest)
     end if
    end if
    Dict.Remove(File.Name)
   else
    if not IsExist(File.Name, Exclude) then
     Logging.WriteLine CopyFile(Source & File.Name, Dest)
    end if
   end if
  next

  for each Key in Dict.Keys
   if Dict.Item(Key).Size <> -1 then
    Logging.WriteLine DeleteFile(Dest & Key)
    Dict.Remove(Key)
   end if
  next

  for each Folder in Folder1.SubFolders
   if Dict.Exists(Folder.Name) then
    Dict.Remove(Folder.Name)
    SyncFolder Logging, Source & Folder.Name & "\", Dest & Folder.Name & "\", Exclude
   else
    if not IsExist(Folder.Name, Exclude) then
     Logging.WriteLine CopyFolder(Source & Folder.Name, Dest & Folder.Name)
    end if
   end if
  next

  for each Key in Dict.Keys
   Logging.WriteLine DeleteFolder(Dest & Key)
  next

 end If
 On Error Goto 0
end sub



const SourceFolder = "D:\Reports\"
const DestinationFolder = "\\computer\Reports\"

dim FSO
set FSO = CreateObject("Scripting.FileSystemObject")
dim Logging
set Logging = FSO.OpenTextFile(SourceFolder & "sync.log", ForAppending, True)

Logging.WriteLine(Now & vbTab & "Starting sync " & SourceFolder & " to " & DestinationFolder)
dim Exclude
Exclude = array(".bak", ".tmp", ".log")
Exclude = ";" & join(Exclude, ";") & ";"
SyncFolder Logging, SourceFolder, DestinationFolder, Exclude


Logging.WriteLine(Now & vbTab & "Done")
Logging.Close
set Logging = Nothing
set FSO = Nothing
