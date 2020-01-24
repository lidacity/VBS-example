Option Explicit

const ForReading = 1, ForAppending = 8
const User = "User", Password = "Password", Domain = "WORKGROUP"


' запісаць у лог-файл нелегальную прыладу
sub WriteLog(SystemName, PnPDeviceID, Description, Name)
 dim FSO
 set FSO = CreateObject("Scripting.FileSystemObject")
 dim Logging
 set Logging = FSO.OpenTextFile("d:\Reports\usb.log", ForAppending, True)
 dim Caption
 if Name = Description then
  Caption = Name
 else
  Caption = Description & " " & Name
 end if
 Logging.WriteLine Now & ";" & SystemName & ";" & PnPDeviceID & ";" & Caption
 Logging.Close
 set Logging = Nothing
 set FSO = Nothing
end sub


' вярнуць імя прылады
function GetDeviceName(Dependent)
 dim Line, DeviceName
 Line = replace(Dependent, chr(34), "")
 DeviceName = split(Line, "=")
 GetDeviceName = DeviceName(1)
end function


' праверыць налады каманднага радка
function IsInit()
 IsInit = False
 if WScript.Arguments.Count = 1 then
  if WScript.Arguments(0) = "--default" then
   IsInit = True
  end if
 end if
end function


' атрымаць спіс прылад, за якімі не трэба назіраць
function GetDefault(Computer)
 dim FSO
 dim FileName
 FileName = "USB\" & Computer & ".def"
 dim File
 set FSO = CreateObject("Scripting.FileSystemObject")
 set File = FSO.OpenTextFile(FileName, ForReading)
 GetDefault = File.ReadAll
 set File = Nothing
 set FSO = Nothing
end function


' захаваць спіс прылад, за якімі не трэба назіраць
sub SaveDefault(Computer)
 dim FSO
 dim FileName
 FileName = "USB\" & Computer & ".def"
 dim Default
 set FSO = CreateObject("Scripting.FileSystemObject")
 if FSO.FileExists(FileName) then
  FSO.DeleteFile(FileName)
 end if
 set Default = FSO.OpenTextFile(FileName, ForAppending, True)
 dim SWbemLocator, WMIService, Device
 set SWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
 set WMIService = SWbemLocator.ConnectServer(Computer, "root\CIMV2", User, Password, "MS_409", "ntlmdomain:" & Domain)
 for each Device in WMIService.ExecQuery("SELECT * FROM Win32_USBControllerDevice")
  Default.WriteLine replace(GetDeviceName(Device.Dependent), "\\", "\")
 next
 set WMIService = Nothing
 set SWbemLocator = Nothing
 Default.Close
 set Default = Nothing
 set FSO = Nothing
end sub


' параўнаць спіс прылад, і запісаць розніцу
sub Check(Computer)
 dim List
 List = GetDefault(Computer)
 dim SWbemLocator, WMIService, Device
 set SWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
 set WMIService = SWbemLocator.ConnectServer(Computer, "root\CIMV2", User, Password, "MS_409", "ntlmdomain:" & Domain)
 for each Device in WMIService.ExecQuery("SELECT * FROM Win32_USBControllerDevice")
  dim DeviceName
  DeviceName = GetDeviceName(Device.Dependent)
  if InStr(List, replace(DeviceName, "\\", "\")) = 0 then
   dim USBDevice
   for each USBDevice in WMIService.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE DeviceID = '" & DeviceName & "'")
    WriteLog USBDevice.SystemName, USBDevice.PnPDeviceID, USBDevice.Description, USBDevice.Name
   next   
  end if
 next
 set WMIService = Nothing
 set SWbemLocator = Nothing
end sub


'On Error Resume Next
dim Computer
for each Computer in array("computer1", "computer2", "computer3", "computer4", "computer5", "computer6")
 if IsInit() then
  SaveDefault(Computer)
 else
  Check(Computer)
 end if
next
