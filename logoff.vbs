Option Explicit


dim Computer

if WScript.Arguments.Count = 1 then
 Computer = WScript.Arguments(0)
else
 WScript.echo "Usage: cscript //nologo logoff.vbs <ComputerName>"
 WScript.Quit(1)
end if


Logoff Computer, "User", "Password"


Sub Logoff(Computer, User, Password)
 const EWX_LOGOFF   = 0 ' останавливает все процессы, связанные с контексте безопасности процесса и отображает диалоговое окно входа в систему
 const EWX_SHUTDOWN = 1 ' Выключить компьютер к точке, где это безопасно выключить питание. Процессы и пользователь уведомляются
 const EWX_REBOOT   = 2 ' Завершает работу, а затем перезагружает компьютер
 const EWX_POWEROFF = 8 ' Выключить компьютер и выключить питание (если поддерживается компьютером).
 const EWX_FORCELOGOFF   = 4 ' (0 + 4) Немедленный выход. Не уведомляет приложения, сессия входа в систему заканчивающиеся. Это может привести к потере данных
 const EWX_FORCESHUTDOWN = 5 ' (1 + 4) Принудительное выключение. Выключить компьютер к точке, где это безопасно выключить питание. Пользователи уведомляются. Все файловые буферы сбрасываются на диск, и все запущенные процессы останавливаются. Из-за этого, вы не сможете получить возвращаемого значения, если вы работаете на удаленном компьютере
 const EWX_FORCEREBOOT   = 6 ' (2 + 4) Принудительная перезагрузка. Завершает работу, а затем перезагружает компьютер. Все запущенные процессы останавливаются. Из-за этого, вы не сможете получить возвращаемого значения, если вы работаете на удаленном компьютере
 const EWX_FORCEPOWEROFF = 12 ' (8 + 4) Принудительное отключение питания. Выключить компьютер и выключить питание (если поддерживается компьютером). Все запущенные процессы останавливаются. Из-за этого, вы не сможете получить возвращаемого значения, если вы работаете на удаленном компьютере.
 dim SWbemLocator, WMIService, OpSys
 set SWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
 set WMIService = SWbemLocator.ConnectServer(Computer, "root\CIMV2", User, Password)
 for each OpSys in WMIService.ExecQuery("select * from Win32_OperatingSystem where Primary=true")
  OpSys.Win32Shutdown EWX_LOGOFF
 next
 set WMIService = Nothing
 set SWbemLocator = Nothing
End Sub
