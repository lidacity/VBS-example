On Error Resume Next

set SNMP = CreateObject("OlePrn.OleSNMP")
SNMP.Open "192.168.0.2", "public", 2, 1000
D = SNMP.Get(".1.3.6.1.4.1.318.2.1.6.1.0")
T = SNMP.Get(".1.3.6.1.4.1.318.2.1.6.2.0")
SNMP.Close
set SNMP = Nothing

if Err.Number = 0 then
 wscript.echo "Date: " & D & " Time: " & T
else
 wscript.echo "ERROR: " & Err.Number & " " & Err.Description
end if

On Error GoTo 0