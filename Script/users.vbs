strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

' �������騥 ������
Wscript.Echo "enabled accounts:"
Set colComputer = objWMIService.ExecQuery _
 	   ("Select * from Win32_UserAccount where Disabled=FALSE")
  For Each objComputer in colComputer
      Wscript.Echo "Name: " & objComputer.Caption
      Wscript.Echo " Desc: " & objComputer.FullName
'      Wscript.Echo " Current status: " & objComputer.Status
  Next

Wscript.Echo "----------------------------"

' �⪫�祭� ������
Wscript.Echo "disabled accounts:"
Set colComputer = objWMIService.ExecQuery _
 	   ("Select * from Win32_UserAccount where Disabled=TRUE")
  For Each objComputer in colComputer
      Wscript.Echo "Name: " & objComputer.Caption
      Wscript.Echo " Desc: " & objComputer.FullName
'      Wscript.Echo " Current status: " & objComputer.Status
  Next


Wscript.Echo "----------------------------"
' �������஢���� ������
Wscript.Echo "locked accounts:"
Set colComputer = objWMIService.ExecQuery _
 	   ("Select * from Win32_UserAccount where Lockout=TRUE")
  For Each objComputer in colComputer
      Wscript.Echo "Name: " & objComputer.Caption
      Wscript.Echo " Desc: " & objComputer.FullName
  Next