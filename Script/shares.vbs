strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

' действующие shares
Wscript.Echo "----------------------------"
Wscript.Echo "Valid share resources:"
Set colComputer = objWMIService.ExecQuery _
 ("Select Name, Caption, Path from Win32_Share")
  For Each objComputer in colComputer
      Wscript.Echo "Name: " & objComputer.Name
      Wscript.Echo " Desc: " & objComputer.Caption
      Wscript.Echo " Current path: " & objComputer.Path
'      Wscript.Echo " Current status: " & objComputer.Type
  Next

