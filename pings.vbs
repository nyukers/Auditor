strMachines = "nyukexp; sziexp; barsnbu; barsrez"

ON ERROR RESUME NEXT

aMachines = split(strMachines, "; ")

For Each machine in aMachines
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        ExecQuery("select * from Win32_PingStatus where address = '"_
            & machine & "'")
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
            WScript.Echo("Ping to " & machine & " is not reachable now.")   
            Else 
	    WScript.Echo("Ping to " & machine & " is Ok!")   

            Err.Clear
            Set objWMIService = GetObject("winmgmts:\\" & machine & "\root\cimv2")
            If Err.Number <> 0 Then
             WScript.Echo("Error # " & CStr(Err.Number) & " " & Err.Description)
             WScript.Echo("Host " & machine & " is unreachable!")               
             Else
             Set IPConfigSet = objWMIService.ExecQuery _
             ("Select IPAddress from Win32_NetworkAdapterConfiguration") 
             For Each IPConfig in IPConfigSet
              If Not IsNull(IPConfig.IPAddress) Then 
               For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
               WScript.Echo IPConfig.IPAddress(i)
               Next
              End If
             Next
' else get parameters

'
            End If
' 
        End If
    Next
Next

WScript.Quit(EXIT_SUCCESS) 