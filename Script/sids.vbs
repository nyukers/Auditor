strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_UserAccount",,48)
For Each objItem in colItems
    Wscript.Echo "User ------------------------------"
'    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "SIDType: " & objItem.SIDType
    Wscript.Echo "Status: " & objItem.Status
'     strInstallDate = WMIDateStringToDate(objItem.InstallDate)
'    Wscript.Echo "Install Date: " & strInstallDate

Next


Set colItems = objWMIService.ExecQuery("Select * from Win32_SystemAccount",,48)
For Each objItem in colItems
    Wscript.Echo "System ----------------------------"
'    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "SIDType: " & objItem.SIDType
    Wscript.Echo "Status: " & objItem.Status
'     strInstallDate = WMIDateStringToDate(objItem.InstallDate)
'    Wscript.Echo "Install Date: " & strInstallDate

Next

Function WMIDateStringToDate(dtmDate)
    WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
         Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
         & " " & Mid (dtmDate, 9, 2) & ":" & _
         Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate, _
         13, 2))
End Function

