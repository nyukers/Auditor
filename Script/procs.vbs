strComputer = "."
Num = 1
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process")
For Each objProcess in colProcessList
    Wscript.Echo "----------------------- N:" & Num
    Wscript.Echo "Process: " & objProcess.Name 
    Wscript.Echo "Process ID: " & objProcess.ProcessID 
    Wscript.Echo "Thread Count: " & objProcess.ThreadCount 
    Wscript.Echo "Page File Size: " & objProcess.PageFileUsage 
    Wscript.Echo "Page Faults: " & objProcess.PageFaults 
    Wscript.Echo "Working Set Size: " & objProcess.WorkingSetSize 
    Wscript.Echo "Command Line: " & objProcess.CommandLine
'    Wscript.Echo "Creation Date: " & WMIDateStringToDate(objProcess.CreationDate)
    Wscript.Echo "Process is owned by " _ 
                  & strUserDomain & "\" & strNameOfUser & "." 
    Num=Num+1
Next

Function WMIDateStringToDate(dtmDate)
    WMIDateStringToDate = CDate(Mid(dtmDate, 7, 2) & "/" & _
         Mid(dtmDate, 5, 2) & "/" & Left(dtmDate, 4) _
         & " " & Mid (dtmDate, 9, 2) & ":" & _
         Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate, _
         13, 2))
End Function
