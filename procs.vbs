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
    Num=Num+1
Next