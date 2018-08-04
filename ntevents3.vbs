strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate, (Security)}!\\" & strComputer & "\root\cimv2")
Set colMonitoredEvents = objWMIService.ExecNotificationQuery _    
("Select * from __instancecreationevent where TargetInstance isa 'Win32_NTLogEvent' and TargetInstance.EventCode = '999'")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
        strAlertToSend = objLatestEvent.TargetInstance.User _ 
            & " Event 999 done."
        Wscript.Echo strAlertToSend
Loop
