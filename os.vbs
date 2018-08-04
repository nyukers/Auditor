Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")

For Each objOS in colOperatingSystems
    dtmConvertedDate.Value = objOS.LastBootUpTime
    dtmLastBootUpTime = dtmConvertedDate.GetVarDate
    dtmSystemUptime = DateDiff("d", dtmLastBootUpTime, Now)
     Wscript.Echo dtmSystemUptime 
     Wscript.Echo objOS.Caption &","& objOS.Version
     Wscript.Echo objOS.SerialNumber     
     Wscript.Echo objOS.BootDevice
     Wscript.Echo "ServicePack: " & objOS.ServicePackMajorVersion &"."& objOS.ServicePackMinorVersion
Next
