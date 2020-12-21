strComputer = "."
Set objArgs = WScript.Arguments   ' Создаем объект WshArguments
Set objNamedArgs = objArgs.Named  ' Создаем объект WshNamed

' Проверяем, существует ли аргумент /Host:
HostName = objNamedArgs("Host")
PathName = objNamedArgs("PathDump")

Wscript.Echo HostName
Wscript.Echo PathName

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Backup, Security)}!\\" & _
        strComputer & "\root\cimv2")

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='Application'")

For Each objLogfile in colLogFiles
    errBackupLog = objLogFile.BackupEventLog(PathName & "\" & HostName &"_app.evt")
    If errBackupLog <> 0 Then        
        Wscript.Echo "The Application eventlog could not be backed up: " & errBackupLog
    End If
Next

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='Security'")

For Each objLogfile in colLogFiles
    errBackupLog = objLogFile.BackupEventLog(PathName & "\" & HostName &"_sec.evt")
    If errBackupLog <> 0 Then        
        Wscript.Echo "The Security eventlog could not be backed up. " & errBackupLog
    End If
Next

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='System'")

For Each objLogfile in colLogFiles
    errBackupLog = objLogFile.BackupEventLog(PathName & "\" & HostName &"_sys.evt")
    If errBackupLog <> 0 Then        
        Wscript.Echo "The System eventlog could not be backed up. " & errBackupLog
    End If
Next


'' SIG '' Begin signature block
'' SIG '' MIIDSgYJKoZIhvcNAQcCoIIDOzCCAzcCAQExDjAMBggq
'' SIG '' hkiG9w0CBQUAMGYGCisGAQQBgjcCAQSgWDBWMDIGCisG
'' SIG '' AQQBgjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIB
'' SIG '' AAIBAAIBAAIBAAIBADAgMAwGCCqGSIb3DQIFBQAEELgl
'' SIG '' pTb2IuqkEDqfi+JF+2WgggGrMIIBpzCCAVGgAwIBAgIQ
'' SIG '' oo7gTQKdIatCpHZVbVHuJDANBgkqhkiG9w0BAQQFADAf
'' SIG '' MR0wGwYDVQQDHhQAUwBaAEkAXwA5AEEAQQBWAEkAUjAe
'' SIG '' Fw0wMzEyMzEyMjAwMDBaFw0wOTEyMzEyMjAwMDBaMB8x
'' SIG '' HTAbBgNVBAMeFABTAFoASQBfADkAQQBBAFYASQBSMFww
'' SIG '' DQYJKoZIhvcNAQEBBQADSwAwSAJBALLAETWbwE31VBQD
'' SIG '' bg7gA0PIDWWcaFjUDXiVT3UDIeOoCK4xo83gH4iCX1nu
'' SIG '' uYLDUp12ptM0BZakTncSrDNqrLUCAwEAAaNpMGcwEwYD
'' SIG '' VR0lBAwwCgYIKwYBBQUHAwMwUAYDVR0BBEkwR4AQ6NZk
'' SIG '' VrOo5+C3bfQY9/zPxKEhMB8xHTAbBgNVBAMeFABTAFoA
'' SIG '' SQBfADkAQQBBAFYASQBSghCijuBNAp0hq0KkdlVtUe4k
'' SIG '' MA0GCSqGSIb3DQEBBAUAA0EAE83ozIWEJviqHKA+ESTI
'' SIG '' VErmU/V8xxpqpxSICysRmwE8SPSFhp4dR/BxanwOr4sy
'' SIG '' uI0qryE9IcvDnF3anSSq1TGCAQkwggEFAgEBMDMwHzEd
'' SIG '' MBsGA1UEAx4UAFMAWgBJAF8AOQBBAEEAVgBJAFICEKKO
'' SIG '' 4E0CnSGrQqR2VW1R7iQwDAYIKoZIhvcNAgUFAKBsMBAG
'' SIG '' CisGAQQBgjcCAQwxAjAAMBkGCSqGSIb3DQEJAzEMBgor
'' SIG '' BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEE
'' SIG '' AYI3AgEVMB8GCSqGSIb3DQEJBDESBBDyfL+1jYYwKZnS
'' SIG '' n4EhIMepMA0GCSqGSIb3DQEBAQUABECpkxTgdYhrsDyd
'' SIG '' RzIPd7rtt1qWBKBq4ElzzKs1kvgYc+Zu2I8HgwnLiXGF
'' SIG '' 2y26M290FI2SLIUJR4TT9ZX4+yKg
'' SIG '' End signature block
