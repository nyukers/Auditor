@echo Query to Registries on Hotfixes
::reg query "HKLM\Software\Microsoft\Updates\" /s
reg query "HKLM\Software\Microsoft\Updates\" /v

@echo _
@echo to WinXP
reg query "HKLM\Software\Microsoft\Updates\Windows XP\" /v
reg query "HKLM\Software\Microsoft\Updates\Windows XP\SP1\" /v
reg query "HKLM\Software\Microsoft\Updates\Windows XP\SP2\" /v
reg query "HKLM\Software\Microsoft\Updates\Windows XP\SP3\" /v

@echo _
@echo to Win2000
reg query "HKLM\Software\Microsoft\Updates\Windows 2000\" /v
reg query "HKLM\Software\Microsoft\Updates\Windows 2000\SP4\" /v
reg query "HKLM\Software\Microsoft\Updates\Windows 2000\SP5\" /v
reg query "HKLM\Software\Microsoft\Updates\Windows 2000\SP6\" /v

reg query "HKLM\Software\Microsoft\Updates\Internet Explorer 6\" /v
reg query "HKLM\Software\Microsoft\Updates\Internet Explorer 6\SP1\" /v
reg query "HKLM\Software\Microsoft\Updates\Internet Explorer 6\SP2\" /v

@echo Query via WMI on Hotfixes
cscript hotfixes.vbs 

@echo Query via MBSA on Hotfixes
mbsacli /hf /v
