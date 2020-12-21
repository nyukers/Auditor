@echo off
if %1.==. goto noname

echo %date%
echo %time%
@echo Get live responce by -=-=-=-=-=--=-=-=-
psgetsid

@echo Host "%1" is -=-=-=-=-=--=-=-=-
psinfo
srvinfo -ns

@echo net -=-=-=-=-=--=-=-=-
@echo net use 
net use 
@echo net session 
net session 
@echo net file 
net file 
@echo net share 
net share 
@echo srvcheck 
srvcheck \\%1

@echo net view 
net view 
@echo net user
net user 
@echo net accounts 
net accounts 
@echo net localgroup 
net localgroup 
@echo net start 
net start 
@echo net config server
net config server
@echo net config workstation
net config workstation
@echo
 
@echo nbtstat -=-=-=-=-=--=-=-=-
nbtstat -c 
nbtstat -n 
nbtstat -s 

@echo netstat -=-=-=-=-=--=-=-=-
netstat -an 
netstat -r 

@echo arp -=-=-=-=-=--=-=-=-
arp -a 

@echo ipconfig -=-=-=-=-=--=-=-=-
ipconfig /all

::@echo doskey -=-=-=-=-=--=-=-=-
::doskey /history 

@echo Ports   -=-=-=-=-=--=-=-=-
fport
openports -lines -path
@echo PortsQuery   -=-=-=-=-=--=-=-=-
PortQry.exe -local -v

@echo proceses pslist -=-=-=-=-=--=-=-=-
pslist
@echo services sclist -=-=-=-=-=--=-=-=-
sclist

@echo Logged On -=-=-=-=-=--=-=-=-
psloggedon

@echo ntlast -=-=-=-=-=--=-=-=-
ntlast
ntlast -v

@echo auditpol -=-=-=-=-=--=-=-=-
auditpol

:@echo List DLLs -=-=-=-=-=--=-=-=-
:listdlls

@echo View Clipboard -=-=-=-=-=--=-=-=-
@echo Text only 
pclip.exe
@echo Text OR Image 
clpview.exe /c
@echo _end of clipboard.

@echo View Recycle Bin -=-=-=-=-=--=-=-=-
@echo AutoCLI tools disable ...(

@echo List of Admins -=-=-=-=-=--=-=-=-
SHOWMBRS.EXE Администраторы
SHOWMBRS.EXE Administrators


@echo View scheduled tasks -=-=-=-=-=--=-=-=-
@echo View AT -=-=-=-=-=--=-=-=-
at
@echo View SCHTASKS -=-=-=-=-=--=-=-=-
schtasks

goto test

@echo View ADS -=-=-=-=-=--=-=-=-
lads  c:\ /s
sfind c:\>%1_ads.log


@echo Files by last accessed times
dir /t:a /o:d /s c:\
@echo Files by last modified times
dir /t:w /o:d /s c:\
@echo Files by creation times
dir /t:c /o:d /s c:\


:test
@echo AutoRunning -=-=-=-=-=--=-=-=-
autorunsc -s -w -a


@echo Query to Registries on autorunning -=-=-=-=-=--=-=-=-
reg query "HKLM\Software\Microsoft\Windows NT\CurrentVersion\WinLogon" 
reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /s 
reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce" /s 
reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnceEx" /s 
reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\RunServices" /s 
reg query "HKLM\Software\Microsoft\Windows NT\CurrentVersion\AeDebug" /v Debugger

reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\URL\Prefixes" /s
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\URL\DefaultPrefix" /s
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects" /s

reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /s 
reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce" /s 
reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnceEx" /s 
reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\RunServices" /s 




goto end
:noname
echo response HOSTNAME ????????

:end
echo %time%
echo %date%




