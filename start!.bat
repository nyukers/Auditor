@echo off
cls
if %1.==. goto noname
if %2.==. goto nopath

goto next
:nopath
md c:\test
set par2=c:\test
@echo %par2%

:next
call responce!.bat %1>%par2%\%1_resp.txt

call hotfixes!.bat >%par2%\%1_hotfixes.txt

goto ie

@echo Eventlogs dump by binary format
cscript evt.vbs /Host:%1 /PathDump:%par2%

@echo Eventlogs dump by text format
dumpel -l security    -f %par2%\%1_sec.txt 
dumpel -l application -f %par2%\%1_app.txt 
dumpel -l system      -f %par2%\%1_sys.txt 

@echo Registries dump
reg export HKLM %par2%\HKLM.reg
reg export HKCU %par2%\HKCU.reg
reg export HKCR %par2%\HKCR.reg
reg export HKCC %par2%\HKCC.reg
reg export HKU  %par2%\HKUsers.reg

:ie
@echo securitypolicy -=-=-=-=-=--=-=-=-
call secpol.bat %par2%\%1

@echo -
@echo index.dat IE history -=-=-=-=-=--=-=-=-
iehistory /c
ren  iehistory.txt %1_ie.log
move %1_ie.log %par2%

move %1_ads.log %par2%
del %par2%\%1_cfg.sec

:n2
md5deep %par2%\*.* >%1_.md5
move %1_.md5 %par2%


:noname
@echo noname!

