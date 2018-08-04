:*************************************************************************
:*
:* Copyright (c) Vinnitsa SZI, 2007. All rights reserved. 
:*
:* Module Name:  Daily.bat 
:*
:* Abstract:     Enables an administrator to cross-query/view all existing
:*               events in a any log(s) by LogParser.
:*
:*************************************************************************

set Source=c:\a

@echo %1
@echo %2
set Today=%1
set Yester=%2
set Before=%3

set week=%Today:~0,2%
set dd=%Today:~3,2%
set mm=%Today:~6,2%
set yy=%Today:~9,4%

set Date2="%yy%-%mm%-%dd% 09:00:00"
set Date51="%dd%.%mm%.%yy%"
set Date31="%yy%-%mm%-%dd%"

set yyUA=%Yester:~9,4%
set mmUA=%Yester:~6,2%
set ddUA=%Yester:~3,2%
set Date5="%ddUA%.%mmUA%.%yyUA%"
set Date3="%yyUA%-%mmUA%-%ddUA%"

set yyb=%Before:~9,4%
set mmb=%Before:~6,2%
set ddb=%Before:~3,2%
set Date1="%yyb%-%mmb%-%ddb% 00:00:00"


@IF %mmUA% EQU 04 set mmUA=äÇí
@IF %mmUA% EQU 05 set mmUA=íêÇ
@IF %mmUA% EQU 06 set mmUA=óêÇ
@IF %mmUA% EQU 07 set mmUA=ãàè
@IF %mmUA% EQU 08 set mmUA=ëêè
set Date4="%ddUA%-%mmUA%-%yyUA%"
@IF %mm% EQU 04 set mm=äÇí
@IF %mm% EQU 05 set mm=íêÇ
@IF %mm% EQU 06 set mm=óêÇ
@IF %mm% EQU 07 set mm=ãàè
@IF %mm% EQU 08 set mm=ëêè
set Date41="%dd%-%mm%-%yy%"


del *.lp /Q

::Windows Eventlog
LogParser file:Secevt3.sql?Src=%Source%+D1=%Date1%+D2=%Date2% -i:EVT -rtp:-1 -fileMode:1
LogParser file:Secevt4.sql?Src=%Source%+D1=%Date1%+D2=%Date2% -i:EVT -rtp:-1 -fileMode:1
LogParser file:Secevt5.sql?Src=%Source%+D1=%Date1%+D2=%Date2% -i:EVT -rtp:-1 -fileMode:1
LogParser file:Appevt1.sql?Src=%Source%+D1=%Date1%+D2=%Date2% -i:EVT -rtp:-1 -fileMode:1
LogParser file:Appevt2.sql?Src=%Source%+D1=%Date1%+D2=%Date2% -i:EVT -rtp:-1 -fileMode:0
LogParser file:Sysevt1.sql?Src=%Source%+D1=%Date1%+D2=%Date2% -i:EVT -rtp:-1 -fileMode:1
LogParser file:ICFevt1.sql?Src=%Source%+D1=%Date1%+D2=%Date2% -i:EVT -rtp:-1 -fileMode:1


::Windows firewall
LogParser file:firewall1.sql?Src=%Source%+D3=%Date3%+D31=%Date31% -rtp:-1 -fileMode:1

::Bars Gateway
: ‚Æ´Ï™Æ §†‚„ ¢Á•‡†
: ÆË®°™® ≠† §†‚„ ¢Á•‡†
LogParser file:gateway1.sql?Src=%Source%+D3=%Date3%+D31=%Date31% -o:CSV -fileMode:1
LogParser file:gateway2.sql?Src=%Source%+D3=%Date3%+D31=%Date31% -o:CSV -fileMode:1
LogParser file:gateway3.sql?Src=%Source%+D3=%Date3%+D31=%Date31% -o:CSV -fileMode:1


::Oracle listener
LogParser file:Listener1.sql?Src=%Source%+D4=%Date4%+D41=%Date41% -rtp:-1 -fileMode:1
LogParser file:Listener2.sql?Src=%Source%+D4=%Date4%+D41=%Date41% -rtp:-1 -fileMode:1
LogParser file:Listener3.sql?Src=%Source%+D4=%Date4%+D41=%Date41% -rtp:-1 -fileMode:1

::Application server CASH
LogParser file:appserv1.sql?Src=%Source%+D5=%Date5%+D51=%Date51% -o:CSV -fileMode:1

:111
