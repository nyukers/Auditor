'*************************************************************************
'*
'* Copyright (c) Vinnitsa SZI, 2007. All rights reserved. 
'*
'* Module Name:  audi2007.vbs 
'*
'* Abstract:     Enables an administrator to cross-query/view all existing
'*               events in a any log(s) by LogParser.
'*
'*************************************************************************


Dim WSHSHell
Set WSHShell = CreateObject("WScript.Shell")

' ‘ормируем beforeyester
y = Cstr(date-2)
mm_y = mid(y,4,2)
dd_y = mid(y,1,2)
yy_y = Year(y)
befores = "xx_"&dd_y&"/"&mm_y&"/"&yy_y

' ‘ормируем вчерашнюю дату
y = Cstr(date-1)
mm_y = mid(y,4,2)
dd_y = mid(y,1,2)
yy_y = Year(y)
yesterday = "xx_"&dd_y&"/"&mm_y&"/"&yy_y

' ‘ормирyeм сегодн€шнюю дату
y = Cstr(date)
mm_t = mid(y,4,2)
dd_t = mid(y,1,2)
yy_t = Year(y)
today = "xx_"&dd_t&"/"&mm_t&"/"&yy_t

WScript.Echo today  
WScript.Echo yesterday 
WScript.Echo befores


Sr="c:\a\Logparser\test.bat"&" "&today&" "&yesterday&" "&befores
code=WSHShell.Run(Sr, 1, True)
WScript.Echo "Retcode: ",code




