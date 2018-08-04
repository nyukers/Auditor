'Option Explicit
'Dim i,Arg,objArgs,s,objNamedArgs,objUnnamedArgs  ' Объявляем переменные

strComputer = "szi01"
strUser =     "adminzah"
strPassword = ""

'Const CONVERT_TO_LOCAL_TIME = True
'Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
'Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime")

Set objArgs = WScript.Arguments  ' Создаем объект WshArguments
Set objNamedArgs=objArgs.Named  ' Создаем объект WshNamed

' Проверяем, существует ли аргумент /Date:
If objNamedArgs.Exists("Date") Then
  s = objNamedArgs("Date") & vbCrLf
End If

'DateToCheck = CDate("05/12/2003")-1
'on Yesterday!
DateToCheck = CDate(objNamedArgs("Date")) - objNamedArgs("DateStart")

WScript.Echo strComputer & ", checkDate is " & DateToCheck

'dtmStartDate.SetVarDate DateToCheck, CONVERT_TO_LOCAL_TIME
'dtmEndDate.SetVarDate DateToCheck + 1, CONVERT_TO_LOCAL_TIME

set Locator = CreateObject("WbemScripting.SWbemLocator")
Set wbemServices = Locator.ConnectServer(_
                   strComputer,_
                   "",_
                   strUser,_
                   strPassword,_
                   "","",0,null)

'Set objWMIService = wbemServices.GetObject("winmgmts:" _
'    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")


Wscript.Echo "--------------------" 
Wscript.Echo "LogFile: System" 
Wscript.Echo "--------------------" 

Set colEvents = wbemServices.ExecQuery _
    ("Select * from Win32_NTLogEvent Where " _
     & "Type = 'error' and Logfile = 'System' and TimeWritten >= '" _ 
        & DateToCheck & "' and TimeWritten < '" & DateToCheck & "'") 

'Set colEvents = wbemServices.ExecQuery _
'    ("Select * from Win32_NTLogEvent Where " _
'     & "Type = 'error' and Logfile = 'System' and TimeWritten >= '" _ 
'        & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'") 

if (colEvents.Count = 0) then WScript.Echo " No Events!"

For each objEvent in colEvents
    Wscript.Echo "Event Code: " & objEvent.EventCode
    Wscript.Echo "Event Type: " & objEvent.Type
    Wscript.Echo "Time Written: " & WMIDateStringToDate(objEvent.TimeWritten)
    Wscript.Echo "Computer Name: " & objEvent.ComputerName
    Wscript.Echo "Source Name: " & objEvent.SourceName
    Wscript.Echo "User: " & objEvent.User
    Wscript.Echo "Message: " & objEvent.Message
Next

Wscript.Echo "--------------------" 
Wscript.Echo "LogFile: Security" 
Wscript.Echo "--------------------" 

Set colEvents = wbemServices.ExecQuery _
    ("Select * from Win32_NTLogEvent Where " _
     & "Type = 'error' and Logfile = 'Security' and TimeWritten >= '" _ 
        & DateToCheck & "' and TimeWritten < '" & DateToCheck & "'") 

if (colEvents.Count = 0) then WScript.Echo " No Events!"

For each objEvent in colEvents
'    Wscript.Echo objEvent.LogFile
    Wscript.Echo "Event Code: " & objEvent.EventCode
    Wscript.Echo "Event Type: " & objEvent.Type
    Wscript.Echo "Time Written: " & WMIDateStringToDate(objEvent.TimeWritten)
'    Wscript.Echo "Category: " & objEvent.Category
'    Wscript.Echo "Record Number: " & objEvent.RecordNumber
    Wscript.Echo "Computer Name: " & objEvent.ComputerName
    Wscript.Echo "Source Name: " & objEvent.SourceName
    Wscript.Echo "User: " & objEvent.User
    Wscript.Echo "Message: " & objEvent.Message
 
Next

Wscript.Echo "--------------------" 
Wscript.Echo "LogFile: Application" 
Wscript.Echo "--------------------" 

Set colEvents = wbemServices.ExecQuery _
    ("Select * from Win32_NTLogEvent Where " _
     & "Type = 'error' and Logfile = 'Application' and TimeWritten >= '" _ 
        & DateToCheck & "' and TimeWritten < '" & DateToCheck & "'") 

if (colEvents.Count = 0) then WScript.Echo " No Events!"

For each objEvent in colEvents
    Wscript.Echo "Event Code: " & objEvent.EventCode
    Wscript.Echo "Event Type: " & objEvent.Type
    Wscript.Echo "Time Written: " & WMIDateStringToDate(objEvent.TimeWritten)
    Wscript.Echo "Computer Name: " & objEvent.ComputerName
    Wscript.Echo "Source Name: " & objEvent.SourceName
    Wscript.Echo "User: " & objEvent.User
    Wscript.Echo "Message: " & objEvent.Message
 
Next

WScript.Quit(EXIT_SUCCESS) 


Function WMIDateStringToDate(dtmDate)
    WMIDateStringToDate = CDate(Mid(dtmDate, 7, 2) & "/" & _
         Mid(dtmDate, 5, 2) & "/" & Left(dtmDate, 4) _
         & " " & Mid (dtmDate, 9, 2) & ":" & _
         Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate, _
         13, 2))
End Function

