' Copyright (c) 1997-1999 Microsoft Corporation
'***************************************************************************
' 
' WMI Sample Script - List event log events of a particular type (VBScript)
'
' This script demonstrates how to retrieve the events of a particular type (Event code)
' from the Win32_NTLogEvent class
'
'***************************************************************************
CONST EVENTTYPE_ERROR             = "1"
CONST EVENTTYPE_WARNING           = "2"
CONST EVENTTYPE_INFORMATION       = "3"
CONST EVENTTYPE_SUCCESSAUDIT      = "4"
CONST EVENTTYPE_FAILUREAUDIT      = "5"

Set EventSet = GetObject("winmgmts:").ExecQuery( _
"select * from Win32_NTLogEvent where (EventCode=1001 and EventType=3)")

if (EventSet.Count = 0) then WScript.Echo "No Events"

for each LogEvent in EventSet
        WScript.Echo "Event Code : " & LogEvent.EventCode
        WScript.Echo "Event Type : " & LogEvent.EventType
        Wscript.Echo "Time: " &  WMIDateStringToDate(LogEvent.TimeGenerated) 
'	WScript.Echo "Event Number: " & LogEvent.RecordNumber
'	WScript.Echo "Log File: " & LogEvent.LogFile
	WScript.Echo "Type: " & LogEvent.Type 
	WScript.Echo "Message: " & LogEvent.Message
'	WScript.Echo "Time: " & LogEvent.TimeGenerated
	WScript.Echo ""
next

Function WMIDateStringToDate(dtmDate)
    WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
         Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
         & " " & Mid (dtmDate, 9, 2) & ":" & _
         Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate, _
         13, 2))
End Function
