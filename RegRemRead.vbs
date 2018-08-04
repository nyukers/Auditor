const HKEY_CURRENT_USER = &H80000001
const HKEY_LOCAL_MACHINE = &H80000002

'strComputer = "."
strComputer = "szi01"
strUser     = "adminzah"
strPassword = ""

'ON ERROR RESUME NEXT
Set StdOut = WScript.StdOut

set Locator = CreateObject("WbemScripting.SWbemLocator")

Set objWMIService = Locator.ConnectServer(_
                   strComputer,_
                   "root\default",_
                   strUser,_
                   strPassword,_
                   "", "", 0, null)


Set oReg=objWMIService.Get("StdRegProv")

'strKeyPath = "Console"
'strValueName = "HistoryBufferSize"
'oReg.GetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
'StdOut.WriteLine "Current History Buffer Size: " & dwValue 


strKeyPath = "SOFTWARE\System Admin Scripting Guide"
strValueName = "String Value Name"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
StdOut.WriteLine "String Value Name: " & strValue 

strValueName = "DWORD Value Name"
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
StdOut.WriteLine "DWORD Value: " & dwValue
