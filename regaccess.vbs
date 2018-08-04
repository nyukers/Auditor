' читаем права на ключи реестра

const KEY_QUERY_VALUE = &H0001
const KEY_SET_VALUE = &H0002
const KEY_CREATE_SUB_KEY = &H0004
const DELETE = &H00010000


const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
'strComputer = "szi01"
Set StdOut = WScript.StdOut

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")

strKeyPath = "SYSTEM\CurrentControlSet"


oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_QUERY_VALUE, bHasAccessRight
If bHasAccessRight = True Then
    StdOut.WriteLine "Have Query Value Access Rights on Key"
Else
    StdOut.WriteLine "Do Not Have Query Value Access Rights on Key"
End If	

oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_SET_VALUE, bHasAccessRight
If bHasAccessRight = True Then
    StdOut.WriteLine "Have Set Value Access Rights on Key"
Else
    StdOut.WriteLine "Do Not Have Set Value Access Rights on Key"
End If	

oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_CREATE_SUB_KEY, bHasAccessRight
If bHasAccessRight = True Then
    StdOut.WriteLine "Have Create SubKey Access Rights on Key"
Else
    StdOut.WriteLine "Do Not Have Create SubKey Access Rights on Key"
End If

oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, DELETE, bHasAccessRight
If bHasAccessRight = True Then
    StdOut.WriteLine "Have Delete Access Rights on Key"
Else
    StdOut.WriteLine "Do Not Have Delete Access Rights on Key"
End If

