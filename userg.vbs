ON ERROR RESUME NEXT

strComputer = "."
Set colGroups = GetObject("WinNT://" & strComputer & "")
colGroups.Filter = Array("group")

   Wscript.Echo "User   :   Group"
   Wscript.Echo "----------------"
For Each objGroup In colGroups
     For Each objUser in objGroup.Members
        If objUser.name = "234459" Then
            Wscript.Echo objUser.Name & ": "& objGroup.Name
        End If
    Next
Next

