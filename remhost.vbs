strComputer = "001"
strUser =     "adminzah"
strPassword = ""

set Locator = CreateObject("WbemScripting.SWbemLocator")

Set wbemServices = Locator.ConnectServer(_
                   strComputer,_
                   "",_
                   strUser,_
                   strPassword,_
                   "","",0,null)

Set wbemObjectSet = wbemServices.InstancesOf("Win32_LogicalMemoryConfiguration")

For Each wbemObject In wbemObjectSet
    WScript.Echo "Total Physical Memory (kb): " & wbemObject.TotalPhysicalMemory
Next

WScript.Quit(EXIT_SUCCESS) 