strComputer = "val01"
strUser =     "adminzah"
strPassword = ""

set Locator = CreateObject("WbemScripting.SWbemLocator")

Set wbemServices = Locator.ConnectServer(_
                   strComputer,_
                   "",_
                   strUser,_
                   strPassword,_
                   "","",0,null)

Set colComputer = wbemServices.ExecQuery("Select * from Win32_ComputerSystem")
  For Each objComputer in colComputer
      Wscript.Echo objComputer.UserName
  Next

WScript.Quit(EXIT_SUCCESS) 
