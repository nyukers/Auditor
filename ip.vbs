strComputer = "."

Set objWMIService = GetObject _
    ("winmgmts:" & "!\\" & strComputer & "\root\cimv2")

' делаем запрос на адаптери через которые сипользуется протокол IP
' без IPEnabled перечўсляются все адаптери

Set colAdapters = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each objAdapter in colAdapters
    Wscript.Echo "-------------------------------"
    Wscript.Echo "Host name query: " & strComputer
    Wscript.Echo "Host name DNS  : " & objAdapter.DNSHostName
    Wscript.Echo "-------------------------------"
    Wscript.Echo "MAC address: " & objAdapter.MACAddress
    Wscript.Echo "Description: " & objAdapter.Description
    
    If Not IsNull(objAdapter.IPAddress) Then
        For i = LBound(objAdapter.IPAddress) To UBound(objAdapter.IPAddress)
            Wscript.Echo "IP address: " & objAdapter.IPAddress(i)
        Next
    End If
    If Not IsNull(objAdapter.IPSubnet) Then
        For i = LBound(objAdapter.IPSubnet) To UBound(objAdapter.IPSubnet)
            Wscript.Echo "Subnet: " & objAdapter.IPSubnet(i)
        Next
    End If
    If Not IsNull(objAdapter.DefaultIPGateway) Then
        For i = LBound(objAdapter.DefaultIPGateway) To UBound(objAdapter.DefaultIPGateway)
            Wscript.Echo "Default gateway: " & objAdapter.DefaultIPGateway(i)
        Next
    End If

    Wscript.Echo "DHCP enabled: " & objAdapter.DHCPEnabled
    Wscript.Echo "DHCP server: " & objAdapter.DHCPServer
    
    If Not IsNull(objAdapter.DNSServerSearchOrder) Then
        For i = LBound(objAdapter.DNSServerSearchOrder) To UBound(objAdapter.DNSServerSearchOrder)
            Wscript.Echo "DNS server: " & objAdapter.DNSServerSearchOrder(i)
        Next
    End If

    Wscript.Echo "DNS domain: " & objAdapter.DNSDomain
'    Wscript.Echo "DNS suffix search list: " & objAdapter.DNSDomainSuffixSearchOrder
    Wscript.Echo "Primary WINS server: " & objAdapter.WINSPrimaryServer
    Wscript.Echo "Secondary WINS server: " & objAdapter.WINSSecondaryServer
'    Wscript.Echo "Lease obtained: " & objAdapter.DHCPLeaseObtained
'    Wscript.Echo "Lease expires: " & objAdapter.DHCPLeaseExpires
Next
