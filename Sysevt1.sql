SELECT TimeGenerated, EventID, EventCategory, EventCategoryName, SourceName, ComputerName, RESOLVE_SID(SID) as User, Strings
INTO BARSNBU1_sys.lp 
---FROM System 
FROM %src%\BARSNBU1_sys.evt 
WHERE ((SourceName like 'W32Time' OR
        SourceName like 'eventlog'  OR
        SourceName like 'tcpip'  OR
        SourceName like 'Application Popup'  OR
        SourceName like 'Windows File Protection'  OR
        SourceName like 'Srv'  OR
        EventTypeName like 'Error%%')  
AND
TimeGenerated >= '%D1%' AND TimeGenerated < '%D2%')