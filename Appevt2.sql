SELECT TimeGenerated, EventID, EventCategory, EventCategoryName, SourceName, ComputerName, RESOLVE_SID(SID) as User, Strings
INTO BARSNBU1_app.lp 
---FROM Application 
FROM %src%\BARSNBU1_app.evt 
WHERE ((SourceName like 'DrWatson' OR
        SourceName like 'SceCli'  OR
        SourceName like 'ISS'  OR
        EventTypeName like 'Error%%')  
AND
TimeGenerated >= '%D1%' AND TimeGenerated < '%D2%')