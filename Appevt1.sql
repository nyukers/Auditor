SELECT TimeGenerated, EventID, EventCategory, EventCategoryName, SourceName, ComputerName, RESOLVE_SID(SID) as User, Strings
INTO BARSNBU1_app.lp 
---FROM Application 
FROM %src%\BARSNBU1_app.evt 
WHERE ((SourceName like 'Oracle.barsrn' OR
        SourceName like 'Oracle.cashrn') AND
TimeGenerated >= '%D1%' AND TimeGenerated < '%D2%')