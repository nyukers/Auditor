--- EventCategoryName like 'Изменение политики' ---
SELECT TimeGenerated, EventID, EventCategory, EventCategoryName, SourceName, ComputerName, RESOLVE_SID(SID) as User, Strings
INTO BARSNBU1_sec.lp 
FROM %src%\BARSNBU1_sec.evt 
WHERE 
 ((EventCategory = 6 OR 
   EventID = 517 OR 
   EventTypeName like 'Failure Audit%%') AND 
   TimeGenerated >= '%D1%' AND TimeGenerated < '%D2%')
