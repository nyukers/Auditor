SELECT TimeGenerated, EventID, EventType, EventCategory, EventCategoryName, SourceName, ComputerName, RESOLVE_SID(SID) as User, Strings
INTO BARSNBU1_term.lp 
FROM %src%\BARSNBU1_sec.evt 
WHERE 
 ((EventID = 682 OR EventID = 683 OR
 ((EventID = 528 OR EventID = 529) AND EventType = 10)) AND
   TimeGenerated >= '%D1%' AND TimeGenerated < '%D2%')
