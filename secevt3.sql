SELECT TimeGenerated, EventID, EventType, EventCategory, EventCategoryName, SourceName, ComputerName, RESOLVE_SID(SID) as User, Strings
INTO BARSNBU1_sec2.lp 
FROM %src%\BARSNBU1_sec.evt 
WHERE 
 ((EventID = 630 OR EventID = 529 OR EventID = 576 OR
   EventID = 531 OR EventID = 534 OR EventID = 601 OR
   EventID = 516 OR EventID = 517 OR EventID = 521 OR
   EventID = 612 OR EventID = 530
  )AND
   TimeGenerated >= '%D1%' AND TimeGenerated < '%D2%')
