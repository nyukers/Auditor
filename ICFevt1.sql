SELECT TimeGenerated, EventID, EventCategory, EventCategoryName, SourceName, ComputerName, RESOLVE_SID(SID) as User, Strings
INTO BARSNBU1_sec_icf.lp 
---FROM Security 
FROM %src%\BARSNBU1_sec.evt 
WHERE 
  (EventID >= 848 AND 
   TimeGenerated >= '%D1%' AND TimeGenerated < '%D2%')
