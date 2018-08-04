SELECT Text
INTO BARSNBU1_lstnrCASH.lp 
FROM %src%\listener.log
WHERE 
 (Text like '%%%CASHRN%%%' AND  
 (Text like '%%%D4%%%' OR Text like '%%%D41%%%'))




