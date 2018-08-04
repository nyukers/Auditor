SELECT Text
INTO BARSNBU1_lstnrBARS.lp 
FROM %src%\listener.log
WHERE 
 (Text like '%%%BARSRN%%%' AND  
 (Text like '%%%D4%%%' OR Text like '%%%D41%%%'))  




