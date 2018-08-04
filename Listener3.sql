SELECT Text
INTO BARSNBU1_listoad.lp 
FROM %src%\listener.log
WHERE 
 (Text like '%%%TOAD%%%' AND  
 (Text like '%%%D4%%%' OR Text like '%%%D41%%%'))  