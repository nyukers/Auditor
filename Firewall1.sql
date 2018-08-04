SELECT Text
INTO BARSNBU1_wall1.lp 
FROM %src%\pfirewall.log
WHERE 
(Text like '%%%D3%%%' OR Text like '%%%D31%%%')
