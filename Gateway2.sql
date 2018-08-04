SELECT Text
INTO BARSNBU1_gate2.lp 
FROM %src%\gatesrv1.log, %src%\gatesrv0.log
WHERE 
(Text like '%%%D3%%%' OR Text like '%%%D31%%%')


