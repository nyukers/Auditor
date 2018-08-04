select Text 
INTO BARSNBU1_gate3.lp 
FROM %src%\gatesrv1.log, %src%\gatesrv0.log
WHERE
 (Text like '%%# %%' AND (Text like '%%%D3%%%' OR Text like '%%%D31%%%'))




