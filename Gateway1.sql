select Index,Text 
INTO BARSNBU1_gate1.lp 
FROM %src%\gatesrv1.log, %src%\gatesrv0.log
 where Index in (
SELECT add(Index,-1)
FROM %src%\gatesrv1.log, %src%\gatesrv0.log
 (Text like '%%! %%' AND (Text like '%%%D3%%%' OR Text like '%%%D31%%%'))
)
or Index in (
SELECT Index
FROM %src%\gatesrv1.log, %src%\gatesrv0.log
 (Text like '%%! %%' AND (Text like '%%%D3%%%' OR Text like '%%%D31%%%'))
)



