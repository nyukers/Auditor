SELECT Text
INTO BARSNBU1_appsrv1.lp 
FROM %src%\appserv.log
WHERE 
 ((Text like '%%%D5%%%' OR Text like '%%%D51%%%') AND 
  (Text like '%%ERROR%%' OR Text like '%%ORA-%%'))


