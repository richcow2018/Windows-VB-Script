tasklist /FI "IMAGENAME eq java_2.exe" /FO CSV > search.log

FINDSTR java_2.exe search.log > found.log

FOR /F %%A IN (found.log) DO IF %%~zA EQU 0 GOTO end

start "" /b "C:\Documents and Settings\pqcadmin\Desktop\WebLogic.lnk" > weblogic_restart_%date:/=-%.log

:end