setlocal EnableDelayedExpansion
for %%f in (*.csv) do (
SET _result=%%f
SET new_file=!_result:~0,-4!
C:\Python27\python.exe py\format.py %%f !new_file!.csv 5 0 1
)
pause

