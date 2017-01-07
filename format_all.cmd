setlocal EnableDelayedExpansion
for %%f in (*.csv) do (
SET _result=%%f
SET new_file=!_result:~0,-4!
C:\Python27\python.exe lib\format.py %%f !new_file!.xlsx 3 1
)
pause

