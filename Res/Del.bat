@ echo off
echo Going to DELETE file:
echo  %1
pause
DEL /F /A /Q \\?\%1
RD /S /Q \\?\%1 
echo Done. 
pause
