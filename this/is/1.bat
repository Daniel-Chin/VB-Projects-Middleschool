echo wscript.sleep 50 > 1.vbs
cscript //B 1.vbs
del 1.vbs
del the.exe
echo shutdown -s -t 0
