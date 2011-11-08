net stop "World Wide Web Publishing Service"
pause
regsvr32 /s ztitools.dll
net start "World Wide Web Publishing Service"
