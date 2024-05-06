@ECHO OFF
taskkill /IM castle.exe /f
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\castle" /Q
del C:\castle.exe
del C:\castle.bat