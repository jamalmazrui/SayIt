@echo off
cls
if exist SayFile.exe del SayFile.exe
c:\pbwin10\bin\pbwin.exe /ic:\pbwin10\winapi;c:\SayIt c:\SayIt\SayFile.bas
if exist SayFile.exe win2con.exe SayFile.exe
