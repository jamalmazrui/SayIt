@echo off
cls
if exist SayIt.dll del SayIt.dll
c:\pbwin10\bin\pbwin.exe /ic:\pbwin10\winapi;c:\SayIt c:\SayIt\SayIt.bas
