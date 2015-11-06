@echo off
cls
if exist SayLine.exe del SayLine.exe
c:\pbwin10\bin\pbwin.exe /ic:\pbwin10\winapi;c:\SayIt c:\SayIt\SayLine.bas
if exist SayLine.exe win2con.exe SayLine.exe
