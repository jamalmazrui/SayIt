Option Explicit

Dim oSay

' Create an instance of The SayIt object
Set oSay = CreateObject("SayIt")

' Check whether a screen reader is active
' True = -1 in the world of COM
MsgBox oSay.ScreenReaderIsActive(), 0, "Result of ScreenReaderIsActive Method"

' Verify the True, default state of the UseSAPIAsBackup property
' This ensures that speech occurs via SAPI if no active screen reader is found
MsgBox oSay.UseSAPIAsBackup, 0, "Default State of UseSAPIAsBackup Property"

' Say text using JAWS, NVDA, System Access, WindowEyes, or SAPI
oSay.Say("Hello world")

' Close the COM object
Set oSay = Nothing
