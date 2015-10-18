Option Explicit

Dim oSay

' Create an instance of The SayIt object
Set oSay = CreateObject("SayIt")

' Check whether a screen reader is active
MsgBox oSay.ScreenReaderIsActive(), 0, "Result of ScreenReaderIsActive Method"

' Verify the False, default state of the UseSAPIAsBackup property
MsgBox oSay.UseSAPIAsBackup, 0, "Default State of UseSAPIAsBackup Property"

' Assign True to that property, so that speech occurs via SAPI if no active screen reader is found
oSay.UseSAPIAsBackup = True

' Verify the new state
MsgBox oSay.UseSAPIAsBackup, 0, "New State of UseSAPIAsBackup Property After Assignment"

' Say text using JAWS, NVDA, System Access, WindowEyes, or SAPI
oSay.Say("Hello world")

' Close the COM object
Set oSay = Nothing
