' SayLine
' Version 1.2
' November 6, 2015
' Copyright 2011 - 2015 by Jamal Mazrui
' GNU Lesser General Public License (LGPL)

#COMPILE EXE
#Include "SayIt.bas"

Function PBMain()
Local oSay As ISayIt
Local sText As String

sText = Command$
If Len(sText) = 0 Then sText = "Blank"
oSay = Class "SayIt"
oSay.Say(sText)
End Function
