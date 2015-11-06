' SayFile
' Version 1.2
' November 6, 2015
' Copyright 2011 - 2015 by Jamal Mazrui
' GNU Lesser General Public License (LGPL)

#COMPILE EXE
#Include "SayIt.bas"

Function FileToString(ByVal zFile As ASCIIZ * 260) As String
Local hFile, iSize As Dword
Local sReturn As String

If IsFalse IsFile(zFile) Then Exit Function

hFile =FreeFile
Open zFile For Binary as hFile
iSize = Lof(hFile)
Get$ hFile, iSize, sReturn
Close hFile
Function = sReturn
End Function

Function PBMain()
Local oSay As ISayIt
Local sFile, sText As String

sFile = Trim$(Command$, ANY $DQ)
sText = "Blank"
If IsFile(sFile) Then sText = FileToString(sFile)
oSay = Class "SayIt"
oSay.Say(sText)
End Function
