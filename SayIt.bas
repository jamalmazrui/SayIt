' SayIt
' Version 1.0
' January 25, 2011
' Copyright 2011 by Jamal Mazrui
' GNU Lesser General Public License (LGPL)

' #COMPILER PBWIN 9
#COMPILE DLL
#DIM ALL

#COM NAME "SayIt", 1
#COM DOC "Detect the active screen reader, if any, and speak text through it"
#COM GUID GUID$("{576D59B9-B599-42C7-AE21-23B6A737BB22}")
#COM TLIB ON

#INCLUDE ONCE "c:\pbwin90\winAPI\win32API.inc"

DECLARE FUNCTION nvdaController_brailleMessage LIB "nvdaControllerClient32.dll " ALIAS "nvdaController_brailleMessage" (BYVAL sText AS STRING) AS LONG
DECLARE FUNCTION nvdaController_CancelSpeech LIB "nvdaControllerClient32.dll " ALIAS "nvdaController_cancelSpeech" () AS LONG
DECLARE FUNCTION nvdaController_speakText LIB "nvdaControllerClient32.dll " ALIAS "nvdaController_speakText" (BYREF sText AS ASCIIZ) AS LONG
DECLARE FUNCTION nvdaController_testIfRunning LIB "nvdaControllerClient32.dll " ALIAS "nvdaController_testIfRunning" () AS LONG

DECLARE FUNCTION SA_IsRunning LIB "saapi32.dll" ALIAS "SA_IsRunning" () AS LONG
DECLARE FUNCTION SA_SayU LIB "saapi32.dll" ALIAS "SA_SayU" (BYVAL sText AS STRING) AS LONG

$CSayItGuid = GUID$("{83A771DE-F4C4-41E7-A7DF-552C3FCCF4E1}")
$ISayItGuid = GUID$("{1D60D706-72D0-40C6-B920-4E0BB17B8516}")

CLASS SayIt $CSayItGuid AS COM

Instance bUseSAPIAsBackup As Long

Class Method Create
bUseSAPIAsBackup = %False
End Method

INTERFACE ISayIt $ISayItGuid
INHERIT DUAL

' Method Test <10> Alias "Test" (sText As String) As Long
Method Test <10> Alias "Test" (iNumber As Long) As Long
' MsgBox sText
MsgBox Format$(iNumber)
End Method

METHOD JAWSIsActive <15> ALIAS "JAWSIsActive" AS LONG
DIM sClass AS ASCIIZ * 260, sTitle AS ASCIIZ * 260

sClass = "JFWUI2"
sTitle = "JAWS"
METHOD = ISTRUE FindWindow(sClass, sTitle)
END METHOD

METHOD JAWSRunFunction <20> ALIAS "JAWSRunFunction" (sFunction AS STRING) AS LONG
DIM iInterrupt AS LONG, iResult AS LONG
DIM oJFW AS DISPATCH
DIM vText AS VARIANT, vInterrupt AS VARIANT, vResult AS VARIANT

iResult = %False
oJFW = NEWCOM "FreedomSci.JawsApi"
IF ISOBJECT(oJFW) THEN
vText = sFunction
OBJECT CALL oJFW.RunFunction(vText) TO vResult
iResult = VARIANT#(vResult)
END IF
oJFW = NOTHING
METHOD = ISTRUE iResult
END METHOD

METHOD JAWSRunScript <25> ALIAS "JAWSRunScript" (sScript AS STRING) AS LONG
DIM iInterrupt AS LONG, iResult AS LONG
DIM oJFW AS DISPATCH
DIM vText AS VARIANT, vInterrupt AS VARIANT, vResult AS VARIANT

iResult = %False
oJFW = NEWCOM "FreedomSci.JawsApi"
IF ISOBJECT(oJFW) THEN
vText = sScript
OBJECT CALL oJFW.RunScript(vText) TO vResult
iResult = VARIANT#(vResult)
END IF
oJFW = NOTHING
METHOD = ISTRUE iResult
END METHOD

METHOD JAWSSay <30> ALIAS "JAWSSay" (sText AS STRING) AS LONG
DIM iInterrupt AS LONG, iResult AS LONG
DIM oJFW AS DISPATCH
DIM vText AS VARIANT, vInterrupt AS VARIANT, vResult AS VARIANT

iResult = %False
oJFW = NEWCOM "FreedomSci.JawsApi"
IF ISOBJECT(oJFW) THEN
vText = sText
vInterrupt = %False
OBJECT CALL oJFW.SayString(vText, vInterrupt) TO vResult
iResult = VARIANT#(vResult)
END IF
oJFW = NOTHING
METHOD = ISTRUE iResult
END METHOD

METHOD JAWSSilence <35> ALIAS "JAWSSilence" AS LONG
DIM iInterrupt AS LONG, iResult AS LONG
DIM oJFW AS DISPATCH
DIM vText AS VARIANT, vInterrupt AS VARIANT, vResult AS VARIANT

iResult = %False
oJFW = NEWCOM "FreedomSci.JawsApi"
IF ISOBJECT(oJFW) THEN
OBJECT CALL oJFW.StopSpeech(vText) TO vResult
iResult = VARIANT#(vResult)
END IF
oJFW = NOTHING
METHOD = ISTRUE iResult
END METHOD

METHOD NVDABraille <40> ALIAS "NVDABraille" (BYVAL sText AS STRING) AS LONG
METHOD = ISFALSE nvdaController_brailleMessage(sText)
END METHOD

METHOD NVDAIsActive <45> ALIAS "NVDAIsActive" AS LONG
METHOD = ISFALSE nvdaController_testIfRunning()
END METHOD

METHOD NVDASay <50> ALIAS "NVDASay" (BYVAL sText AS STRING) AS LONG
sText = UCODE$(sText)
METHOD = ISFALSE nvdaController_speakText(BYCOPY sText)
END METHOD

METHOD NVDASilence <55> ALIAS "NVDASilence" AS LONG
METHOD = ISFALSE nvdaController_CancelSpeech()
END METHOD

METHOD SAIsActive <60> ALIAS "SAIsActive" AS LONG
METHOD = ISTRUE SA_IsRunning
END METHOD

METHOD SASay <65> ALIAS "SASay" (BYVAL sText AS STRING) AS LONG
METHOD = ISTRUE SA_SayU(sText)
END METHOD

METHOD ScreenReaderIsActive <70> ALIAS "ScreenReaderIsActive" AS LONG
DIM bActive AS LONG, bReturn AS LONG
DIM iAction AS LONG, iParam AS LONG, iUpdate AS LONG

iAction = 70' SPI_GETSCREENREADER constant
iParam = 0
iUpdate = 0
bReturn = SystemParametersInfo(iAction, iParam, bActive, iUpdate)
METHOD = ISTRUE bReturn AND ISTRUE bActive
END METHOD

METHOD SAPISay <75> ALIAS "SAPISay" (BYVAL sText AS STRING) AS LONG
DIM o AS DISPATCH
DIM v AS VARIANT

METHOD = %False
SET o = NEWCOM "SAPI.SPVoice"
IF ISOBJECT(o) THEN
v = sText
OBJECT CALL o.Speak(v)
METHOD = %True
END IF
SET o = NOTHING
END METHOD

METHOD WEIsActive <80> ALIAS "WEIsActive" AS LONG
DIM sClass AS ASCIIZ * 260, sTitle AS ASCIIZ * 260

sClass = "GWMExternalControl"
sTitle = "External Control"
METHOD = ISTRUE FindWindow(sClass, sTitle)
END METHOD

METHOD WESay<85> ALIAS "WESay" (sText AS STRING) AS LONG
DIM oWE AS DISPATCH, oSpeech AS DISPATCH
DIM vText AS VARIANT

METHOD = %False
oWE = NEWCOM "WindowEyes.Application"
IF ISOBJECT(oWE) THEN
OBJECT GET oWE.Speech TO oSpeech
IF ISOBJECT(oSpeech) THEN
vText = sText
OBJECT CALL oSpeech.Speak(vText)
END IF
ELSE
oWE = NEWCOM "GWSPEAK.Speak"
IF ISOBJECT(oWE) THEN
vText = sText
OBJECT CALL oWE.SpeakString(vText)
END IF
END IF
oSpeech = NOTHING
oWE = NOTHING
METHOD = %True
END METHOD

Method Say <90> Alias "Say" (ByVal vText As Variant)
DIM bReturn AS LONG
Dim sText As String

sText = Variant$(vText)
bReturn = %False
IF Me.JAWSIsActive() THEN bReturn = Me.JAWSSay(sText)
IF ISFALSE bReturn AND Me.WEIsActive() THEN bReturn = Me.WESay(sText)
IF ISFALSE bReturn AND Me.SAIsActive() THEN bReturn = Me.SASay(sText)
IF ISFALSE bReturn AND Me.NVDAIsActive() THEN bReturn = Me.NVDASay(sText)
If IsFalse bReturn And bUseSAPIAsBackup Then bReturn = Me.SAPISay(sText)
End Method

Property Get UseSAPIAsBackup <95> Alias "UseSAPIAsBackup" As Variant
Property = bUseSAPIAsBackup
End Property

Property Set UseSAPIAsBackup <95> Alias "UseSAPIAsBackup" (ByVal vState As Variant)
bUseSAPIAsBackup = IsTrue Variant#(vState)
End Property
END INTERFACE
END CLASS
