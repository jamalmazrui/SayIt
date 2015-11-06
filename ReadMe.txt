SayIt
Version 1.2
November 6, 2015
Copyright 2011 - 2015 by Jamal Mazrui
GNU Lesser General Public License (LGPL)

SayIt is a 32-bit, dual mode COM server, supporting both early and late binding to a few, targeted methods.  It enables an application to speak text directly via the API of an active screen reader, if found, or optionally via the default SAPI engine.  

The "SayIt" ProgID may be used to instantiate a COM object with this functionality.  The "Say" method is the main one, which accepts a Unicode string variant as a parameter, and then speaks accordingly.  The "UseSAPIAsBackup" property determines whether SAPI is used in the absence of a screen reader -- the default being False.

Other methods are as follows:

ScreenReaderIsActive() = Check whether any screen reader is active in memory.

JAWSIsActive() = Check whether JAWS is active.

NVDAIsActive() = Check whether NVDA is active.

SAIsActive() = Check whether System Access is active.

WEIsActive() = Check whether Window-Eyes is active.


The following files need to be placed in the same directory for SayIt to work:

SayIt.dll = the main dynamic link library (DLL).

SayIt.tlb = An accompanying type library, containing method signatures for early binding via COM.

nvdaControllerClient32.dll = DLL for NVDA support.

saapi32.dll = DLL for System Access support.


These files may be copied into any directory on a Windows computer.  They may be placed in a shared location for multiple applications to use, or be placed in the main program directory of a client application.

The COM server needs to be registered on the computer.  This may be done as an installation step of an application (installers include this capability).  It may also be done using syntax at a Windows command prompt like the following:

RegSvr32.exe SayIt.dll

PowerBASIC is a commercial compiler from powerbasic.com.  The PowerBASIC source code for this COM server is in the file SayIt.bas, which may be compiled, using PowerBASIC 10, with the batch file CompileSayIt.bat  A demonstration VBScript program is in the file TestSayIt.vbs, which may be run with the batch file RunTestSayIt.bat.

In addition to the SayIt COM server, the distribution now includes two 32-bit,  console-mode executables that use the same code base.  SayLine.exe says a string of text passed on the command line, e.g.,
SayLine.exe Hello world

SayFile.exe says the contents of a text file passed on the command line, e.g.,
SayFile.exe temp.txt

Thus, an application can produce speech messages either by calling the SayIt COM server or by running one of the SayIt executables.  SayLine.exe and SayFile.exe require two additional files in the same directory:  nvdaControllerClient32.dll and saapi32.dll.  Although the SayIt COM server cannot be invoked by a 64-bit application, the SayIt executables can.
