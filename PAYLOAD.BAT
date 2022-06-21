::Comment - Example file for how to use DET in command mode
::
::This example reflects the need of an ECU to first do an Automated SWDL.
::After performing the automated SWDL, the tool then needs to write new values to certain DIDs based upon existing values.
:: The tool first checks to see if DID 1234 = a and DID 4567 = k. If so, it updates the values in each to b and l.
:: If the first DID checks don't match, the tool then checks to see if DID 1234 = c and DID 4567 = m, and updates to d and n if so.
:: ::The batch file loads a specific DCS file so that the DET will use the associated hot keys, DVS file, etc.
::
::Note that when checking error level (i.e., IF ERRORLEVEL X), this evaluates to true if the ERRORLEVEL is >= X.
::
@echo off
@echo.
@echo Running DET in command line mode
@echo.
:: Using -01 runs the DVS file (referenced in new DefaultSettings.dcs) to actually do the download
::"C:\Program Files (x86)\Ford Netcom\Diagnostic Engineering Tool\diagnosticengineeringtool.exe" –f "C:\SDLCCANFLASH\BLU.vbs" -02
::IF ERRORLEVEL 1 GOTO PROBLEM
:: Then load file that runs a DVS that checks for Type 1 (DID 1234 = a, DID 4567 = k),
:: and has hotkeys to write desired replacement values (e.g., DID 1234 = b, DID 4567 = l)
:: and has VBS ACTION of SENDALLHOTKEYS
:: and has post SWDL verification of (e.g., DID 1234 = b, DID 4567 = l)
:: Using -01 runs the DVS file
"C:\Program Files (x86)\Ford Netcom\Diagnostic Engineering Tool\diagnosticengineeringtool.exe"  "C:\SDLCCANFLASH\BLU.vbs" -02
IF ERRORLEVEL 1 GOTO CHECKTYPE2
:: It will go to CHECKTYPE2 if it did not find type 1, or if a problem writing
echo Successfully update to Type 1
GOTO END
:CHECKTYPE2
:: Then copy file that runs a DVS that checks for Type 2 (DID 1234 = c, DID 4567 = m),
:: and has hotkeys to write desired replacement values (e.g., DID 1234 = d, DID 4567 = n)
:: and has VBS ACTION of SENDALLHOTKEYS
:: and has post SWDL verification of (e.g., DID 1234 = d, DID 4567 = n)
IF ERRORLEVEL 1 GOTO FILECOPYPROBLEM
:: Using -01 runs the DVS file
IF ERRORLEVEL 1 GOTO UNKNOWNTYPE
echo Successfully update to Type 2
GOTO END
:: This is reached if not Type 1 or Type 2 match
:UNKNOWNTYPE
echo Was not type 1 or type 2
GOTO END
:FILECOPYPROBLEM
echo Error copying file
GOTO END
:PROBLEM
IF ERRORLEVEL 9 GOTO UMSFAIL
IF ERRORLEVEL 8 GOTO DCSNOTFOUND
IF ERRORLEVEL 7 GOTO ECDFAIL
IF ERRORLEVEL 6 GOTO BUSQUERYFAIL
IF ERRORLEVEL 5 GOTO HOTKEYFAIL
IF ERRORLEVEL 4 GOTO VBSFAIL
IF ERRORLEVEL 3 GOTO DVSFAIL
IF ERRORLEVEL 2 GOTO INVPARAM
IF ERRORLEVEL 1 GOTO INVFORMAT
:UMSFAIL
echo User Message sequence Transmittal Failed
GOTO END
:DCSNOTFOUND
echo DCS File Not Found
GOTO END
:ECDFAIL
echo ECD Transmittal Failed
GOTO END
:BUSQUERYFAIL
echo Bus Query Transmittal Failed
GOTO END
:HOTKEYFAIL
echo Hot Key Transmittal Failed
GOTO END
:VBSFAIL
echo VBS Failed
GOTO END
:DVSFAIL
echo DVS Failed
GOTO END
:INVPARAM
echo Invalid Parameters
GOTO END
:INVFORMAT
echo Invalid format
GOTO END
:END
pause