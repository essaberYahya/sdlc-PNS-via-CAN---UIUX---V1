@echo off
@echo.
@echo Running DET in command line mode
@echo.
:: Using -01 runs the DVS file (referenced in new DefaultSettings.dcs) to actually do the download
"C:\Program Files (x86)\Ford Netcom\Diagnostic Engineering Tool\diagnosticengineeringtool.exe" -f "C:\sdlc PNS via CAN - UIUX - V1\autoload\5CTH12F10_WRITE_PN.dcs" -03
IF ERRORLEVEL 1 GOTO PROBLEM

:: Then copy file that runs a DVS that checks for Type 2 (DID 1234 = c, DID 4567 = m),
:: and has hotkeys to write desired replacement values (e.g., DID 1234 = d, DID 4567 = n)
:: and has VBS ACTION of SENDALLHOTKEYS
:: and has post SWDL verification of (e.g., DID 1234 = d, DID 4567 = n)
IF ERRORLEVEL 1 GOTO FILECOPYPROBLEM
:: Using -01 runs the DVS file
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
EXIT