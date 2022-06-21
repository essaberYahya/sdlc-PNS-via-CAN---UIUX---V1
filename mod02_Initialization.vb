Imports System.Threading
Imports System.Reflection
Imports System.IO
Imports ApiCall

Module mod02_Initialization

    ' *** Public VB.net Variables
    Public rpVbNetIsDebug As Boolean = False                        ' ***
    Public rpVbNetStartupPath As String = Nothing                   ' ***
    Public rpVbNetFileName As String = Nothing                      ' ***
    Public rpVbNetBuildDate As String = Nothing                     ' ***
    Public rpVbNetCurrentDirectory As String = Nothing

    ' *** Public RunPack Variables
    Public rpRpkSystemSerialNumber As String = Nothing              ' ***
    Public rpRpkOperatorName As String = Nothing                    ' ***
    Public rpRpkTpgmName As String = Nothing                        ' ***
    Public rpRpkTpgmVariant As String = Nothing                     ' ***
    Public rpRpkSerialNumberIsEnable As String = NO                 ' ***
    Public rpRpkIsDualStage As String                               ' ***
    Public rpRpkBayNumber As Integer = 0                            ' ***
    Public rpRpkBoardNumber As Integer = 0                          ' ***
    Public rpRpkSiteOnBay() As Integer = Nothing                   ' ***
    Public rpRpkBayInUse As Integer = 0                             ' ***
    Public rpRpkSiteInUse As Integer = 0                            ' ***
    Public rpMsgLogClass As MsgPrintLogClass = New MsgPrintLogClass(2, 98)                           ' ***

    ' *** Public INI File Variables
    Public rpIniTpgmDebugMode As String                             ' ***
    Public rpIniBayToBeDebug As Integer                             ' ***
    Public rpIniDebugPrint As String                                ' ***
    Public rpIniPrintInterTestResult As String                      ' ***
    Public rpIniPrintErrorOnFile As String                          ' ***
    Public rpIniPrintTraceOnFile As String                          ' ***

    Public rpIniITAC_Enable As String
    Public rpIniICT_Enable As String                                ' ***
    Public rpIniOBP_Enable As String                                ' ***
    Public rpIniFCT_Enable As String                                ' ***

    ' *** Public UUT Variables
    Public rpUutIctSerialNumber(11) As String                         ' ***
    Public rpUutIctResult(11) As Long                                 ' ***

    'Public rpUutRealSerialNumber(3) As String
    Public rpUutFctSerialNumber(11) As String                         ' ***
    Public rpUutFctResult(11) As Long                                 ' ***
    Public rVBFFilePath As String                         ' ***

    Public Function fTPGM_Initializiation() As Long

        Dim rError As String = Nothing
        Dim rIndex As Integer

        Try

            ' --- Set use SITE 0
            UseSiteWrite(0)

            Thread.Sleep(100)

            ' *** Set Function Result fail
            fTPGM_Initializiation = FAIL

            ' *** Set New Current Directory
            Directory.SetCurrentDirectory(Application.StartupPath)

            ' --- Get RunPack Bay in use
            UseBayRead(rpRpkBayInUse)

            ' --- Check if VB.net project is compiled in Debug Mode
#If CONFIG = "Debug" Then
            'If rpRpkBayInUse = 2 Then MsgPrintLogIfEn("--- OBP VB.Net Mode:                         DEBUG", 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then rpMsgLogClass.MsgPrintLogMultyBay("--- OBP VB.Net Mode", rpRpkBayInUse, "DEBUG")
            rpVbNetIsDebug = True
#End If
            ' --- Check if VB.net project is compiled in Released Mode
#If CONFIG = "Release" Then
            'If rpRpkBayInUse = 2 Then MsgPrintLogIfEn("--- OBP VB.Net Mode:         @FG{Navy}RELEASE", 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then rpMsgLogClass.MsgPrintLogMultyBay("--- OBP VB.Net Mode", rpRpkBayInUse, "@FG{Navy}RELEASE")
            rpVbNetIsDebug = False
#End If
            ' ***********************************************************************************************************************************************
            ' ***  Get Assembly Information
            ' ***********************************************************************************************************************************************
            Dim rVbNetFileNameInfo As IO.FileInfo

            ' --- Get Tpgm StartUp Path
            rpVbNetStartupPath = Application.StartupPath
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- TPGM StartUp Path:           " & rpVbNetStartupPath, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- TPGM StartUp Path", rpRpkBayInUse, rpVbNetStartupPath)

            ' --- Get Tpgm Name
            rpVbNetFileName = System.Reflection.Assembly.GetExecutingAssembly.Location
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- TPGM File Name:              " & rpVbNetFileName, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- TPGM File Name", rpRpkBayInUse, rpVbNetFileName)

            ' --- Get Tpgm Build Date
            rVbNetFileNameInfo = New IO.FileInfo(rpVbNetFileName)
            rpVbNetBuildDate = Format(rVbNetFileNameInfo.LastWriteTime, "yyyy/MM/dd - HH:mm:ss")
            'If rpRpkBayInUse = 2 Then MsgPrintLogIfEn("--- TPGM Build Date:         " & rpVbNetBuildDate, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then rpMsgLogClass.MsgPrintLogMultyBay("--- TPGM Build Date", rpRpkBayInUse, rpVbNetBuildDate)

            'If rpRpkBayInUse = 2 Then MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)


            ' ***********************************************************************************************************************************************
            ' *** VB.net Project Initialization
            ' ***********************************************************************************************************************************************
            ' --- Get RunPack Bay in use
            UseBayRead(rpRpkBayInUse)

            'If rpVbNetIsDebug Then MsgPrintLogIfEn("--- RPK Bay in use:              " & rpRpkBayInUse, 0, "ALWAYS")
            If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- RPK Bay in use", rpRpkBayInUse, "" & rpRpkBayInUse)

            If rpRpkBayInUse = 1 Then
                fTPGM_Initializiation = FAIL
                GoTo fTPGM_Initializiation_End
            End If

            ' --- Get RunPack Site in use
            'UseSiteRead(rpRpkSiteInUse)
            'If rpVbNetIsDebug Then MsgPrintLog("--- RPK Site in use:             " & rpRpkSiteInUse, 0)


            ' ***********************************************************************************************************************************************
            ' *** SYSTEM Initialization
            ' ***********************************************************************************************************************************************
            Dim rRpkTpgmNameLength As Integer = 0
            Dim rRpkOperatorNameLength As Integer = 0
            Dim rRpkSerialNumberLength As Integer = 0
            Dim rRpkTpgmVariantLength As Integer = 0

            ' --- Check if RunPack System Mode is Debug
            If rpVbNetIsDebug Then
                If IsDebug() Then
                    'If rpRpkBayInUse = 2 Then MsgPrintLogIfEn("--- SYS Test Mode:               @FG{Purple}DEBUG MODE", 0, "ALWAYS")
                    If rpRpkBayInUse = 2 Then rpMsgLogClass.MsgPrintLogMultyBay("--- SYS Test Mode", rpRpkBayInUse, "" & "@FG{Purple}DEBUG MODE")
                Else
                    'If rpRpkBayInUse = 2 Then MsgPrintLogIfEn("--- SYS Test Mode:               @FG{Purple}PRODUCTION MODE", 0, "ALWAYS")
                    If rpRpkBayInUse = 2 Then rpMsgLogClass.MsgPrintLogMultyBay("--- SYS Test Mode", rpRpkBayInUse, "" & "@FG{Purple}PRODUCTION MODE")
                End If
            End If

            ' --- Check if RunPack Automation Mode is Manual
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then If IsSystemManual() Then MsgPrintLogIfEn("--- SYS Automation Mode:         MANUAL MODE", 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then If IsSystemManual() Then rpMsgLogClass.MsgPrintLogMultyBay("--- SYS Automation Mode", rpRpkBayInUse, "" & "@FG{Red}MANUAL MODE")


            ' --- Get RunPack System SN
            rpRpkSystemSerialNumber = Space(64)
            rRpkSerialNumberLength = SystemSerialNumberRead(rpRpkSystemSerialNumber)
            rpRpkSystemSerialNumber = Left(rpRpkSystemSerialNumber, rRpkSerialNumberLength)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- RPK System Serial Number:    " & rpRpkSystemSerialNumber, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- RPK System Serial Number", rpRpkBayInUse, "" & rpRpkSystemSerialNumber)


            ' --- Get RunPack Operator
            rpRpkOperatorName = Space(64)
            rRpkOperatorNameLength = OperatorRead(rpRpkOperatorName)
            rpRpkOperatorName = Left(rpRpkOperatorName, rRpkOperatorNameLength)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- RPK Operator Name:           " & rpRpkOperatorName, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- RPK Operator Name", rpRpkBayInUse, "" & rpRpkOperatorName)


            ' --- Get RunPack Name
            rpRpkTpgmName = Space(64)
            rRpkTpgmNameLength = TpgmNameRead(rpRpkTpgmName)
            rpRpkTpgmName = Left(rpRpkTpgmName, rRpkTpgmNameLength)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- RPK TPGM Name:               " & rpRpkTpgmName, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- RPK TPGM Name", rpRpkBayInUse, "" & rpRpkTpgmName)


            ' --- Get RunPack Variant composition
            rpRpkTpgmVariant = Space(20)
            rRpkTpgmVariantLength = VariantCompositionRead(rpRpkTpgmVariant)
            rpRpkTpgmVariant = Left(rpRpkTpgmVariant, rRpkTpgmVariantLength)
            If rpRpkBayInUse = 2 Then If rpRpkTpgmVariant = "" Then rpRpkTpgmVariant = "@FG{Purple}NO VARIANT"
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- RPK TPGM Variant:            " & rpRpkTpgmVariant, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- RPK TPGM Variant", rpRpkBayInUse, "" & rpRpkTpgmVariant)


            ' --- Get RunPack Number of Board
            fRpkGetNumberOfBoard(rpRpkBoardNumber)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- RPK Board Number:            " & rpRpkBoardNumber, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- RPK Board Number", rpRpkBayInUse, "" & "" & rpRpkBoardNumber)


            ' --- Get Runpack is Serial Number Enabled
            fRpkSerialNumberIsEnable(rpRpkSerialNumberIsEnable)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- RPK SN is Enable:            " & rpRpkSerialNumberIsEnable, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- RPK SN is Enable", rpRpkBayInUse, "" & rpRpkSerialNumberIsEnable)


            ' --- Get RunPack Site on Bay Correlation
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)


            ' ***********************************************************************************************************************************************
            ' *** INI File Initialization
            ' ***********************************************************************************************************************************************

            ' --- Declaration Configuration INI File Path
            Dim rIniFilePath As String = "Configuration\Configuration.ini"

            ' --- Check Existing Configuration INI File
            If File.Exists(Application.StartupPath & "\" & rIniFilePath) = FAIL Then
                rIniFilePath = "@FG{Purple}INI FILE NOT FOUND"
                'If rpRpkBayInUse = 2 Then MsgPrintLog("--- INI File Path:               " & rIniFilePath, 0)
                If rpRpkBayInUse = 2 Then rpMsgLogClass.MsgPrintLogMultyBay("--- INI File Path", rpRpkBayInUse, rIniFilePath)
                fTPGM_Initializiation = False
                GoTo fTPGM_Initializiation_End
            End If

            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLog("--- INI File Path:               " & rIniFilePath, 0)

            ' --- Get INI TPGM Parameter [DEBUG] 
            rpIniTpgmDebugMode = fGetStringFromINIFile("DEBUG", "DebugMode", rIniFilePath)
            rpIniBayToBeDebug = fGetStringFromINIFile("DEBUG", "BayToBeDebug", rIniFilePath)
            rpIniDebugPrint = fGetStringFromINIFile("DEBUG", "DebugPrint", rIniFilePath)
            rpDebugPrint = CInt(rpIniDebugPrint)
            rpIniPrintInterTestResult = fGetStringFromINIFile("DEBUG", "InterTestResult", rIniFilePath)
            rpIniPrintErrorOnFile = fGetStringFromINIFile("DEBUG", "PrintErrorOnFile", rIniFilePath)
            rpIniPrintTraceOnFile = fGetStringFromINIFile("DEBUG", "PrintTraceOnFile", rIniFilePath)

            If rpIniDebugPrint = "NO" Or rpIniDebugPrint = "0" Then rpMsgLogClass.SetDebugEnable(False)

            ' --- Get INI TPGM Parameter [ITAC]
            rpITACIniFilePath = fGetStringFromINIFile("ITAC", "ITAC_IniFilePath ", rIniFilePath)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLog("--- INI ITAC_IniFilePath:        " & rpITACIniFilePath, 0)

            rpIniITAC_Enable = fGetStringFromINIFile("ITAC", "ITAC_Enable", rIniFilePath).ToUpper
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLog("--- INI ITAC_Check_Enable:       " & rpIniITAC_Enable, 0)

            ' --- Get INI TPGM Parameter [ICT]
            'rpIniICT_Enable = fGetStringFromINIFile("ICT", "ICT_Test_Enable", rIniFilePath)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLog("--- INI ICT_Test_Enable:         " & rpIniICT_Enable, 0)

            ' --- Get INI TPGM Parameter [OBP]
            rpIniOBP_Enable = fGetStringFromINIFile("OBP", "OBP_Test_Enable", rIniFilePath)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- INI OBP_Test_Enable:         " & rpIniOBP_Enable, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- INI OBP_Test_Enable", rpRpkBayInUse, rpIniOBP_Enable)

            ' --- Get INI TPGM Parameter [FCT]
            rpIniFCT_Enable = fGetStringFromINIFile("FCT", "FCT_Test_Enable", rIniFilePath)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- INI FCT_Test_Enable:         " & rpIniFCT_Enable, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- INI FCT_Test_Enable", rpRpkBayInUse, rpIniFCT_Enable)

            ' --- Get INI TPGM Parameter [VBFPath]
            rVBFFilePath = fGetStringFromINIFile("FCT", "VBF_File_Path ", rIniFilePath)
            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--- INI FCT_Test_Enable:         " & rpIniFCT_Enable, 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("--- INI VBF FILE", rpRpkBayInUse, rVBFFilePath)


            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)

            ' ***********************************************************************************************************************************************
            ' *** Redim Status variables
            ' ***********************************************************************************************************************************************
            If rpRpkBoardNumber > 0 Then
                ReDim rpUutFctResult(rpRpkBoardNumber - 1)
                ReDim rpUutFctSerialNumber(rpRpkBoardNumber - 1)
            Else
                MsgBox("Invalid board number: " & CStr(rpRpkBoardNumber))
                fTPGM_Initializiation = FAIL
                GoTo fTPGM_Initializiation_End
            End If

            ' ***********************************************************************************************************************************************
            ' *** GET previus ICT Result
            ' ***********************************************************************************************************************************************
            If rpVbNetIsDebug = True Then

                ' --- ONLY FOR DEBUG
                For x As Integer = 0 To rpUutFctResult.Length - 1
                    rpUutFctResult(x) = RESULT_PASS
                Next

                'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("@BG{RED}@FG{WHITE}DEBUG MODE ENABLE ALL RESULT FORCE - PASS", 0, "ALWAYS")
                If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("@BG{RED}@FG{WHITE}DEBUG MODE ENABLE ALL RESULT FORCE - PASS", rpRpkBayInUse)

            Else

                If rpRpkBayInUse = 2 Then

                    For x As Integer = 0 To rpUutFctResult.Length - 1
                        rpUutFctResult(x) = SerialNumberResultRead(x + rpRpkBoardNumber + 1)
                    Next

                End If

            End If

            ' ***********************************************************************************************************************************************
            ' *** GET Serial Number
            ' ***********************************************************************************************************************************************
            If rpVbNetIsDebug = True Then
                ' --- ONLY FOR DEBUG
                'Dim rSNmask As String = "D010006995"
                'For x As Integer = 0 To rpUutFctSerialNumber.Length - 1
                '    rpUutFctSerialNumber(x) = Left(rSNmask, rSNmask.Length - 2) & x.ToString("00")
                'Next
                rpUutFctSerialNumber(0) = "D010006995"
                rpUutFctSerialNumber(1) = "D010006996"
                rpUutFctSerialNumber(2) = "D010006997"
                rpUutFctSerialNumber(3) = "D010006998"
                rpUutFctSerialNumber(4) = "D010006999"
                rpUutFctSerialNumber(5) = "D010007000"
                rpUutFctSerialNumber(6) = "D010007001"
                rpUutFctSerialNumber(7) = "D010007002"
                rpUutFctSerialNumber(8) = "D010007003"
                rpUutFctSerialNumber(9) = "D010007004"
                rpUutFctSerialNumber(10) = "D010007005"
                rpUutFctSerialNumber(11) = "D010007006"
		
		'rpUutFctSerialNumber(0) = "D010007235"
                'rpUutFctSerialNumber(1) = "D010007236"
                'rpUutFctSerialNumber(2) = "D010007237"
                'rpUutFctSerialNumber(3) = "D010007238"
                'rpUutFctSerialNumber(4) = "D010007239"
                'rpUutFctSerialNumber(5) = "D010007240"
                'rpUutFctSerialNumber(6) = "D010007241"
                'rpUutFctSerialNumber(7) = "D010007242"
                'rpUutFctSerialNumber(8) = "D010007243"
                'rpUutFctSerialNumber(9) = "D010007244"
                'rpUutFctSerialNumber(10) = "D010007245"
                'rpUutFctSerialNumber(11) = "D010007246"
            Else

                If rpRpkBayInUse = 2 Then

                    For x As Integer = 0 To rpUutFctSerialNumber.Length - 1
                        rpUutFctSerialNumber(x) = Space(10)
                        SerialNumberRead(x + rpRpkBoardNumber + 1, rpUutFctSerialNumber(x))
                    Next

                End If

            End If




            ' ***********************************************************************************************************************************************
            '  *** Check SN for Panel Present in OBP stage
            ' ***********************************************************************************************************************************************
            For rIndex = 0 To rpUutFctSerialNumber.Length - 1
                If rpUutFctSerialNumber(rIndex) = "0000000000" Or rpUutFctSerialNumber(rIndex) = Nothing Then
                    Exit For
                End If
                ' *** Baco di runpack ritorna primo carattere del serial number come NULL se SN non presente
                If rpUutFctSerialNumber(rIndex) = vbNullChar & "         " Then
                    Exit For
                End If
            Next

            If rIndex < rpUutFctSerialNumber.Length Then
                'MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - OBP Board Presence Verify ........... @BG{Red}@FG{White}FAIL", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("BAY " & rpRpkBayInUse & " - OBP Board Presence Verify", rpRpkBayInUse, FAIL)
                fTPGM_Initializiation = FAIL
                GoTo fTPGM_Initializiation_End
            ElseIf rIndex = rpUutFctSerialNumber.Length Then
                'MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - OBP Board Presence Verify ........... @BG{Green}@FG{White}PASS", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("BAY " & rpRpkBayInUse & " - OBP Board Presence Verify", rpRpkBayInUse, PASS)
                fTPGM_Initializiation = PASS
            End If

            ' *** Result Debug Print
            fResultPrintDebug(rpUutFctSerialNumber, rpUutFctResult)

            ' ***********************************************************************************************************************************************
            ' *** ITAC MANAGE
            ' ***********************************************************************************************************************************************
            If fItac_Check("OBP", rpUutFctSerialNumber, rpUutFctResult) = FAIL Then
                fTPGM_Initializiation = FAIL
                GoTo fTPGM_Initializiation_End
            End If

            'If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
            If rpRpkBayInUse = 2 Then If rpVbNetIsDebug Then rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)


        Catch ex As Exception
            rError = ex.Message
        End Try

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)


fTPGM_Initializiation_End:

    End Function

End Module
