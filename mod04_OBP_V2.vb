Imports System.Threading

Module mod04_OBP_V2

    Public aUsedChList(31) As Short

    Public Function fObp_U100_V2(ByVal StartSite As Integer, ByVal StopSite As Integer) As Long

        Dim rError As String = Nothing
        Dim rErrorCode As Long = FAIL

        Dim aUsedChList(64) As Short
        Dim rObpChToBeConnect As String = Nothing

        Dim rPPSu As Integer = Nothing
        Dim rPPsu_RearbackVoltage As Double = Nothing
        Dim rPPsu_RearbackCorrent As Double = Nothing

        Dim rSiteEnabled As String = Nothing
        Dim rNumSiteEnable As Integer = 0
        Dim aUserFlagArray(32) As Integer
        Dim aUsedChForStuck(2) As Short

        fObp_U100_V2 = FAIL


        Try



            rNumSiteEnable = 0
            For i As Integer = StartSite To StopSite
                If rpUutFctResult(i) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            Next

            ' ************************************************************************************************************************
            ' *** Select programming 
            ' ************************************************************************************************************************
            If ObpProgrammingSelect(OBP500A, "U100") <> PASS Then GoTo EndWithFail

            ' ************************************************************************************************************************
            ' *** Site Selection
            ' ************************************************************************************************************************
            rSiteEnabled = ""
            For i As Integer = StartSite To StopSite
                If rpUutFctResult(i) = PASS Then rSiteEnabled = rSiteEnabled & (i + 1) & ","
            Next

            ' *** Remove last comma and Check if there aren't site enabled
            If rSiteEnabled.Length > 1 Then
                rSiteEnabled = Left(rSiteEnabled, rSiteEnabled.Length - 1)
            Else
                GoTo EndWithFail
            End If

            ' *** Debug Print 
            'MsgPrintLogIfEn("Site Enabled: " & rSiteEnabled, 0, "IF ENABLED", rpIniDebugPrint)
            rpMsgLogClass.MsgPrintLogMultyBay("Site Enabled", rpRpkBayInUse, rSiteEnabled, True)

            ' *** Set Active site
            If ObpSiteSet(OBP500A, rSiteEnabled) <> PASS Then GoTo EndWithFail

            ' *** Debug Print 
            Console.WriteLine("BAY" & rpRpkBayInUse & "-OBP500A-Site Enabled: " & rSiteEnabled)

            ' ************************************************************************************************************************
            ' *** Configuration 
            ' ************************************************************************************************************************
            If ObpModelOptionsSet(OBP500A, "DEVICE=SPC584C70;COMMFREQ=9M") <> PASS Then GoTo EndWithFail
            ' ************************************************************************************************************************
            ' *** Setting and connection of used channels
            ' ************************************************************************************************************************
            rObpChToBeConnect = ""
            If rpUutFctResult(0) = PASS And (StartSite <= 0 And 0 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3073,3075,3077,3079,3081,"
            If rpUutFctResult(1) = PASS And (StartSite <= 1 And 1 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3083,3085,3087,3089,3091,"
            If rpUutFctResult(2) = PASS And (StartSite <= 2 And 2 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3093,3095,3097,3099,3101,"
            If rpUutFctResult(3) = PASS And (StartSite <= 3 And 3 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3103,3105,3107,3109,3111,"
            If rpUutFctResult(4) = PASS And (StartSite <= 4 And 4 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3113,3115,3117,3119,3121,"
            If rpUutFctResult(5) = PASS And (StartSite <= 5 And 5 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3123,3125,3127,3129,3131,"
            If rpUutFctResult(6) = PASS And (StartSite <= 6 And 6 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3074,3076,3078,3080,3082,"
            If rpUutFctResult(7) = PASS And (StartSite <= 7 And 7 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3084,3086,3088,3090,3092,"
            If rpUutFctResult(8) = PASS And (StartSite <= 8 And 8 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3094,3096,3098,3100,3102,"
            If rpUutFctResult(9) = PASS And (StartSite <= 9 And 9 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3104,3106,3108,3110,3112,"
            If rpUutFctResult(10) = PASS And (StartSite <= 10 And 10 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3114,3116,3118,3120,3122,"
            If rpUutFctResult(11) = PASS And (StartSite <= 11 And 11 <= StopSite) Then rObpChToBeConnect = rObpChToBeConnect & "3124,3126,3128,3130,3132,"



            ' *** Remove last comma and Check if there aren't site enabled
            If rObpChToBeConnect.Length > 1 Then
                rObpChToBeConnect = Left(rObpChToBeConnect, rObpChToBeConnect.Length - 1)
            Else
                GoTo EndWithFail
            End If

            ' *** Convert String to Array
            S2A(rObpChToBeConnect, aUsedChList(0), 60)

            ' *** Channel Logical Level Set
            If ObpChLevelSet(OBP500A, aUsedChList(0), 4.8, 0.0) <> PASS Then GoTo EndWithFail

            ' *** Channel Logical Threshold Level Set
            If ObpChSensorSet(OBP500A, aUsedChList(0), 2) <> PASS Then GoTo EndWithFail

            ' *** Channel Connect
            If ObpChConnect(OBP500A, aUsedChList(0)) <> PASS Then GoTo EndWithFail

            '' '' ''aUsedChList(0) = 3082
            '' '' ''aUsedChList(1) = 0

            '' '' ''ObpChStuckSet(OBP500A, aUsedChList(0), LOW)
            '' '' ''ObpChStuckSet(OBP500A, aUsedChList(0), HIGH)

            ' *** Debug Print 
            Console.WriteLine("BAY" & rpRpkBayInUse & "-OBP500A-Channel Connected: " & rObpChToBeConnect)

            ' ************************************************************************************************************************
            ' *** Interface Enable
            ' ************************************************************************************************************************
            rpNtest = 1200                                                          ' *** Set Test Number
            rpDrawRef = "OBP"                                                      ' *** Set Drawing Reference
            rpTestName = "Interface"                                                    ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            ' ************************************************************************************************************************
            If ObpInterfaceEnable(OBP500A) Then GoTo EndWithFail
            '************************************************************************************************************************
            'Manage result of previous operation
            '************************************************************************************************************************
            For rIndex As Integer = StartSite To StopSite 
                If rpUutFctResult(rIndex) = PASS Then
                    ObpGetSiteResult(OBP500A, rIndex + 1, rpUutFctResult(rIndex))
                    MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rpUutFctResult(rIndex), rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
                End If
            Next

            ' ************************************************************************************************************************
            ' *** Manage Resource & PowerOn
            ' ************************************************************************************************************************
            Dim rUserFlagArrayString As String = ""
            If rpUutFctResult(0) = PASS And (StartSite <= 0 And 0 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "1,13,"
            If rpUutFctResult(1) = PASS And (StartSite <= 1 And 1 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "2,14,"
            If rpUutFctResult(2) = PASS And (StartSite <= 2 And 2 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "3,15,"
            If rpUutFctResult(3) = PASS And (StartSite <= 3 And 3 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "4,16,"
            If rpUutFctResult(4) = PASS And (StartSite <= 4 And 4 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "5,17,"
            If rpUutFctResult(5) = PASS And (StartSite <= 5 And 5 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "6,18,"
            If rpUutFctResult(6) = PASS And (StartSite <= 6 And 6 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "7,19,"
            If rpUutFctResult(7) = PASS And (StartSite <= 7 And 7 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "8,20,"
            If rpUutFctResult(8) = PASS And (StartSite <= 8 And 8 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "9,21,"
            If rpUutFctResult(9) = PASS And (StartSite <= 9 And 9 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "10,22,"
            If rpUutFctResult(10) = PASS And (StartSite <= 10 And 10 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "11,23,"
            If rpUutFctResult(11) = PASS And (StartSite <= 11 And 11 <= StopSite) Then rUserFlagArrayString = rUserFlagArrayString & "12,24,"

            ' *** Remove last comma and Check if there aren't site enabled
            If rUserFlagArrayString.Length > 1 Then
                rUserFlagArrayString = Left(rUserFlagArrayString, rUserFlagArrayString.Length - 1)
            Else
                GoTo EndWithFail
            End If

            ' *** Convert String to Array
            S2DWA(rUserFlagArrayString, aUserFlagArray(0), 24)

            ' *** Close UFL CPU for PowerOn
            UserFlagsSet(RLYNOA, aUserFlagArray(0))

            ' ************************************************************************************************************************
            ' *** PowerOn HW1-FXSEZ200 and Setup Relay (PPS1 [5V])
            ' ************************************************************************************************************************
            rPPSu = PPS1
            rPPsu_RearbackVoltage = Nothing
            rPPsu_RearbackCorrent = Nothing

            ' *** PowerON Instrument
            PpsuOn(rPPSu, 5, 1)

            ' *** Rearback Voltage Instrument
            PpsuReadbackV(rPPSu, rPPsu_RearbackVoltage)

            ' *** Rearback Current Instrument
            PpsuReadbackI(rPPSu, rPPsu_RearbackCorrent)

            ' ************************************************************************************************************************
            ' *** PowerOn  for UUT (PPS4 [13.5V])
            ' ************************************************************************************************************************
            rPPSu = PPS4
            rPPsu_RearbackVoltage = Nothing
            rPPsu_RearbackCorrent = Nothing

            ' *** PowerON Instrument
            PpsuOn(rPPSu, 13.5, 0.12 * rNumSiteEnable)

            ' *** Wait For Stabilization
            Thread.Sleep(1000)

            ' *** Rearback Voltage Instrument
            PpsuReadbackV(rPPSu, rPPsu_RearbackVoltage)

            ' *** Rearback Current Instrument
            PpsuReadbackI(rPPSu, rPPsu_RearbackCorrent)

            ' ************************************************************************************************************************
            ' *** Stuck for OBP
            ' ************************************************************************************************************************
            aUsedChForStuck(0) = 3133
            aUsedChForStuck(1) = 3134
            aUsedChForStuck(2) = 0
            ' *** Channel Logical Level Set
            If ObpChLevelSet(OBP500A, aUsedChForStuck(0), 5.0, 0.0) <> PASS Then GoTo EndWithFail
            ' *** Channel Connect
            If ObpChConnect(OBP500A, aUsedChForStuck(0)) <> PASS Then GoTo EndWithFail
            ' *** Channel Stuck High
            If ObpChStuckSet(OBP500A, aUsedChForStuck(0), HIGH) <> PASS Then GoTo EndWithFail
			
			Thread.Sleep(50)

            ' *** FOR DEBUG
            'ObpLogPrintEnable(OBP500A, 20)
            ' *******************************



            ' ************************************************************************************************************************
            ' *** ERASE
            ' ************************************************************************************************************************
            rpNtest = 1200                                                          ' *** Set Test Number
            rpDrawRef = "OBP"                                                      ' *** Set Drawing Reference
            rpTestName = "Erase"                                                    ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            ' ************************************************************************************************************************
            If ObpChipErase(OBP500A, TYP_DWORD) Then GoTo EndWithFail

            ' ************************************************************************************************************************
            ' Manage result of previous operation
            ' ************************************************************************************************************************
            For rIndex As Integer = StartSite To StopSite
                If rpUutFctResult(rIndex) = PASS Then
                    ObpGetSiteResult(OBP500A, rIndex + 1, rpUutFctResult(rIndex))
                    MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rpUutFctResult(rIndex), rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
                End If
            Next



            ' ************************************************************************************************************************
            rpNtest = 1201                                                          ' *** Set Test Number
            rpDrawRef = "OBP"                                                       ' *** Set Drawing Reference
            rpTestName = "FILE 1 - Write Verify"                                           ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            ' ************************************************************************************************************************
            If ObpWriteVerifyFile(OBP500A, TYP_DWORD, 1, 0, 0, ALLFILE, NOSKIP) Then GoTo EndWithFail
            ' ************************************************************************************************************************
            ' Manage result of previous operation
            ' ************************************************************************************************************************
            For rIndex As Integer = StartSite To StopSite
                If rpUutFctResult(rIndex) = PASS Then
                    ObpGetSiteResult(OBP500A, rIndex + 1, rpUutFctResult(rIndex))
                    MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rpUutFctResult(rIndex), rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
                End If
            Next



            '' ************************************************************************************************************************
            'rpNtest = 1202                                                          ' *** Set Test Number
            'rpDrawRef = "OBP"                                                       ' *** Set Drawing Reference
            'rpTestName = "FILE 1 - Verify"                                          ' *** Set Test Name
            'rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            'rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            '' ************************************************************************************************************************
            'If ObpVerifyFile(OBP500A, TYP_DWORD, 1, 0, 0, ALLFILE) Then GoTo EndWithFail

            '' ************************************************************************************************************************
            '' Manage result of previous operation
            '' ************************************************************************************************************************
            'For rIndex As Integer = StartSite To StopSite
            '    If rpUutFctResult(rIndex) = PASS Then
            '        ObpGetSiteResult(OBP500A, rIndex + 1, rpUutFctResult(rIndex))
            '        MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rpUutFctResult(rIndex), rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            '    End If
            'Next




            Dim TempResult As Boolean = False
            For i As Integer = StartSite To StopSite
                If rpUutFctResult(i) = RESULT_PASS Then
                    TempResult = True
                End If
            Next
            If TempResult Then GoTo EndWithPass

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fObp_U100_V2 = FAIL
        ' *** Set All Site Fail
        For i As Integer = StartSite To StopSite
            rpUutFctResult(i) = FAIL
        Next
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fObp_U100_V2 = PASS
Exit_Function:

        ObpChDisconnect(OBP500A, aUsedChForStuck(0))

        ' *** PowerOff Board
        PpsOff(PPS4, 5)
        ' *** PowerOff HW1 FXSEZ200
        PpsOff(PPS1, 5)

        ' *** UserFlag Reset
        UserFlagsReset(RLYNOA, aUserFlagArray(0))


        ' *** Obp Ch disconnect
        ObpChDisconnect(OBP500A, aUsedChList(0))



        ' *** Board Power Off
        'fPowerOffBoard()

        ' *** Print Function Result
        'MsgPrintLogIfEn("fObp_U100 - Function execution result: " & fCRes(fObp_U100_V2), 0, "IF ENABLED", rpIniDebugPrint)
        rpMsgLogClass.MsgPrintLogMultyBay("fObp_U100 - Function execution result", rpRpkBayInUse, fObp_U100_V2, True)

        ' *** Print error if occurred
        If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function

End Module
