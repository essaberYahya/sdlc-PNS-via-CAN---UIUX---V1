Imports System.Threading

Module mod04_OBP

    Public aUsedChList(31) As Short

    Public Function fObp_U100() As Long

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

        fObp_U100 = FAIL


        Try

            rpUutFctResult(0) = FAIL
            rpUutFctResult(1) = FAIL
            rpUutFctResult(2) = FAIL
            rpUutFctResult(3) = FAIL
            rpUutFctResult(4) = FAIL
            rpUutFctResult(5) = FAIL
            rpUutFctResult(6) = PASS
            rpUutFctResult(7) = PASS
            rpUutFctResult(8) = PASS
            rpUutFctResult(9) = PASS
            rpUutFctResult(10) = PASS
            rpUutFctResult(11) = PASS

            rNumSiteEnable = 0
            If rpUutFctResult(0) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(1) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(2) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(3) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(4) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(5) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(6) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(7) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(8) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(9) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(10) = PASS Then rNumSiteEnable = rNumSiteEnable + 1
            If rpUutFctResult(11) = PASS Then rNumSiteEnable = rNumSiteEnable + 1

            ' ************************************************************************************************************************
            ' *** Select programming 
            ' ************************************************************************************************************************
            If ObpProgrammingSelect(OBP500A, "U100") <> PASS Then GoTo EndWithFail

            ' ************************************************************************************************************************
            ' *** Site Selection
            ' ************************************************************************************************************************
            If rpUutFctResult(0) = PASS Then rSiteEnabled = 1 & ","
            If rpUutFctResult(1) = PASS Then rSiteEnabled = rSiteEnabled & 2 & ","
            If rpUutFctResult(2) = PASS Then rSiteEnabled = rSiteEnabled & 3 & ","
            If rpUutFctResult(3) = PASS Then rSiteEnabled = rSiteEnabled & 4 & ","
            If rpUutFctResult(4) = PASS Then rSiteEnabled = rSiteEnabled & 5 & ","
            If rpUutFctResult(5) = PASS Then rSiteEnabled = rSiteEnabled & 6 & ","
            If rpUutFctResult(6) = PASS Then rSiteEnabled = rSiteEnabled & 7 & ","
            If rpUutFctResult(7) = PASS Then rSiteEnabled = rSiteEnabled & 8 & ","
            If rpUutFctResult(8) = PASS Then rSiteEnabled = rSiteEnabled & 9 & ","
            If rpUutFctResult(9) = PASS Then rSiteEnabled = rSiteEnabled & 10 & ","
            If rpUutFctResult(10) = PASS Then rSiteEnabled = rSiteEnabled & 11 & ","
            If rpUutFctResult(11) = PASS Then rSiteEnabled = rSiteEnabled & 12 & ","

            ' *** Check if there aren't site enabled
            If rSiteEnabled Is Nothing Then GoTo EndWithFail

            ' *** Remove last comma
            rSiteEnabled = Left(rSiteEnabled, rSiteEnabled.Length - 1)

            ' *** Debug Print 
            MsgPrintLogIfEn("Site Enabled: " & rSiteEnabled, 0, "IF ENABLED", rpIniDebugPrint)

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
            If rpUutFctResult(0) = PASS Then rObpChToBeConnect = "3073,3075,3077,3079,3081,"
            If rpUutFctResult(1) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3083,3085,3087,3089,3091,"
            If rpUutFctResult(2) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3093,3095,3097,3099,3101,"
            If rpUutFctResult(3) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3103,3105,3107,3109,3111,"
            If rpUutFctResult(4) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3113,3115,3117,3119,3121,"
            If rpUutFctResult(5) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3123,3125,3127,3129,3131,"
            If rpUutFctResult(6) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3074,3076,3078,3080,3082,"
            If rpUutFctResult(7) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3084,3086,3088,3090,3092,"
            If rpUutFctResult(8) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3094,3096,3098,3100,3102,"
            If rpUutFctResult(9) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3104,3106,3108,3110,3112,"
            If rpUutFctResult(10) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3114,3116,3118,3120,3122,"
            If rpUutFctResult(11) = PASS Then rObpChToBeConnect = rObpChToBeConnect & "3124,3126,3128,3130,3132,"



            ' *** Check if there aren't channel to be connect
            If rObpChToBeConnect Is Nothing Then GoTo EndWithFail

            ' *** Remove last comma
            rObpChToBeConnect = Left(rObpChToBeConnect, rObpChToBeConnect.Length - 1)

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
            For rIndex As Integer = 0 To 11
                If rpUutFctResult(rIndex) = PASS Then
                    ObpGetSiteResult(OBP500A, rIndex + 1, rpUutFctResult(rIndex))
                    MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rpUutFctResult(rIndex), rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
                End If
            Next

            ' ************************************************************************************************************************
            ' *** Manage Resource & PowerOn
            ' ************************************************************************************************************************
            Dim rUserFlagArrayString As String = Nothing

            If rpUutFctResult(0) = PASS Then rUserFlagArrayString = "1,13,"
            If rpUutFctResult(1) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "2,14,"
            If rpUutFctResult(2) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "3,15,"
            If rpUutFctResult(3) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "4,16,"
            If rpUutFctResult(4) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "5,17,"
            If rpUutFctResult(5) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "6,18,"
            If rpUutFctResult(6) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "7,19,"
            If rpUutFctResult(7) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "8,20,"
            If rpUutFctResult(8) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "9,21,"
            If rpUutFctResult(9) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "10,22,"
            If rpUutFctResult(10) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "11,23,"
            If rpUutFctResult(11) = PASS Then rUserFlagArrayString = rUserFlagArrayString & "12,24,"

            ' *** Check if there aren't channel to be connect
            If rUserFlagArrayString Is Nothing Then GoTo EndWithFail


            ' *** Remove last comma
            rUserFlagArrayString = Left(rUserFlagArrayString, rUserFlagArrayString.Length - 1)

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
            For rIndex As Integer = 0 To 11
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
            For rIndex As Integer = 0 To 11
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
            'For rIndex As Integer = 0 To 11
            '    If rpUutFctResult(rIndex) = PASS Then
            '        ObpGetSiteResult(OBP500A, rIndex + 1, rpUutFctResult(rIndex))
            '        MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rpUutFctResult(rIndex), rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            '    End If
            'Next






            If rpUutFctResult(0) = RESULT_PASS Or
               rpUutFctResult(1) = RESULT_PASS Or
               rpUutFctResult(2) = RESULT_PASS Or
               rpUutFctResult(3) = RESULT_PASS Or
               rpUutFctResult(4) = RESULT_PASS Or
               rpUutFctResult(5) = RESULT_PASS Or
               rpUutFctResult(6) = RESULT_PASS Or
               rpUutFctResult(7) = RESULT_PASS Or
               rpUutFctResult(8) = RESULT_PASS Or
               rpUutFctResult(9) = RESULT_PASS Or
               rpUutFctResult(10) = RESULT_PASS Or
               rpUutFctResult(11) = RESULT_PASS Then
                GoTo EndWithPass
            End If

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fObp_U100 = FAIL
        ' *** Set All Site Fail
        rpUutFctResult(0) = FAIL
        rpUutFctResult(1) = FAIL
        rpUutFctResult(2) = FAIL
        rpUutFctResult(3) = FAIL
        rpUutFctResult(4) = FAIL
        rpUutFctResult(5) = FAIL
        rpUutFctResult(6) = FAIL
        rpUutFctResult(7) = FAIL
        rpUutFctResult(8) = FAIL
        rpUutFctResult(9) = FAIL
        rpUutFctResult(10) = FAIL
        rpUutFctResult(11) = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fObp_U100 = PASS
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
        MsgPrintLogIfEn("fObp_U100 - Function execution result: " & fCRes(fObp_U100), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)

    End Function

End Module
