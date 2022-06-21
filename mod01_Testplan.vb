Option Strict Off
Option Explicit On

Module modTestplan

  ' ******************************************************************
  ' *** Kernel32 Variables
  ' ******************************************************************
  Public Declare Function GetTickCount Lib "kernel32" () As Int32

  Public Function Testplan() As Integer

    Dim rError As String = Nothing
    Dim FailFlag As Integer

    Dim rTimeStart As Object
    Dim rTime_OBP As Object
    Dim rTime_FCT As Object

        Try

            FailFlag = PASS

            ' **************************************************************************************************
            ' *** INITIALIZATION
            ' **************************************************************************************************
            rTimeStart = GetTickCount

            If fTPGM_Initializiation() <> PASS Then
                ' *** OBP Disable
                rpIniOBP_Enable = "NO"
                ' *** FCT Disable
                rpIniFCT_Enable = "NO"
                ' *** ITAC Disable
                rpIniITAC_Enable = "NO"
                ' *** SET fail flag
                FailFlag = FAIL
                GoTo END_TPGM
                'rpUutFctResult(0) = PASS
                'rpUutFctResult(1) = PASS
                'rpUutFctResult(2) = PASS
                'rpUutFctResult(3) = PASS
                'rpUutFctResult(4) = PASS
                'rpUutFctResult(5) = PASS
                'rpUutFctResult(6) = PASS
                'rpUutFctResult(7) = PASS
                'rpUutFctResult(8) = PASS
                'rpUutFctResult(9) = PASS
                'rpUutFctResult(10) = PASS
                'rpUutFctResult(11) = PASS

                Threading.Thread.Sleep(10)

                If CheckStop() Then
                    Exit Function
                End If


            End If

            ' *** Result Print
            fResultPrint(rpUutFctSerialNumber, rpUutFctResult)
            ' **************************************************************************************************
            ' *** OBP
            ' **************************************************************************************************
            Dim OBPOk As Boolean = True
            If (rpIniOBP_Enable = "YES") And rpRpkBayInUse = 2 Then

                ' *** Select Common Site
                UseSiteWrite(0)
                OBPOk = False

                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)
                'MsgPrintLogIfEn("ON BOARD PROGRAMMING : TestPlan .......... @FG{Blue}Running", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("ON BOARD PROGRAMMING", rpRpkBayInUse, 3) 'Running
                rpMsgLogClass.MsgPrintLogMultyBay("Part Number under test :" & rpRpkTpgmVariant, rpRpkBayInUse, 3) 'Running
                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)

                rTime_OBP = GetTickCount
                If fObp_U100_V2(0, 5) <> PASS Then
                    ' *** SET fail flag
                    FailFlag = FAIL
                Else
                    ' *** Enable FCT if PASS
                    OBPOk = True
                End If

                If fObp_U100_V2(6, 11) <> PASS Then
                    ' *** SET fail flag
                    FailFlag = FAIL
                Else
                    ' *** Enable FCT if PASS
                    OBPOk = True
                End If

                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)
                'MsgPrintLogIfEn("ON BOARD PROGRAMMING : TestPlan ......... @FG{Purple}[" & Format((GetTickCount - rTime_OBP) / 1000, "00.00") & "s]", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("ON BOARD PROGRAMMING", rpRpkBayInUse, "@FG{Purple}[" & Format((GetTickCount - rTime_OBP) / 1000, "00.00") & "s]")
                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)


                ' *** Result Debug Print
                fResultPrintDebug(rpUutFctSerialNumber, rpUutFctResult)


            End If

            If CheckStop() Then
                Exit Function
            End If





            ' **************************************************************************************************
            ' *** FCT
            ' **************************************************************************************************
            'rpIniFCT_Enable = NO
            If (rpIniFCT_Enable = "YES" And OBPOk And rpRpkBayInUse = 2) Then

                ' *** Select Common Site
                UseSiteWrite(0)

                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)
                'MsgPrintLogIfEn("FUNCTIONAL TESTPLAN : TestPlan ........... @FG{Blue}Running", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("FUNCTIONAL TESTPLAN", rpRpkBayInUse, 3) 'Running
                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)


                rTime_FCT = GetTickCount
                If fFCT() <> PASS Then
                    ' *** SET fail flag
                    FailFlag = FAIL
                End If


                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)
                'MsgPrintLogIfEn("FUNCTIONAL TESTPLAN : TestPlan ......... @FG{Purple}[" & Format((GetTickCount - rTime_FCT) / 1000, "00.00") & "s]", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("FUNCTIONAL TESTPLAN", rpRpkBayInUse, "@FG{Purple}[" & Format((GetTickCount - rTime_FCT) / 1000, "00.00") & "s]")
                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)


            End If

END_TPGM:
            If rpRpkBayInUse = 2 Then
                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)
                'MsgPrintLogIfEn("TOTAL TEST TIME BAY" & rpRpkBayInUse & " : ................. @FG{Purple}[" & Format((GetTickCount - rTimeStart) / 1000, "00.00") & "s]", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("TOTAL TEST TIME BAY" & rpRpkBayInUse, rpRpkBayInUse, "@FG{Purple}[" & Format((GetTickCount - rTimeStart) / 1000, "00.00") & "s]")
                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)
            End If

            If CheckStop() Then
                Exit Function
            End If



            ' **************************************************************************************************
            ' *** ITAC Store Result
            ' **************************************************************************************************
            If rpITACIsOnline = True And rpIniITAC_Enable = "YES" Then
                If fItac_StoreResult("OBP", rpUutFctSerialNumber, rpUutFctResult, (GetTickCount - rTimeStart) / 1000) = PASS Then
                    'MsgPrintLogIfEn("BAY" & rpRpkBayInUse & " - ITAC Store Result .................... @BG{Green}@FG{White}PASS", 0, "ALWAYS")
                    rpMsgLogClass.MsgPrintLogMultyBay(" - ITAC Store Result", rpRpkBayInUse, PASS)
                Else
                    'MsgPrintLogIfEn("BAY" & rpRpkBayInUse & " - ITAC Store Result .................... @BG{Red}@FG{White}FAIL", 0, "ALWAYS")
                    rpMsgLogClass.MsgPrintLogMultyBay(" - ITAC Store Result", rpRpkBayInUse, FAIL)
                End If
                'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)
            End If


                ' **************************************************************************************************
                ' *** Set Site Result
                ' **************************************************************************************************
                SiteResultWrite(1, rpUutFctResult(0))
            SiteResultWrite(2, rpUutFctResult(1))
            SiteResultWrite(3, rpUutFctResult(2))
            SiteResultWrite(4, rpUutFctResult(3))
            SiteResultWrite(5, rpUutFctResult(4))
            SiteResultWrite(6, rpUutFctResult(5))
            SiteResultWrite(7, rpUutFctResult(6))
            SiteResultWrite(8, rpUutFctResult(7))
            SiteResultWrite(9, rpUutFctResult(8))
            SiteResultWrite(10, rpUutFctResult(9))
            SiteResultWrite(11, rpUutFctResult(10))
            SiteResultWrite(12, rpUutFctResult(11))

            If (rpIniOBP_Enable = "YES") Or (rpIniFCT_Enable = "YES") Then
                ' *** Result Debug Print
                fResultPrint(rpUutFctSerialNumber, rpUutFctResult)
            End If
            ' **************************************************************************************************
            ' *** Set Tpgm Result
            ' **************************************************************************************************
            FailFlag = rpUutFctResult(0) Or rpUutFctResult(1) Or rpUutFctResult(2) Or rpUutFctResult(3) Or rpUutFctResult(4) Or rpUutFctResult(5) Or rpUutFctResult(6) Or rpUutFctResult(7) Or rpUutFctResult(8) Or rpUutFctResult(9) Or rpUutFctResult(10) Or rpUutFctResult(11)

			SyncParallelExec()
			
            'If FailFlag = 1 Then MsgBox("FKT STAGE - BOARD FAIL - Press OK to continue!")

        Catch ex As Exception
            rError = ex.Message
      FailFlag = FAIL
    End Try

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)



        'If FailFlag = FAIL Then
        '    SemaphoreSet(SEMAPHORE_LIGHT_RED_BLINK, SEMAPHORE_BUZZER_ON)
        '    MsgBox("FAIL DETECTED!")
        'End If
        SyncParallelExec()
        TplanResultSet(FailFlag)

    ' *** Set Function Result
    Testplan = 1


  End Function


  ' ----------------------------------------------------------------------------
  '
  ' --- TESTPLAN Initialization
  '
  ' This function is executed only one time when the test program is loaded.
  '
  Public Function TestplanInit() As Integer

    '
    ' --- INSERT YOUR CODE HERE ...
    '

    TestplanInit = 1
  End Function


  ' ----------------------------------------------------------------------------
  '
  ' --- TESTPLAN End
  '
  ' This function is executed only one time when the test program ends.
  '
  Public Function TestplanEnd() As Integer

    '
    ' --- INSERT YOUR CODE HERE ...
    '

    TestplanEnd = 1
  End Function

End Module
