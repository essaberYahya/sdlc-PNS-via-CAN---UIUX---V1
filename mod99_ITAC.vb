Module mod99_ITAC

    ' *** Public ITAC Variables
    Public ITAC As New ApiCall.iTAC_Call

    'Public rpITACLogin As Boolean
    'Public rpITACLogout As Boolean
    Public rpItacStage As String

    Public rpITACIsOnline As Boolean
    Public rpITACIniFilePath As String

    Public apITACFailure() As String = Nothing
    Public rpITACBookingOK As Boolean

    Public Sub sCollectITACFailure(ByVal rSite As Integer, ByVal str1 As String, ByVal str2 As String, ByVal str3 As String, ByVal str4 As String)

        Dim rError As String = Nothing

        Try

            Dim rSeparator As Char = vbCrLf

            If apITACFailure(rSite) = Nothing Then
                apITACFailure(rSite) = str1 & ";" & str2 & ";" & str3 & ";" & str4 & vbCrLf & vbCrLf
            Else
                apITACFailure(rSite) = apITACFailure(rSite) & ";" & str1 & ";" & str2 & ";" & str3 & ";" & str4 & vbCrLf & vbCrLf
            End If

        Catch ex As Exception
            rError = ex.Message
        End Try

        ' *** Print error if occurred
        If rError IsNot Nothing Then MsgPrintLog("sCollectITACFailure - Error: " & rError, 0)

    End Sub

    Public Function fItac_Check(ByVal rStage As String, ByVal rUutSerialNumber() As String, ByRef rUutResult() As Long) As Long

        fItac_Check = FAIL

        Dim rError As String = Nothing
        Dim rItac_Enabled As Long = FAIL
        rpITACIsOnline = False

        Try

            ' ************************************************************************************************************************
            ' *** ITAC Initialization of Failures collect variable
            ' ************************************************************************************************************************
            ReDim apITACFailure(rUutSerialNumber.Length - 1)

            ' ************************************************************************************************************************
            ' *** ITAC Enabled Check
            ' ************************************************************************************************************************
            rpNtest = 5000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                      ' *** Set Drawing Reference
            rpTestName = "ITAC STATUS ENABLED"                                      ' *** Set Test Name
            rpDiagRemark = rpTestName & "-FAIL"                                     ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure

            If rpIniITAC_Enable = "YES" Then
                'MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - ITAC Enable Check ................... @BG{Green}@FG{White}PASS", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("BAY " & rpRpkBayInUse & " - ITAC Enable Check", rpRpkBayInUse, PASS)
                rItac_Enabled = PASS
            Else
                'MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - ITAC Enable Check ................... @BG{Red}@FG{White}FAIL", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("BAY " & rpRpkBayInUse & " - ITAC Enable Check", rpRpkBayInUse, FAIL)
            End If

            For rIndex As Integer = 0 To rUutSerialNumber.Length - 1
                MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rItac_Enabled, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "NO PRINT")

            Next

            ' *** If Itac Disabled Force All Pass
            If rpIniITAC_Enable = "NO" Then
                For rIndex As Integer = 0 To rUutSerialNumber.Length - 1
                    rUutResult(rIndex) = PASS
                Next

                'MsgPrintLogIfEn("@BG{RED}@FG{WHITE}ITAC DISABLED - ALL RESULT FORCE - PASS", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("@BG{RED}@FG{WHITE}ITAC DISABLED - ALL RESULT FORCE - PASS", rpRpkBayInUse)
                GoTo EndWithPass
            End If

            ' ************************************************************************************************************************
            ' *** ITAC Online Check
            ' ************************************************************************************************************************
            rpNtest = 5001                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                      ' *** Set Drawing Reference
            rpTestName = "ITAC Online Check"                                        ' *** Set Test Name
            rpDiagRemark = rpTestName & "-FAIL"                                     ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure

            rpITACIsOnline = ITAC.iTAC_Init(rpITACIniFilePath, rStage)

            If (rpITACIsOnline = True) Then
                'MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - ITAC Online Check ................... @BG{Green}@FG{White}PASS", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("BAY " & rpRpkBayInUse & " - ITAC Online Check", rpRpkBayInUse, PASS)
            Else
                'MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - ITAC Online Check ................... @BG{Red}@FG{White}FAIL", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("BAY " & rpRpkBayInUse & " - ITAC Online Check", rpRpkBayInUse, FAIL)
            End If

            For rIndex As Integer = 0 To rUutSerialNumber.Length - 1
                MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rUutResult(rIndex), rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "NO PRINT")
            Next

            If (rpITACIsOnline = False) Then GoTo EndWithFail

            ' ************************************************************************************************************************
            ' *** ITAC Serial Number
            ' ************************************************************************************************************************
            rpNtest = 5002                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                         ' *** Set Drawing Reference
            rpTestName = "ITAC Serial Numbers Check"                                ' *** Set Test Name
            rpDiagRemark = rpTestName & "-FAIL"                                     ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure

            Dim rItacResult As Long = fItac_SerialNuberCheck(rStage, rUutSerialNumber, rUutResult)

            For rIndex As Integer = 0 To rUutSerialNumber.Length - 1
                MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rUutResult(rIndex), rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            Next

            If rItacResult = FAIL Then GoTo EndWithFail

            ' *** At least one Pass Board
            GoTo EndWithPass
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fItac_Check = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fItac_Check = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fItac_Check - Function execution result: " & fCRes(fItac_Check), 0, "IF ENABLED", rpIniDebugPrint)
        rpMsgLogClass.MsgPrintLogMultyBay("fItac_Check - Function execution result", rpRpkBayInUse, fItac_Check, True)


        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("fItac_Check - Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("fItac_Check - Error: " & rError, rpRpkBayInUse)

    End Function

    Public Function fItac_SerialNuberCheck(ByVal rStage As String, ByVal rUutSerialNumber() As String, ByRef rUutResult() As Long) As Long

        ' *** Set Function Fail
        fItac_SerialNuberCheck = FAIL

        Dim rError As String = Nothing
        Dim aITACCheckResult() As String

        Try

            If rStage = "ICT" Then rpItacStage = ITAC.getStationStg1
            If rStage = "OBP" Then rpItacStage = ITAC.getStationStg2

            ' *** ITAC Request Panel Status 
            aITACCheckResult = ITAC.iTAC_Check_PNL(rpItacStage, ITAC.getLayer, rpRpkBoardNumber, rUutSerialNumber(0))

            ' *** ITAC Panel Rejected Check
            If aITACCheckResult.Length = 2 Then
                'MsgPrintLogIfEn("   @FG{RED}PANEL REJECTED", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("   @FG{RED}PANEL REJECTED", rpRpkBayInUse)
                'MsgPrintLogIfEn("   ERROR CODE: " & aITACCheckResult(0), 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("   ERROR CODE: " & aITACCheckResult(0), rpRpkBayInUse)
                'MsgPrintLogIfEn("   ERROR: " & aITACCheckResult(1), 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("   ERROR: " & aITACCheckResult(1), rpRpkBayInUse)
                rError = aITACCheckResult(1)
                GoTo EndWithFail
            End If

            ' *** ITAC Board Check
            For BoardNumber As Integer = 0 To rUutSerialNumber.Length - 1

                Dim rItacMessage As String = aITACCheckResult(BoardNumber * 4)
                Dim rItacSerialNumber As String = aITACCheckResult(BoardNumber * 4 + 1)
                Dim rItacBoardPosition As String = aITACCheckResult(BoardNumber * 4 + 2)
                Dim rItacBoardResult As String = aITACCheckResult(BoardNumber * 4 + 3)

                If rUutSerialNumber(BoardNumber) = rItacSerialNumber And rItacMessage = "0" Then
                    rUutResult(BoardNumber) = PASS
                Else

                    If rItacMessage = "0" Then
                        rError = "PANEL INVERTED / WRONG SN" ' TO BE DONE FOR BOTH CASES
                        GoTo EndWithFail
                    End If

                    rUutResult(BoardNumber) = FAIL

                    MsgPrintReport("ERROR: " & rItacMessage, 0)
                    sCollectITACFailure(BoardNumber, "0", "N.A", "ITAC Check SN", rItacMessage)

                End If
            Next

            ' *** If all site Fail function return Fail
            For rIndex As Integer = 0 To rUutSerialNumber.Length - 1
                If rUutResult(rIndex) = PASS Then GoTo EndWithPass
            Next

            rError = "ITAC ALL SITE FAIL"
            GoTo EndWithFail
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:

        MsgPrintReport("   ERROR: " & rError, 0)

        For rIndex As Integer = 0 To rUutSerialNumber.Length - 1
            rUutResult(rIndex) = FAIL
        Next

        ' *** Store Result Fail
        fItac_SerialNuberCheck = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fItac_SerialNuberCheck = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fItac_SerialNuberCheck - Function execution result: " & fCRes(fItac_SerialNuberCheck), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("fItac_SerialNuberCheck - Error: " & rError, 0)

    End Function

    Public Function fItac_StoreResult(ByVal rStage As String, ByVal rUutSerialNumber() As String, ByRef rUutResult() As Long, ByVal rCycleTime As Single) As Long

        fItac_StoreResult = FAIL

        Dim rError As String = Nothing
        Dim rItacBookingResult(rUutSerialNumber.Length - 1) As Long

        Try

            If rStage = "ICT" Then rpItacStage = ITAC.getStationStg1
            If rStage = "OBP" Then rpItacStage = ITAC.getStationStg2

            ' ************************************************************************************************************************
            ' *** ITAC Store Result
            ' ************************************************************************************************************************
            rpNtest = 5100                                                          ' *** Set Test Number
            rpDrawRef = " "                                                         ' *** Set Drawing Reference
            rpTestName = "ITAC STORE RESULT"                                        ' *** Set Test Name
            rpDiagRemark = rpTestName & "-FAIL"                                     ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure

            For rIndex As Integer = 0 To rUutSerialNumber.Length - 1

                If rUutResult(rIndex) = 0 Then
                    rpITACBookingOK = ITAC.iTAC_Upload_Status(rpItacStage, ITAC.getLayer, rUutSerialNumber(rIndex), "-1", True, rCycleTime, 0)
                Else
                    Dim rItacFailureDescription() As String = Split(apITACFailure(rIndex), ";", , CompareMethod.Text)
                    rpITACBookingOK = ITAC.iTAC_Upload_Failures(rpItacStage, ITAC.getLayer, rUutSerialNumber(rIndex), "-1", False, rCycleTime, rItacFailureDescription)
                End If

                If rpITACBookingOK = True Then
                    rItacBookingResult(rIndex) = PASS
                Else
                    rUutResult(rIndex) = FAIL
                    rItacBookingResult(rIndex) = FAIL
                End If

                MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rIndex + 1, rpNtest, rpDrawRef, rpTestName, rItacBookingResult(rIndex), rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            Next

            ' **************************************************************************************************
            ' *** ITAC LOG OUT
            ' **************************************************************************************************
            ITAC.iTAC_End()

            'MsgPrintLogIfEn("--------------------------------------------------", 0, "ALWAYS")
            rpMsgLogClass.MsgPrintLogMultyBay("-", rpRpkBayInUse)

            ' *** Search for Booking Fail
            For rIndex As Integer = 0 To rUutSerialNumber.Length - 1
                If rItacBookingResult(rIndex) = FAIL Then
                    GoTo EndWithFail
                End If
            Next

            GoTo EndWithPass
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:

        MsgPrintReport("   ERROR: " & rError, 0)

        'For rIndex As Integer = 0 To rUutSerialNumber.Length - 1
        '    rUutResult(rIndex) = FAIL
        'Next

        ' *** Store Result Fail
        fItac_StoreResult = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fItac_StoreResult = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fItac_StoreResult - Function execution result: " & fCRes(fItac_StoreResult), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("fItac_StoreResult - Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("fItac_StoreResult - Error: " & rError, rpRpkBayInUse)

    End Function

End Module


'Public Function fItac_Check() As Long

'    fItac_Check = FAIL

'    Dim rError As String = Nothing


'    Try














'        GoTo EndWithPass
'        ' *************************************************************************************************************    

'    Catch ex As Exception
'        rError = ex.Message
'        GoTo EndWithFail
'    End Try

'EndWithFail:
'    ' *** Store Result Fail
'    fItac_Check = FAIL
'    GoTo Exit_Function
'EndWithPass:
'    ' *** Store Result Pass
'    fItac_Check = PASS
'Exit_Function:

'    ' *** Print Function Result
'    MsgPrintLogIfEn("fItac_Check - Function execution result: " & fCRes(fItac_Check), 0, "IF ENABLED", rpIniDebugPrint)

'    ' *** Print error if occurred
'    If rError IsNot Nothing Then MsgPrintLog("fItac_Check - Error: " & rError, 0)























'    ' --- Set default ITAC Status
'    rpITACIsOnline = False
'    rpIniITAC_Enable = rpIniITAC_Enable.ToUpper

'    ' --- ITAC Enabled Check
'    If rpIniITAC_Enable = "YES" Then
'        MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - ITAC STATUS: ............. @FG{Green}ENABLED", 0, "ALWAYS")
'    Else
'        MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - ITAC STATUS: ............. @FG{Red}DISABLED FROM INI FILE", 0, "ALWAYS")
'    End If

'    ' --- ITAC Online Check
'    rpITACIsOnline = ITAC.iTAC_Init(rpITACIniFilePath, "OBP")

'    If (rpITACIsOnline = True) And (rpIniITAC_Enable = "YES") Then
'        MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - ITAC STATUS: ............................. @FG{Blue}Online", 0, "ALWAYS")
'    Else
'        MsgPrintLogIfEn("BAY " & rpRpkBayInUse & " - ITAC STATUS: ............................. @FG{Red}Offline", 0, "ALWAYS")
'        fTPGM_Initializiation = FAIL
'        GoTo fTPGM_Initializiation_End
'    End If

'    ' --- Initialization of ITAC Failures collect variable
'    ReDim apITACFailure(rpUutFctSerialNumber.Length - 1)
'    For rIndex = 0 To rpUutFctSerialNumber.Length - 1
'        apITACFailure(rIndex) = {}
'    Next



'End Function

