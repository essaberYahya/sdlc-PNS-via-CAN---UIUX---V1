Imports AtosF
Imports System.Threading
Imports System.Windows.Forms
Imports System.IO
Imports HCMUnlocker16bytes
Imports FordKeyUp
Imports System.Linq

Module mod05_FCT

    Public Function fFCT() As Long

        Dim rUserFlagArrayString As String = Nothing
        Dim rUserFlagCPUArrayString As String = Nothing
        Dim aUserFlagArray(32) As Integer
        Dim aUserFlagCPUArray(32) As Integer
        Dim rPPSu As Integer = Nothing
        Dim rPPsu_RearbackVoltage As Double = Nothing
        Dim rPPsu_RearbackCorrent As Double = Nothing
        Dim rError As String = Nothing
        Dim aUsedChForStuck(2) As Short
        rpRpkTpgmVariant = "5CTH1AF02"

        fFCT = PASS

        ' *** Set New Current Directory
        Directory.SetCurrentDirectory(Application.StartupPath)

        Try


            rpCanPort = 1
            rpCanSpeed = 500000
            rpUutFctResult(0) = PASS
            For rSite As Integer = 0 To 0

                If rpUutFctResult(rSite) = PASS Then
                    ' ********************************************************************************************************************************************
                    ' *** CAN Port Open
                    ' ********************************************************************************************************************************************
                    'GoTo WRITEPNS

                    If fCanPortOpen(rpCanPortHandle, rpCanPort, rpCanSpeed) <> 0 Then
                        fFCT = FAIL

                    End If

                    If fFCT <> FAIL Then


                        ' ************************************************************************************************************************
                        ' *** Sink Test BEFORE FKT
                        ' ************************************************************************************************************************
                        Dim rpFCTNtest As Integer = 3000                                                          ' *** Set Test Number
                        Dim rpFCTDrawRef As String = "FCT"                                                      ' *** Set Drawing Reference
                        Dim rpFCTTestName As String = "Sink Before FKT"                                           ' *** Set Test Name
                        Dim rpFCTDiagRemark As String = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                        rpTh = 0.07                                                          ' *** Set High Threshold 
                        rpTl = 0.04                                                            ' *** Set Low Threshold
                        Dim rpFCTUnitMeasure As String = "A"                                                     ' *** Set Unit Measure

                        'GoTo Unlock

                        ' ********************************************************************************************************************************************
                        ' *** SERIAL NUMBER READ
                        ''********************************************************************************************************************************************
                        foo.filout("Read SRN", "Read serial number from the part micro", "Nan", False)
                        rpFCTNtest = 3007                                                          ' *** Set Test Number
                        rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                        rpFCTTestName = "Read SN"                                                   ' *** Set Test Name
                        rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                        rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                        Dim rSerialNumberRead As String = Nothing
                        If fSerialNumberRead(rSite, rSerialNumberRead) <> PASS Then
                            rpUutFctResult(rSite) = FAIL
                            'MsgBox("RAED SN FAIL")
                            foo.filout("RAED SN FAIL", "Unable to read serial number from the part", "Nan", True)
                            'MsgPrintLogIfEn("SN: " & rSerialNumberRead, 0, "ALWAYS")
                            'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                            GoTo NEXT_SITE
                        End If
                        'MsgPrintLogIfEn("SN: " & rSerialNumberRead, 0, "ALWAYS")
                        'rpMsgLogClass.MsgPrintLogMultyBay("SN", rpRpkBayInUse, rSerialNumberRead)
                        'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                        'rSerialNumberRead = "D171234567"
                        rpUutFctSerialNumber(0) = rSerialNumberRead

                        If rSerialNumberRead = "" Or Left(rpUutFctSerialNumber(0), 1) <> "D" Then


                            MsgBox("                        POSITION1" & Chr(10) & Chr(10) & "SN :  " & rSerialNumberRead & Chr(10) & Chr(10) & "part to be discard")

                            fFCT = FAIL
                            GoTo NEXT_SITE
                        End If

                        'rpUutFctSerialNumber(rSite) = "D131234567"

                        MsgBox("POSITION1" & Chr(10) & Chr(10) & "SN:  " & rSerialNumberRead & Chr(10) & Chr(10) & "PLEASE REFLASH BLU + APP THEN PRESS OK")

                        GoTo UnlockL3


                        com("PAYLOADFC")

                        com("PAYLOADSC")


                        ' ********************************************************************************************************************************************
                        ' *** Unlock Device LEVEL 3 WRITE DE14
                        ' ********************************************************************************************************************************************
UnlockL3:               rpFCTNtest = 3001                                                          ' *** Set Test Number
                            rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            rpFCTTestName = "Unlock L3 Device"                                                 ' *** Set Test Name
                            rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                            rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                        If fDUTUnlock3(0) <> PASS Then
                            rpUutFctResult(0) = FAIL
                            fFCT = FAIL
                            MsgBox("UNCLOCK L3 FAIL")
                            'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                            GoTo NEXT_SITE
                        End If

                        'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")




                        ' ********************************************************************************************************************************************
                        ' *** Unlock Device LEVEL 1
                        ' ********************************************************************************************************************************************
                        rpFCTNtest = 3001                                                          ' *** Set Test Number
                            rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            rpFCTTestName = "Unlock L1 Device"                                                 ' *** Set Test Name
                            rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                            rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                            If fDUTUnlock(rSite) <> PASS Then
                                fFCT = FAIL
                            MsgBox("UNLOCK L1 FAIL")
                            'MsgBox("                        POSITION1" & Chr(10) & Chr(10) & "UNLOCK L1 FAIL" & Chr(10) & Chr(10) & "PLEASE REFLASH BLU + APP THEN PRESS OK")
                            rpUutFctResult(rSite) = FAIL
                            'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                            GoTo NEXT_SITE
                        End If
                            'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")


                            ' ********************************************************************************************************************************************
                            ' *** Secondary Bootloader Download 
                            ' ********************************************************************************************************************************************
                            rpFCTNtest = 3002                                                          ' *** Set Test Number
                            rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            rpFCTTestName = "SBL Download"                             ' *** Set Test Name
                            rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                            rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                            If fDUTSBLDownload(rSite) <> PASS Then
                                fFCT = FAIL
                                MsgBox(" DOWNLOAD SBL FAIL")
                                rpUutFctResult(rSite) = FAIL
                                'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                                GoTo NEXT_SITE
                            End If
                            'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")


                            ' ********************************************************************************************************************************************
                            ' *** Secondary Bootloader Startup
                            ' ********************************************************************************************************************************************
                            rpFCTNtest = 3003                                                          ' *** Set Test Number
                            rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            rpFCTTestName = "SBL Startup"                             ' *** Set Test Name
                            rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                            rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                            If fDUTSBLStartup(rSite) <> PASS Then
                                rpUutFctResult(rSite) = FAIL
                                fFCT = FAIL
                                MsgBox(" Bootloader Startup FAIL")
                                'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                                GoTo NEXT_SITE
                            End If
                            'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                            'GoTo PROVA1

                            ' ********************************************************************************************************************************************
                            ' *** Write 0xF111 0xF113
                            ' ********************************************************************************************************************************************
WRITEPNS:                   rpFCTNtest = 3004                                                          ' *** Set Test Number
                            rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            rpFCTTestName = "Write 0xF111 0xF113"                             ' *** Set Test Name
                            rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                            rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                            If fDUTWriteByID(rSite) <> PASS Then
                                rpUutFctResult(rSite) = FAIL
                                fFCT = FAIL
                                MsgBox("write F111/F113 FAIL")
                                'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                                GoTo NEXT_SITE
                            End If
                            'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")


                            ' ********************************************************************************************************************************************
                            ' *** Write Julian Date
                            ' ********************************************************************************************************************************************
                            rpFCTNtest = 3005                                                          ' *** Set Test Number
                            rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            rpFCTTestName = "Write Julian Date"                             ' *** Set Test Name
                            rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                            rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                            If fDUTWriteJulianDate(rSite) <> PASS Then
                                rpUutFctResult(rSite) = FAIL
                                fFCT = FAIL
                                MsgBox("write DATE FAIL")
                                'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                                GoTo NEXT_SITE
                            End If
                            'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")


                            ' ********************************************************************************************************************************************
                            ' *** SERIAL NUMBER WRITE
                            '********************************************************************************************************************************************
                            rpFCTNtest = 3006                                                          ' *** Set Test Number
                            rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            rpFCTTestName = "Write SN"                             ' *** Set Test Name
                            rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                            rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                            If fSerialNumberWrite(rSite) <> PASS Then
                                rpUutFctResult(rSite) = FAIL
                                fFCT = FAIL
                                MsgBox("write SN FAIL")
                                'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                                GoTo NEXT_SITE
                            End If
                            'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")

                            'PROVA1:
                            ' ********************************************************************************************************************************************
                            ' *** SERIAL NUMBER READ
                            ''********************************************************************************************************************************************
                            'rpFCTNtest = 3007                                                          ' *** Set Test Number
                            'rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            'rpFCTTestName = "Read SN"                                                   ' *** Set Test Name
                            'rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                            'rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                            'Dim rSerialNumberRead As String = Nothing
                            'If fSerialNumberRead(rSite, rSerialNumberRead) <> PASS Then
                            '    rpUutFctResult(rSite) = FAIL
                            '    MsgBox("RAED SN FAIL")
                            '    'MsgPrintLogIfEn("SN: " & rSerialNumberRead, 0, "ALWAYS")
                            '    'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")
                            '    GoTo NEXT_SITE
                            'End If
                            ''MsgPrintLogIfEn("SN: " & rSerialNumberRead, 0, "ALWAYS")
                            ''rpMsgLogClass.MsgPrintLogMultyBay("SN", rpRpkBayInUse, rSerialNumberRead)
                            ''MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")




NEXT_SITE:

                            '' *** PowerOff Board
                            'PpsuOff(PPS4)
                            '' *** UserFlag Reset
                            'UserFlagsReset(RLYNOCPU, aUserFlagCPUArray(0))
                            'UserFlagsReset(RLYNOA, aUserFlagArray(0))
                            '' *** PowerOff HW1/HW2 - FXSEZ200
                            'PpsOff(PPS1, 5)

                            ' ********************************************************************************************************************************************
                            ' *** CAN Port Close
                            ' ********************************************************************************************************************************************
                            If fCanPortClose(rpCanPortHandle, rpCanPort) <> 0 Then
                                fFCT = FAIL
                            End If



                        End If

                    End If



            Next

        Catch ex As Exception
            rError = ex.Message

            If fCanPortClose(rpCanPortHandle, rpCanPort) <> 0 Then
                fFCT = FAIL
            End If
        End Try

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        'If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function


    Public Sub com(ByVal payload As String)
        Dim proc As Process = Nothing
        Dim batDir As String = String.Format(Application.StartupPath + "\autoload")
        proc = New Process()
        proc.StartInfo.WorkingDirectory = batDir
        proc.StartInfo.FileName = payload + ".bat"
        proc.StartInfo.CreateNoWindow = False
        proc.Start()
        proc.WaitForExit()
    End Sub
    Public Function fDUTUnlock3(ByVal rSite As Integer) As Long


        fDUTUnlock3 = FAIL

        Dim rError As String = Nothing

        Dim rFrameToTx As String = ""
        Dim rFrameToTx10 As String = ""
        Dim rFrameToTx0 As String = ""
        Dim rFrameToTx1 As String = ""
        Dim rFrameToTx2 As String = ""
        Dim rFrameToTx3 As String = ""
        Dim rFrameToTxDE14 As String = ""

        Dim rExpectedFrameRx(7) As String
        Dim rExpectedFrameRx0(7) As String
        Dim rExpectedFrameRx1(7) As String
        Dim rExpectedFrameRx2(7) As String
        Dim rExpectedFrameRx3(7) As String
        Dim rExpectedFrameRx4(7) As String

        Dim rFrameRx(7) As String
        Dim rFrameRx0(7) As String
        Dim rFrameRx1(7) As String
        Dim rFrameRx2(7) As String
        Dim rFrameRx3(7) As String
        Dim rFrameRx4(7) As String

        Dim rRetry As Integer
        Dim rRetryr As Integer
        Dim rRetryMax As Integer
        Dim aKeyBytes(18) As Byte

        Try

            ' ********************************************************************************************************************************************
            ' *** Load File and check data
            ' ********************************************************************************************************************************************
            'rpRpkTpgmVariant = "Variant"
            Dim rDataFile As StreamReader = File.OpenText("C:\Inifiles\SDLC\" & rpRpkTpgmVariant & ".txt")

            Dim rvar_Data As String = Nothing

            Do While rDataFile.Peek() > -1
                Dim Data As String = rDataFile.ReadLine()
                If Data.Contains("DE14:") And Data.Length > 5 Then
                    Data = Data.Substring(5)
                    Data = RemoveWhitespace(Data)
                    If Data.Length = 1 Then
                        rvar_Data = Data
                    Else
                        rError = "DE14 has the wrong length"
                        GoTo EndWithFail
                    End If
                    '////////////////////////////////////////////////////////////////////////////////////////////////
                End If
            Loop

            If rvar_Data Is Nothing Then
                rError = "DE14 IS missing!"
                GoTo EndWithFail
            End If

            ' *** Message to transmitt
            If rvar_Data = "D" Then
                rFrameToTxDE14 = "042EDE146F000000"
            ElseIf rvar_Data = "E" Then
                rFrameToTxDE14 = "042EDE146F000000"
            ElseIf rvar_Data = "F" Then
                rFrameToTxDE14 = "042EDE146B000000"
            Else
                rpTestResult = PASS
                GoTo Var_OTHER
            End If
            ' ********************************************************************************************************************************************
            ' *** REQUEST START SESSION (EXTENDED SESSION)
            '            ' ********************************************************************************************************************************************
            '            rRetry = 0
            '            rRetryMax = 10

            'TEST_MODE_RETRY:

            '            rpNtest = 2000                                                          ' *** Set Test Number
            '            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            '            rpTestName = "Test Mode Enable"                                         ' *** Set Test Name
            '            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            '            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            '            rpTestResult = FAIL

            '            ' *** ECU RESET
            '            rFrameToTx10 = "0211010000000000"
            '            ' *** Message to transmitt
            '            rFrameToTx = "0210030000000000"
            '            ' *** Message to be received
            '            rExpectedFrameRx = {"06", "50", "03", "00", "32", "01", "F4", "00"}
            '            ' *** Clear Can Buffer
            '            fCanBufferClearFlush(50)
            '            ' *** Send Command
            '            'Console.WriteLine("SITE " & rSite & " port" & CanPort & ". before tx")
            '            'MySemaphore.WaitOne()
            '            'If fCanTxFrames(rpCanPort, "716", rFrameToTx10) = FAIL Then GoTo EndWithFail
            '            Thread.Sleep(60)
            '            If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
            '            'MySemaphore.Release()
            '            'Console.WriteLine("SITE " & rSite & " port" & CanPort & ". after tx")

            ' *** Wait Response
            'Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            'If fCanRxFrames(rpCanPort, "71E", rFrameRx) = FAIL Then
            '    If rRetry < rRetryMax Then
            '        rRetry = rRetry + 1
            '        GoTo TEST_MODE_RETRY
            '    Else
            '        GoTo EndWithFail
            '    End If
            'End If

            ' *** Check  Message
            'For rIndex As Integer = 0 To 7
            '    If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
            '        rpTestResult = PASS
            '    Else
            '        rpTestResult = FAIL
            '        Exit For
            '    End If
            'Next
            '' *** Retry Manage
            'If rRetry < rRetryMax And rpTestResult = FAIL Then
            '    rRetry = rRetry + 1
            '    GoTo TEST_MODE_RETRY
            'ElseIf rpTestResult = FAIL Then
            '    GoTo EndWithFail
            'End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** SEED REQUEST SECURITY LEVEL 3
            ' ********************************************************************************************************************************************
            '            rRetry = 0
            '            rRetryMax = 10

            'SEED_REQUEST_RETRY:

            '            rpNtest = 2001                                                          ' *** Set Test Number
            '            rpDrawRef = "FCT"                                                      ' *** Set Drawing Reference
            '            rpTestName = "Seed Request"                                             ' *** Set Test Name
            '            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            '            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            '            rpTestResult = FAIL

            '            ' *** Message to transmitt 1
            '            rFrameToTx1 = "0227030000000000"
            '            ' *** Message to be received 1
            '            rExpectedFrameRx1 = {"10", "12", "67", "03", "seed0", "seed1", "seed2", "seed3"}
            '            ' *** Message to transmitt 2
            '            rFrameToTx0 = "3000000000000000"
            '            ' *** Message to be received 2
            '            rExpectedFrameRx2 = {"21", "seed4", "seed5", "seed6", "seed7", "seed8", "seed9", "seed10"}
            '            ' *** Message to be received 3
            '            rExpectedFrameRx3 = {"22", "seed11", "seed12", "seed13", "seed14", "seed15", "00", "00"}


            '            ' *** Clear Can Buffer
            '            fCanBufferClearFlush(15)

            '            ' *** Send Command ---------------------------------------------------
            '            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            '            ' *** Wait Response
            '            Thread.Sleep(60)
            '            ' *** Check Buffer For Specific ID 
            '            If fCanRxFrames(rpCanPort, "071E", rFrameRx1) = FAIL Then
            '                If rRetry < rRetryMax Then
            '                    rRetry = rRetry + 1
            '                    GoTo SEED_REQUEST_RETRY
            '                Else
            '                    GoTo EndWithFail
            '                End If
            '            End If

            '            ' *** Send Command ---------------------------------------------------
            '            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            '            ' *** Wait Response
            '            Thread.Sleep(60)
            '            ' *** Check Buffer For Specific ID 
            '            If fCanRxFrames(rpCanPort, "071E", rFrameRx2) = FAIL Then
            '                If rRetry < rRetryMax Then
            '                    rRetry = rRetry + 1
            '                    GoTo SEED_REQUEST_RETRY
            '                Else
            '                    GoTo EndWithFail
            '                End If
            '            End If
            '            ' *** Wait Response
            '            Thread.Sleep(60)
            '            ' *** Check Buffer For Specific ID 
            '            If fCanRxFrames(rpCanPort, "071E", rFrameRx3) = FAIL Then
            '                If rRetry < rRetryMax Then
            '                    rRetry = rRetry + 1
            '                    GoTo SEED_REQUEST_RETRY
            '                Else
            '                    GoTo EndWithFail
            '                End If
            '            End If

            '            ' *** Check  Reply 1
            '            For rIndex As Integer = 0 To 7
            '                If (rIndex <> 4) And (rIndex <> 5) And (rIndex <> 6) And (rIndex <> 7) Then
            '                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
            '                        rpTestResult = PASS
            '                    Else
            '                        rpTestResult = FAIL
            '                        Exit For
            '                    End If
            '                End If
            '            Next


            '            ' *** Check  Reply 2
            '            If rpTestResult = PASS Then
            '                For rIndex As Integer = 0 To 7
            '                    If (rIndex <> 1) And (rIndex <> 2) And (rIndex <> 3) And (rIndex <> 4) And (rIndex <> 5) And (rIndex <> 6) And (rIndex <> 7) Then
            '                        If rFrameRx2(rIndex) = rExpectedFrameRx2(rIndex) Then
            '                            rpTestResult = PASS
            '                        Else
            '                            rpTestResult = FAIL
            '                            Exit For
            '                        End If
            '                    End If
            '                Next
            '            End If

            '            ' *** Check  Reply 3
            '            If rpTestResult = PASS Then
            '                For rIndex As Integer = 0 To 7
            '                    If (rIndex <> 1) And (rIndex <> 2) And (rIndex <> 3) And (rIndex <> 4) And (rIndex <> 5) Then
            '                        If rFrameRx3(rIndex) = rExpectedFrameRx3(rIndex) Then
            '                            rpTestResult = PASS
            '                        Else
            '                            rpTestResult = FAIL
            '                            Exit For
            '                        End If
            '                    End If
            '                Next
            '            End If

            '            ' *** Retry Manage
            '            If rRetry < rRetryMax And rpTestResult = FAIL Then
            '                rRetry = rRetry + 1
            '                GoTo SEED_REQUEST_RETRY
            '            ElseIf rpTestResult = FAIL Then
            '                GoTo EndWithFail
            '            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************


            ' ********************************************************************************************************************************************
            ' *** CALCULATE THE KEY
            ' ********************************************************************************************************************************************


            Dim SecurityKey(17) As Byte
            Dim aSeedBytes(15) As Byte
            Dim FixedBytes As Byte() = {&HAC, &H6E, &H88, &H7B, &H11, &HB9, &H92, &HBF, &H2D, &H1C, &H1E, &HB1}
            'Dim unlock_key As Byte() = {&HA8, &HDB, &H6A, &H4D, &H12, &H4C, &HBE, &H79, &H1, &HEF, &H69, &HD5, &H4F, &H7B, &HF7, &H8F, &H47, &HC4, &H58, &HEB, &H2D, &HCC, &HF1, &H64, &HD, &HA3, &H52, &HF0, &H3B, &H11, &HB3, &HED}
            Dim unlock_key As Byte() = {&HD9, &HBD, &H8E, &HA2, &HBC, &H36, &H3B, &H6E, &HAF, &HB8, &H68, &HCB, &H91, &H3B, &H36, &HA7, &H69, &H6E, &HEC, &H2E, &H87, &H31, &HC7, &HC3, &HC, &H35, &H99, &HD5, &H89, &HE4, &HB2, &H9E}

            aSeedBytes(0) = CUInt(fConv_HexToDec(rFrameRx1(4)))
            aSeedBytes(1) = CUInt(fConv_HexToDec(rFrameRx1(5)))
            aSeedBytes(2) = CUInt(fConv_HexToDec(rFrameRx1(6)))
            aSeedBytes(3) = CUInt(fConv_HexToDec(rFrameRx1(7)))

            aSeedBytes(4) = CUInt(fConv_HexToDec(rFrameRx2(1)))
            aSeedBytes(5) = CUInt(fConv_HexToDec(rFrameRx2(2)))
            aSeedBytes(6) = CUInt(fConv_HexToDec(rFrameRx2(3)))
            aSeedBytes(7) = CUInt(fConv_HexToDec(rFrameRx2(4)))
            aSeedBytes(8) = CUInt(fConv_HexToDec(rFrameRx2(5)))
            aSeedBytes(9) = CUInt(fConv_HexToDec(rFrameRx2(6)))
            aSeedBytes(10) = CUInt(fConv_HexToDec(rFrameRx2(7)))

            aSeedBytes(11) = CUInt(fConv_HexToDec(rFrameRx3(1)))
            aSeedBytes(12) = CUInt(fConv_HexToDec(rFrameRx3(2)))
            aSeedBytes(13) = CUInt(fConv_HexToDec(rFrameRx3(3)))
            aSeedBytes(14) = CUInt(fConv_HexToDec(rFrameRx3(4)))
            aSeedBytes(15) = CUInt(fConv_HexToDec(rFrameRx3(5)))


            Dim myKeyUnlocker As FordKeyUp.Security = New FordKeyUp.Security()


            Dim unlock As Boolean = myKeyUnlocker.Povolit32(unlock_key)
            If Not unlock Then
                MsgBox("Impossible to unlock Ford Dll!!!", 0, "ALWAYS")
                'rpMsgLogClass.MsgPrintLogMultyBay("@BG{Red}@bFG{White}Impossible to unlock Ford Dll!!!", rpRpkBayInUse)
                GoTo EndWithFail
            End If
            Dim AlgorithmType As Integer = 1
            Dim Result As Byte = myKeyUnlocker.GetKey(aSeedBytes, FixedBytes, AlgorithmType, SecurityKey)
            If Result <> 1 Then
                'MsgPrintLogIfEn("Impossible to get a valid Key!!!", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("@BG{Red}@bFG{White}Impossible to get a valid Key!!!", rpRpkBayInUse)
                GoTo EndWithFail
            End If

            ' '' *** ONLY DEBUG
            ''If rpVbNetIsDebug Then

            ''    MsgPrintLogIfEn("--- aSeedBytes ---", 0, "ALWAYS")
            ''    For Z As Integer = 0 To aSeedBytes.Length - 1
            ''        MsgPrintLogIfEn(Z & " - " & aSeedBytes(Z), 0, "ALWAYS")
            ''    Next

            ''    MsgPrintLogIfEn("--- aKeyBytes ---", 0, "ALWAYS")
            ''    For Z As Integer = 0 To aKeyBytes.Length - 1
            ''        MsgPrintLogIfEn(Z & " - " & aKeyBytes(Z), 0, "ALWAYS")
            ''    Next

            ''End If

            ' ********************************************************************************************************************************************
            ' *** SEND KEY
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 3

SEND_KEY_RETRY:

            rpNtest = 2002                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                      ' *** Set Drawing Reference
            rpTestName = "Key Send (Unlock Request)"                                                 ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt 1
            rFrameToTx0 = "10" & "12" & "27" & "04" &
                                    fConv_DecToHex(CDbl(SecurityKey(0))) &
                                    fConv_DecToHex(CDbl(SecurityKey(1))) &
                                    fConv_DecToHex(CDbl(SecurityKey(2))) &
                                    fConv_DecToHex(CDbl(SecurityKey(3)))

            rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}


            ' *** Message to transmitt 2

            rFrameToTx1 = "21" & fConv_DecToHex(CDbl(SecurityKey(4))) &
                                    fConv_DecToHex(CDbl(SecurityKey(5))) &
                                    fConv_DecToHex(CDbl(SecurityKey(6))) &
                                    fConv_DecToHex(CDbl(SecurityKey(7))) &
                                    fConv_DecToHex(CDbl(SecurityKey(8))) &
                                    fConv_DecToHex(CDbl(SecurityKey(9))) &
                                    fConv_DecToHex(CDbl(SecurityKey(10)))


            ' *** Message to transmitt 3
            rFrameToTx2 = "22" & fConv_DecToHex(CDbl(SecurityKey(11))) &
                                    fConv_DecToHex(CDbl(SecurityKey(12))) &
                                    fConv_DecToHex(CDbl(SecurityKey(13))) &
                                    fConv_DecToHex(CDbl(SecurityKey(14))) &
                                    fConv_DecToHex(CDbl(SecurityKey(15))) &
                                    "00" & "00"

            ' *** Message to be received 1
            rExpectedFrameRx1 = {"02", "67", "04", "00", "00", "00", "00", "00"}


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail

            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo SEND_KEY_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx2) = FAIL Then GoTo EndWithFail

            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo SEND_KEY_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If


            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next


            ' *** Check  Reply 1
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If


            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo SEND_KEY_RETRY
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If



            ' ********************************************************************************************************************************************
            ' *** Write 0XDE14
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

WRITE_0xDE14:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Write 0xDE14"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to be received
            rExpectedFrameRx = {"03", "6E", "DE", "14", "00", "00", "00", "00"}
            rExpectedFrameRx0 = {"03", "7F", "2E", "78", "00", "00", "00", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTxDE14) = FAIL Then GoTo EndWithFail

            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo WRITE_0xDE14
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Check  Reply 1
            'If rpTestResult = PASS Then
            If rFrameRx(1) = "7F" And rFrameRx(3) = "78" Then
                ' *** Wait Response
                rRetryr = 0
READ_0xDE14:
                Thread.Sleep(250)
                If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                    ' *** Retry Manage
                    If rRetryr < rRetryMax Then
                        rRetryr = rRetryr + 1
                        GoTo READ_0xDE14
                        'ElseIf rpTestResult = FAIL Then
                        '    GoTo EndWithFail
                    End If
                    'GoTo EndWithFail
                Else
                    For rIndex As Integer = 0 To 7
                        If rFrameRx0(rIndex) = rExpectedFrameRx(rIndex) Then
                            rpTestResult = PASS
                        Else
                            rpTestResult = FAIL
                            Exit For
                        End If
                    Next
                End If
            Else
                For rIndex As Integer = 0 To 7
                    If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If
            ' End If

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo WRITE_0xDE14
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' 
            ' *** Message to transmitt
            rFrameToTx = "0210010000000000"
            ' *** Message to be received
            rExpectedFrameRx = {"06", "50", "01", "00", "32", "01", "F4", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(50)
            ' *** Send Command
            'Console.WriteLine("SITE " & rSite & " port" & CanPort & ". before tx")
            'MySemaphore.WaitOne()
            If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
            'MySemaphore.Release()
            'Console.WriteLine("SITE " & rSite & " port" & CanPort & ". after tx")

            ' *** Wait Response
            Thread.Sleep(100)

Var_OTHER:
            '////////////////////////////////////////////////////////////////////////////////////////////////

            ' *** COMPLETED CORRECTLY
            If rpTestResult = PASS Then GoTo EndWithPass
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fDUTUnlock3 = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fDUTUnlock3 = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fObp_FCT - Function execution result: " & fCRes(fDUTUnlock), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function
    Public Function fDUTUnlock(ByVal rSite As Integer) As Long

        fDUTUnlock = FAIL

        Dim rError As String = Nothing

        Dim rFrameToTx As String = ""
        Dim rFrameToTx0 As String = ""
        Dim rFrameToTx1 As String = ""
        Dim rFrameToTx2 As String = ""
        Dim rFrameToTx3 As String = ""

        Dim rExpectedFrameRx(7) As String
        Dim rExpectedFrameRx0(7) As String
        Dim rExpectedFrameRx1(7) As String
        Dim rExpectedFrameRx2(7) As String
        Dim rExpectedFrameRx3(7) As String
        Dim rExpectedFrameRx4(7) As String

        Dim rFrameRx(7) As String
        Dim rFrameRx0(7) As String
        Dim rFrameRx1(7) As String
        Dim rFrameRx2(7) As String
        Dim rFrameRx3(7) As String
        Dim rFrameRx4(7) As String

        Dim rRetry As Integer
        Dim rRetryMax As Integer
        Dim aKeyBytes(18) As Byte

        Try

            ' ********************************************************************************************************************************************
            ' *** REQUEST START SESSION (PROGRAMMING SESSION)
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

TEST_MODE_RETRY:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Test Mode Enable"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt
            rFrameToTx = "0210020000000000"
            ' *** Message to be received
            rExpectedFrameRx = {"06", "50", "02", "00", "19", "01", "F4", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(50)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo TEST_MODE_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Check  Message
            For rIndex As Integer = 0 To 7
                If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo TEST_MODE_RETRY
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** Request DataIdentifier 0xF162
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

READ_DATA1_RETRY:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Read Data &HF162"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL

            ' *** Message to transmitt
            rFrameToTx = "0322F16200000000"
            ' *** Message to be received
            rExpectedFrameRx = {"04", "62", "F1", "62", "06", "00", "00", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo READ_DATA1_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Check  Message
            For rIndex As Integer = 0 To 7
                If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo READ_DATA1_RETRY
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** Request DataIdentifier 0xF111
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

READ_DATA2_RETRY:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Read Data 0xF111"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL

            ' *** Message to transmitt 1
            rFrameToTx1 = "0322F11100000000"
            ' *** Message to be received 1
            rExpectedFrameRx1 = {"10", "1B", "62", "F1", "11", "FF", "FF", "FF"}
            ' *** Message to transmitt 2
            rFrameToTx0 = "3000000000000000"
            ' *** Message to be received 2
            rExpectedFrameRx2 = {"21", "FF", "FF", "FF", "FF", "FF", "FF", "FF"}
            ' *** Message to be received 3
            rExpectedFrameRx3 = {"22", "FF", "FF", "FF", "FF", "FF", "FF", "FF"}
            ' *** Message to be received 3
            rExpectedFrameRx4 = {"23", "FF", "FF", "FF", "FF", "FF", "FF", "FF"}


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo READ_DATA2_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx2) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo READ_DATA2_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx3) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo READ_DATA2_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx4) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo READ_DATA2_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Check  Reply 1
            For rIndex As Integer = 0 To 7
                If (rIndex <> 4) And (rIndex <> 5) And (rIndex <> 6) And (rIndex <> 7) Then
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                End If
            Next

            ' *** Check  Reply 2
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If (rIndex <> 1) And (rIndex <> 2) And (rIndex <> 3) And (rIndex <> 4) And (rIndex <> 5) And (rIndex <> 6) And (rIndex <> 7) Then
                        If rFrameRx2(rIndex) = rExpectedFrameRx2(rIndex) Then
                            rpTestResult = PASS
                        Else
                            rpTestResult = FAIL
                            Exit For
                        End If
                    End If
                Next
            End If

            ' *** Check  Reply 3
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If (rIndex <> 1) And (rIndex <> 2) And (rIndex <> 3) And (rIndex <> 4) And (rIndex <> 5) Then
                        If rFrameRx3(rIndex) = rExpectedFrameRx3(rIndex) Then
                            rpTestResult = PASS
                        Else
                            rpTestResult = FAIL
                            Exit For
                        End If
                    End If
                Next
            End If
            ' *** Check  Reply 4
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If (rIndex <> 1) And (rIndex <> 2) And (rIndex <> 3) And (rIndex <> 4) And (rIndex <> 5) Then
                        If rFrameRx4(rIndex) = rExpectedFrameRx4(rIndex) Then
                            rpTestResult = PASS
                        Else
                            rpTestResult = FAIL
                            Exit For
                        End If
                    End If
                Next
            End If
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo READ_DATA2_RETRY
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** SEED REQUEST SECURITY LEVEL 1
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

SEED_REQUEST_RETRY:

            rpNtest = 2001                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                      ' *** Set Drawing Reference
            rpTestName = "Seed Request"                                             ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL

            ' *** Message to transmitt 1
            rFrameToTx1 = "0227010000000000"
            ' *** Message to be received 1
            rExpectedFrameRx1 = {"10", "12", "67", "01", "seed0", "seed1", "seed2", "seed3"}
            ' *** Message to transmitt 2
            rFrameToTx0 = "3000000000000000"
            ' *** Message to be received 2
            rExpectedFrameRx2 = {"21", "seed4", "seed5", "seed6", "seed7", "seed8", "seed9", "seed10"}
            ' *** Message to be received 3
            rExpectedFrameRx3 = {"22", "seed11", "seed12", "seed13", "seed14", "seed15", "00", "00"}


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo SEED_REQUEST_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx2) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo SEED_REQUEST_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx3) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo SEED_REQUEST_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Check  Reply 1
            For rIndex As Integer = 0 To 7
                If (rIndex <> 4) And (rIndex <> 5) And (rIndex <> 6) And (rIndex <> 7) Then
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                End If
            Next


            ' *** Check  Reply 2
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If (rIndex <> 1) And (rIndex <> 2) And (rIndex <> 3) And (rIndex <> 4) And (rIndex <> 5) And (rIndex <> 6) And (rIndex <> 7) Then
                        If rFrameRx2(rIndex) = rExpectedFrameRx2(rIndex) Then
                            rpTestResult = PASS
                        Else
                            rpTestResult = FAIL
                            Exit For
                        End If
                    End If
                Next
            End If

            ' *** Check  Reply 3
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If (rIndex <> 1) And (rIndex <> 2) And (rIndex <> 3) And (rIndex <> 4) And (rIndex <> 5) Then
                        If rFrameRx3(rIndex) = rExpectedFrameRx3(rIndex) Then
                            rpTestResult = PASS
                        Else
                            rpTestResult = FAIL
                            Exit For
                        End If
                    End If
                Next
            End If

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo SEED_REQUEST_RETRY
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************


            ' ********************************************************************************************************************************************
            ' *** CALCULATE THE KEY
            ' ********************************************************************************************************************************************


            Dim SecurityKey(17) As Byte
            Dim aSeedBytes(15) As Byte
            Dim FixedBytes As Byte() = {&HAC, &H6E, &H88, &H7B, &H11, &HB9, &H92, &HBF, &H2D, &H1C, &H1E, &HB1}
            Dim unlock_key As Byte() = {&HA8, &HDB, &H6A, &H4D, &H12, &H4C, &HBE, &H79, &H1, &HEF, &H69, &HD5, &H4F, &H7B, &HF7, &H8F, &H47, &HC4, &H58, &HEB, &H2D, &HCC, &HF1, &H64, &HD, &HA3, &H52, &HF0, &H3B, &H11, &HB3, &HED}

            aSeedBytes(0) = CUInt(fConv_HexToDec(rFrameRx1(4)))
            aSeedBytes(1) = CUInt(fConv_HexToDec(rFrameRx1(5)))
            aSeedBytes(2) = CUInt(fConv_HexToDec(rFrameRx1(6)))
            aSeedBytes(3) = CUInt(fConv_HexToDec(rFrameRx1(7)))

            aSeedBytes(4) = CUInt(fConv_HexToDec(rFrameRx2(1)))
            aSeedBytes(5) = CUInt(fConv_HexToDec(rFrameRx2(2)))
            aSeedBytes(6) = CUInt(fConv_HexToDec(rFrameRx2(3)))
            aSeedBytes(7) = CUInt(fConv_HexToDec(rFrameRx2(4)))
            aSeedBytes(8) = CUInt(fConv_HexToDec(rFrameRx2(5)))
            aSeedBytes(9) = CUInt(fConv_HexToDec(rFrameRx2(6)))
            aSeedBytes(10) = CUInt(fConv_HexToDec(rFrameRx2(7)))

            aSeedBytes(11) = CUInt(fConv_HexToDec(rFrameRx3(1)))
            aSeedBytes(12) = CUInt(fConv_HexToDec(rFrameRx3(2)))
            aSeedBytes(13) = CUInt(fConv_HexToDec(rFrameRx3(3)))
            aSeedBytes(14) = CUInt(fConv_HexToDec(rFrameRx3(4)))
            aSeedBytes(15) = CUInt(fConv_HexToDec(rFrameRx3(5)))


            Dim myKeyUnlocker As FordKeyUp.Security = New FordKeyUp.Security()


            Dim unlock As Boolean = myKeyUnlocker.Povolit32(unlock_key)
            If Not unlock Then
                'MsgPrintLogIfEn("Impossible to unlock Ford Dll!!!", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("@BG{Red}@bFG{White}Impossible to unlock Ford Dll!!!", rpRpkBayInUse)
                GoTo EndWithFail
            End If
            Dim AlgorithmType As Integer = 2
            Dim Result As Byte = myKeyUnlocker.GetKey(aSeedBytes, FixedBytes, AlgorithmType, SecurityKey)
            If Result <> 1 Then
                'MsgPrintLogIfEn("Impossible to get a valid Key!!!", 0, "ALWAYS")
                rpMsgLogClass.MsgPrintLogMultyBay("@BG{Red}@bFG{White}Impossible to get a valid Key!!!", rpRpkBayInUse)
                GoTo EndWithFail
            End If

            ' '' *** ONLY DEBUG
            ''If rpVbNetIsDebug Then

            ''    MsgPrintLogIfEn("--- aSeedBytes ---", 0, "ALWAYS")
            ''    For Z As Integer = 0 To aSeedBytes.Length - 1
            ''        MsgPrintLogIfEn(Z & " - " & aSeedBytes(Z), 0, "ALWAYS")
            ''    Next

            ''    MsgPrintLogIfEn("--- aKeyBytes ---", 0, "ALWAYS")
            ''    For Z As Integer = 0 To aKeyBytes.Length - 1
            ''        MsgPrintLogIfEn(Z & " - " & aKeyBytes(Z), 0, "ALWAYS")
            ''    Next

            ''End If

            ' ********************************************************************************************************************************************
            ' *** SEND KEY
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

SEND_KEY_RETRY:

            rpNtest = 2002                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                      ' *** Set Drawing Reference
            rpTestName = "Key Send (Unlock Request)"                                                 ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt 1
            rFrameToTx0 = "10" & "14" & "27" & "02" &
                                    fConv_DecToHex(CDbl(SecurityKey(0))) &
                                    fConv_DecToHex(CDbl(SecurityKey(1))) &
                                    fConv_DecToHex(CDbl(SecurityKey(2))) &
                                    fConv_DecToHex(CDbl(SecurityKey(3)))

            rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}


            ' *** Message to transmitt 2

            rFrameToTx1 = "21" & fConv_DecToHex(CDbl(SecurityKey(4))) &
                                    fConv_DecToHex(CDbl(SecurityKey(5))) &
                                    fConv_DecToHex(CDbl(SecurityKey(6))) &
                                    fConv_DecToHex(CDbl(SecurityKey(7))) &
                                    fConv_DecToHex(CDbl(SecurityKey(8))) &
                                    fConv_DecToHex(CDbl(SecurityKey(9))) &
                                    fConv_DecToHex(CDbl(SecurityKey(10)))


            ' *** Message to transmitt 3
            rFrameToTx2 = "22" & fConv_DecToHex(CDbl(SecurityKey(11))) &
                                    fConv_DecToHex(CDbl(SecurityKey(12))) &
                                    fConv_DecToHex(CDbl(SecurityKey(13))) &
                                    fConv_DecToHex(CDbl(SecurityKey(14))) &
                                    fConv_DecToHex(CDbl(SecurityKey(15))) &
                                    fConv_DecToHex(CDbl(SecurityKey(16))) &
                                    fConv_DecToHex(CDbl(SecurityKey(17)))

            ' *** Message to be received 1
            rExpectedFrameRx1 = {"02", "67", "02", "00", "00", "00", "00", "00"}


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail

            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo SEND_KEY_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail

            ' *** Send Command ---------------------------------------------------
            If fCanTxFrames(rpCanPort, "716", rFrameToTx2) = FAIL Then GoTo EndWithFail

            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo SEND_KEY_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If


            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next


            ' *** Check  Reply 1
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If


            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo SEND_KEY_RETRY
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")

            ' *** COMPLETED CORRECTLY
            If rpTestResult = PASS Then GoTo EndWithPass
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fDUTUnlock = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fDUTUnlock = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fObp_FCT - Function execution result: " & fCRes(fDUTUnlock), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function


    Public Structure VBFFileDataBlock
        Public Address As String
        Public Length As String
        Public Data As String
        Public Checksum As String
    End Structure
    Public Function VBFParser(ByVal VBFFilePath As String, ByRef VBFDataBlocks() As VBFFileDataBlock) As Integer
        Dim bytes As Byte() = My.Computer.FileSystem.ReadAllBytes(VBFFilePath)

        Dim BracketsCount As Integer = 0
        Dim BracketsFound As Boolean = False

        Dim i As Integer = 0
        While (BracketsCount <> 0 Or (Not BracketsFound)) And i < bytes.Length
            If Chr(bytes(i)) = "{"c Then
                BracketsFound = True
                BracketsCount = BracketsCount + 1
            ElseIf Chr(bytes(i)) = "}"c Then
                BracketsCount = BracketsCount - 1
            End If
            i = i + 1
        End While

        If BracketsCount <> 0 Or (Not BracketsFound) Or i >= bytes.Length Then
            VBFDataBlocks = Nothing
            Return FAIL
        End If

        Dim nDataBlock As Integer = 0
        While i < bytes.Length
            Try
                Dim Address As String = bytes(i).ToString("X02") & bytes(i + 1).ToString("X02") & bytes(i + 2).ToString("X02") & bytes(i + 3).ToString("X02")
                i = i + 4
                Dim Length As Long = bytes(i) * 256 * 256 * 256 + bytes(i + 1) * 256 * 256 + bytes(i + 2) * 256 + bytes(i + 3)
                i = i + 4



                Dim LengthString As String = Length.ToString("X08")

                Dim Data As String = ""
                For i = i To i + Length - 1
                    Data &= bytes(i).ToString("X02")
                Next

                Dim CheckSum As String = bytes(i).ToString("X02") & bytes(i + 1).ToString("X02")
                i = i + 2


                ReDim Preserve VBFDataBlocks(nDataBlock)
                VBFDataBlocks(nDataBlock).Address = Address
                VBFDataBlocks(nDataBlock).Length = LengthString
                VBFDataBlocks(nDataBlock).Data = Data
                VBFDataBlocks(nDataBlock).Checksum = CheckSum
                nDataBlock = nDataBlock + 1
            Catch ex As Exception
                VBFDataBlocks = Nothing
                Return FAIL
            End Try

        End While

        Return PASS

    End Function

    Public Function fDUTSBLDownloadV2(ByVal rSite As Integer) As Long


        fDUTSBLDownloadV2 = FAIL

        Dim rError As String = Nothing

        Dim rFrameToTx As String = ""
        Dim rFrameToTx0 As String = ""
        Dim rFrameToTx1 As String = ""
        Dim rFrameToTx2 As String = ""
        Dim rFrameToTx3 As String = ""

        Dim rExpectedFrameRx(7) As String
        Dim rExpectedFrameRx0(7) As String
        Dim rExpectedFrameRx1(7) As String
        Dim rExpectedFrameRx2(7) As String
        Dim rExpectedFrameRx3(7) As String
        Dim rExpectedFrameRx4(7) As String

        Dim rFrameRx(7) As String
        Dim rFrameRx0(7) As String
        Dim rFrameRx1(7) As String
        Dim rFrameRx2(7) As String
        Dim rFrameRx3(7) As String
        Dim rFrameRx4(7) As String

        Dim rRetry As Integer
        Dim rRetryMax As Integer
        Dim aKeyBytes(18) As Byte

        Try
            ' ********************************************************************************************************************************************
            ' *** LOAD FILE
            ' ********************************************************************************************************************************************
            Dim myVBF() As VBFFileDataBlock
            VBFParser(rVBFFilePath, myVBF)
            If myVBF Is Nothing Then GoTo EndWithFail

            For Each VBFDataBlock As VBFFileDataBlock In myVBF

                ' ********************************************************************************************************************************************
                ' *** REQUEST TO DOWNLOAD DATA
                ' ********************************************************************************************************************************************
                rRetry = 0
                rRetryMax = 10

REQUEST_DOWNLOAD_DATA:

                rpNtest = 2000                                                          ' *** Set Test Number
                rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
                rpTestName = "Request to download data"                                         ' *** Set Test Name
                rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
                rpUnitMeasure = ""                                                      ' *** Set Unit Measure
                rpTestResult = FAIL


                ' *** Message to transmitt
                rFrameToTx0 = "100B341044" & VBFDataBlock.Address.Substring(0, 6)
                ' *** Message to be received
                rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}
                ' *** Message to transmitt
                rFrameToTx1 = "21" & VBFDataBlock.Address.Substring(6, 2) & VBFDataBlock.Length & "0000"
                ' *** Message to be received
                rExpectedFrameRx1 = {"04", "74", "20", "XX", "XX", "00", "00", "00"}


                ' *** Clear Can Buffer
                fCanBufferClearFlush(15)
                ' *** Send Command
                If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
                ' *** Wait Response
                Thread.Sleep(60)
                ' *** Check Buffer For Specific ID 
                If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                    If rRetry < rRetryMax Then
                        rRetry = rRetry + 1
                        GoTo REQUEST_DOWNLOAD_DATA
                    Else

                        GoTo EndWithFail
                    End If
                End If


                ' *** Send Command
                If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
                ' *** Wait Response
                Thread.Sleep(60)
                ' *** Check Buffer For Specific ID 
                If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
                    If rRetry < rRetryMax Then
                        rRetry = rRetry + 1
                        GoTo REQUEST_DOWNLOAD_DATA
                    Else

                        GoTo EndWithFail
                    End If
                End If





                ' *** Check  Reply 0
                For rIndex As Integer = 0 To 7
                    If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                        rpTestResult = PASS

                    Else
                        rpTestResult = FAIL

                        Exit For
                    End If
                Next

                ' *** Check  Reply 1
                If rpTestResult = PASS Then
                    For rIndex As Integer = 0 To 2
                        If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                            rpTestResult = PASS

                        Else
                            rpTestResult = FAIL

                            Exit For
                        End If
                    Next
                End If
                Dim MaxAcceptedLengthString As String = rFrameRx1(3) & rFrameRx1(4)
                Dim MaxAcceptedLength As Integer = Convert.ToInt32(MaxAcceptedLengthString, 16)

                ' *** Retry Manage
                If rRetry < rRetryMax And rpTestResult = FAIL Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_DOWNLOAD_DATA
                ElseIf rpTestResult = FAIL Then
                    GoTo EndWithFail
                End If
                ' *** Result Manage 
                ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
                ' ********************************************************************************************************************************************
                ' ********************************************************************************************************************************************


                Dim BlocksToSend As Integer = (VBFDataBlock.Data.Length \ 2 - 1) \ MaxAcceptedLength + 1
                Dim BlocksSent As Integer = 0
                While BlocksToSend > BlocksSent
                    Dim DataIndex As Integer = 0
                    Dim BlockData As String
                    If BlocksSent + 1 = BlocksToSend Then
                        Dim RemainingLength As Integer = (VBFDataBlock.Data.Length - MaxAcceptedLength * 2 * BlocksSent) \ 2
                        BlockData = VBFDataBlock.Data.Substring(MaxAcceptedLength * 2 * BlocksSent, RemainingLength * 2)
                    Else
                        BlockData = VBFDataBlock.Data.Substring(MaxAcceptedLength * 2 * BlocksSent, MaxAcceptedLength * 2)
                    End If

                    ' ********************************************************************************************************************************************
                    ' *** Send First Data
                    ' ********************************************************************************************************************************************
                    rpNtest = 2000                                                          ' *** Set Test Number
                    rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
                    rpTestName = "Send First Data"                                         ' *** Set Test Name
                    rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
                    rpUnitMeasure = ""                                                      ' *** Set Unit Measure
                    rpTestResult = FAIL


                    Dim Data As String
                    Dim SingleFrame As Boolean = False
                    If BlockData.Length < 4 * 2 Then
                        Dim MissingBytes As Integer = 4 - (BlockData.Length - DataIndex) \ 2
                        Data = BlockData.Substring(DataIndex, 4 * 2 - MissingBytes * 2) & New String("0"c, MissingBytes * 2)
                        DataIndex = DataIndex + (4 * 2 - MissingBytes * 2)
                        SingleFrame = True
                    Else
                        Data = BlockData.Substring(DataIndex, 4 * 2)
                        DataIndex = DataIndex + 4 * 2
                    End If


                    ' *** Message to transmitt
                    If SingleFrame Then
                        rFrameToTx = "0" & ((BlockData.Length \ 2) + 2).ToString("X03") & "36" & (BlocksSent + 1).ToString("X02") & Data
                    Else
                        rFrameToTx = "1" & ((BlockData.Length \ 2) + 2).ToString("X03") & "36" & (BlocksSent + 1).ToString("X02") & Data
                    End If

                    ' *** Message to be received
                    rExpectedFrameRx = {"30", "XX", "00", "00", "00", "00", "00", "00"}
                    ' *** Clear Can Buffer
                    fCanBufferClearFlush(15)
                    ' *** Send Command
                    If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
                    ' *** Wait Response
                    Thread.Sleep(60)
                    ' *** Check Buffer For Specific ID 
                    If fCanRxFrames(rpCanPort, "071E", rFrameRx) = FAIL Then
                        If rRetry < rRetryMax Then
                            rRetry = rRetry + 1

                            GoTo REQUEST_DOWNLOAD_DATA
                        Else

                            GoTo EndWithFail
                        End If






                    End If
                    If SingleFrame Then
                        GoTo MULTI_FRAME_SKIP
                    End If
                    ' *** Check  Message
                    For rIndex As Integer = 0 To 7
                        If rIndex <> 1 Then
                            If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                                rpTestResult = PASS
                            Else
                                rpTestResult = FAIL
                                Exit For




                            End If
                        End If
                    Next

                    Dim FlowControlNumber As Integer = Convert.ToInt32(rFrameRx(1), 16)


                    ' *** Retry Manage
                    If rRetry < rRetryMax And rpTestResult = FAIL Then
                        rRetry = rRetry + 1
                        GoTo REQUEST_DOWNLOAD_DATA
                    ElseIf rpTestResult = FAIL Then
                        GoTo EndWithFail
                    End If

                    ' *** Result Manage 
                    ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
                    ' ********************************************************************************************************************************************
                    ' ********************************************************************************************************************************************


                    ' ********************************************************************************************************************************************
                    ' *** DOWNLOAD SBL DATA
                    ' ********************************************************************************************************************************************
                    rpNtest = 2000                                                          ' *** Set Test Number
                    rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
                    rpTestName = "Download SBL data"                                         ' *** Set Test Name
                    rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
                    rpUnitMeasure = ""                                                      ' *** Set Unit Measure
                    rpTestResult = FAIL



                    ' *** Clear Can Buffer
                    fCanBufferClearFlush(15)
                    Dim FramesSent As Integer = 0

                    Dim SBLFileLine As Integer = 1
                    Do While DataIndex < BlockData.Length
                        If (DataIndex + 7 * 2) <= BlockData.Length Then
                            Data = BlockData.Substring(DataIndex, 7 * 2)
                            DataIndex = DataIndex + 7 * 2
                        Else
                            Dim MissingBytes As Integer = 7 - (BlockData.Length - DataIndex) \ 2
                            Data = BlockData.Substring(DataIndex, 7 * 2 - MissingBytes * 2) & New String("0"c, MissingBytes * 2)
                            DataIndex = DataIndex + (7 * 2 - MissingBytes * 2)
                        End If


                        Dim SBLFileLineHex As String = Hex(SBLFileLine).ToUpper()
                        rFrameToTx = "2" & SBLFileLineHex.Substring(SBLFileLineHex.Length - 1) & Data

                        If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail


                        FramesSent = FramesSent + 1

                        If FramesSent >= FlowControlNumber And FlowControlNumber <> 0 Then

                            FramesSent = 0

                            rExpectedFrameRx = {"30", "XX", "00", "00", "00", "00", "00", "00"}
                            '' *** Wait Response
                            'Thread.Sleep(60)
                            '' *** Check Buffer For Specific ID 
                            'If fCanRxFrames(rpCanPort, "071E", rFrameRx) = 1 Then
                            '    If rRetry < rRetryMax Then
                            '        rRetry = rRetry + 1
                            '        GoTo REQUEST_DOWNLOAD_DATA
                            '    Else
                            '        GoTo EndWithFail
                            '    End If
                            'End If
                            rFrameRx = {"30", "00", "00", "00", "00", "00", "00", "00"}

                            ' *** Check  Message
                            For rIndex As Integer = 0 To 7
                                If rIndex <> 1 Then
                                    If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                                        rpTestResult = 0
                                    Else
                                        rpTestResult = 1
                                        Exit For
                                    End If
                                End If
                            Next

                            FlowControlNumber = Convert.ToInt32(rFrameRx(1), 16)

                            ' *** Retry Manage
                            If rRetry < rRetryMax And rpTestResult = 1 Then
                                rRetry = rRetry + 1
                                GoTo REQUEST_DOWNLOAD_DATA
                            ElseIf rpTestResult = 1 Then
                                GoTo EndWithFail
                            End If
                        End If

                        SBLFileLine = SBLFileLine + 1
                    Loop

MULTI_FRAME_SKIP:

                    ' *** Message to be received 1
                    rExpectedFrameRx = {"02", "76", (BlocksSent + 1).ToString("X02"), "00", "00", "00", "00", "00"}

                    ' *** Wait Response
                    Thread.Sleep(60)
                    ' *** Check Buffer For Specific ID 
                    If fCanRxFrames(rpCanPort, "071E", rFrameRx) = FAIL Then
                        If rRetry < rRetryMax Then
                            rRetry = rRetry + 1

                            GoTo REQUEST_DOWNLOAD_DATA
                        Else

                            GoTo EndWithFail
                        End If
                    End If


                    ' *** Check  Reply
                    For rIndex As Integer = 0 To 7
                        If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                            rpTestResult = PASS

                        Else
                            rpTestResult = FAIL

                            Exit For
                        End If
                    Next

                    ' *** Retry Manage
                    If rRetry < rRetryMax And rpTestResult = FAIL Then
                        rRetry = rRetry + 1
                        GoTo REQUEST_DOWNLOAD_DATA
                    ElseIf rpTestResult = FAIL Then
                        GoTo EndWithFail
                    End If

                    BlocksSent = BlocksSent + 1
                    ' *** Result Manage 
                    ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
                    ' ********************************************************************************************************************************************


                End While



                ' ********************************************************************************************************************************************
                ' *** Request Trasfer Exit
                ' ********************************************************************************************************************************************
                rRetry = 0
                rRetryMax = 10

REQUEST_TRANSFER_EXIT:
                rpNtest = 2000                                                          ' *** Set Test Number
                rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
                rpTestName = "Request Trasfer Exit"                                         ' *** Set Test Name
                rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
                rpUnitMeasure = ""                                                      ' *** Set Unit Measure
                rpTestResult = FAIL


                ' *** Message to transmitt
                rFrameToTx = "0137000000000000"
                ' *** Message to be received
                rExpectedFrameRx0 = {"03", "7F", "37", "78", "00", "00", "00", "00"}
                ' *** Message to be received
                rExpectedFrameRx1 = {"03", "77", "XX", "XX", "00", "00", "00", "00"}
                ' *** Clear Can Buffer
                fCanBufferClearFlush(15)
                ' *** Send Command
                If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
                ' *** Wait Response
                Thread.Sleep(60)
                ' *** Check Buffer For Specific ID 
                If fCanRxFrames(rpCanPort, "071E", rFrameRx0) = FAIL Then
                    If rRetry < rRetryMax Then
                        rRetry = rRetry + 1
                        GoTo REQUEST_TRANSFER_EXIT
                    Else

                        GoTo EndWithFail
                    End If
                End If


                ' *** Wait Response
                Thread.Sleep(60)
                ' *** Check Buffer For Specific ID 
                If fCanRxFrames(rpCanPort, "071E", rFrameRx1) = FAIL Then
                    If rRetry < rRetryMax Then
                        rRetry = rRetry + 1
                        GoTo REQUEST_TRANSFER_EXIT
                    Else

                        GoTo EndWithFail
                    End If
                End If


                ' *** Check  Reply 0
                For rIndex As Integer = 0 To 7
                    If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                        rpTestResult = PASS

                    Else
                        rpTestResult = FAIL

                        Exit For
                    End If
                Next

                ' *** Check  Reply 1
                If rpTestResult = PASS Then
                    For rIndex As Integer = 0 To 1
                        If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                            rpTestResult = PASS

                        Else
                            rpTestResult = FAIL

                            Exit For
                        End If
                    Next
                End If

                ' *** Check CheckSum
                Dim ECUchecksum As String = rFrameRx1(2) & rFrameRx1(3)
                If rpTestResult = PASS Then
                    If ECUchecksum <> VBFDataBlock.Checksum Then
                        rpTestResult = FAIL

                    End If
                End If


                ' *** Retry Manage
                If rRetry < rRetryMax And rpTestResult = FAIL Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_TRANSFER_EXIT
                ElseIf rpTestResult = FAIL Then
                    GoTo EndWithFail
                End If

                ' *** Result Manage 
                ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
                ' ********************************************************************************************************************************************
                ' ********************************************************************************************************************************************

            Next



            ' *** COMPLETED CORRECTLY
            If rpTestResult = PASS Then GoTo EndWithPass
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fDUTSBLDownloadV2 = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fDUTSBLDownloadV2 = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fObp_FCT - Function execution result: " & fCRes(fDUTSBLDownload), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function

    Public Function fDUTSBLDownload(ByVal rSite As Integer) As Long

        fDUTSBLDownload = FAIL

        Dim rError As String = Nothing

        Dim rFrameToTx As String = ""
        Dim rFrameToTx0 As String = ""
        Dim rFrameToTx1 As String = ""
        Dim rFrameToTx2 As String = ""
        Dim rFrameToTx3 As String = ""

        Dim rExpectedFrameRx(7) As String
        Dim rExpectedFrameRx0(7) As String
        Dim rExpectedFrameRx1(7) As String
        Dim rExpectedFrameRx2(7) As String
        Dim rExpectedFrameRx3(7) As String
        Dim rExpectedFrameRx4(7) As String

        Dim rFrameRx(7) As String
        Dim rFrameRx0(7) As String
        Dim rFrameRx1(7) As String
        Dim rFrameRx2(7) As String
        Dim rFrameRx3(7) As String
        Dim rFrameRx4(7) As String

        Dim rRetry As Integer
        Dim rRetryMax As Integer
        Dim aKeyBytes(18) As Byte

        Try
            ' ********************************************************************************************************************************************
            ' *** LOAD FILE
            ' ********************************************************************************************************************************************
            Dim rSBLFile1 As StreamReader = File.OpenText("C:\Inifiles\SDLC\SBL1.txt")
            Dim rSBLFile2 As StreamReader = File.OpenText("C:\Inifiles\SDLC\SBL2.txt")


            ' ********************************************************************************************************************************************
            ' *** REQUEST TO DOWNLOAD DATA (Address 0x400AA000, Size: 0x0000028F)
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

REQUEST_DOWNLOAD_DATA_1:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Request to download data 1"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt
            rFrameToTx0 = "100B341044400AA0"
            ' *** Message to be received
            rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}
            ' *** Message to transmitt
            rFrameToTx1 = "21000000028F0000"
            ' *** Message to be received
            rExpectedFrameRx1 = {"04", "74", "20", "0F", "FF", "00", "00", "00"}


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_DOWNLOAD_DATA_1
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_DOWNLOAD_DATA_1
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Check  Reply 1
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_DOWNLOAD_DATA_1
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** Send First Data
            ' ********************************************************************************************************************************************
            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Send First Data"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            Dim Data As String
            Data = rSBLFile1.ReadLine()
            If Data.Length <> 8 Then
                rError = "First Line of data is not 4 bytes!"
                GoTo EndWithFail
            End If

            ' *** Message to transmitt
            rFrameToTx = "12913601" & Data
            ' *** Message to be received
            rExpectedFrameRx = {"30", "00", "00", "00", "00", "00", "00", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_DOWNLOAD_DATA_1
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Check  Message
            For rIndex As Integer = 0 To 7
                If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_DOWNLOAD_DATA_1
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** DOWNLOAD SBL FILE 1
            ' ********************************************************************************************************************************************
            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Download SBL file 1"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)

            Dim SBLFileLine As Integer = 1
            Do While rSBLFile1.Peek() > -1
                Data = rSBLFile1.ReadLine()
                If Data.Length <> 14 Then
                    rError = "Line " & (SBLFileLine + 1) & "of data is not 7 bytes!"
                    GoTo EndWithFail
                End If

                Dim SBLFileLineHex As String = Hex(SBLFileLine).ToUpper()
                rFrameToTx = "2" & SBLFileLineHex.Substring(SBLFileLineHex.Length - 1) & Data

                If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail

                SBLFileLine = SBLFileLine + 1
            Loop

            ' *** Message to be received 1
            rExpectedFrameRx = {"02", "76", "01", "00", "00", "00", "00", "00"}

            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_DOWNLOAD_DATA_1
                Else
                    GoTo EndWithFail
                End If
            End If


            ' *** Check  Reply
            For rIndex As Integer = 0 To 7
                If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_DOWNLOAD_DATA_1
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' *
            ' ********************************************************************************************************************************************
            ' *** Request Trasfer Exit
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

REQUEST_TRANSFER_EXIT_1:
            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Request Trasfer Exit"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL

            ' *** Message to transmitt
            rFrameToTx = "0137000000000000"
            ' *** Message to be received
            rExpectedFrameRx0 = {"03", "7F", "37", "78", "00", "00", "00", "00"}
            ' *** Message to be received
            rExpectedFrameRx1 = {"03", "77", "8C", "AC", "00", "00", "00", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_TRANSFER_EXIT_1
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_TRANSFER_EXIT_1
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Check  Reply 1
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_TRANSFER_EXIT_1
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************



            ' ********************************************************************************************************************************************
            ' *** REQUEST TO DOWNLOAD DATA (Address 0x400AADF8, Size: 0x00000151)
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

REQUEST_DOWNLOAD_DATA_2:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Request to download data 2"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt
            rFrameToTx0 = "100B341044400AAD"
            ' *** Message to be received
            rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}
            ' *** Message to transmitt
            rFrameToTx1 = "21F8000001510000"
            ' *** Message to be received
            rExpectedFrameRx1 = {"04", "74", "20", "0F", "FF", "00", "00", "00"}


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_DOWNLOAD_DATA_2
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_DOWNLOAD_DATA_2
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Check  Reply 1
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_DOWNLOAD_DATA_2
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** Send First Data
            ' ********************************************************************************************************************************************
            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Send First Data"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            Data = rSBLFile2.ReadLine()
            If Data.Length <> 8 Then
                rError = "First Line of data is not 4 bytes!"
                GoTo EndWithFail
            End If

            ' *** Message to transmitt
            rFrameToTx = "11533601" & Data
            ' *** Message to be received
            rExpectedFrameRx = {"30", "00", "00", "00", "00", "00", "00", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_DOWNLOAD_DATA_2
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Check  Message
            For rIndex As Integer = 0 To 7
                If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_DOWNLOAD_DATA_2
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** DOWNLOAD SBL FILE 2
            ' ********************************************************************************************************************************************
            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Download SBL file 2"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)

            SBLFileLine = 1
            Do While rSBLFile2.Peek() > -1
                Data = rSBLFile2.ReadLine()
                If Data.Length <> 14 Then
                    rError = "Line " & (SBLFileLine + 1) & "of data is not 7 bytes!"
                    GoTo EndWithFail
                End If

                Dim SBLFileLineHex As String = Hex(SBLFileLine).ToUpper()
                rFrameToTx = "2" & SBLFileLineHex.Substring(SBLFileLineHex.Length - 1) & Data

                If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail

                SBLFileLine = SBLFileLine + 1
            Loop

            ' *** Message to be received 1
            rExpectedFrameRx = {"02", "76", "01", "00", "00", "00", "00", "00"}

            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_DOWNLOAD_DATA_2
                Else
                    GoTo EndWithFail
                End If
            End If


            ' *** Check  Reply
            For rIndex As Integer = 0 To 7
                If rFrameRx(rIndex) = rExpectedFrameRx(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_DOWNLOAD_DATA_2
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' *
            ' ********************************************************************************************************************************************
            ' *** Request Trasfer Exit
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

REQUEST_TRANSFER_EXIT_2:
            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Request Trasfer Exit"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL

            ' *** Message to transmitt
            rFrameToTx = "0137000000000000"
            ' *** Message to be received
            rExpectedFrameRx0 = {"03", "7F", "37", "78", "00", "00", "00", "00"}
            ' *** Message to be received
            rExpectedFrameRx1 = {"03", "77", "3E", "9E", "00", "00", "00", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_TRANSFER_EXIT_2
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "071E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_TRANSFER_EXIT_2
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Check  Reply 1
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_TRANSFER_EXIT_2
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************


            ' *** COMPLETED CORRECTLY
            If rpTestResult = PASS Then GoTo EndWithPass
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fDUTSBLDownload = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fDUTSBLDownload = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fObp_FCT - Function execution result: " & fCRes(fDUTSBLDownload), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function

    Public Function fDUTSBLStartup(ByVal rSite As Integer) As Long

        fDUTSBLStartup = FAIL

        Dim rError As String = Nothing

        Dim rFrameToTx As String = ""
        Dim rFrameToTx0 As String = ""
        Dim rFrameToTx1 As String = ""
        Dim rFrameToTx2 As String = ""
        Dim rFrameToTx3 As String = ""

        Dim rExpectedFrameRx(7) As String
        Dim rExpectedFrameRx0(7) As String
        Dim rExpectedFrameRx1(7) As String
        Dim rExpectedFrameRx2(7) As String
        Dim rExpectedFrameRx3(7) As String
        Dim rExpectedFrameRx4(7) As String

        Dim rFrameRx(7) As String
        Dim rFrameRx0(7) As String
        Dim rFrameRx1(7) As String
        Dim rFrameRx2(7) As String
        Dim rFrameRx3(7) As String
        Dim rFrameRx4(7) As String

        Dim rRetry As Integer
        Dim rRetryMax As Integer

        Try
            ' ********************************************************************************************************************************************
            ' *** REQUEST TO START ROUTINE 0x0301
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

REQUEST_START_ROUTINE_0301:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Request to start routine 0x0301"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt
            rFrameToTx0 = "100831010301400A"
            ' *** Message to be received
            rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}
            ' *** Message to transmitt
            rFrameToTx1 = "21A0000000000000"
            ' *** Message to be received
            rExpectedFrameRx1 = {"03", "7F", "31", "78", "00", "00", "00", "00"}
            ' *** Message to be received
            rExpectedFrameRx2 = {"05", "71", "01", "03", "01", "10", "00", "00"}


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_START_ROUTINE_0301
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_START_ROUTINE_0301
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Look for a responce every 500 ms. Max 3 seconds
            For i As Integer = 1 To 6
                Thread.Sleep(500)
                If fCanRxFrames(rpCanPort, "71E", rFrameRx2) = PASS Then Exit For
                If i = 6 Then GoTo EndWithFail
            Next

            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Check  Reply 1
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If

            ' *** Check  Reply 2
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rFrameRx2(rIndex) = rExpectedFrameRx2(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_START_ROUTINE_0301
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** REQUEST TO START ROUTINE 0x0304
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

REQUEST_START_ROUTINE_0304:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Request to start routine 0x0304"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt
            rFrameToTx0 = "0431010304000000"
            ' *** Message to be received
            rExpectedFrameRx0 = {"06", "71", "01", "03", "04", "10", "02", "00"}


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_START_ROUTINE_0304
                Else
                    GoTo EndWithFail
                End If
            End If


            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_START_ROUTINE_0304
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************


            ' *** COMPLETED CORRECTLY
            If rpTestResult = PASS Then GoTo EndWithPass
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fDUTSBLStartup = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fDUTSBLStartup = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fObp_FCT - Function execution result: " & fCRes(fDUTSBLStartup), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function

    Function RemoveWhitespace(fullString As String) As String
        Return New String(fullString.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray())
    End Function

    Public Function fDUTWriteByID(ByVal rSite As Integer) As Long

        fDUTWriteByID = FAIL

        Dim rError As String = Nothing

        Dim rFrameToTx As String = ""
        Dim rFrameToTx0 As String = ""
        Dim rFrameToTx1 As String = ""
        Dim rFrameToTx2 As String = ""
        Dim rFrameToTx3 As String = ""

        Dim rExpectedFrameRx(7) As String
        Dim rExpectedFrameRx0(7) As String
        Dim rExpectedFrameRx1(7) As String
        Dim rExpectedFrameRx2(7) As String
        Dim rExpectedFrameRx3(7) As String
        Dim rExpectedFrameRx4(7) As String

        Dim rFrameRx(7) As String
        Dim rFrameRx0(7) As String
        Dim rFrameRx1(7) As String
        Dim rFrameRx2(7) As String
        Dim rFrameRx3(7) As String
        Dim rFrameRx4(7) As String

        Dim rRetry As Integer
        Dim rRetryMax As Integer

        Try
            ' ********************************************************************************************************************************************
            ' *** Load File and check data
            ' ********************************************************************************************************************************************
            rpRpkTpgmVariant = "5CTH1AF02"

            Dim rDataFile As StreamReader = File.OpenText("C:\Inifiles\SDLC\" & rpRpkTpgmVariant & ".txt")

            Dim rF111_Data As String = Nothing
            Dim rF113_Data As String = Nothing
            Dim rvar_Data As String = Nothing

            Do While rDataFile.Peek() > -1
                Dim Data As String = rDataFile.ReadLine()
                If Data.Contains("F111:") And Data.Length > 5 Then
                    Data = Data.Substring(5)
                    Data = RemoveWhitespace(Data)
                    If Data.Length = 15 Then
                        rF111_Data = Data
                    Else
                        rError = "F111 has the wrong length"
                        GoTo EndWithFail
                    End If
                ElseIf Data.Contains("F113:") And Data.Length > 5 Then
                    Data = Data.Substring(5)
                    Data = RemoveWhitespace(Data)
                    If Data.Length > 10 Then
                        rF113_Data = Data
                    Else
                        rError = "F113 has the wrong length"
                        GoTo EndWithFail
                    End If
                ElseIf Data.Contains("DE14:") And Data.Length > 5 Then
                    Data = Data.Substring(5)
                    Data = RemoveWhitespace(Data)
                    If Data.Length = 1 Then
                        rvar_Data = Data
                    Else
                        rError = "DE14 has the wrong length"
                        GoTo EndWithFail
                    End If
                End If
            Loop

            If rF111_Data Is Nothing Or rF113_Data Is Nothing Or rvar_Data Is Nothing Then
                rError = "F111 or F113 are missing!"
                GoTo EndWithFail
            End If

            'GoTo WRITE_0xF113

            ' ********************************************************************************************************************************************
            ' *** Write 0xF111
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

WRITE_0xF111:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Write 0xF111"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL
            Dim F111 As String = Nothing
            'Dim F113 As String = Nothing
            Dim FD14 As String = Nothing
            Dim i As Integer

            For i = 1 To rF111_Data.Length
                F111 = F111 & Hex(CStr(Asc(Mid$(rF111_Data, i, 1))))
            Next i


            For i = 1 To rvar_Data.Length
                FD14 = FD14 & Hex(CStr(Asc(Mid$(rvar_Data, i, 1))))
            Next i



            ' *** Message to transmitt
            rFrameToTx0 = "101B2EF111" & F111.Substring(0, 6)
            ' *** Message to be received
            rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}
            ' *** Message to transmitt
            rFrameToTx1 = "21" & F111.Substring(6, 14)
            ' *** Message to transmitt
            rFrameToTx2 = "22" & F111.Substring(20, 10) & "0000"
            ' *** Message to transmitt
            rFrameToTx3 = "2300000000000000"
            ' *** Message to be received
            rExpectedFrameRx1 = {"03", "6E", "F1", "11", "00", "00", "00", "00"}



            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(100)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo WRITE_0xF111
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx2) = FAIL Then GoTo EndWithFail
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx3) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(100)
            ' *** Check Buffer For Specific ID 
            'If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
            '    If rRetry < rRetryMax Then
            '        rRetry = rRetry + 1
            '        GoTo WRITE_0xF111
            '    Else
            '        GoTo EndWithFail
            '    End If
            'End If


            ' *** Check  Reply 0
            'For rIndex As Integer = 0 To 7
            '    If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
            '        rpTestResult = PASS
            '    Else
            '        rpTestResult = FAIL
            '        Exit For
            '    End If
            'Next

            ' *** Check  Reply 1
            'If rpTestResult = PASS Then
            '    For rIndex As Integer = 0 To 7
            '        If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
            '            rpTestResult = PASS
            '        Else
            '            rpTestResult = FAIL
            '            Exit For
            '        End If
            '    Next
            'End If

            ' *** Check  Reply 2
            'If rpTestResult = PASS Then
            '    For rIndex As Integer = 0 To 7
            '        If rFrameRx2(rIndex) = rExpectedFrameRx2(rIndex) Then
            '            rpTestResult = PASS
            '        Else
            '            rpTestResult = FAIL
            '            Exit For
            '        End If
            '    Next
            'End If
            ' *** Retry Manage
            'If rRetry < rRetryMax And rpTestResult = FAIL Then
            '    rRetry = rRetry + 1
            '    GoTo WRITE_0xF111
            'ElseIf rpTestResult = FAIL Then
            '    GoTo EndWithFail
            'End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************

            ' ********************************************************************************************************************************************
            ' *** Write 0xF113
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

WRITE_0xF113:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Write 0xF113"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL

            Dim F113 As String = Nothing

            For i = 1 To rF113_Data.Length
                F113 = F113 & Hex(CStr(Asc(Mid$(rF113_Data, i, 1))))
            Next i

            rFrameToTx0 = "101B2EF113" & F113.Substring(0, 6)
            ' *** Message to be received
            rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}
            ' *** Message to transmitt
            rFrameToTx1 = "21" & F113.Substring(6, 14)
            ' *** Message to transmitt
            rFrameToTx2 = "22" & F113.Substring(20, 8) & "000000"
            ' *** Message to transmitt
            rFrameToTx3 = "2300000000000000"
            ' *** Message to be received
            rExpectedFrameRx1 = {"03", "6E", "F1", "13", "00", "00", "00", "00"}



            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail

            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo WRITE_0xF113
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)

            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx2) = FAIL Then GoTo EndWithFail
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx3) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            rpTestResult = PASS
            ' *** Check Buffer For Specific ID 
            'If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
            '    If rRetry < rRetryMax Then
            '        rRetry = rRetry + 1
            '        GoTo WRITE_0xF113
            '    Else
            '        GoTo EndWithFail
            '    End If
            'End If

            ' *** Check  Reply 0
            'For rIndex As Integer = 0 To 7
            '    If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
            '        rpTestResult = PASS
            '    Else
            '        rpTestResult = FAIL
            '        Exit For
            '    End If
            'Next

            '' *** Check  Reply 1
            'If rpTestResult = PASS Then
            '    For rIndex As Integer = 0 To 7
            '        If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
            '            rpTestResult = PASS
            '        Else
            '            rpTestResult = FAIL
            '            Exit For
            '        End If
            '    Next
            'End If

            ' *** Check  Reply 2
            'If rpTestResult = PASS Then
            '    For rIndex As Integer = 0 To 7
            '        If rFrameRx2(rIndex) = rExpectedFrameRx2(rIndex) Then
            '            rpTestResult = PASS
            '        Else
            '            rpTestResult = FAIL
            '            Exit For
            '        End If
            '    Next
            'End If
            '' *** Retry Manage
            'If rRetry < rRetryMax And rpTestResult = FAIL Then
            '    rRetry = rRetry + 1
            '    GoTo WRITE_0xF113
            'ElseIf rpTestResult = FAIL Then
            '    GoTo EndWithFail
            'End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************
            '            ', ********************************************************************************************************************************************
            '            ' *** Write 0XDE14
            '            ' ********************************************************************************************************************************************
            '            rRetry = 0
            '            rRetryMax = 10

            'WRITE_0xDE14:

            '            rpNtest = 2000                                                          ' *** Set Test Number
            '            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            '            rpTestName = "Write 0xDE14"                                         ' *** Set Test Name
            '            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            '            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            '            rpTestResult = FAIL


            '            ' *** Message to transmitt
            '            If rvar_Data = "D" Then
            '                rFrameToTx0 = "042EDE1467000000"
            '            ElseIf rvar_Data = "E" Then
            '                rFrameToTx0 = "042EDE146F000000"
            '            ElseIf rvar_Data = "F" Then
            '                rFrameToTx0 = "042EDE146B000000"
            '            Else
            '                rpTestResult = PASS
            '                GoTo Var_OTHER
            '            End If

            '            ' *** Message to be received
            '            rExpectedFrameRx1 = {"03", "6E", "DE", "14", "00", "00", "00", "00"}

            '            ' *** Clear Can Buffer
            '            fCanBufferClearFlush(15)
            '            ' *** Send Command
            '            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail

            '            ' *** Wait Response
            '            Thread.Sleep(60)
            '            ' *** Check Buffer For Specific ID 
            '            If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
            '                If rRetry < rRetryMax Then
            '                    rRetry = rRetry + 1
            '                    GoTo WRITE_0xDE14
            '                Else
            '                    GoTo EndWithFail
            '                End If
            '            End If

            '            ' *** Check  Reply 1
            '            If rpTestResult = PASS Then
            '                For rIndex As Integer = 0 To 7
            '                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
            '                        rpTestResult = PASS
            '                    Else
            '                        rpTestResult = FAIL
            '                        Exit For
            '                    End If
            '                Next
            '            End If

            '            ' *** Retry Manage
            '            If rRetry < rRetryMax And rpTestResult = FAIL Then
            '                rRetry = rRetry + 1
            '                GoTo WRITE_0xDE14
            '            ElseIf rpTestResult = FAIL Then
            '                GoTo EndWithFail
            '            End If
            '            ' *** Result Manage 
            '            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            '            ' ********************************************************************************************************************************************
            '            ' 
            'Var_OTHER:



            ' *** COMPLETED CORRECTLY
            If rpTestResult = PASS Then GoTo EndWithPass
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fDUTWriteByID = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fDUTWriteByID = PASS
Exit_Function:

        ' *** Print Function Result		PASS	0	Long

        MsgPrintLogIfEn("fObp_FCT - Function execution result: " & fCRes(fDUTWriteByID), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function

    Public Function fDUTWriteJulianDate(ByVal rSite As Integer) As Long

        fDUTWriteJulianDate = FAIL

        Dim rError As String = Nothing

        Dim rFrameToTx As String = ""
        Dim rFrameToTx0 As String = ""
        Dim rFrameToTx1 As String = ""
        Dim rFrameToTx2 As String = ""
        Dim rFrameToTx3 As String = ""

        Dim rExpectedFrameRx(7) As String
        Dim rExpectedFrameRx0(7) As String
        Dim rExpectedFrameRx1(7) As String
        Dim rExpectedFrameRx2(7) As String
        Dim rExpectedFrameRx3(7) As String
        Dim rExpectedFrameRx4(7) As String

        Dim rFrameRx(7) As String
        Dim rFrameRx0(7) As String
        Dim rFrameRx1(7) As String
        Dim rFrameRx2(7) As String
        Dim rFrameRx3(7) As String
        Dim rFrameRx4(7) As String

        Dim rRetry As Integer
        Dim rRetryMax As Integer

        Try
            ' ********************************************************************************************************************************************
            ' *** Compute Julian Date
            ' ********************************************************************************************************************************************
            Dim rCurrentTime As DateTime = Date.Now()
            Dim JulianDate As String = String.Format("{0}{1}", rCurrentTime.ToString("yy"), rCurrentTime.DayOfYear.ToString("000")).Substring(1)

            Dim Bytes() As Byte = Text.Encoding.ASCII.GetBytes(JulianDate)

            Dim JulianDateHex As String = ""
            Dim JulianDateBuilder As System.Text.StringBuilder = New System.Text.StringBuilder(Bytes.Length * 2)
            For Each B As Byte In Bytes
                JulianDateBuilder.AppendFormat("{0:x2}", B).ToString.Trim()
            Next
            JulianDateHex = JulianDateBuilder.ToString()




            ' ********************************************************************************************************************************************
            ' *** Write Julian Date
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

WRITE_JULIAN_DATE:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Write Julian Date"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt
            'rFrameToTx0 = "072EFD0B" & JulianDateHex
            rFrameToTx0 = "072EFD0B" & "32313533"
            ' *** Message to be received
            rExpectedFrameRx0 = {"03", "6E", "FD", "0B", "00", "00", "00", "00"}


            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo WRITE_JULIAN_DATE
                Else
                    GoTo EndWithFail
                End If
            End If

            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo WRITE_JULIAN_DATE
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************


            ' *** COMPLETED CORRECTLY
            If rpTestResult = PASS Then GoTo EndWithPass
            ' *************************************************************************************************************    

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fDUTWriteJulianDate = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fDUTWriteJulianDate = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fObp_FCT - Function execution result: " & fCRes(fDUTWriteJulianDate), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function

    Public Function fSerialNumberWrite(ByVal SN As String) As Integer

        fSerialNumberWrite = FAIL

        Dim rError As String = Nothing

        Dim rFrameToTx0 As String = ""
        Dim SESTERPRESENT As String = "023E800000000000"
        Dim rFrameToTx1 As String = ""
        Dim rFrameToTx2 As String = ""
        Dim rFrameToTx3 As String = ""

        Dim rExpectedFrameRx0(7) As String
        Dim rExpectedFrameRx1(7) As String
        Dim rExpectedFrameRx2(7) As String
        Dim rExpectedFrameRx3(7) As String

        Dim rFrameRx0(7) As String
        Dim rFrameRx1(7) As String
        Dim rFrameRx2(7) As String
        Dim rFrameRx3(7) As String

        Dim rRetry As Integer
        Dim rRetryMax As Integer

        Try

            rpTestResult = FAIL

            ' *** Prepare data for SN write
            Dim SerialNumber As String = SN ' & New String("0", 16 - rpUutFctSerialNumber(rSite).Length)
            Dim J As Integer
            Dim Bytes() As Byte = Text.Encoding.ASCII.GetBytes(SerialNumber)
            Dim SerialNumberHex As String = ""
            Dim SerialNumberBuilder As System.Text.StringBuilder = New System.Text.StringBuilder(Bytes.Length * 2)
            For Each B As Byte In Bytes
                SerialNumberBuilder.AppendFormat("{0:x2}", B).ToString.Trim()
            Next
            For i As Integer = Bytes.Length To 16
                SerialNumberBuilder.AppendFormat("{0:x2}", 0).ToString.Trim()
            Next
            SerialNumberHex = SerialNumberBuilder.ToString()

            rFrameToTx0 = "10132EF18C" & SerialNumberHex.Substring(0, 3 * 2)

            rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}

            rFrameToTx1 = "21" & SerialNumberHex.Substring(3 * 2, 7 * 2)

            rFrameToTx2 = "22" & SerialNumberHex.Substring(10 * 2, 7 * 2)

            For J = 0 To 2
                If fCanTxFrames(rpCanPort, "7DF", SESTERPRESENT) = FAIL Then GoTo EndWithFail
                sWait(100)
            Next
            'rFrameToTx0 = "10132EF18C" & CStr(Hex(Asc(Mid(SerialNumber, 1, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 2, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 3, 1))))

            'rExpectedFrameRx0 = {"30", "00", "00", "00", "00", "00", "00", "00"}

            'rFrameToTx1 = "21" & CStr(Hex(Asc(Mid(SerialNumber, 4, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 5, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 6, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 7, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 8, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 9, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 10, 1))))


            'rFrameToTx2 = "22" & CStr(Hex(Asc(Mid(SerialNumber, 11, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 12, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 13, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 14, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 15, 1)))) &
            '                                CStr(Hex(Asc(Mid(SerialNumber, 16, 1)))) &
            '                                "00"


            rExpectedFrameRx1 = {"03", "6E", "F1", "8C", "00", "00", "00", "00"}

            rRetry = 0
            rRetryMax = 10

WRITE_RETRY:

            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo WRITE_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx2) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo WRITE_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If


            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next

            ' *** Check  Reply 1
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                Next
            End If

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo WRITE_RETRY
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")

            ' *** COMPLETED CORRECTLY
            If rpTestResult = PASS Then GoTo EndWithPass
            ' ************************************************************************************************************* 

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:
        ' *** Store Result Fail
        fSerialNumberWrite = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fSerialNumberWrite = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fObp_FCT - Function execution result: " & fCRes(fSerialNumberWrite), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function

    Public Function fSerialNumberRead(ByVal rSite As Integer, ByRef rSerialNumberRead As String) As Integer

        fSerialNumberRead = FAIL

        rSerialNumberRead = ""

        Dim rError As String = Nothing

        Dim rFrameToTx0 As String = ""
        Dim rFrameToTx1 As String = ""
        Dim rFrameToTx2 As String = ""
        Dim rFrameToTx3 As String = ""

        Dim rExpectedFrameRx0(7) As String
        Dim rExpectedFrameRx1(7) As String
        Dim rExpectedFrameRx2(7) As String
        Dim rExpectedFrameRx3(7) As String

        Dim rFrameRx0(7) As String
        Dim rFrameRx1(7) As String
        Dim rFrameRx2(7) As String
        Dim rFrameRx3(7) As String

        Dim rRetry As Integer
        Dim rRetryMax As Integer

        Try

            rpTestResult = FAIL

            ' ********************************************************************************************************************************************
            ' *** REQUEST START DEFAULT SESSION (0x01)
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

REQUEST_START_DEFAULT_SESSION:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Default Session"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt
            rFrameToTx0 = "0210010000000000"
            ' *** Message to be received
            rExpectedFrameRx0 = {"06", "50", "01", "00", "32", "01", "F4", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_START_DEFAULT_SESSION
                Else
                    GoTo EndWithFail
                End If
            End If


            ' *** Check  Message
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_START_DEFAULT_SESSION
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************


            ' ********************************************************************************************************************************************
            ' *** REQUEST START EXTENDED SESSION (0x03)
            ' ********************************************************************************************************************************************
            rRetry = 0
            rRetryMax = 10

REQUEST_START_EXTENDED_SESSION:

            rpNtest = 2000                                                          ' *** Set Test Number
            rpDrawRef = "FCT"                                                       ' *** Set Drawing Reference
            rpTestName = "Default Session"                                         ' *** Set Test Name
            rpDiagRemark = rpTestName & "- FAIL"                                    ' *** Set Diag Remark
            rpUnitMeasure = ""                                                      ' *** Set Unit Measure
            rpTestResult = FAIL


            ' *** Message to transmitt
            rFrameToTx0 = "0210030000000000"
            ' *** Message to be received
            rExpectedFrameRx0 = {"06", "50", "03", "00", "32", "01", "F4", "00"}
            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo REQUEST_START_EXTENDED_SESSION
                Else
                    GoTo EndWithFail
                End If
            End If


            ' *** Check  Message
            For rIndex As Integer = 0 To 7
                If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                    rpTestResult = PASS
                Else
                    rpTestResult = FAIL
                    Exit For
                End If
            Next
            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo REQUEST_START_EXTENDED_SESSION
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If
            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")
            ' ********************************************************************************************************************************************
            ' ********************************************************************************************************************************************


            ' *** Prepare data for SN read
            'Dim SerialNumber As String = rpUutFctSerialNumber(rSite) & New String(vbNullChar, 16 - rpUutFctSerialNumber(rSite).Length)



            rFrameToTx0 = "0322F18C00000000"

            rExpectedFrameRx0 = {"10", "13", "62", "F1", "8C", "XX", "XX", "XX"}

            rFrameToTx1 = "3000000000000000"

            rExpectedFrameRx1 = {"21", "XX", "XX", "XX", "XX", "XX", "XX", "XX"}

            rExpectedFrameRx2 = {"22", "XX", "XX", "XX", "XX", "XX", "XX", "XX"}

            rRetry = 0
            rRetryMax = 10

READ_RETRY:

            ' *** Clear Can Buffer
            fCanBufferClearFlush(15)
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx0) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx0) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo READ_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Send Command
            If fCanTxFrames(rpCanPort, "716", rFrameToTx1) = FAIL Then GoTo EndWithFail
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx1) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo READ_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If
            ' *** Wait Response
            Thread.Sleep(60)
            ' *** Check Buffer For Specific ID 
            If fCanRxFrames(rpCanPort, "71E", rFrameRx2) = FAIL Then
                If rRetry < rRetryMax Then
                    rRetry = rRetry + 1
                    GoTo READ_RETRY
                Else
                    GoTo EndWithFail
                End If
            End If


            ' *** Check  Reply 0
            For rIndex As Integer = 0 To 7
                If rIndex < 5 Then
                    If rFrameRx0(rIndex) = rExpectedFrameRx0(rIndex) Then
                        rpTestResult = PASS
                    Else
                        rpTestResult = FAIL
                        Exit For
                    End If
                End If
            Next

            ' *** Check  Reply 1
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rIndex = 0 Then
                        If rFrameRx1(rIndex) = rExpectedFrameRx1(rIndex) Then
                            rpTestResult = PASS
                        Else
                            rpTestResult = FAIL
                            Exit For
                        End If
                    End If
                Next
            End If

            ' *** Check  Reply 2
            If rpTestResult = PASS Then
                For rIndex As Integer = 0 To 7
                    If rIndex = 0 Then
                        If rFrameRx2(rIndex) = rExpectedFrameRx2(rIndex) Then
                            rpTestResult = PASS
                        Else
                            rpTestResult = FAIL
                            Exit For
                        End If
                    End If
                Next
            End If

            ' *** Retry Manage
            If rRetry < rRetryMax And rpTestResult = FAIL Then
                rRetry = rRetry + 1
                GoTo READ_RETRY
            ElseIf rpTestResult = FAIL Then
                GoTo EndWithFail
            End If

            Dim rSerialNumberReadHex As String = ""
            For rIndex As Integer = 5 To 7
                rSerialNumberReadHex = rSerialNumberReadHex & rFrameRx0(rIndex)
            Next
            For rIndex As Integer = 1 To 7
                rSerialNumberReadHex = rSerialNumberReadHex & rFrameRx1(rIndex)
            Next
            'For rIndex As Integer = 1 To 6
            '    rSerialNumberReadHex = rSerialNumberReadHex & rFrameRx2(rIndex)
            'Next

            For i As Integer = 0 To rSerialNumberReadHex.Length - 1 Step 2
                Dim k As String = rSerialNumberReadHex.Substring(i, 2)
                rSerialNumberRead &= System.Convert.ToChar(System.Convert.ToUInt32(k, 16)).ToString()
            Next

            'If rSerialNumberRead <> SerialNumber Then
            '    rError = "Wrong Serial Number!!!"
            '    GoTo EndWithFail
            'End If

            ' *** Result Manage 
            ' MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpNtest, rpDrawRef, rpTestName, rpTestResult, rpDiagRemark, -1000000, -1000000, -1000000, rpUnitMeasure, "ALWAYS")

            ' *** COMPLETED CORRECTLY
            If rpTestResult = PASS Then GoTo EndWithPass
            ' *************************************************************************************************************   

        Catch ex As Exception
            rError = ex.Message
            GoTo EndWithFail
        End Try

EndWithFail:

        rSerialNumberRead = "Read Fail"

        ' *** Store Result Fail
        fSerialNumberRead = FAIL
        GoTo Exit_Function
EndWithPass:
        ' *** Store Result Pass
        fSerialNumberRead = PASS
Exit_Function:

        ' *** Print Function Result
        'MsgPrintLogIfEn("fObp_FCT - Function execution result: " & fCRes(fSerialNumberRead), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Print error if occurred
        'If rError IsNot Nothing Then MsgPrintLog("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, 0)
        If rError IsNot Nothing Then rpMsgLogClass.MsgPrintLogMultyBay("STAGE " & rpRpkBayInUse & "   -   Error: " & rError, rpRpkBayInUse)

    End Function


    ' *** DEBUG ONLY --- Try CAN between CH1 and CH2
    Public Sub sCANTryCH()

        Dim FrameRx(7) As String

        rpCanPort1 = 1
        rpCanSpeed1 = 250000
        ' ********************************************************************************************************************************************
        ' *** TX CAN Port Open
        ' ********************************************************************************************************************************************
        fCanPortOpen(rpCanPortHandle1, rpCanPort1, rpCanSpeed1)

        rpCanPort2 = 2
        rpCanSpeed2 = 250000
        ' ********************************************************************************************************************************************
        ' *** RX CAN Port Open
        ' ********************************************************************************************************************************************
        fCanPortOpen(rpCanPortHandle2, rpCanPort2, rpCanSpeed2)

        fCanBufferClearFlush()

        For rindex As Integer = 0 To 9

            MsgPrintLog("TX > " & "010203040506070" & CStr(rindex), 0)
            fCanTxFrames(rpCanPort1, "100", "010203040506070" & CStr(rindex))
            Thread.Sleep(100)

            If fCanRxFrames(rpCanPort2, "100", FrameRx) = PASS Then
                MsgPrintLog("RX < " & FrameRx(0) &
                                    FrameRx(1) &
                                    FrameRx(2) &
                                    FrameRx(3) &
                                    FrameRx(4) &
                                    FrameRx(5) &
                                    FrameRx(6) &
                                    FrameRx(7), 0)
            Else
                MsgPrintLog("@BG{Red}No message received", 0)
            End If

        Next

        ' ********************************************************************************************************************************************
        ' *** TX CAN Port Close
        ' ********************************************************************************************************************************************
        fCanPortClose(rpCanPortHandle1, rpCanPort1)

        ' ********************************************************************************************************************************************
        ' *** RX CAN Port Close
        ' ********************************************************************************************************************************************
        fCanPortClose(rpCanPortHandle2, rpCanPort2)

    End Sub



End Module

