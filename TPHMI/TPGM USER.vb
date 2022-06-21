Imports System.Collections.Generic
Imports System.Configuration
Imports System.IO
Imports System.Threading

Public Class TPGM_USER
    Dim _task As Thread
    Dim running As Boolean = False

    Public Property listActions() As New List(Of actions)

    Private Sub BunifuCustomDataGrid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles ActionDataList.CellContentClick

    End Sub

    Private Sub TPGM_USER_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadFilesNames()
        ReflectionTimer.Start()
    End Sub

    Public Sub loadFilesNames()
        lblCurrentVar.Text += ConfigurationManager.AppSettings("var").ToString()
        'For Each foundFile As String In My.Computer.FileSystem.GetFiles(Application.StartupPath + "\var")
        '    Dim foo As String
        '    foo = New DirectoryInfo(foundFile).Name
        '    cmbFiles.AddItem(foo)
        'Next

    End Sub

    Public Sub initAll()

        If running = True Then
            If listActions.Count > 0 Then
                txtCurrentProc.Text = listActions(listActions.Count - 1).Procname1.ToString()
                txtDescription.Text = listActions(listActions.Count - 1).ProcDescription1.ToString()
            End If
            btnActions.Text = "Running..."
            btnActions.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(208, Byte), Integer), CType(CType(194, Byte), Integer))
            'If listActions(listActions.Count - 1).ProcState1 Then
            '    'imgstate.Image = My.Resources.circle1
            'Else
            '    'imgstate.Image = My.Resources.circle__1_
            'End If
        Else
            btnActions.Text = "Init..."
            btnActions.BackColor = System.Drawing.Color.FromArgb(CType(CType(66, Byte), Integer), CType(CType(101, Byte), Integer), CType(CType(235, Byte), Integer))
            If txtCurrentProc.Text <> String.Empty Then
                txtCurrentProc.Text = ""
                txtDescription.Text = ""
                'imgstate.Image = My.Resources.circle_emp
                lbltxtflag.Text = ""

            Else
                txtCurrentProc.Text = "Waiting ..."
                txtDescription.Text = "Waiting ..."
                lbltxtflag.Text = "Waiting ..."
                'imgstate.Visible = True
                'imgstate.Image = My.Resources.circle
                panelflag.BackColor = Color.LemonChiffon
            End If
        End If
    End Sub
    Public Sub invokeComponent(ByVal com As Control)

    End Sub
    Public Sub fillDataGridView(ByVal item As actions)

        If item.ProcState1 = 0 Then
            ActionDataList.Invoke(Sub()
                                      ActionDataList.Rows.Insert(ActionDataList.Rows.Count, item.Procname1, item.ProcDescription1, My.Resources.ongoing, item.ProcCreatedAt1)
                                  End Sub)
        ElseIf item.ProcState1 = 1 Then
            ActionDataList.Invoke(Sub()
                                      ActionDataList.Rows.Insert(ActionDataList.Rows.Count, item.Procname1, item.ProcDescription1, My.Resources.good, item.ProcCreatedAt1)
                                  End Sub)
        Else
            ActionDataList.Invoke(Sub()
                                      ActionDataList.Rows.Insert(ActionDataList.Rows.Count, item.Procname1, item.ProcDescription1, My.Resources.ng, item.ProcCreatedAt1)
                                  End Sub)
        End If



    End Sub

    Public Sub filout(ByVal title As String,
                      ByVal description As String,
                      ByVal serial As String,
                      ByVal state As Integer)
        Dim act As New actions(title, description, serial, state,
                               DateTime.Now)

        listActions.Add(act)

        fillDataGridView(act)
    End Sub

    Public Sub putTheMaterials()
        running = True


        'rpUutFctSerialNumber(0) = InputBox("Please scan the Serial Number", "Write Data").ToUpper
        'rpUutFctSerialNumber(0) = "D101234567"
        'Call fSerialNumberRead(0, sn)
        'MsgBox("PUT NEW PART ON POSITION 1")
        'If rpUutFctSerialNumber(0) = "" Or Left(rpUutFctSerialNumber(0), 1) <> "D" Or rpUutFctSerialNumber(0).Length <> 10 Then
        '    MsgBox("Le SN scanné est erroné :( !!")

        '    Exit Do
        '    'ElseIf rpUutFctSerialNumber(0) = "E330062910" Or rpUutFctSerialNumber(0) = "E330062913" Then
        '    '    MsgBox("PUT IT A SIDE")
        'End If
        lbltxtflag.Invoke(Sub()
                              lbltxtflag.Text = "FLASH ongoing"
                          End Sub
        )

        If fFCT() <> PASS Then
            ' *** SET fail flag
            filout("ERROR FLASH FAILED", "Unable to flash the part", "Nan", -1)
            panelflag.Invoke(Sub()
                                 panelflag.BackColor = Color.Red
                             End Sub
        )
            lbltxtflag.Invoke(Sub()
                                  lbltxtflag.Text = "FLASH failed"
                              End Sub
                )
            _task.Abort()
            running = False
            'MsgBox("STATION1 : failled!!!! ")
        Else
            'MsgBox("Write operation Done with " & rpUutFctResult(0) & " error")
            'MsgBox("<<<<<<<<<<<<<<<     POSITION1    >>>>>>>>>>>>>" & Chr(10) & Chr(10) & "  PART FLASHED CORRECTLY  ")
            'filout("POSITION1", "PART FLASHED CORRECTLY ", "Nan", True)
            panelflag.Invoke(Sub()
                                 panelflag.BackColor = Color.LawnGreen
                             End Sub
      )
            lbltxtflag.Invoke(Sub()
                                  lbltxtflag.Text = "FLASH passed"
                              End Sub
                )
        End If


    End Sub

    Private Sub ReflectionTimer_Tick(sender As Object, e As EventArgs) Handles ReflectionTimer.Tick
        initAll()
    End Sub

    Private Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        ActionDataList.Rows.Clear()
        _task = New Thread(Sub()
                               putTheMaterials()
                           End Sub)
        _task.Start()
    End Sub

    Public Function fFCT() As Long
        Dim SN_AFTER As String = Nothing
        Dim SN_BEFORE As String = Nothing
        Dim rUserFlagArrayString As String = Nothing
        Dim rUserFlagCPUArrayString As String = Nothing
        Dim aUserFlagArray(32) As Integer
        Dim aUserFlagCPUArray(32) As Integer
        Dim rPPSu As Integer = Nothing
        Dim rPPsu_RearbackVoltage As Double = Nothing
        Dim rPPsu_RearbackCorrent As Double = Nothing
        Dim rError As String = Nothing
        Dim aUsedChForStuck(2) As Short
        rpRpkTpgmVariant = "5CTH12F10"

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
                    filout("Start opening can", "Start opening can port", "going", 0)
                    If fCanPortOpen(rpCanPortHandle, rpCanPort, rpCanSpeed) <> 0 Then
                        filout("Error", "Unable to open can port", "FAIL", -1)
                        fFCT = FAIL
                    Else
                        filout("Done Open can", "Open can communication", "PASS", 1)
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
                        'filout("Read SRN", "Read serial number from the part micro", "Nan", False)
                        rpFCTNtest = 3007                                                          ' *** Set Test Number
                        rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                        rpFCTTestName = "Read SN"                                                   ' *** Set Test Name
                        rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                        rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                        Dim rSerialNumberRead As String = Nothing
                        'rSerialNumberRead = "D121234567"
                        filout("Start read SN", "trying to read serial number", "Nan", 0)
                        If fSerialNumberRead(rSite, rSerialNumberRead) <> PASS Then
                            rpUutFctResult(rSite) = FAIL
                            'MsgBox("RAED SN FAIL")
                            filout("Error RAED SN FAIL", "Unable to read serial number from the part", "Nan", -1)

                        End If

                        SN_BEFORE = rSerialNumberRead

                        'MsgBox(Strings.Left(rpUutFctSerialNumber(0), 1))
                        lblsrn.Invoke(Sub()
                                          lblsrn.Text = "Serial number : " + rSerialNumberRead
                                      End Sub)

                        If rSerialNumberRead = "" Or Strings.Left(SN_BEFORE, 1) <> "D" Then

                            filout("error SN READ", "SERIAL NUMBER READ NOT CORRECT", rSerialNumberRead & Chr(10) & Chr(10), -1)
                            'lblsrn.Invoke(Sub()
                            '                  lblsrn.Text = "Serial number : "
                            '              End Sub)
                            fFCT = FAIL

                            GoTo NEXT_SITE
                        Else
                            filout("Done SN READ", "SN READ BEFORE REFLASH CORRECT", rSerialNumberRead & Chr(10) & Chr(10), 1)

                            fFCT = PASS


                        End If


                        'filout("POSITION1", "part to be discard", rSerialNumberRead & Chr(10) & Chr(10), True)
                        'MsgBox("POSITION1" & Chr(10) & Chr(10) & "SN:  " & rSerialNumberRead & Chr(10) & Chr(10) & "PLEASE REFLASH BLU + APP THEN PRESS OK")

                        'GoTo FLASHSBL
                        filout("START FLASH BLU", "START UPDATE OF BOOTLOADER", "" & Chr(10) & Chr(10), 0)
                        com("PAYLOAD_BLU")
                        filout("Done FLASH BLU", "UPDATE OF BOOTLOADER", "" & Chr(10) & Chr(10), 1)

                        filout("START FLASH APPLICATION", "START FLASH OF SW E13411010_SW-AL ", "" & Chr(10) & Chr(10), 0)
                        com(rpRpkTpgmVariant & "_APP")
                        filout("Done FLASH APPLICATION", "FLASH OF SW E13411010_SW-AL ", "" & Chr(10) & Chr(10), 1)

FLASHSBL:               filout("START FLASH SBL", "START FLASH SECONDARY BOOTLOADER", "" & Chr(10) & Chr(10), 0)
                        com("PAYLOADSBL")
                        filout("Done FLASH SBL", "FLASH SECONDARY BOOTLOADER", "" & Chr(10) & Chr(10), 1)

                        filout("START FLASH PNS", "START FLASH CORE ASS\DELIV ASS\PROD DATE", "" & Chr(10) & Chr(10), 0)
                        com(rpRpkTpgmVariant & "_WR_PNS")
                        filout("Done FLASH PNS", "FLASH CORE ASS\DELIV ASS\PROD DATE", "" & Chr(10) & Chr(10), 1)


                        ' ********************************************************************************************************************************************
                        ' *** SERIAL NUMBER WRITE
                        '********************************************************************************************************************************************
                        filout("START WRITE SN", "START WRITE SN", "" & Chr(10) & Chr(10), 0)
                        rpFCTNtest = 3006                                                          ' *** Set Test Number
                        rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            rpFCTTestName = "Write SN"                             ' *** Set Test Name
                            rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                            rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                            If fSerialNumberWrite(SN_BEFORE) <> PASS Then
                                rpUutFctResult(rSite) = FAIL
                                fFCT = FAIL
                            filout("Error WRITE SN", "UNABLE TO WRITE SN", "" & Chr(10) & Chr(10), -1)

                            GoTo NEXT_SITE
                        Else
                            filout("Done WRITE SN", "SN WRITTEN SUCCESSFULLY", "" & Chr(10) & Chr(10), 1)

                        End If
                        'MsgPrintLogIfEn_v0_1(rpRpkBayInUse, rSite + 1, rpFCTNtest, rpFCTDrawRef, rpFCTTestName, rpUutFctResult(rSite), rpFCTDiagRemark, -1000000, -1000000, -1000000, rpFCTUnitMeasure, "ALWAYS")

                        filout("START CAN CHANNEL CONFIG", "START CAN CHANNEL CONFIG", "" & Chr(10) & Chr(10), 0)
                        com(rpRpkTpgmVariant & "_WR_DE14")
                        filout("Done CAN CHANNEL CONFIG", "CAN CHANNEL CONFIG", "" & Chr(10) & Chr(10), 1)

                        'PROVA1:
                        ' ********************************************************************************************************************************************
                        ' *** SERIAL NUMBER READ
                        ''********************************************************************************************************************************************
                        ' ********************************************************************************************************************************************
                        ' ***  READ SN AFTER REWRITTING
                        ''********************************************************************************************************************************************
                        rpFCTNtest = 3007                                                          ' *** Set Test Number
                            rpFCTDrawRef = "FCT"                                                      ' *** Set Drawing Reference
                            rpFCTTestName = "Read SN"                                                   ' *** Set Test Name
                            rpFCTDiagRemark = rpFCTTestName & "- FAIL"                                    ' *** Set Diag Remark
                        'rpFCTUnitMeasure = ""                                                      ' *** Set Unit Measure
                        'Dim rSerialNumberRead As String = Nothing
                        filout("START READ SN", "START READ SN", "" & Chr(10) & Chr(10), 0)
                        If fSerialNumberRead(rSite, SN_AFTER) <> PASS Or SN_AFTER <> SN_BEFORE Then
                            rpUutFctResult(rSite) = FAIL
                            fFCT = FAIL
                            filout("Error READ SN", "UNABLE TO READ SN ", rSerialNumberRead & Chr(10) & Chr(10), -1)
                            GoTo NEXT_SITE
                        Else
                            filout("Done READ SN", "READ SN AFTER FLASH SUCCESSFULLY", rSerialNumberRead & Chr(10) & Chr(10), 1)
                        End If


                        filout("START READ F111/F113/DE14", "START MATCH F111/F113/DE14 WITH 5CTH12F10 PNs", "" & Chr(10) & Chr(10), 0)
                        com("PAYLOAD_CHECK_DIDs")
                        filout("Done READ F111/F113/DE14", "MATCH F111/F113/DE14 WITH 5CTH12F10 PNs", "" & Chr(10) & Chr(10), 1)


NEXT_SITE:

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

    Private Sub BunifuFlatButton3_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton3.Click
        _task.Abort()
        running = False
    End Sub

    Private Sub cmbFiles_onItemSelected(sender As Object, e As EventArgs)

    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub imgstate_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnActions_Click(sender As Object, e As EventArgs) Handles btnActions.Click

    End Sub
End Class