
Option Explicit On

Imports AtosF
Imports vxlapi_NET20
Imports System.Threading
Imports System.Globalization

Module mod08_CanRoutineV2

    Public Const XL_BUS_TYPE_CAN As Integer = 1
    Public Const XL_ACTIVATE_RESET_CLOCK As Integer = 8
    Public Const XL_ERR_QUEUE_IS_EMPTY As Integer = 10
    Public Const XL_TRANSMIT_MSG As Integer = 10
    Public Const XL_CAN_STD As Integer = 1
    Public Const XL_CAN_EXT As Integer = 2

    Public rIsFirstTime As Boolean = True
    Public rpCanDevice1 As New XLDriver
    Public rpCanDevice2 As New XLDriver


    Public Function fCanPortOpenV2(ByRef myCanPortHandle As Integer, ByVal myCanPort As ULong, ByVal myCanSpeed As UInteger) As Integer

        Dim rXlStatus As XLClass.XLstatus

        Try

            ' *** Inizializzazione porta CAN
            If myCanPort = 1 Then
                rXlStatus = rpCanDevice1.XL_OpenDriver()
            Else
                rXlStatus = rpCanDevice2.XL_OpenDriver()
            End If

            If rXlStatus <> 0 Then
                MsgBox("XL_OpenDriver: FAIL. Error code: " & rXlStatus.ToString & " [" & rXlStatus & "]")
                Return (1)
            End If

            If myCanPort = 1 Then
                rXlStatus = rpCanDevice1.XL_OpenPort(myCanPortHandle, "xlCANcontrol", myCanPort, myCanPort, 256, XL_BUS_TYPE_CAN)
            Else
                rXlStatus = rpCanDevice2.XL_OpenPort(myCanPortHandle, "xlCANcontrol", myCanPort, myCanPort, 256, XL_BUS_TYPE_CAN)
            End If
            If rXlStatus <> 0 Then
                MsgBox("XL_OpenPort: FAIL. Error code: " & rXlStatus.ToString & " [" & rXlStatus & "]")
                Return (1)
            End If

            If myCanPort = 1 Then
                rXlStatus = rpCanDevice1.XL_CanSetChannelBitrate(myCanPortHandle, myCanPort, myCanSpeed)
            Else
                rXlStatus = rpCanDevice2.XL_CanSetChannelBitrate(myCanPortHandle, myCanPort, myCanSpeed)
            End If
            If rXlStatus <> 0 Then
                MsgBox("XL_CanSetChannelBitrate: FAIL. Error code: " & rXlStatus.ToString & " [" & rXlStatus & "]")
                Return (1)
            End If

            If myCanPort = 1 Then
                rXlStatus = rpCanDevice1.XL_ActivateChannel(myCanPortHandle, myCanPort, XL_BUS_TYPE_CAN, XL_ACTIVATE_RESET_CLOCK)

            Else
                rXlStatus = rpCanDevice2.XL_ActivateChannel(myCanPortHandle, myCanPort, XL_BUS_TYPE_CAN, XL_ACTIVATE_RESET_CLOCK)

            End If
            If rXlStatus <> 0 Then
                MsgBox("XL_ActivateChannel: FAIL. Error code: " & rXlStatus.ToString & " [" & rXlStatus & "]")
                Return (1)
            End If

            'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.fCanPortOpen_Port       : " & CStr(rCanPort), 0, "IF ENABLED", rpDebugPrint)
            'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.fCanPortOpen_Speed      : " & CStr(rCanSpeed), 0, "IF ENABLED", rpDebugPrint)
            rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.fCanPortOpen_Port", rpRpkBayInUse, CStr(myCanPort), True)
            rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.fCanPortOpen_Speed", rpRpkBayInUse, CStr(myCanPort), True)

            Return (0)

        Catch ex As Exception

            MsgBox("fCanPortOpen -> Run-Time error: " & ex.Message)
            Return (1)

        End Try

    End Function

    Public Function fCanPortCloseV2(ByVal myCanPortHandle As Integer, ByVal myCanPort As ULong) As Integer

        Dim rXlStatus As XLClass.XLstatus

        Try

            If myCanPort = 1 Then
                rXlStatus = rpCanDevice1.XL_DeactivateChannel(myCanPortHandle, myCanPort)
            Else
                rXlStatus = rpCanDevice2.XL_DeactivateChannel(myCanPortHandle, myCanPort)
            End If
            If rXlStatus <> 0 Then
                MsgBox("XL_DeactivateChannel: FAIL. Error code: " & rXlStatus.ToString & " [" & rXlStatus & "]")
                Return (1)
            End If
            If myCanPort = 1 Then
                rXlStatus = rpCanDevice1.XL_ClosePort(myCanPortHandle)
            Else
                rXlStatus = rpCanDevice2.XL_ClosePort(myCanPortHandle)
            End If
            If rXlStatus <> 0 Then
                MsgBox("XL_ClosePort: FAIL. Error code: " & rXlStatus.ToString & " [" & rXlStatus & "]")
                Return (1)
            End If
            If myCanPort = 1 Then
                rXlStatus = rpCanDevice1.XL_CloseDriver
            Else
                rXlStatus = rpCanDevice2.XL_CloseDriver
            End If
            If rXlStatus <> 0 Then
                MsgBox("XL_CloseDriver: FAIL. Error code: " & rXlStatus.ToString & " [" & rXlStatus & "]")
                Return (1)
            End If

            'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.0_fCanPortClose_Port    : " & CStr(rCanPort), 0, "IF ENABLED", rpDebugPrint)
            rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.0_fCanPortClose_Port", rpRpkBayInUse, CStr(myCanPort), True)

            Return (0)

        Catch ex As Exception

            MsgBox("fCanPortClose -> Run-Time error: " & ex.Message)
            Return (1)

        End Try

        sFileTxtAppend("---------------------------------------------------------------------------------") ' *** End of Can Trace file

    End Function

    Public Function fCanRxFramesV2(ByVal myCanPortHandle As Integer, ByVal myCanPort As ULong, ByVal rRxID As String,
                                 ByRef rFrameRx() As String, Optional ByRef rTimeStamp As ULong = 0,
                                 Optional ByVal rOutputType As Integer = 0) As Long

        ' *** For choose OutputType refer to this table

        ' *** HEX = Typ_Hexdec  = 0 ' *** HEX ARRAY STRING
        ' *** BIN = Typ_Binary  = 1 ' *** BIN ARRAY STRING
        ' *** DEC = Typ_Decimal = 2 ' *** DEC ARRAY STRING

        Dim rXlStatus As XLClass.XLstatus
        Dim rXlEventTx As New XLClass.xl_event
        Dim rXlEventRx As XLClass.xl_event
        Dim rByte As Integer
        Dim rRxData As String
        Dim rRxDataLen As Long
        Dim rId As Long
        Dim rRxId_Dec As Long

        'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.0_fCanRxFrames_START", 0, "IF ENABLED", rpDebugPrint)
        rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.0_fCanRxFrames_START", rpRpkBayInUse, True)

        fCanRxFramesV2 = FAIL

        'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.1_fCanRxFrames_CAN PORT : " & rCanPort.ToString, 0, "IF ENABLED", rpDebugPrint)
        rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.1_fCanRxFrames_CAN PORT", rpRpkBayInUse, True)
        'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.2_fCanRxFrames_ID_TO_RX : " & rRxID, 0, "IF ENABLED", rpDebugPrint)
        rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.2_fCanRxFrames_ID_TO_RX : " & rRxID, rpRpkBayInUse, True)


        ' *** Convert Id from Hex to Decimal
        rRxId_Dec = fConv_HexToDec(rRxID)

        ' *** Check if ID is Standard o Extended
        If rRxId_Dec > &HFFF Then rRxId_Dec = rRxId_Dec + 2147483648

        ' *** Filtro di ricezione
        If rRxId_Dec > &HFFF Then
            rpCanDevice.XL_CanResetAcceptance(myCanPortHandle, myCanPort, XL_CAN_EXT)
        Else
            rpCanDevice.XL_CanResetAcceptance(myCanPortHandle, myCanPort, XL_CAN_STD)
        End If

        ' *** Ricevo Dato aggiornato
        Dim rMaxTry As Integer = 0
        Do
            rXlStatus = rpCanDevice.XL_Receive(myCanPortHandle, rXlEventRx)

            ' *** Se ho ricevuto dei dati
            If rXlStatus <> XL_ERR_QUEUE_IS_EMPTY Then

                rId = rXlEventRx.tagData.can_Msg.id

                If rId <> 0 Then
                    'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.3_fCanRxFrames_ID_RX    : " & fConv_DecToHex(rId), 0, "IF ENABLED", rpDebugPrint)
                    rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.3_fCanRxFrames_ID_RX", rpRpkBayInUse, fConv_DecToHex(rId), True)
                End If


            End If
            If rMaxTry = 10000 Then Exit Do
            rMaxTry = rMaxTry + 1
        Loop Until (rXlStatus = XL_ERR_QUEUE_IS_EMPTY) Or (rRxId_Dec = rId)

        ' *** Se ho ricevuto dei dati dall'id richiesto
        If rRxId_Dec = rId Then

            rTimeStamp = rXlEventRx.timeStamp

            'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.3_fCanRxFrames_MATCH ID : PASS", 0, "IF ENABLED", rpDebugPrint)
            rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.3_fCanRxFrames_MATCH ID", rpRpkBayInUse, PASS, True)

            rRxData = ""
            rRxDataLen = rXlEventRx.tagData.can_Msg.dlc

            Dim rNum As Long
            Dim rRxDataDec As Long

            For rByte = 0 To rRxDataLen - 1

                rRxDataDec = rXlEventRx.tagData.can_Msg.data(rByte)

                Select Case rOutputType

                    Case 0 ' *** HEX ARRAY STRING

                        rFrameRx(rByte) = fConv_DecToHex(rRxDataDec)

                    Case 1 ' *** BIN ARRAY STRING

                        rFrameRx(rByte) = fConv_DecToHex(rRxDataDec)
                        rNum = Long.Parse(rFrameRx(rByte), NumberStyles.HexNumber)
                        rFrameRx(rByte) = Convert.ToString(rNum, 2).PadLeft(8, "0")

                        'rFrameRx(rByte) = Convert.ToString(rNum, 2).PadLeft(8, "0")

                    Case 2 ' *** DEC ARRAY STRING

                        rFrameRx(rByte) = CStr(rRxDataDec)

                End Select

                rRxData = rRxData & fConv_DecToHex(rXlEventRx.tagData.can_Msg.data(rByte))

            Next

            If rRxData <> " " Then sFileTxtAppend(CStr(Date.Today) & " - TestID:" & CStr(rpNtest) & " - [RX]-[ID: &H" & fConv_DecToHex(rRxId_Dec) & "]-[DATA: &H" & rRxData & "]") 'scrivo il frame ricevuto sul file di trace

            'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.4_fCanRxFrames_RX DATA  : " & rRxData, 0, "IF ENABLED", rpDebugPrint)
            rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.4_fCanRxFrames_RX DATA", rpRpkBayInUse, rRxData, True)

            fCanRxFramesV2 = PASS

            ' *** riprovo a ricevere 100 volte
            'rRxRetry = rRxRetry + 1
            'If rRxRetry < 2 Then GoTo RxRetry
            'rFunctionResult = 1

        Else

            'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.5_fCanRxFrames_MATCH ID : FAIL", 0, "IF ENABLED", rpDebugPrint)
            rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.5_fCanRxFrames_MATCH ID", rpRpkBayInUse, FAIL, True)
            'MsgPrintLogIfEn("@FG{Purple}Can Rx Frames ID: [" & rRxID & "] not Found", 0, "IF ENABLED", rpDebugPrint)
            rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}Can Rx Frames ID: [" & rRxID & "] not Found", rpRpkBayInUse, True)


        End If

        'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.6_fCanRxFrames_END", 0, "IF ENABLED", rpDebugPrint)
        rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.6_fCanRxFrames_END", rpRpkBayInUse, True)


    End Function

    Public Function fCanTxFramesV2(ByVal myCanPortHandle As Integer, ByVal myCanPort As ULong, ByVal rTxID As String, ByRef rFrameTx As String) As Long

        Dim rXlStatus As XLClass.XLstatus
        Dim rXlEventTx As New XLClass.xl_event

        Dim rByte As Integer
        Dim rByteStr As Integer
        Dim rFrameLen As Integer
        Dim rTxData As String
        Dim rId As Long

        fCanTxFramesV2 = PASS

        'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.0_fCanTxFrames_CAN_PORT : " & rCanPort.ToString, 0, "IF ENABLED", rpDebugPrint)
        rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.0_fCanTxFrames_CAN_PORT", rpRpkBayInUse, myCanPort.ToString, True)
        'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.1_fCanTxFrames_TX_ID    : " & rTxID.ToString, 0, "IF ENABLED", rpDebugPrint)
        rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.1_fCanTxFrames_TX_ID", rpRpkBayInUse, rTxID.ToString, True)
        'MsgPrintLogIfEn("@FG{Purple}DBG_C_1.2_fCanTxFrames_TX_DATA  : " & rFrameTx, 0, "IF ENABLED", rpDebugPrint)
        rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_C_1.2_fCanTxFrames_TX_DATA", rpRpkBayInUse, rFrameTx, True)

        Try

            ' *** Convert Id from Hex to Decimal
            rId = fConv_HexToDec(rTxID)

            ' *** Check if ID is Standard o Extended
            If rId > &HFFF Then rId = rId + 2147483648

            If rFrameTx(0) <> "" Then

                ' *** Impostazioni di trasmissione
                rXlEventTx.tagData.can_Msg.id = rId
                rXlEventTx.tag = XL_TRANSMIT_MSG
                rXlEventTx.tagData.can_Msg.flags = 0

                ' *** Check if ID is Standard o Extended
                If rId > &HFFF Then
                    If myCanPort = 1 Then
                        rpCanDevice1.XL_CanResetAcceptance(myCanPortHandle, myCanPort, XL_CAN_EXT)
                    Else
                        rpCanDevice2.XL_CanResetAcceptance(myCanPortHandle, myCanPort, XL_CAN_EXT)
                    End If
                Else
                    If myCanPort = 1 Then
                        rpCanDevice1.XL_CanResetAcceptance(myCanPortHandle, myCanPort, XL_CAN_STD)
                    Else
                        rpCanDevice2.XL_CanResetAcceptance(myCanPortHandle, myCanPort, XL_CAN_STD)
                    End If
                End If

                ' *** Filtro di ricezione
                If myCanPort = 1 Then
                    rpCanDevice1.XL_CanAddAcceptanceRange(myCanPortHandle, myCanPort, rId, rId)
                Else
                    rpCanDevice2.XL_CanAddAcceptanceRange(myCanPortHandle, myCanPort, rId, rId)
                End If

                ' *** Lunghezza caratteri esadecimali da trasmettere
                rFrameLen = Len(rFrameTx)

                ' *** Set lunghezza dato decimale da trasmettere
                rXlEventTx.tagData.can_Msg.dlc = rFrameLen / 2

                ' *** Preparazione dato da trasmettere
                rByte = 0
                For rByteStr = 0 To rFrameLen - 1 Step 2
                    rTxData = Mid(rFrameTx, rByteStr + 1, 2)
                    rXlEventTx.tagData.can_Msg.data(rByte) = fConv_HexToDec(rTxData)
                    rByte = rByte + 1
                Next

                ' *** Send data
                If myCanPort = 1 Then
                    rXlStatus = rpCanDevice1.XL_CanTransmit(myCanPortHandle, myCanPort, rXlEventTx)
                Else
                    rXlStatus = rpCanDevice2.XL_CanTransmit(myCanPortHandle, myCanPort, rXlEventTx)
                End If
                If rXlStatus <> 0 Then
                    MsgBox("XL_CanTransmit: FAIL. Error code: " & rXlStatus.ToString & " [" & rXlStatus & "]")
                    Return (1)
                End If

                ' *** Format Data for Trace
                If rTxID.Length < 8 Then rTxID = rTxID & Space(8 - rTxID.Length)

                sFileTxtAppend(CStr(Date.Today) & " - TestID:" & CStr(rpNtest) & " - [TX]-[ID: &H" & rTxID & "]-[DATA: &H" & rFrameTx & "]") ' *** scrivo il frame trasmesso sul file di trace

            End If

        Catch ex As Exception

            MsgBox("fCanTxFrames -> Run-Time error: " & ex.Message)
            Return (1)

        End Try

    End Function

    Public Function fCanBufferClearV2(ByVal myCanPortHandle As Integer, ByVal myCanPort As ULong, Optional ByVal rWaitAfterClear As Integer = 1, Optional ByVal rMsgToBePreserved As Integer = 5) As Long

        Dim rXlStatus As XLClass.XLstatus
        Dim rXlEventRx As XLClass.xl_event

        Dim rMsgNumber As Integer = 0
        Dim rMsgCleared As Integer = 0

        Dim rId As Long = 0

        ' *** Set Function Result
        fCanBufferClearV2 = PASS

        ' *** Debug Print
        MsgPrintLogIfEn("@FG{Purple}DBG_X_1.0_fCanBufferClear_Remin : " & rMsgNumber.ToString, 0, "IF ENABLED", rpDebugPrint)

        ' *** Get Dimension Buffer
        rpCanDevice.XL_GetReceiveQueueLevel(myCanPortHandle, rMsgNumber)

        ' *** Debug Print
        MsgPrintLogIfEn("@FG{Purple}DBG_X_1.0_fCanBufferClear_Size  : " & rMsgNumber.ToString, 0, "IF ENABLED", rpDebugPrint)

        ' Check if clear buffer or not
        If rMsgNumber > rMsgToBePreserved Then

            ' *** Svuota Buffer Can
            Do
                rXlStatus = rpCanDevice.XL_Receive(myCanPortHandle, rXlEventRx)
                rId = rXlEventRx.tagData.can_Msg.id
                rMsgNumber = rMsgNumber - 1
                rMsgCleared = rMsgCleared + 1

            Loop Until (rMsgNumber = rMsgToBePreserved)

            'MsgPrintLogIfEn("@FG{Purple}DBG_X_1.0_fCanBufferClear_Clear : " & rMsgCleared.ToString, 0, "IF ENABLED", rpDebugPrint)
            rpMsgLogClass.MsgPrintLogMultyBay("@FG{Purple}DBG_X_1.0_fCanBufferClear_Clear", rpRpkBayInUse, rMsgCleared.ToString, True)


        End If

        Thread.Sleep(rWaitAfterClear)

    End Function

    Public Function fCanBufferClearFlushV2(ByVal myCanPortHandle As Integer, ByVal myCanPort As ULong, Optional ByVal rWaitAfterClear As Integer = 1, Optional ByVal rMsgToBePreserved As Integer = 5) As Long

        Dim rXlStatus As XLClass.XLstatus
        Dim rXlEventRx As XLClass.xl_event

        Dim rMsgNumber As Integer = 0
        Dim rMsgCleared As Integer = 0

        Dim rId As Long = 0

        ' *** Set Function Result
        fCanBufferClearFlushV2 = PASS

        rXlStatus = rpCanDevice.XL_FlushReceiveQueue(rpCanPortHandle)

        Thread.Sleep(rWaitAfterClear)

    End Function

End Module