
Imports System.Threading

Module modAutomation

    Public Function fReceiverMiddle(ByVal rTimeOutMs As Integer) As Integer

        Dim rtime As Integer
        Dim rtimeout As Integer = 5000
        Dim rStepTime As Integer = 1

        ' *** Set Function Result
        fReceiverMiddle = PASS

        ' *** Command
        ReceiverMiddle()

        ' *** Debug Print
        MsgPrintLogIfEn("@FG{Purple}DBG_D_1.0_fReceiverMiddle_ReceiverMiddle", 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Lopp Until receiver in Middle position
        Do Until (IsReceiverMiddle() = 1) Or (rtime = rtimeout)
            ' *** 0 = not middle
            ' *** 1 = middle
            Thread.Sleep(rStepTime)
            rtime = rtime + rStepTime

        Loop

        ' *** Check if timeout is occurred
        If rtime = rtimeout Then
            ' *** Debug Print
            MsgPrintLogIfEn("@FG{Purple}DBG_D_1.0_fReceiverMiddle_TimeOutOccurred", 0, "IF ENABLED", rpIniDebugPrint)
            ' *** ERROR MESSAGE
            MsgPrintLogIfEn("@BG{Red}@FG{White}RECEIVER NOT IN MIDDLE POSITION...Result Fail", 0, "ALWAYS")
            ' *** Set Function Result
            fReceiverMiddle = FAIL
        Else

            ' *** Debug Print
            MsgPrintLogIfEn("@FG{Purple}DBG_D_1.0_fReceiverMiddle_CorrectlyPositioned", 0, "IF ENABLED", rpIniDebugPrint)

        End If

    End Function

    Public Function fReceiverDown(ByVal rTimeOutMs As Integer) As Integer

        Dim rtime As Integer
        Dim rtimeout As Integer = 5000
        Dim rStepTime As Integer = 1

        ' *** Set Function Result
        fReceiverDown = PASS

        ' *** Command
        ReceiverDown()

        ' *** Debug Print
        MsgPrintLogIfEn("@FG{Purple}DBG_D_1.0_fReceiverDown_ReceiverDown", 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Lopp Until receiver in Middle position
        Do Until (IsReceiverDown() = 1) Or (rtime = rtimeout)
            ' *** 0 = not middle
            ' *** 1 = middle
            Thread.Sleep(rStepTime)
            rtime = rtime + rStepTime

        Loop

        ' *** Check if timeout is occurred
        If rtime = rtimeout Then
            ' *** Debug Print
            MsgPrintLogIfEn("@FG{Purple}DBG_D_1.0_fReceiverDown_TimeOutOccurred", 0, "IF ENABLED", rpIniDebugPrint)
            ' *** ERROR MESSAGE
            MsgPrintLogIfEn("@BG{Red}@FG{White}RECEIVER NOT IN DOWN POSITION...Result Fail", 0, "ALWAYS")
            ' *** Set Function Result
            fReceiverDown = FAIL
        Else

            ' *** Debug Print
            MsgPrintLogIfEn("@FG{Purple}DBG_D_1.0_fReceiverDown_CorrectlyPositioned", 0, "IF ENABLED", rpIniDebugPrint)

        End If

    End Function

    Public Function fReceiverUp(ByVal rTimeOutMs As Integer) As Integer

        Dim rtime As Integer
        Dim rtimeout As Integer = 5000
        Dim rStepTime As Integer = 100

        ' *** Set Function Result
        fReceiverUp = PASS

        ' *** Command
        ReceiverUp()

        ' *** Debug Print
        MsgPrintLogIfEn("@FG{Purple}DBG_D_1.0_fReceiverUp_ReceiverUp", 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Lopp Until receiver in Middle position
        Do Until (IsReceiverUp() = 1) Or (rtime = rtimeout)
            ' *** 0 = not middle
            ' *** 1 = middle
            Thread.Sleep(rStepTime)
            rtime = rtime + rStepTime
        Loop

        ' *** Check if timeout is occurred
        If rtime = rtimeout Then
            ' *** Debug Print
            MsgPrintLogIfEn("@FG{Purple}DBG_D_1.0_fReceiverUp_TimeOutOccurred", 0, "IF ENABLED", rpIniDebugPrint)
            ' *** ERROR MESSAGE
            MsgPrintLogIfEn("@BG{Red}@FG{White}RECEIVER NOT IN UP POSITION...Result Fail", 0, "ALWAYS")
            ' *** Set Function Result
            fReceiverUp = FAIL
        Else
            ' *** Debug Print
            MsgPrintLogIfEn("@FG{Purple}DBG_D_1.0_fReceiverUp_CorrectlyPositioned", 0, "IF ENABLED", rpIniDebugPrint)
        End If

    End Function

    Public Function fSyncParallelExec(ByVal rTPGMSectorName As String) As Long

        ' *** Set Function Result
        fSyncParallelExec = FAIL

        ' *** Debug Print
        'MsgPrintLogIfEn("@FG{Blue}DBG_D_1.0_" & rTPGMSectorName & "_SyncBegin   : BAY" & CStr(rpRpkBayInUse), 0, "IF ENABLED", rpIniDebugPrint)

        'MsgBox(rTPGMSectorName)

        ' *** Waiting Bay Syncronization
        SyncParallelExec()

        ' *** Debug Print
        'MsgPrintLogIfEn("@FG{Blue}DBG_D_1.0_" & rTPGMSectorName & "_SyncEnd     : BAY" & CStr(rpRpkBayInUse), 0, "IF ENABLED", rpIniDebugPrint)

        ' *** Set Function Result
        fSyncParallelExec = PASS

    End Function

End Module
