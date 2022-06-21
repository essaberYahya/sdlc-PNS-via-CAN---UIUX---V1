
Imports System.Threading

Module mod11_ErrorMessage_v0_4

    ''' <summary>
    ''' Function Description
    ''' </summary>
    ''' <param name="rMessage">Messaggio da stampare[String]</param>
    ''' <param name="rAvoidCrLf">A Capo [0=SI/1=NO]</param>
    ''' <param name="rPrintCondition">Condizione di stampa [SI/NO]</param>
    ''' <param name="rEnableVarValue">Parametro che abilita/disabilita la stampa condizionata["ALWAYS"/"IF ENABLED"]</param>
    ''' <returns></returns>
    ''' no returns
    ''' <remarks></remarks>

    Public Function MsgPrintLogIfEn(ByVal rMessage As String, ByVal rAvoidCrLf As Long, ByVal rPrintCondition As String, Optional ByVal rEnableVarValue As Long = NO) As Long

        'rMessage = "@FG{Black}" & rMessage


        Dim rShift As String = 50
        Dim rMessageLen As Integer
        Dim rEndPosition As Integer
        Dim rOccurence(10) As String

        rMessageLen = rMessage.Length

        ' *** Check if string constains color settings
        If (rMessage.Contains("@FG")) Or (rMessage.Contains("@BG")) Then

            ' *** How many "@" there are?
            rOccurence = Split(rMessage, "@")

            For rIndex As Integer = 1 To rOccurence.Length - 1
                rEndPosition = rOccurence(rIndex).IndexOf("}") + 2
                rMessageLen = rMessageLen - rEndPosition
            Next

        End If

        ' *** Select Shift of Massage
        If rpRpkBayInUse = 1 Then

            ' *** Set Print Shift Only For Bay 1
            If rMessageLen > rShift Then
                If rMessageLen = ((rShift + 1) * 2) Then
                    rMessage = "@FG{Black}" & rMessage ' *** Stampa titoli
                Else
                    MsgBox("MsgPrintLogIfEn - Internal Function Error - Message length is too big")
                End If
            Else
                rMessage = "@FG{Black}" & rMessage & "@BG{White}@FG{Black}" & Space(rShift - rMessageLen) & "|" & Space(rShift) & "|"
            End If

        ElseIf rpRpkBayInUse = 2 Then

            ' *** Set Print Shift Only For Bay 2
            If rMessageLen > rShift Then
                MsgBox("MsgPrintLogIfEn - Internal Function Error - Message length is too big")
            Else
                rMessage = "@FG{Black}" & Space(rShift) & "@FG{Black}|" & rMessage & "@BG{White}@FG{Black}" & Space(rShift - rMessageLen) & "@BG{White}@FG{Black}" & "|"
            End If

        End If

        ' **************************************************************************************************************************************
        ' *** SYSTEM LOG - Check Print Contition
        ' **************************************************************************************************************************************
        If rPrintCondition = "ALWAYS" Or (rPrintCondition = "IF ENABLED" And rEnableVarValue = YES) Then

            Call MsgPrintLog(rMessage, rAvoidCrLf)            ' *** Print Message

        End If

    End Function

    Public Function MsgPrintLogIfEn_v0_1(ByVal rBay As Integer, ByVal rSite As Integer, ByVal rTestNumbr As Integer,
                                         ByVal rDrawingReference As String, ByVal rTestName As String, ByRef rResult As Integer, ByVal rRemark As String,
                                         ByVal rTl As Double, ByVal rMeasuredValue As Double, ByVal rTh As Double, ByVal rUnitMeasure As String,
                                         ByVal rPrintCondition As String, Optional ByVal rEnableVarValue As Long = NO,
                                         Optional ByVal rAvoidCrLf As Long = 0) As Long

        Dim rMessage As String = rTestName
        Dim rMessageLen As Integer = rMessage.Length
        Dim rShift As String = 50

        ' *** Select Shift of Massage

        'If rpCurrentBay = 1 Then

        '    ' *** Set Print Shift Only For Bay 1
        '    If rMessageLen > rShift Then
        '        If rMessageLen = ((rShift + 1) * 2) Then
        '            rMessage = "@FG{Black}" & rMessage ' *** Stampa titoli
        '        Else
        '            MsgBox("MsgPrintLogIfEn - Internal Function Error - Message length is too big")
        '        End If
        '    Else
        '        rMessage = "@FG{Black}" & rMessage & "@BG{White}@FG{Black}" & Space(rShift - rMessageLen) & "|" & Space(rShift) & "@BG{White}@FG{Black}" & "|"
        '    End If

        'ElseIf rpCurrentBay = 2 Then

        '    ' *** Set Print Shift Only For Bay 2
        '    If rMessageLen > rShift Then
        '        MsgBox("MsgPrintLogIfEn - Internal Function Error - Message length is too big")
        '    Else
        '        rMessage = "@FG{Black}" & Space(rShift) & "@FG{Black}|" & rMessage & "@BG{White}@FG{Black}" & Space(rShift - rMessageLen) & "@BG{White}@FG{Black}" & "|"
        '    End If

        'End If

        ' **************************************************************************************************************************************
        ' *** SYSTEM LOG - Check Print Contition
        ' **************************************************************************************************************************************

        If rPrintCondition = "ALWAYS" Or (rPrintCondition = "IF ENABLED" And rEnableVarValue = YES) Then 'Or (rResult = FAIL) Then

            ' *** Formatting Message
            'Dim rDotToBeInsered As Integer = 28 - rMessage.Length

            'If rDotToBeInsered < 0 Then MsgBox("Messaggio troppo lungo")
            'rMessage = rMessage & " "
            'For rindex As Integer = 0 To rDotToBeInsered

            'rMessage = rMessage & "."
            'Next
            'rMessage = rMessage & " "
            rMessage = "@FG{Black}BAY" & rBay & " - SITE " & rSite.ToString("00") & " - " & rMessage ' & fCRes(rResult)

            rpMsgLogClass.MsgPrintLogMultyBay(rMessage, rpRpkBayInUse, rResult)

            'rMessageLen = rMessage.Length


            'If rpRpkBayInUse = 1 Then

            '    rMessage = rMessage & "@BG{White}@FG{Black}" & Space(rShift - rMessageLen) & "@BG{White}@FG{Black}" & "|"

            'End If


            'If rpRpkBayInUse = 2 Then

            '    rMessage = "@FG{Black}" & Space(rShift) & "@FG{Black}|" & rMessage & "@BG{White}@FG{Black}" & "|"

            'End If

            'MsgPrintLog(rMessage, rAvoidCrLf)                ' *** If PrintVariable is Enabled


        End If

        ' **************************************************************************************************************************************
        ' *** STORE MEASURE & TEST REPORT - Check Print Contition
        ' **************************************************************************************************************************************

        fMeasureStoreError(rSite, rTestNumbr, rDrawingReference, rTestName, rRemark, rResult, rTl, rMeasuredValue, rTh, rUnitMeasure)

    End Function

    Public Function fMeasureStoreError(ByVal rNsite As Long, ByVal rTestNumber As Long, ByVal rDrawingReference As String,
                                       ByVal rTestName As String, ByVal rRemark As String, ByVal rResult As Long,
                                       Optional ByVal rTl As Double = 0, Optional ByVal rMeasVal As Double = 0, Optional ByVal rTh As Double = 0,
                                       Optional ByVal rUnit As String = "") As Long

        If rResult = FAIL Then

            MsgPrintReport("--------------------------------------------------------", 0)
            MsgPrintReport(" Esito:     @FG{Red}FAIL", 0)

            If rpRpkBayInUse = 1 Then
                MsgPrintReport(" Stage:     In Circuit", 0)                                     ' Print Stage
            Else
                MsgPrintReport(" Stage:     Functional", 0)                                     ' Print Stage
            End If

            If rNsite <> 0 Then MsgPrintReport(" Site:     " & " FIGURA " & rNsite, 0) ' Print Site

            If rTestNumber <> 0 Then MsgPrintReport(" Test N:    " & rTestNumber, 0) ' Print Test Number

            If rTestName <> "" Then MsgPrintReport(" Test Name: " & rTestName, 0) ' Print Test Number

            If rRemark <> "" Then MsgPrintReport(" Remark:    " & rRemark, 0) ' Print Message

            If rMeasVal <> -1000000 Then

                MsgPrintReport("--------------------------------------------------------", 0)

                If rMeasVal = 1000000 Then
                    MsgPrintReport(" Misura: OVERRANGE[+]", 0)                                          ' Print Overrange if MeasValue = 10000
                Else
                    MsgPrintReport(" Misura:    " & Math.Round(rMeasVal, 2) & rUnit, 1)                                ' Print Measured Value
                End If

            End If

            If (rTl <> -1000000) Then MsgPrintReport("     [ Tl: " & rTl & rUnit & " ]", 1) ' Print High Threshold

            If (rTh <> -1000000) Then
                MsgPrintReport("   [ Th: " & rTh & rUnit & " ]", 0)                                     ' Print Low Threshold
            Else
                MsgPrintReport("", 0)
            End If

            MsgPrintReport("--------------------------------------------------------", 0)

        End If

        ' **************************************************************************************************************************************
        ' *** Itac Store Fail
        ' **************************************************************************************************************************************
        If rResult = FAIL Then
            If rpITACIsOnline = True And rpIniITAC_Enable = "YES" Then

                sCollectITACFailure(rNsite - 1, "0", rpDrawRef, rpTestName, rpDiagRemark)

            End If
        End If

        ' **************************************************************************************************************************************
        ' *** CDCOLL Store Measure
        ' **************************************************************************************************************************************

        ' *** Check Paramiters for Qsoft
        If rDrawingReference = "" Then
            MsgDispService(" - DRAWING REFERENCE:" & rDrawingReference & ", NON COMPATIBILE QSOFT - ", 1)
        ElseIf Len(rDrawingReference) > 6 Then
            MsgDispService(" - DRAWING REFERENCE:" & rDrawingReference & ", NON COMPATIBILE QSOFT - ", 1)
        End If

        If rTestNumber < 1 Then MsgDispService(" - NUMERO DI TEST:" & rTestNumber & ", NON COMPATIBILE QSOFT - ", 1)

        If rUnit <> "" Then
            Select Case rUnit
                Case "V"
                Case "A"
                Case "Hz"
                Case "Ohm"
                Case "S"
                Case "F"
                Case "H"
                Case Else
                    MsgDispService(" - UNITA' DI MISURA:" & rUnit & ", NON COMPATIBILE QSOFT - ", 1)
            End Select
        End If

        If rNsite > 12 Then MsgDispService(" - OVERFLOW UNITÁ, NON COMPATIBILE QSOFT - ", 1)

        ' *******************************************************************************************************************************************

        Dim rMeasValue_ToStore As String = Nothing

        If rMeasVal = -1000000 Then
            rMeasValue_ToStore = ""
        Else
            rMeasValue_ToStore = CStr(rMeasVal) & rUnit
        End If

        Dim rTh_ToStore As String = Nothing

        If rTh = -1000000 Then
            rTh_ToStore = ""
        Else
            rTh_ToStore = CStr(rTh) & rUnit
        End If

        Dim rTl_ToStore As String = Nothing

        If rTl = -1000000 Then
            rTl_ToStore = ""
        Else
            rTl_ToStore = CStr(rTl) & rUnit
        End If

        ' *** Send Error String To QSoft Server
        'MsgBox("FCT-" & rTestNumber & "-" & rDrawingReference & "-" & rTestName & "-" & rResult & "-" & rMeasValue_ToStore & "-" & rTl_ToStore & "-" & rTh_ToStore & "-" & rNsite)
        ObsDatalogTest(rTestNumber, rDrawingReference, rTestName, rResult, rMeasValue_ToStore, rTl_ToStore, rTh_ToStore, "0", rNsite + 12)
        'FuncDatalogTest(1, 1, 1, rTestNumber, rDrawingReference, rTestName, rResult, rMeasValue_ToStore, rTl_ToStore, rTh_ToStore, "0", rNsite)
        'MsgBox("DESTROY FCT")

        'Dim rpSerialNumberFct As String = Space(20)

        ' *** Get Serial Number Actual Site
        'Call SerialNumberRead(rNsite, rpSerialNumberFct)

        'If rpPrintErrorOnFile = YES Then

        '    sFileTxtWrite(Now & " - Section OBP/FCT - Site: " & rNsite & " - SN:" & CStr(rpSerialNumberFct) & " - " & rTestNumber & " - " & rTestType & " - " & rMessage & _
        '                    "-MEAS:" & CStr(rMeasVal) & rUnit & "-Tl:" & CStr(rTl) & rUnit & "-Th:" _
        '                    & CStr(rTh) & rUnit, "Configuration\ErrorList[" & rNsite & "] .txt")
        'End If

    End Function

    'Public Sub sPrintSpecFailValueQsoft(ByVal pDrawingRef As String, ByVal pMessage As String, ByVal pTestNumberQsoft As Long,
    '                                    Optional ByVal pSiteNumber As Long = 1, Optional ByVal pSiteOffset As Long = 0, Optional ByVal pExpectValue As String = "",
    '                                    Optional ByVal pReceiveValue As String = "")

    '    'Routine name     : sPrintSpecFailValueQsoft
    '    'Description      : Stampa del messaggio di Fail e Qsoft datalogging
    '    'Input            :
    '    '                  pDrawingRef = deve essere un componente (U1, IC5, D8, R11) oppure la fase (OBP, FCT),
    '    '                                deve essere al max lungo 6 caratteri.
    '    '                  pMessage = messaggio di errore, deve essere al max lungo 40 caratteri.
    '    '                  pTestNumberQsoft = numero di test: deve essere maggiore dell'ultimo numero presente
    '    '                                     nel test plan analogico, sempre diverso e coerente col messaggio.
    '    '                                     es: SI -> 1001 = errore tx, 1002 = errore rx; NO -> 1001 = errore tx, 1001 = errore rx
    '    '                  pSiteNumber = numero di uut sotto test, non deve essere 0 e neanche maggiore dell'ultima uut.
    '    '                                es: pannello di 4 figure, non deve essere 5
    '    '                  pSiteOffset = offset figura da usare nel caso di dual stage.
    '    '                                es: pannello di 4 figure, se fail stage A pSiteOffset deve essere 0,
    '    '                                se fail stage B pSiteOffset deve essere 4.
    '    'Release          : 1.00
    '    'Word originator  : Zirpoli C.

    '    Dim rSpace As Integer
    '    Dim rPrimi40Carat As String

    '    If pDrawingRef = "" Then
    '        MsgDispService(" - DRAWING REFERENCE:" & pDrawingRef & ", NON COMPATIBILE QSOFT - ", 1)
    '    ElseIf Len(pDrawingRef) > 6 Then
    '        MsgDispService(" - DRAWING REFERENCE:" & pDrawingRef & ", NON COMPATIBILE QSOFT - ", 1)
    '    End If

    '    If pTestNumberQsoft < 1 Then MsgDispService(" - NUMERO DI TEST:" & pTestNumberQsoft & ", NON COMPATIBILE QSOFT - ", 1)

    '    MsgDispService(" ", 0)

    '    'se il messaggio supera i 40 caratteri lo mando a capo senza spezzare le parole
    '    If Len(pMessage) > 40 Then
    '        rPrimi40Carat = Left(pMessage, 40)
    '        If Right(rPrimi40Carat, 1) <> " " Then
    '            rSpace = InStrRev(rPrimi40Carat, " ")
    '            MsgDispService(Left(pMessage, rSpace), 0)
    '            If pMessage Like "*G{*" Then
    '                MsgDispService(Left(pMessage, 10) & Right(pMessage, Len(pMessage) - rSpace), 0)
    '            Else
    '                MsgDispService(Right(pMessage, Len(pMessage) - rSpace), 0)
    '            End If
    '        Else
    '            MsgDispService(Left(pMessage, 40), 0)
    '            If pMessage Like "*G{*" Then
    '                MsgDispService(Left(pMessage, 10) & Right(pMessage, Len(pMessage) - 40), 0)
    '            Else
    '                MsgDispService(Right(pMessage, Len(pMessage) - 40), 0)
    '            End If
    '        End If
    '    Else
    '        MsgDispService(pMessage, 0)
    '    End If

    '    If pExpectValue <> "" Then MsgDispService("Expect Value: " _
    '        + "'" + pExpectValue + "'", 0)

    '    If pReceiveValue <> "" Then MsgDispService("Receive Value: " _
    '        + "'" + pReceiveValue + "'", 0)

    '    If (pExpectValue <> "") Or (pReceiveValue <> "") Then _
    '     pMessage = pMessage & " - Expect Value: " & pExpectValue & "Receive Value: " & pReceiveValue

    '    'MsgDispService String(40, "-"), 0

    '    If Len(pMessage) > 40 Then pMessage = Left(pMessage, 40)

    '    If pSiteNumber = 0 Then
    '        UseSiteRead(pSiteNumber)
    '        If pSiteNumber = 0 Then pSiteNumber = 1
    '    End If

    '    pSiteNumber = pSiteNumber + pSiteOffset

    '    ObsDatalogTest(pTestNumberQsoft, pDrawingRef, pMessage, FAIL, "0", "0", "0", "0", pSiteNumber)

    'End Sub

    Public Function fResultPrintDebug(ByVal rUutSerialNumber() As String, ByVal rUutResult() As Long) As Long

        If rpVbNetIsDebug Then rpIniPrintInterTestResult = 1

        ' *** Print Panel Status Before Test

        For rIndex As Integer = 0 To rUutSerialNumber.Length - 1

            'MsgPrintLogIfEn("@BG{Yellow}SITE " & rIndex + 1 & " - Serial Number: " & rUutSerialNumber(rIndex) & " - " & fCRes(rUutResult(rIndex)), 0, "IF ENABLED", rpIniPrintInterTestResult)
            rpMsgLogClass.MsgPrintLogMultyBay("@BG{Yellow}SITE " & (rIndex + 1).ToString("00") & " - Serial Number: " & rUutSerialNumber(rIndex), rpRpkBayInUse, rUutResult(rIndex))
        Next

    End Function

    Public Function fResultPrint(ByVal rUutSerialNumber() As String, ByVal rUutResult() As Long) As Long

    ' If rpVbNetIsDebug Then rpIniPrintInterTestResult = 1

    ' *** Print Panel Status Before Test

    For rIndex As Integer = 0 To rUutSerialNumber.Length - 1

            'MsgPrintLogIfEn("SITE " & rIndex + 1 & " - Serial Number: " & rUutSerialNumber(rIndex) & " - " & fCRes(rUutResult(rIndex)), 0, "ALWAYS", rpIniPrintInterTestResult)
            rpMsgLogClass.MsgPrintLogMultyBay("SITE " & (rIndex + 1).ToString("00") & " - Serial Number: " & rUutSerialNumber(rIndex), rpRpkBayInUse, rUutResult(rIndex))
        Next

End Function
End Module

