'Module Release          : 1.3
'Last Update             : 03/04/2015

Option Explicit On

Imports AtosF
Imports System
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Globalization

Module mod10_UtilityRoutine



    Public Declare Function GetTickCount Lib "kernel32" () As Int32

    Public Function fCheckTotalResult(ByVal rFase As String) As Long

        fCheckTotalResult = PASS

        Dim rpResultBay1 As Integer = fGetStringFromINIFile(rFase, "BAY1_RESULT", "Configuration\TempResult.INI")
        Dim rpResultBay2 As Integer = fGetStringFromINIFile(rFase, "BAY2_RESULT", "Configuration\TempResult.INI")

        If (rpCurrentBay = 1) And (rpResultBay1 <> FAIL) Then

            If (rpResultBay2 = FAIL) Then

                fCheckTotalResult = FAIL

            End If

        ElseIf (rpCurrentBay = 2) And (rpResultBay2 <> FAIL) Then

            If (rpResultBay1 = FAIL) Then

                fCheckTotalResult = FAIL

            End If

        End If

    End Function

    Function HexStringToBinary(ByVal hexString As String) As String
        Dim num As Long = Long.Parse(hexString, NumberStyles.HexNumber)
        Return Convert.ToString(num, 2)
    End Function

    Function fCalcCapVoltageAtTime(ByVal rSupplyCapVoltage As Double, ByVal rT As Double, _
                               ByVal rTau As Double, ByRef rCapVoltage As Double) As Long

        fCalcCapVoltageAtTime = FAIL

        Try
            rCapVoltage = rSupplyCapVoltage * (1 - (Math.E ^ (-(rT / rTau))))
            fCalcCapVoltageAtTime = PASS

        Catch ex As Exception


        End Try

    End Function

    Function fCalcTau(ByVal rSupplyVoltage As Double, ByVal rT As Double, _
                       ByRef rTau As Double, ByVal rCapVoltage As Double) As Long

        fCalcTau = FAIL

        Try
            rTau = -rT / (Math.Log(1 - (rCapVoltage / (rSupplyVoltage))))
            fCalcTau = PASS

        Catch ex As Exception

        End Try

    End Function

    Function fCalcCapVoltageAtTimeDiscarge(ByVal rCapVoltageStart As Double, ByVal rT As Double, _
                               ByVal rTau As Double, ByRef rCapVoltage As Double) As Long

        fCalcCapVoltageAtTimeDiscarge = FAIL

        Try
            rCapVoltage = rCapVoltageStart * (Math.E ^ (-(rT / rTau)))
            fCalcCapVoltageAtTimeDiscarge = PASS

        Catch ex As Exception


        End Try

    End Function

    Function fCalcTauDischarge(ByVal rCapVoltageStart As Double, ByVal rCapVoltageEnd As Double, ByVal rT As Double, _
                       ByRef rTau As Double) As Long

        fCalcTauDischarge = FAIL

        Try
            rTau = -rT / (Math.Log(rCapVoltageEnd / rCapVoltageStart))
            fCalcTauDischarge = PASS

        Catch ex As Exception

        End Try

    End Function

    Function fConv_BinToDec(ByVal pNumeroBinario As String) As Double

        Dim a As Double
        Dim rNumDec As Double = 0
        Dim rNumBin As String = ""

        ' inverto la stringa
        For a = 1 To Len(pNumeroBinario)
            rNumBin = Mid(pNumeroBinario, a, 1) & rNumBin
        Next

        For a = 0 To Len(pNumeroBinario) - 1
            rNumDec = 2 ^ a * CInt(Mid(rNumBin, a + 1, 1)) + rNumDec
        Next

        fConv_BinToDec = rNumDec

    End Function

    Function fConv_DecToHex(ByVal pDecNum As Double) As String

        'Routine name     : fConversionDecToHex
        'Description      : conversion from decimal to hexadecimal
        'Input            : Decimal number to convert
        'Release          : 2.30
        'Word originator  : Cristian Zirpoli

        Dim HighPart As Double
        Dim HexDouble As String
        Dim Padding As Integer
        Dim rByte(12) As String
        Dim rIndex As Integer
        Dim rTwo As Integer

        If (Len(CStr(pDecNum))) > 4 Then
            pDecNum = Fix(pDecNum)
            HighPart = Int(pDecNum / 65536)
            If HighPart < 0 Then HighPart = HighPart - 1
            pDecNum = pDecNum - HighPart * 65536
            HexDouble = Hex$(HighPart) & Right$("0000" & Hex$(pDecNum), 4)
            If Padding > 0 Then HexDouble = Right$(HexDouble, Padding)
        Else
            HexDouble = Hex(pDecNum)
        End If

        'se la lunghezza della risposta e` pari
        If Len(HexDouble) \ 2 = Len(HexDouble) / 2 Then
            'scompongo la risposta esadec. in byte
            rIndex = 1
            For rTwo = Len(HexDouble) - 1 To 0 Step -2
                rByte(rIndex) = Mid(HexDouble, Len(HexDouble) - rTwo, 2)
                rIndex = rIndex + 1
            Next
            fConv_DecToHex = HexDouble
        Else
            HexDouble = "0" & HexDouble
            fConv_DecToHex = HexDouble
        End If

    End Function

    Function fConv_HexToDec(ByVal pHexNum As String) As Long

        Dim rByte As Long
        Dim rByteStr As String
        Dim rIndex As Integer
        Dim rHexNumNoZero As Boolean = False
        Dim rNumeroDecimale As Long = 0

        'se i numeri da convertire sono solo zeri esco, es: "00"
        For rByte = 1 To Len(pHexNum)
            If Mid(pHexNum, rByte, 1) <> "0" Then
                rHexNumNoZero = True
                Exit For
            End If
        Next
        If rHexNumNoZero = False Then Return (0)

        For rIndex = Len(pHexNum) To 1 Step -1
            If IsNumeric(Mid(pHexNum, rIndex, 1)) = False Then
                rByteStr = Mid(pHexNum, rIndex, 1)
                Select Case UCase(rByteStr)
                    Case "A"
                        rByte = 10
                    Case "B"
                        rByte = 11
                    Case "C"
                        rByte = 12
                    Case "D"
                        rByte = 13
                    Case "E"
                        rByte = 14
                    Case "F"
                        rByte = 15
                End Select
                rNumeroDecimale = CLng(rNumeroDecimale + (16 ^ (Len(pHexNum) - rIndex) * rByte))
            Else
                rByte = CLng(Val(Mid(pHexNum, rIndex, 1)))
                If rByte <> 0 Then rNumeroDecimale = CLng(rNumeroDecimale + (16 ^ (Len(pHexNum) - rIndex) * rByte))
            End If
        Next

        Return (rNumeroDecimale)

    End Function

    Public Function fString2Hex(ByVal pInputString As String, ByRef pOutputString As String) As Integer

        Dim i As Integer
        Dim rHexChar As String
        Dim rConvString As String = Nothing
        For i = 1 To Len(pInputString)
            rHexChar = Hex(Asc(Mid(pInputString, i, 1)))
            If Len(rHexChar) < 2 Then
                rHexChar = "0" & rHexChar
            End If
            rConvString = rConvString & rHexChar
        Next
        pOutputString = rConvString
        Return (0)

    End Function

    Function fPariDispari(ByVal pNumero As Double) As Integer

        If (pNumero / 2) = (pNumero \ 2) Then     'numero pari
            Return (1)
        Else                                      'numero dispari
            Return (0)
        End If

    End Function

    Function fSerialNumberRead(ByVal pUseSite As Long) As String

        Dim rSerialNumberLen As Long
        Dim rSerialNumber As String

        'acquisizione e conversione del serial number
        rSerialNumber = Space(100)
        rSerialNumberLen = SerialNumberRead(pUseSite, rSerialNumber)
        fSerialNumberRead = Left$(rSerialNumber, rSerialNumberLen)

    End Function

    Function fRitornaPercorso(ByVal pPercorsoDaTagliare As String, ByVal pNumeroDiSlash As Integer, _
                              ByVal pRitornaParteDiSinistra As Boolean) As String

        'funzione di accorciamento percorso. Es: "c:\$spea\wksinfo\prjsel.ini"

        Dim rSlash As Integer
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer

        b = 1
        c = 1
        If pRitornaParteDiSinistra Then
            'ritorna la parte di sinistra. Es: "c:\$spea"
RestartLeft:
            rSlash = InStr(b, pPercorsoDaTagliare, "\")
            If c <> pNumeroDiSlash Then
                b = b + (rSlash - b) + 1
                c = c + 1
                GoTo RestartLeft
            End If
            If rSlash = 0 Then Return (pPercorsoDaTagliare)
            pPercorsoDaTagliare = Microsoft.VisualBasic.Left(pPercorsoDaTagliare, rSlash - 1)
        Else
            'ritorna la parte di destra. Es: "wksinfo\prjsel.ini"
            For a = 1 To pNumeroDiSlash
                rSlash = InStr(pPercorsoDaTagliare, "\")
                pPercorsoDaTagliare = Microsoft.VisualBasic.Right(pPercorsoDaTagliare, Len(pPercorsoDaTagliare) - rSlash)
            Next
        End If

        Return (pPercorsoDaTagliare)

    End Function

    Public Sub sPrintFailValue(ByVal pMessage As String, Optional ByVal pThresholdHighMeasure As Double = 0, _
                               Optional ByVal pThresholdLowMeasure As Double = 0, Optional ByVal pMeasureValue As Double = 0, Optional ByVal pUnitMeasure As String = "V")

        'Routine name     : sPrintFailValue
        'Description      : Print Fail Measure
        'Input            : Message to show, Measure Value, High Threshold of Measure,
        'Low Threshold of Measure, Unit Measure
        'Release          : 1.00
        'Word originator  : Olivetti Mauro

        Dim rSpace As Integer
        Dim rPrimi40Carat As String
        Dim rThresholdHighMeasure As String
        Dim rThresholdLowMeasure As String
        Static rTestNumber As Long

        MsgDispService(" ", 0)

        'se il messaggio supera i 40 caratteri lo mando a capo senza spezzare le parole
        If Len(pMessage) > 40 Then
            rPrimi40Carat = Left(pMessage, 40)
            If Right(rPrimi40Carat, 1) <> " " Then
                rSpace = InStrRev(rPrimi40Carat, " ")
                MsgDispService(Left(pMessage, rSpace), 0)
                MsgDispService(Right(pMessage, Len(pMessage) - rSpace), 0)
            Else
                MsgDispService(Left(pMessage, 40), 0)
                MsgDispService(Right(pMessage, Len(pMessage) - 40), 0)
            End If
        Else
            MsgDispService(pMessage, 0)
        End If

        rThresholdLowMeasure = CStr(pThresholdLowMeasure)
        rThresholdHighMeasure = CStr(pThresholdHighMeasure)

        If (rThresholdHighMeasure <> "") Or (rThresholdLowMeasure <> "") Then
            MsgDispService("Threshold: " + "Min " & rThresholdLowMeasure & " / " & "Max " & rThresholdHighMeasure _
            + " " + pUnitMeasure, 0)

            MsgDispService("Measured Value: " + CStr(pMeasureValue) + " " + pUnitMeasure, 0) '+ " /" + Str(SiteinUse), 0
        End If


        ObsDatalogTest(rTestNumber, Left(pMessage, 6), pMessage, FAIL, CStr(pMeasureValue), CStr(rThresholdLowMeasure), _
         CStr(rThresholdHighMeasure), "0", 1)
        rTestNumber = rTestNumber + 1

    End Sub

    Public Sub sPrintSpecFailValue(ByVal pMessage As String, Optional ByVal pExpectValue As String = "", _
                                   Optional ByVal pReceiveValue As String = "")

        'Routine name     : sPrintSpecFailValue
        'Description      : Print Fail Measure
        'Input            : Message to show, Expect  Value, Receive Value
        'Release          : 1.00
        'Word originator  : Cristian Zirpoli

        Dim rPrimi40Carat As String
        Dim rSpace As Integer

        MsgDispService(" ", 0)

        'se il messaggio supera i 40 caratteri lo mando a capo senza spezzare le parole
        If Len(pMessage) > 40 Then
            rPrimi40Carat = Left(pMessage, 40)
            If Right(rPrimi40Carat, 1) <> " " Then
                rSpace = InStrRev(rPrimi40Carat, " ")
                MsgDispService(Left(pMessage, rSpace), 0)
                MsgDispService(Right(pMessage, Len(pMessage) - rSpace), 0)
            Else
                MsgDispService(Left(pMessage, 40), 0)
                MsgDispService(Right(pMessage, Len(pMessage) - 40), 0)
            End If
        Else
            MsgDispService(pMessage, 0)
        End If

        If pExpectValue <> "" Then MsgDispService("Expect  Value: " _
            + "'" + pExpectValue + "'", 0)

        If pReceiveValue <> "" Then MsgDispService("Receive Value: " _
            + "'" + pReceiveValue + "'", 0)

        'MsgDispService String(40, "-"), 0

    End Sub

    Public Sub sPrintSpecFailValueQsoft(ByVal pDrawingRef As String, ByVal pMessage As String, ByVal pTestNumberQsoft As Long, _
                                        Optional ByVal pSiteNumber As Long = 1, Optional ByVal pSiteOffset As Long = 0, Optional ByVal pExpectValue As String = "", _
                                        Optional ByVal pReceiveValue As String = "")

        'Routine name     : sPrintSpecFailValueQsoft
        'Description      : Stampa del messaggio di Fail e Qsoft datalogging
        'Input            :
        '                  pDrawingRef = deve essere un componente (U1, IC5, D8, R11) oppure la fase (OBP, FCT),
        '                                deve essere al max lungo 6 caratteri.
        '                  pMessage = messaggio di errore, deve essere al max lungo 40 caratteri.
        '                  pTestNumberQsoft = numero di test: deve essere maggiore dell'ultimo numero presente
        '                                     nel test plan analogico, sempre diverso e coerente col messaggio.
        '                                     es: SI -> 1001 = errore tx, 1002 = errore rx; NO -> 1001 = errore tx, 1001 = errore rx
        '                  pSiteNumber = numero di uut sotto test, non deve essere 0 e neanche maggiore dell'ultima uut.
        '                                es: pannello di 4 figure, non deve essere 5
        '                  pSiteOffset = offset figura da usare nel caso di dual stage.
        '                                es: pannello di 4 figure, se fail stage A pSiteOffset deve essere 0,
        '                                se fail stage B pSiteOffset deve essere 4.
        'Release          : 1.00
        'Word originator  : Zirpoli C.

        Dim rSpace As Integer
        Dim rPrimi40Carat As String

        If pDrawingRef = "" Then
            MsgDispService(" - DRAWING REFERENCE:" & pDrawingRef & ", NON COMPATIBILE QSOFT - ", 1)
        ElseIf Len(pDrawingRef) > 6 Then
            MsgDispService(" - DRAWING REFERENCE:" & pDrawingRef & ", NON COMPATIBILE QSOFT - ", 1)
        End If

        If pTestNumberQsoft < 1 Then MsgDispService(" - NUMERO DI TEST:" & pTestNumberQsoft & ", NON COMPATIBILE QSOFT - ", 1)

        MsgDispService(" ", 0)

        'se il messaggio supera i 40 caratteri lo mando a capo senza spezzare le parole
        If Len(pMessage) > 40 Then
            rPrimi40Carat = Left(pMessage, 40)
            If Right(rPrimi40Carat, 1) <> " " Then
                rSpace = InStrRev(rPrimi40Carat, " ")
                MsgDispService(Left(pMessage, rSpace), 0)
                If pMessage Like "*G{*" Then
                    MsgDispService(Left(pMessage, 10) & Right(pMessage, Len(pMessage) - rSpace), 0)
                Else
                    MsgDispService(Right(pMessage, Len(pMessage) - rSpace), 0)
                End If
            Else
                MsgDispService(Left(pMessage, 40), 0)
                If pMessage Like "*G{*" Then
                    MsgDispService(Left(pMessage, 10) & Right(pMessage, Len(pMessage) - 40), 0)
                Else
                    MsgDispService(Right(pMessage, Len(pMessage) - 40), 0)
                End If
            End If
        Else
            MsgDispService(pMessage, 0)
        End If

        If pExpectValue <> "" Then MsgDispService("Expect Value: " _
            + "'" + pExpectValue + "'", 0)

        If pReceiveValue <> "" Then MsgDispService("Receive Value: " _
            + "'" + pReceiveValue + "'", 0)

        If (pExpectValue <> "") Or (pReceiveValue <> "") Then _
         pMessage = pMessage & " - Expect Value: " & pExpectValue & "Receive Value: " & pReceiveValue

        'MsgDispService String(40, "-"), 0

        If Len(pMessage) > 40 Then pMessage = Left(pMessage, 40)

        If pSiteNumber = 0 Then
            UseSiteRead(pSiteNumber)
            If pSiteNumber = 0 Then pSiteNumber = 1
        End If

        pSiteNumber = pSiteNumber + pSiteOffset

        ObsDatalogTest(pTestNumberQsoft, pDrawingRef, pMessage, FAIL, "0", "0", "0", "0", pSiteNumber)

    End Sub

    Public Sub sWait(ByVal pAttesaTempomsec As Long)

        'Nome della procedura   : sAttendiTempo
        'Descrizione            : Attende il Tempo impostato in msec
        'Parametri              : pAttesaTempomsec   :
        '                                                Numeric Value
        '                                                Unit msec
        'Ritorno della funzione :
        'Release                : 1.50
        'Autore                 : Di Carlo Davide
        'Note                   :

        Dim rStartTempo As Long
        Dim rStopTempo As Long

        rStartTempo = GetTickCount

        Do
            System.Windows.Forms.Application.DoEvents()
            rStopTempo = GetTickCount - rStartTempo
        Loop Until rStopTempo >= pAttesaTempomsec

    End Sub

    Function fFileTxtRead(ByRef pRowsRead() As String, Optional ByVal pPathFile As String = "") As Integer

        Const ForReading As Integer = 1
        'Const ForWriting As Integer = 2
        'Const ForAppending As Integer = 8
        Dim FileSystemObject As Object
        Dim FileTxt As Object
        Dim a As Long

        'apro il file di testo
        FileSystemObject = CreateObject("Scripting.FileSystemObject")
        If pPathFile = "" Then
            FileTxt = FileSystemObject.OpenTextFile(Application.StartupPath & "\FctTrace.txt", ForReading, True)
        Else
            FileTxt = FileSystemObject.OpenTextFile(pPathFile, ForReading, True)
        End If

        'leggo dal file di testo
        Do While FileTxt.AtEndOfStream <> True
            pRowsRead(a) = FileTxt.ReadLine
            a = a + 1
        Loop
        FileTxt.Close()

        Return (a)

    End Function

    Sub sFileTxtWrite(ByVal pRowToWrite As String, Optional ByVal pPathFile As String = "")

        'Const ForReading As Integer = 1
        Const ForWriting As Integer = 2
        Const ForAppending As Integer = 8
        Dim FileSystemObject As Object
        Dim FileTxt As Object

        'apro il file di testo
        FileSystemObject = CreateObject("Scripting.FileSystemObject")
        If pPathFile = "" Then
            FileTxt = FileSystemObject.OpenTextFile(Application.StartupPath & "\FctTrace.txt", ForWriting, True)
        Else
            FileTxt = FileSystemObject.OpenTextFile(pPathFile, ForAppending, True)
        End If
        'scrivo nel file di testo
        FileTxt.Writeline(pRowToWrite)
        FileTxt.Close()

    End Sub

    Sub sFileTxtAppend(ByVal pRowToWrite As String, Optional ByVal pPathFile As String = "")

        'Const ForReading As Integer = 1
        'Const ForWriting As Integer = 2
        Const ForAppending As Integer = 8
        Dim FileSystemObject As Object
        Dim FileTxt As Object

        'apro il file di testo
        FileSystemObject = CreateObject("Scripting.FileSystemObject")
        If pPathFile = "" Then
            FileTxt = FileSystemObject.OpenTextFile(Application.StartupPath & "\Configuration\FctTrace[" & rpCurrentBay & "].txt", ForAppending, True)
        Else
            FileTxt = FileSystemObject.OpenTextFile(pPathFile, ForAppending, True)
        End If
        'scrivo nel file di testo
        FileTxt.Writeline(pRowToWrite)
        FileTxt.Close()

    End Sub

    Sub sFileBinWrite(ByVal pFileName As String, ByVal pData As String)

        Dim i As Integer
        Dim a As Byte
        Dim rFileNum As Integer

        rFileNum = FreeFile()

        FileSystem.FileOpen(rFileNum, pFileName, OpenMode.Binary, OpenAccess.Write)
        For i = 1 To Len(pData)
            a = Asc(Mid(pData, i, 1))
            FileSystem.FilePut(rFileNum, a)
        Next
        FileSystem.FileClose(rFileNum)

    End Sub

    Sub sFileBinRead(ByVal pFileName As String, ByRef pData As String)

        Dim i As Integer
        Dim a As Byte
        Dim rFileNum As Integer
        Dim rFileLen As Integer
        Dim rCarattere As String

        rFileNum = FreeFile()

        FileSystem.FileOpen(rFileNum, pFileName, OpenMode.Binary, OpenAccess.Read)
        rFileLen = FileSystem.LOF(rFileNum)
        For i = 1 To rFileLen
            FileSystem.FileGet(rFileNum, a)
            rCarattere = Chr(a)
            pData = pData & rCarattere
        Next
        FileSystem.FileClose(rFileNum)

    End Sub

    Public Sub sLaunchBatchFile1(ByVal pFileName As String)

        Dim procID As Integer
        Dim newProc As Diagnostics.Process

        'MsgPrintLog("Starting " & pFileName, NO)
        newProc = Diagnostics.Process.Start(pFileName)
        procID = newProc.Id
        'MsgPrintLog("Process ID = " & procID, NO)
        newProc.WaitForExit()

        Dim procEC As Integer = -1

        If newProc.HasExited Then
            procEC = newProc.ExitCode
            'MsgPrintLog("Exit Code = " & procEC, NO)
        End If


        'x funzionare il file batch deve essere cosí:
        ' @echo off

        'path("%programfiles%\Vacon\Dce")

        'cd\
        'cd "c:\Batches\MIURA CB\Eeprom\"

        'EEIMAGE_FILE = CB01227EepromImage.bin
        'IMGBASE_FILE = CB01227EepromImageBase.bin
        'set ORGBASE_FILE=.\Base\%IMGBASE_FILE%

        'if exist %ORGBASE_FILE% (
        '   copy %ORGBASE_FILE% %EEIMAGE_FILE%
        ') else (
        '   copy %IMGBASE_FILE% %EEIMAGE_FILE%
        ')

        'XML_CONFIG = Config_Create_Image.xml
        'XML_BOARD = CB_70CVB01227_Calibration.xml

        'dce config=%XML_CONFIG% batch=%XML_BOARD%

        'if %ERRORLEVEL% == 0 goto OK
        '@echo.
        '@echo Failure while creating image (see log).
        '@echo.
        'pause()
        ': OK()


    End Sub

    Public Sub sLaunchBatchFile2(ByVal pFileName As String)

        Dim rBatchExecute As New Process
        Dim rBatchExecuteInfo As New ProcessStartInfo(pFileName)
        Dim rDirBatch As String
        Dim rSlash As Integer

        Try

            'acquisizione cartella in cui risiede l'eseguibile da lanciare
            rSlash = InStrRev(pFileName, "\")
            rDirBatch = Left(pFileName, rSlash - 1)

            rBatchExecuteInfo.UseShellExecute = True
            rBatchExecuteInfo.CreateNoWindow = False
            rBatchExecuteInfo.WorkingDirectory = rDirBatch
            rBatchExecuteInfo.FileName = pFileName
            rBatchExecute.StartInfo = rBatchExecuteInfo

            rBatchExecute.Start()

            rBatchExecute.WaitForExit(6000)
            'se il batch non si chiude entro il time-out, lo uccido
            If rBatchExecute.HasExited = False Then rBatchExecute.Kill()

        Catch ex As Exception

            MsgBox("sLaunchBatchFile2. Run-Time error: " & ex.Message)

        End Try

        'x funzionare il file batch deve essere cosí:
        ' @echo off
        'C:
        'path("%programfiles%\Vacon\Dce")

        'EEIMAGE_FILE = CB01227EepromImage.bin
        'IMGBASE_FILE = CB01227EepromImageBase.bin
        'set ORGBASE_FILE=.\Base\%IMGBASE_FILE%

        'if exist %ORGBASE_FILE% (
        '   copy %ORGBASE_FILE% %EEIMAGE_FILE%
        ') else (
        '   copy %IMGBASE_FILE% %EEIMAGE_FILE%
        ')

        'XML_CONFIG = Config_Create_Image.xml
        'XML_BOARD = CB_70CVB01227_Calibration.xml

        'dce config=%XML_CONFIG% batch=%XML_BOARD%

        'if %ERRORLEVEL% == 0 goto OK
        '@echo.
        '@echo Failure while creating image (see log).
        '@echo.
        'pause()
        ': OK()

    End Sub

    Public Function fLaunchBatchFile2(ByVal pFileName As String, Optional ByVal pTimeOut As Integer = 10000, Optional ByVal pErrorFile As String = "") As Integer

        Dim rBatchExecute As New Process
        Dim rBatchExecuteInfo As New ProcessStartInfo(pFileName)
        Dim rDirBatch As String
        Dim rSlash As Integer

        Try

            'acquisizione cartella in cui risiede l'eseguibile da lanciare
            rSlash = InStrRev(pFileName, "\")
            rDirBatch = Left(pFileName, rSlash - 1)

            'cancellazione file di errore
            If pErrorFile <> "" Then
                If Dir(rDirBatch & "\" & pErrorFile) <> "" Then
                    Kill(rDirBatch & "\" & pErrorFile)
                    sWait(10)
                End If
            End If

            rBatchExecuteInfo.UseShellExecute = True
            rBatchExecuteInfo.CreateNoWindow = False
            rBatchExecuteInfo.WorkingDirectory = rDirBatch
            rBatchExecuteInfo.FileName = pFileName
            rBatchExecute.StartInfo = rBatchExecuteInfo
#If CONFIG <> "Debug" Then
            rBatchExecute.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
#End If
            rBatchExecute.Start()

            rBatchExecute.WaitForExit(pTimeOut)
            'se il batch non si chiude entro il time-out, lo uccido
            If rBatchExecute.HasExited = False Then
                rBatchExecute.Kill()
                MsgBox("fLaunchBatchFile2: Time-out")
                Return (1)
            End If

            'verifica presenza file di errore
            sWait(10)
            If pErrorFile <> "" Then
                If Dir(rDirBatch & "\" & pErrorFile) <> "" Then
                    MsgBox("fLaunchBatchFile2: FAIL")
                    Return (1)
                End If
            End If

        Catch ex As Exception

            MsgBox("fLaunchBatchFile2. Run-Time error: " & ex.Message)
            Return (1)

        End Try

        Return (0)

    End Function

    Public Function fGetFolders(ByVal pPathToGetFolders As String, ByRef pFoldersGot() As String) As Integer

        Dim Dir As New System.IO.DirectoryInfo(pPathToGetFolders)
        Dim Folders As System.IO.DirectoryInfo() = Dir.GetDirectories()
        Dim i As Integer

        For i = 0 To Folders.Length - 1
            pFoldersGot(i) = Folders(i).Name
        Next i

        Return (Folders.Length)

    End Function

    Public Function fGetDriverFree() As String

        'ricerca prima lettera libera
        Dim rDrivesFreeFound As Boolean = False
        Dim rDrivesAsc As Integer
        Dim rDrives As String
        'Dim getInfo = System.IO.DriveInfo.GetDrives()

        rDrives = ""
        For rDrivesAsc = 68 To 90
            rDrivesFreeFound = True
            rDrives = Chr(rDrivesAsc) & ":\"
            'For Each info In getInfo
            '  If rDrives = info.Name Then
            '    rDrivesFreeFound = False
            '    Exit For
            '  End If
            'Next
            If rDrivesFreeFound = True Then Exit For
            rDrives = ""
        Next

        Return (rDrives)

    End Function

    Public Function fGetMapNetDriver(ByRef pNetDriverFound() As String) As Integer

        Dim rDrivesNum As Integer
        'Dim rDrives As String
        'Dim getInfo = System.IO.DriveInfo.GetDrives()
        'Dim rProcess As New Process()

        'For Each info In getInfo
        '  rDrives = ""
        '  If info.DriveType = IO.DriveType.Network Then
        '    rDrives = info.Name
        '    rDrives = Left(rDrives, 2)

        '    With rProcess.StartInfo
        '      .FileName = "net"
        '      .Arguments = "use " & rDrives
        '      .UseShellExecute = False
        '      .RedirectStandardOutput = True
        '      .CreateNoWindow = True
        '    End With
        '    rProcess.Start()

        '    Dim rBatchRows = rProcess.StandardOutput.ReadToEnd()

        '    rProcess.WaitForExit()
        '    For Each Line In Split(rBatchRows, vbNewLine)
        '      If Line.StartsWith("Remote name") Then
        '        If Line.Contains("rogetto") Then
        '          pNetDriverFound(rDrivesNum) = rDrives
        '          rDrivesNum = rDrivesNum + 1
        '        End If
        '      End If
        '    Next
        '  End If
        'Next

        Return (rDrivesNum)

    End Function

    Public Function fRunDosCommand(ByVal pCommand As String, ByVal pParameters As String) As Integer

        Dim rProcess As New Process()

        With rProcess.StartInfo
            .FileName = pCommand '"net"
            .Arguments = pParameters '"use G:"
            .UseShellExecute = False
            .RedirectStandardOutput = True
            .CreateNoWindow = True
        End With
        rProcess.Start()

        'Dim rBatchRows = rProcess.StandardOutput.ReadToEnd()

        rProcess.WaitForExit()
        'For Each Line In Split(rBatchRows, vbNewLine)
        '  If Line.StartsWith("Remote name") Then Return (0)
        'Next

        Return (1)

    End Function

    Sub sSwapBytes(ByVal pFileNameToSwap As String, ByVal pFileNameSwapped As String)

        Dim rByte(65536) As Byte
        Dim rByteSwap(65536) As Byte
        Dim i As Integer
        Dim rFileNum As Integer
        Dim FileMaxSize As Long
        rFileNum = FreeFile()

        Try

            'Cancella il file swap anche se non esiste.
            My.Computer.FileSystem.DeleteFile(pFileNameSwapped)

            'lettura file sorgente
            FileSystem.FileOpen(rFileNum, pFileNameToSwap, OpenMode.Binary, OpenAccess.Read)
            FileMaxSize = LOF(rFileNum)   ' Legge la dimensione del file in byte.
            For i = 1 To FileMaxSize
                FileSystem.FileGet(rFileNum, rByte(i), i)
            Next
            FileSystem.FileClose(rFileNum)

            'inversione byte
            For i = 1 To FileMaxSize Step 2
                rByteSwap(i) = rByte(i + 1)
                rByteSwap(i + 1) = rByte(i)
            Next

            'scrittura file con byte invertiti
            FileSystem.FileOpen(rFileNum, pFileNameSwapped, OpenMode.Binary, OpenAccess.Write)
            For i = 1 To FileMaxSize
                FileSystem.FilePut(rFileNum, rByteSwap(i))
            Next
            FileSystem.FileClose(rFileNum)

        Catch ex As Exception
        End Try

    End Sub

    Function fKillProcess(ByVal pName As String, Optional ByVal pProcess As Boolean = True) As Integer

        Dim rNumProcessPc As Integer
        Dim i As Long
        Dim rProcess() As Diagnostics.Process

        'array contenete tutti i processi in esecuzione sul pc
        rProcess = Diagnostics.Process.GetProcesses()
        rNumProcessPc = UBound(rProcess, 1)

        If pProcess = True Then
            'chiusura processi 
            For i = 0 To rNumProcessPc
                If UCase(rProcess(i).ProcessName) = UCase(pName) Then
                    rProcess(i).Kill()
                End If
            Next i
        Else
            'chiusura windows
            For i = 1 To rNumProcessPc
                If UCase(rProcess(i).MainWindowTitle) = UCase(pName) Then
                    rProcess(i).Kill()
                End If
            Next i
        End If

        ' verifica chiusura eseguibile
        If pProcess = True Then
            'chiusura processi 
            For i = 0 To rNumProcessPc
                If UCase(rProcess(i).ProcessName) = UCase(pName) Then
                    Return (1)
                End If
            Next i
        Else
            'chiusura windows
            For i = 1 To rNumProcessPc
                If UCase(rProcess(i).MainWindowTitle) = UCase(pName) Then
                    Return (1)
                End If
            Next i
        End If

        Return (0)

    End Function

End Module
