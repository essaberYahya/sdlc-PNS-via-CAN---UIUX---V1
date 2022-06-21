Module mod99_RunPackUtility_v0_1

    'Function fRpkResultConvert(ByVal rResultToBeConvert As Long) As String

    '    Select rResultToBeConvert

    '        Case -1
    '            fRpkResultConvert = "NONE"
    '        Case 0
    '            fRpkResultConvert = "PASS"
    '        Case 1
    '            fRpkResultConvert = "FAIL"
    '        Case Else
    '            fRpkResultConvert = "ERRORE"
    '    End Select

    'End Function

    Function fRpkGetStageConfiguration(ByRef rIsDualStage As String) As Long

        fRpkGetStageConfiguration = FAIL

        Dim rTPGM_Path As String
        Dim aTPGM_Path(10) As String

        rTPGM_Path = Application.StartupPath
        aTPGM_Path = Split(rTPGM_Path, "\", 8, CompareMethod.Text)
        rTPGM_Path = Nothing

        For rIndex As Integer = 0 To 6
            rTPGM_Path = rTPGM_Path & aTPGM_Path(rIndex) & "\"
        Next

        rIsDualStage = fGetStringFromINIFile("ReceiverILR", "Dual_Site", rTPGM_Path & "TPGMCFG.INI")

        If rIsDualStage = "0" Then
            rIsDualStage = "SINGLE STAGE"
        ElseIf rIsDualStage = "1" Then
            rIsDualStage = "DUAL STAGE"
        End If


    End Function


    ''' <summary>
    ''' Riceve da RunPack lo stato di abilitazione del SN
    ''' </summary>
    ''' <param name="rSerialNumberIsEnable">[String]Stato di abilitazione del SerialNumber[SI/NO]</param>
    ''' <returns></returns>
    ''' Ritorna una variabile che contiene lo stato di abilitazione del SerialNumber
    ''' <remarks></remarks>

    Function fRpkSerialNumberIsEnable(ByRef rSerialNumberIsEnable As String) As Long

        fRpkSerialNumberIsEnable = FAIL

        Dim rTPGM_Path As String
        Dim aTPGM_Path(10) As String

        rTPGM_Path = Application.StartupPath
        aTPGM_Path = Split(rTPGM_Path, "\", 8, CompareMethod.Text)
        rTPGM_Path = Nothing

        For rIndex As Integer = 0 To 6
            rTPGM_Path = rTPGM_Path & aTPGM_Path(rIndex) & "\"
        Next

        rSerialNumberIsEnable = fGetStringFromINIFile("SerialNumber", "SN", rTPGM_Path & "TPGMCFG.INI")

    End Function

    ''' <summary>
    ''' Riceve da RunPack il numero di Bay del TPGM
    ''' </summary>
    ''' <param name="rBayNumber">[Integer]Numero di bay presenti nel TPGM</param>
    ''' <returns></returns>
    ''' Ritorna una variabile che contiene il numero di Bay settati nel TPGM
    ''' <remarks></remarks>

    Function fRpkGetBayNumber(ByRef rBayNumber As Integer) As Long

        fRpkGetBayNumber = FAIL

        Dim rTPGM_Path As String
        Dim aTPGM_Path(10) As String

        rTPGM_Path = Application.StartupPath
        aTPGM_Path = Split(rTPGM_Path, "\", 8, CompareMethod.Text)
        rTPGM_Path = Nothing

        For rIndex As Integer = 0 To 6
            rTPGM_Path = rTPGM_Path & aTPGM_Path(rIndex) & "\"
        Next

        rBayNumber = fGetStringFromINIFile("ParallelTest", "SysRackUsed", rTPGM_Path & "TPGMCFG.INI")


    End Function

    ''' <summary>
    ''' Riceve da RunPack il SerialNumebr del site richiesto
    ''' </summary>
    ''' <param name="pUseSite">[Integer]Site del quale si desidera leggere il SerialNumber</param>
    ''' <returns></returns>
    ''' Ritorna una stringa contenente il SerialNumber
    ''' <remarks></remarks>

    Function fRpkSerialNumberRead(pUseSite As Integer) As String

        Dim rUUTSerialNumber As String
        Dim rUUTSerialNumberLength As Integer

        fRpkSerialNumberRead = "Serial number not read [Error code: d10m3rd4]"

        ' --- Get UUT SN
        rUUTSerialNumber = Space(64)
        rUUTSerialNumberLength = SerialNumberRead(pUseSite, rUUTSerialNumber)
        rUUTSerialNumber = Left(rUUTSerialNumber, rUUTSerialNumberLength)

        fRpkSerialNumberRead = rUUTSerialNumber

    End Function

    ''' <summary>
    ''' Riceve da RunPack il numero di schede presenti nel pannello
    ''' </summary>
    ''' <returns></returns>
    ''' Ritorna una valore long che indica se la funzione é stata eseguita correttamente
    ''' <remarks></remarks>

    Public Function fRpkGetNumberOfBoard(ByRef rBoardNumber As Integer) As Long

        fRpkGetNumberOfBoard = FAIL

        Dim rTPGM_Path As String
        Dim aTPGM_Path(10) As String

        rTPGM_Path = Application.StartupPath
        aTPGM_Path = Split(rTPGM_Path, "\", 8, CompareMethod.Text)
        rTPGM_Path = Nothing

        For rIndex As Integer = 0 To 6
            rTPGM_Path = rTPGM_Path & aTPGM_Path(rIndex) & "\"
        Next

        rBoardNumber = fGetStringFromINIFile("PanelOfBoard", "BoardNum", rTPGM_Path & "PANEL.INI")

    End Function

    Public Function fRpkGetSiteListOnBay(ByVal rBayNumber As Integer, ByRef rSiteList(,) As Integer) As Long

        Dim rTPGM_Path As String
        Dim aTPGM_Path(10) As String
        Dim rBoardNumber As Integer = 0

        ReDim rSiteList(rBayNumber, 100)

        fRpkGetSiteListOnBay = FAIL

        rTPGM_Path = Application.StartupPath
        aTPGM_Path = Split(rTPGM_Path, "\", 8, CompareMethod.Text)
        rTPGM_Path = Nothing

        For rIndex As Integer = 0 To 6
            rTPGM_Path = rTPGM_Path & aTPGM_Path(rIndex) & "\"
        Next

        rBoardNumber = fGetStringFromINIFile("PanelOfBoard", "BoardNum", rTPGM_Path & "PANEL.INI")

        Dim i As Integer
        Dim Index_Max As Integer

        For rBayCounter As Integer = 1 To rBayNumber

            If Index_Max < i Then Index_Max = i ' *** Storo il numero massimo di site in un bay

            i = 0

            For rIndex As Integer = 1 To rBoardNumber

                If fGetStringFromINIFile("PanelOfBoard", "BAY" & CStr(rIndex), rTPGM_Path & "PANEL.INI") = rBayCounter Then
                    rSiteList(rBayCounter, rIndex) = rIndex
                    i = i + 1
                End If
            Next

        Next

        ReDim Preserve rSiteList(rBayNumber, Index_Max - 1)

    End Function

End Module
