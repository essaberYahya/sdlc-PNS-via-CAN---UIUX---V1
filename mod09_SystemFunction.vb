Module mod09_SystemFUnction

    Public Function fTpConnectAbus(ByVal rTpOnRow1 As Integer, ByVal rTpOnRow2 As Integer,
                                   ByVal rTpOnRow3 As Integer, ByVal rTpOnRow4 As Integer) As Long


        Dim rError As String = Nothing
        Dim rErrorCode As Long = 0

        Dim aTpList(1) As Integer


        fTpConnectAbus = FAIL                                 ' *** Set Default FAIL
        ' **************************************************************************

        Try
            If rTpOnRow1 <> 0 Then
                aTpList(0) = rTpOnRow1
                aTpList(1) = 0
                rErrorCode = TpConnectAbus(aTpList(0), ABUS1)
            End If

            If rTpOnRow2 <> 0 Then
                aTpList(0) = rTpOnRow2
                aTpList(1) = 0
                rErrorCode = TpConnectAbus(aTpList(0), ABUS2)
            End If

            If rTpOnRow3 <> 0 Then
                aTpList(0) = rTpOnRow3
                aTpList(1) = 0
                rErrorCode = TpConnectAbus(aTpList(0), ABUS3)
            End If

            If rTpOnRow4 <> 0 Then
                aTpList(0) = rTpOnRow4
                aTpList(1) = 0
                rErrorCode = TpConnectAbus(aTpList(0), ABUS4)
            End If

            fTpConnectAbus = PASS                                 ' *** Set Default PASS
            ' **************************************************************************

        Catch ex As Exception
            rError = ex.Message
        End Try

    End Function

    Public Function fTpDisconnectAbus(ByVal rTpOnRow1 As Integer, ByVal rTpOnRow2 As Integer,
                               ByVal rTpOnRow3 As Integer, ByVal rTpOnRow4 As Integer) As Long


        Dim rError As String = Nothing
        Dim rErrorCode As Long = 0

        fTpDisconnectAbus = FAIL                              ' *** Set Default FAIL
        ' **************************************************************************

        Try
            If rTpOnRow1 <> 0 Then
                rErrorCode = TpDisconnectAbus(rTpOnRow1, ABUS1)
            End If

            If rTpOnRow2 <> 0 Then
                rErrorCode = TpDisconnectAbus(rTpOnRow2, ABUS2)
            End If

            If rTpOnRow3 <> 0 Then
                rErrorCode = TpDisconnectAbus(rTpOnRow3, ABUS3)
            End If

            If rTpOnRow4 <> 0 Then
                rErrorCode = TpDisconnectAbus(rTpOnRow4, ABUS4)
            End If

            fTpDisconnectAbus = PASS                              ' *** Set Default PASS
            ' **************************************************************************

        Catch ex As Exception
            rError = ex.Message
        End Try

    End Function

    Public Function fCRes(ByVal rResult As Long) As String

        Select Case rResult
            Case 0
                fCRes = "@BG{Green}@FG{White}PASS"
            Case 1
                fCRes = "@BG{Red}@FG{White}FAIL"
            Case 2
                fCRes = "@BG{Purple}@FG{White}RETR"
            Case -1
                fCRes = "@BG{Blue}@FG{White}NONE"
            Case Else
                fCRes = "@BG{PURPLE}@FG{White}VALUE NOT ALLOWED"
        End Select

    End Function

    Public Function S2A_Integer(ByRef rStringtoBeConvert As String, ByVal rArray() As Integer, ByVal rMaxElementNumber As Integer) As Integer

        S2A_Integer = FAIL

        ReDim rArray(rMaxElementNumber)

        Dim rarraystr() As String

        rarraystr = Split(rStringtoBeConvert, ",", rMaxElementNumber, CompareMethod.Text)

        For rindex As Integer = 0 To rarraystr.Length - 1
            rArray(rindex) = CInt(rarraystr(rindex))
        Next

        S2A_Integer = PASS

    End Function

End Module
