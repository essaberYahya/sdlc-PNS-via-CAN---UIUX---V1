Option Strict Off
Option Explicit On

Module modMain
    Dim AddTestplan As TestplanDelegate = AddressOf Testplan
    Dim AddTestplanInit As TestplanDelegate = AddressOf TestplanInit
    Dim AddTestplanEnd As TestplanDelegate = AddressOf TestplanEnd

    Public Sub Main()

        Dim l As Integer
        Dim gmCode As Integer

        Dim sn As String

        'Do

        '    'rpUutFctSerialNumber(0) = InputBox("Please scan the Serial Number", "Write Data").ToUpper
        '    'rpUutFctSerialNumber(0) = "D101234567"
        '    'Call fSerialNumberRead(0, sn)
        '    MsgBox("PUT NEW PART ON POSITION 1")
        '    'If rpUutFctSerialNumber(0) = "" Or Left(rpUutFctSerialNumber(0), 1) <> "D" Or rpUutFctSerialNumber(0).Length <> 10 Then
        '    '    MsgBox("Le SN scanné est erroné :( !!")

        '    '    Exit Do
        '    '    'ElseIf rpUutFctSerialNumber(0) = "E330062910" Or rpUutFctSerialNumber(0) = "E330062913" Then
        '    '    '    MsgBox("PUT IT A SIDE")
        '    'End If

        '    If fFCT() <> PASS Then
        '        ' *** SET fail flag
        '        MsgBox("STATION1 : failled!!!! ")
        '    Else
        '        'MsgBox("Write operation Done with " & rpUutFctResult(0) & " error")
        '        MsgBox("<<<<<<<<<<<<<<<     POSITION1    >>>>>>>>>>>>>" & Chr(10) & Chr(10) & "                             PART FLASHED CORRECTLY  ")
        '    End If

        'Loop While True



        End

    End Sub

End Module