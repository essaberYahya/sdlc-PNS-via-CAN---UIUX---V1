Option Strict Off
Option Explicit On

Module modMain
    Dim AddTestplan As TestplanDelegate = AddressOf Testplan
    Dim AddTestplanInit As TestplanDelegate = AddressOf TestplanInit
    Dim AddTestplanEnd As TestplanDelegate = AddressOf TestplanEnd

    Public Sub Main()
        Dim l As Integer
        Dim gmCode As Integer
        
        Call vbTplanSetControlFunctions(AddTestplan, AddTestplanInit, AddTestplanEnd)
        Call TplanSetVB()

        l = TplanCreateWindow(VB6.GetHInstance.ToInt32(), False)

        Do
            gmCode = TplanWinMsgLoop()
        Loop While (gmCode <> 0) And (gmCode <> -1)

        End

    End Sub

End Module