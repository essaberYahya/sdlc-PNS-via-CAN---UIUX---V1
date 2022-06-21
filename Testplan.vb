Option Strict Off
Option Explicit On

Module modTestplan

    Public Function Testplan() As Integer

        Dim FailFlag As Integer

        FailFlag = PASS

        '
        ' --- INSERT YOUR CODE HERE ...
        '



        '
        ' --- Test Result management
        '
        TplanResultSet(FailFlag)


        Testplan = 1
    End Function


    ' ----------------------------------------------------------------------------
    '
    ' --- TESTPLAN Initialisation
    '
    ' This function is executed only one time when the test program is loaded.
    '
    Public Function TestplanInit() As Integer

        '
        ' --- INSERT YOUR CODE HERE ...
        '

        TestplanInit = 1
    End Function


    ' ----------------------------------------------------------------------------
    '
    ' --- TESTPLAN End
    '
    ' This function is executed only one time when the test program ends.
    '
    Public Function TestplanEnd() As Integer

        '
        ' --- INSERT YOUR CODE HERE ...
        '

        TestplanEnd = 1
    End Function

End Module
