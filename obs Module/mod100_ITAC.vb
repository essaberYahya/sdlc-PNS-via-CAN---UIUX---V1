Module mod100_ITAC

    ' *** Public ITAC Variables
    Public ITAC As New ApiCall.iTAC_Call
    Public rpITACIsOnline As Boolean
    Public rpITACLogin As Boolean
    Public rpITACLogout As Boolean
    Public rpITACIniFilePath As String
    Public apITACCheckResult() As String
    Public apITACFailure()() As String = Nothing
    Public rpITACBookingOK As Boolean

    Public Sub sCollectITACFailure(ByVal rSite As Integer, ByVal str1 As String, ByVal str2 As String, ByVal str3 As String, ByVal str4 As String)

        Dim arraysize As Integer
        Dim newarraysize As Integer

        arraysize = apITACFailure(rSite).Length
        newarraysize = arraysize + 3
        ReDim Preserve apITACFailure(rSite)(newarraysize)

        apITACFailure(rSite)(arraysize) = str1
        apITACFailure(rSite)(arraysize + 1) = str2
        apITACFailure(rSite)(arraysize + 2) = str3
        apITACFailure(rSite)(arraysize + 3) = str4

    End Sub

End Module
