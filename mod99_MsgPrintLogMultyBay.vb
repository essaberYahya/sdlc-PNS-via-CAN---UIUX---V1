Module mod99_MsgPrintLogMultyBay
    'Author: Marco Pomponio
    '
    'Usage:
    '	TotalBaysCount -> number of columns
    '	TotalLineDimension -> dimension of the runpack window
    '
    '   Constructors:
    '       Dim MyMsgPrintLogClass As MsgPrintLogClass = New MsgPrintLogClass()
    '       Dim MyMsgPrintLogClass As MsgPrintLogClass = New MsgPrintLogClass(4, 150)   -> Specify TotalBaysCount and TotalLineDimension
    '
    '   MyMsgPrintLogClass.SetBayCount(2)                                               -> Specify TotalBaysCoun
    '   MyMsgPrintLogClass.SetTotalLineDimension(77)                                    -> Specify SetTotalLineDimension
    '   MyMsgPrintLogClass.SetDebugEnable(True)                                         -> Enable or disable debug prints
    '
    '	MyMsgPrintLogClass.MsgPrintLogMultyBay("My Message", 1, True) 	                -> Print the message in the column indicated (this print has been specified as debug)
    '   MyMsgPrintLogClass.MsgPrintLogMultyBay("My Message", 1, "My second Message") 	-> Print the message in the column indicated. The second message is padded to the right
    '	MyMsgPrintLogClass.MsgPrintLogMultyBay("My Message", 2, 0)	                    -> Print the message in the column indicated with a PASS, FAIL, WARN or RUNN message at the end
    '										                                         0 	-> PASS
    '										                                         1 	-> FAIL
    '										                                         2 	-> WARN
    '										                                         3 	-> RUNN
    '	MyMsgPrintLogClass.MsgPrintLogMultyBay("-", 1) 			                        -> Print a "-----" filler in the column indicated
    '
    '   MyMsgPrintLogClass.MsgPrintLogSingleBay(...)                                    -> Like MsgPrintLogMultyBay but TotalBaysCoun is forced temporarily to 1
    '
    'Notes:
    '	Special Commands for Foreground and Background colors are supported 
    Class MsgPrintLogClass

        Public Sub New()
            TotalBaysCount = 2
            TotalLineDimension = 77
            DebugEnable = True
        End Sub
        Public Sub New(ByVal nBay As Integer, ByVal nDim As Integer)
            TotalBaysCount = nBay
            TotalLineDimension = nDim
            DebugEnable = True
        End Sub

        Private TotalBaysCount As Integer
        Private TotalLineDimension As Integer
        Private DebugEnable As Boolean

        Sub SetBayCount(ByVal nBay As Integer)
            TotalBaysCount = nBay
        End Sub

        Sub SetDebugEnable(ByVal Enable As Boolean)
            DebugEnable = Enable
        End Sub
        Sub SetTotalLineDimension(ByVal nDim As Integer)
            TotalLineDimension = nDim
        End Sub

        Sub MsgPrintLogSingleBay(ByVal Message As String, ByVal nBay As Integer, ByVal Result As Integer, Optional DebugPrint As Boolean = False)
            If DebugPrint And (Not DebugEnable) Then Return
            Dim RightPadding As String
            If Result = 0 Then
                RightPadding = "@BG{Green}@FG{White}PASS"
            ElseIf Result = 2 Then
                RightPadding = "@BG{Yellow}@FG{Black}WARN"
            ElseIf Result = 3 Then
                RightPadding = "@BG{Blue}@FG{White}RUNN"
            Else
                RightPadding = "@BG{Red}@FG{White}FAIL"
            End If

            MsgPrintLogSingleBay(Message, nBay, RightPadding)

        End Sub

        Sub MsgPrintLogSingleBay(ByVal Message As String, ByVal nBay As Integer, ByVal Result As Long, Optional DebugPrint As Boolean = False)
            If DebugPrint And (Not DebugEnable) Then Return
            Dim RightPadding As String
            If Result = 0 Then
                RightPadding = "@BG{Green}@FG{White}PASS"
            ElseIf Result = 2 Then
                RightPadding = "@BG{Yellow}@FG{Black}WARN"
            ElseIf Result = 3 Then
                RightPadding = "@BG{Blue}@FG{White}RUNN"
            Else
                RightPadding = "@BG{Red}@FG{White}FAIL"
            End If

            MsgPrintLogSingleBay(Message, nBay, RightPadding)

        End Sub

        Sub MsgPrintLogSingleBay(ByVal Message As String, ByVal nBay As Integer, Optional DebugPrint As Boolean = False)
            If DebugPrint And (Not DebugEnable) Then Return
            Dim MyTotalBaysCount As Integer = 1

            MsgPrintLogMultyBay_Line(Message, (TotalLineDimension - (MyTotalBaysCount - 1)) \ MyTotalBaysCount, nBay, MyTotalBaysCount, "", True, "")

        End Sub

        Sub MsgPrintLogSingleBay(ByVal Message As String, ByVal nBay As Integer, ByVal RightPadding As String, Optional DebugPrint As Boolean = False)
            If DebugPrint And (Not DebugEnable) Then Return
            Dim MyTotalBaysCount As Integer = 1

            MsgPrintLogMultyBay_Line(Message, (TotalLineDimension - (MyTotalBaysCount - 1)) \ MyTotalBaysCount, nBay, MyTotalBaysCount, "", True, RightPadding)

        End Sub

        Sub MsgPrintLogMultyBay(ByVal Message As String, ByVal nBay As Integer, ByVal Result As Long, Optional DebugPrint As Boolean = False)
            If DebugPrint And (Not DebugEnable) Then Return
            Dim RightPadding As String
            If Result = 0 Then
                RightPadding = "@BG{Green}@FG{White}PASS"
            ElseIf Result = 2 Then
                RightPadding = "@BG{Yellow}@FG{Black}WARN"
            ElseIf Result = 3 Then
                RightPadding = "@BG{Blue}@FG{White}RUNN"
            Else
                RightPadding = "@BG{Red}@FG{White}FAIL"
            End If

            MsgPrintLogMultyBay(Message, nBay, RightPadding)

        End Sub
        Sub MsgPrintLogMultyBay(ByVal Message As String, ByVal nBay As Integer, ByVal Result As Integer, Optional DebugPrint As Boolean = False)
            If DebugPrint And (Not DebugEnable) Then Return
            Dim RightPadding As String
            If Result = 0 Then
                RightPadding = "@BG{Green}@FG{White}PASS"
            ElseIf Result = 2 Then
                RightPadding = "@BG{Yellow}@FG{Black}WARN"
            ElseIf Result = 3 Then
                RightPadding = "@BG{Blue}@FG{White}RUNN"
            Else
                RightPadding = "@BG{Red}@FG{White}FAIL"
            End If

            MsgPrintLogMultyBay(Message, nBay, RightPadding)

        End Sub

        Sub MsgPrintLogMultyBay(ByVal Message As String, ByVal nBay As Integer, Optional DebugPrint As Boolean = False)
            If DebugPrint And (Not DebugEnable) Then Return

            MsgPrintLogMultyBay_Line(Message, (TotalLineDimension - (TotalBaysCount - 1)) \ TotalBaysCount, nBay, TotalBaysCount, "", True, "")

        End Sub

        Sub MsgPrintLogMultyBay(ByVal Message As String, ByVal nBay As Integer, ByVal RightPadding As String, Optional DebugPrint As Boolean = False)
            If DebugPrint And (Not DebugEnable) Then Return

            MsgPrintLogMultyBay_Line(Message, (TotalLineDimension - (TotalBaysCount - 1)) \ TotalBaysCount, nBay, TotalBaysCount, "", True, RightPadding)

        End Sub

        Private Sub MsgPrintLogMultyBay_Line(ByVal Message As String, ByVal MessageMaxLength As Integer, ByVal nBay As Integer, ByVal nTotalBay As Integer, ByVal PreviousSpecialCommandsString As String, ByVal MakeMessage As Boolean, Optional RightPaddingMessage As String = "")

            If nBay < 1 Or nBay > nTotalBay Then
                Dim ErrorMessage As String = "nBay is out of bounds"
                Message = ""
                For i As Integer = 1 To MessageMaxLength - ErrorMessage.Length \ 2
                    Message = Message & " "
                Next
                MsgPrintLog(Message & ErrorMessage, 0)
                Return
            End If

            Dim MessageIndex As Integer = 0
            Dim RealMessageIndex As Integer = 0
            Dim SpecialCommand As Boolean = False
            Dim SpecialCommandsString As String = ""

            'Make the entire message
            If MakeMessage Then
                If Message = "-" Then
                    Message = New String("-"c, MessageMaxLength)
                ElseIf RightPaddingMessage <> "" Then 'If there is a right padding message fill with . up to the closest MessageMaxLength multiple
                    If Message = "" Then Message = "."
                    Dim ActualLength As Integer = RealMessageLength(Message) + RealMessageLength(RightPaddingMessage) + 1
                    Dim Multiple As Integer = (ActualLength - 1) \ MessageMaxLength + 1
                    Message = Message & "@BG{White}@FG{Black}" & New String("."c, MessageMaxLength * Multiple - ActualLength + 1) & RightPaddingMessage & "@BG{White}@FG{Black}"
                Else 'Check if the length is a multiple of MessageMaxLength. If not add spaces
                    If Message = "" Then Message = " "
                    Dim ActualLength As Integer = RealMessageLength(Message)
                    If ActualLength Mod MessageMaxLength <> 0 Then
                        Dim Multiple As Integer = (ActualLength - 1) \ MessageMaxLength + 1
                        Message = Message & "@BG{White}@FG{Black}" & New String(" "c, MessageMaxLength * Multiple - ActualLength)
                    End If
                End If
            End If

            While RealMessageIndex <= MessageMaxLength
                If MessageIndex = Message.Length Then
                    MsgPrintLogMultyBay_LinePrint(PreviousSpecialCommandsString & Message, MessageMaxLength, nBay, nTotalBay)
                    Return
                End If

                Dim myChr As Char = Message.Chars(MessageIndex)

                If myChr = "@" Then
                    SpecialCommand = True
                    SpecialCommandsString = SpecialCommandsString & myChr
                ElseIf myChr = "}" Then
                    SpecialCommand = False
                    SpecialCommandsString = SpecialCommandsString & myChr
                Else
                    If SpecialCommand = True Then
                        SpecialCommandsString = SpecialCommandsString & myChr
                    Else
                        RealMessageIndex = RealMessageIndex + 1
                    End If
                End If

                MessageIndex = MessageIndex + 1
            End While

            MsgPrintLogMultyBay_LinePrint(PreviousSpecialCommandsString & Message.Substring(0, MessageIndex - 1) & "@BG{White}@FG{Black}", MessageMaxLength, nBay, nTotalBay)

            MsgPrintLogMultyBay_Line(Message.Substring(MessageIndex - 1), MessageMaxLength, nBay, nTotalBay, PreviousSpecialCommandsString & SpecialCommandsString, False)
        End Sub

        Private Sub MsgPrintLogMultyBay_LinePrint(ByVal Message As String, ByVal MessageMaxLength As Integer, ByVal nBay As Integer, ByVal nTotalBay As Integer)
            Dim ToDisplay As String = ""

            For i As Integer = 1 To nTotalBay
                If i > 1 Then
                    ToDisplay = ToDisplay & "|"
                End If

                If i = nBay Then
                    ToDisplay = ToDisplay & Message
                Else
                    For j As Integer = 1 To MessageMaxLength
                        ToDisplay = ToDisplay & " "
                    Next
                End If
            Next

            MsgPrintLog(ToDisplay, 0)
        End Sub

        Private Function RealMessageLength(ByVal Message As String) As Integer
            Dim RealMessageIndex As Integer = 0
            Dim MessageIndex As Integer = 0

            Dim SpecialCommand As Boolean = False

            While MessageIndex < Message.Length
                Dim myChr As Char = Message.Chars(MessageIndex)

                If myChr = "@" Then
                    SpecialCommand = True
                ElseIf myChr = "}" Then
                    SpecialCommand = False
                ElseIf SpecialCommand = False Then
                    RealMessageIndex = RealMessageIndex + 1
                End If

                MessageIndex = MessageIndex + 1
            End While
            Return RealMessageIndex
        End Function
    End Class
End Module
