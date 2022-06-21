Module mod05_FCT

    Public Function fFCT() As Long

        fFCT = FAIL

        Dim aUsedChList(16) As Short

        ' *** Channels Stuck only for debug
        'S2A("1090", aUsedChList(0), 5)

        'ObpChLevelSet(OBP500A, aUsedChList(0), 5.0, 0.0)
        'ObpChSensorSet(OBP500A, aUsedChList(0), 1.8)
        'ObpChConnect(OBP500A, aUsedChList(0))

        'ObpChStuckSet(OBP500A, aUsedChList(0), HIGH)
        'ObpChStuckSet(OBP500A, aUsedChList(0), LOW)

        ObpProgrammingSelect(OBP500A, "U4")
        ObpSiteSet(OBP500A, "1")
        ObpModelOptionsSet(OBP500A, "COMMFREQ=9600;DATABITS=8;PARITY=NONE;STOPBITS=1")

        Dim aTxFrame As String
        Dim aRxFrame As String
        Dim rReadDataLen As Integer = 1

        ' *** Setting and connection of used channels
        S2A("1089,1090", aUsedChList(0), 5)
        ObpChLevelSet(OBP500A, aUsedChList(0), 5.0, 0.0)
        ObpChSensorSet(OBP500A, aUsedChList(0), 1.8)
        ObpChConnect(OBP500A, aUsedChList(0))

        ' *** Board Power On
    'fPowerOnBoard()

        If ObpInterfaceEnable(OBP500A) = 0 Then

            ' *** Send and receive data
            aTxFrame = Chr(&H18)

            ' *** Send 5 byte
            'If ObpSerialPortSend(OBP500A, aTxFrame, 1) = 0 Then

            'aRxFrame = Space(20)

            'If ObpSerialPortReceive(OBP500A, 1, aRxFrame, 1, -1, 1, 1000) = 0 Then fFCT = PASS

            'ObpInterfaceDisable(OBP500A)

            'End If

        End If

        ' *** Board Power Off
    '  fPowerOffBoard()

        ' *** Disconnect used channels
    '  ObpChDisconnect(OBP500A, aUsedChList(0))

    End Function




End Module
