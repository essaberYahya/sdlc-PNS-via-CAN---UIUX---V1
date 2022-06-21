Imports System.Threading.Thread

Module mod07_Stimuli

    Public Function fPowerOnBSTV1() As Long

        ' *** PowerOn Booster V1

        sPwrOnBst(BSTV1, 0, 0.1, False)

        sSetBst(BSTV1, 0, 0.5, False)

        sSetBst(BSTV1, 0, 1, False)

        sSetBst(BSTV1, 5, 1, False)

        Sleep(1000)

    End Function

    Public Function fPowerOnBSTI1() As Long

        ' *** PowerOn Booster I1

        sPwrOnBst(BSTI1, 0, 0.1, True)

        sSetBst(BSTI1, 0, 0.5, True)

        sSetBst(BSTI1, 0, 1, True)

        sSetBst(BSTI1, 5, 1, True)

        Sleep(100)

    End Function

    Public Function fPowerOffBSTV1() As Long

        ' *** PowerOn Booster V1

        BstSourceSet(BSTV1, 0, 1, R1A, CONT_ON)
        Sleep(100)
        BstSourceSet(BSTV1, 0, 0.5, R1A, CONT_ON)
        Sleep(10)
        BstSourceSet(BSTV1, 0, 0.1, R1A, CONT_ON)
        Sleep(10)
        BstSourceSet(BSTV1, 0, 0.01, R100mA, CONT_ON)

        BstDisable(BSTV1)
        Sleep(5)

        BstDisconnectAll(BSTV1)

        'BstDisconnectInterf(BSTV1, HOT)
        'BstDisconnectInterf(BSTV1, COLD)

        Sleep(10)

    End Function

    Public Function fPowerOffBSTI1() As Long

        ' *** PowerOn Booster I1

        BstSourceSet(BSTI1, 0, 1, R1A, CONT_ON)
        Sleep(100)
        BstSourceSet(BSTI1, 0, 0.5, R1A, CONT_ON)
        Sleep(10)
        BstSourceSet(BSTI1, 0, 0.1, R1A, CONT_ON)
        Sleep(10)
        BstSourceSet(BSTI1, 0, 0.01, R100mA, CONT_ON)

        BstDisable(BSTI1)
        Sleep(5)

        BstDisconnectInterf(BSTI1, HOT)
        BstDisconnectInterf(BSTI1, COLD)

        Sleep(10)

    End Function

End Module

Module Driver

    Public Sub sPwrOnDrv(ByVal pDrv As Long, ByVal pVoltageVolt As Double, ByVal pCurrentAmp As Double, _
                         ByVal pExternal As Boolean)

        If pExternal Then
            DriConnectInterf(pDrv, HOT)
            DriConnectInterf(pDrv, COLD)
        Else
            DriConnectAbus(pDrv, ABUS1, ABUS4)
        End If

        Sleep(5)

        If pCurrentAmp <= 0.1 Then
            DriSourceSet(pDrv, pVoltageVolt, pCurrentAmp, R100mA, DIRECT, CONT_ON, NORMAL, BY_PASS_OFF)
        Else
            DriSourceSet(pDrv, pVoltageVolt, pCurrentAmp, R1A, DIRECT, CONT_ON, NORMAL, BY_PASS_OFF)
        End If

        Sleep(5)

        DriEnable(pDrv)
        Sleep(5)

    End Sub

    Public Sub sSetDrv(ByVal pDrv As Long, ByVal pVoltageVolt As Double, ByVal pCurrentAmp As Double, _
                       ByVal pExternal As Boolean)

        If pCurrentAmp <= 0.1 Then
            DriSourceSet(pDrv, pVoltageVolt, pCurrentAmp, R100mA, DIRECT, CONT_ON, NORMAL, BY_PASS_OFF)
        Else
            DriSourceSet(pDrv, pVoltageVolt, pCurrentAmp, R1A, DIRECT, CONT_ON, NORMAL, BY_PASS_OFF)
        End If

        Sleep(5)

    End Sub

    Public Sub sPwrOffDrv(ByVal pDrv As Long)

        DriSourceSet(pDrv, 0, 0.01, R100mA, DIRECT, CONT_ON, NORMAL, BY_PASS_OFF)
        Sleep(5)

        DriDisable(pDrv)
        Sleep(5)

        DriDisconnectAll(pDrv)
        Sleep(5)

        DriClear(pDrv)
        Sleep(5)

    End Sub

End Module

Module Booster

    Public Sub sPwrOnBst(ByVal pBst As Long, ByVal pVoltageVolt As Double, ByVal pCurrentAmp As Double, _
                         ByVal pExternal As Boolean)

        If pExternal = True Then
            BstConnectInterf(pBst, HOT)
            BstConnectInterf(pBst, COLD)
            BstConnectInterf(pBst, SENSE)
        ElseIf pExternal = False Then
            BstConnectAbus(pBst, ABUS1, ABUS4)
        End If

        Sleep(5)

        BstSourceSet(pBst, 0, 0.01, R100mA, CONT_ON)

        If pCurrentAmp < 0.1 Then
            BstSourceSet(pBst, pVoltageVolt, pCurrentAmp, R100mA, CONT_ON)
        ElseIf pCurrentAmp <= 1 Then
            BstSourceSet(pBst, 0, 0.1, R1A, CONT_ON)
            BstSourceSet(pBst, pVoltageVolt, pCurrentAmp, R1A, CONT_ON)
        ElseIf pCurrentAmp <= 3 Then
            BstSourceSet(pBst, 0, 0.1, R1A, CONT_ON)
            BstSourceSet(pBst, 0, 1.1, R3A, CONT_ON)
            BstSourceSet(pBst, pVoltageVolt, pCurrentAmp, R3A, CONT_ON)
        End If

        Sleep(5)

        BstEnable(pBst)
        Sleep(5)

    End Sub

    Public Sub sSetBst(ByVal pBst As Long, ByVal pVoltageVolt As Double, ByVal pCurrentAmp As Double, _
                       ByVal pExternal As Boolean)

        If pCurrentAmp < 0.1 Then
            BstSourceSet(pBst, pVoltageVolt, pCurrentAmp, R100mA, CONT_ON)
        Else
            BstSourceSet(pBst, pVoltageVolt, pCurrentAmp, R1A, CONT_ON)
        End If

        Sleep(5)

    End Sub

    Public Sub sPwrOffBst(ByVal pBst As Long, Optional ByVal pCurrent As Double = 0.1)

        If pCurrent <= 0.1 Then
            BstSourceSet(pBst, 0, 0.1, R100mA, CONT_ON)
        ElseIf pCurrent <= 1 Then
            BstSourceSet(pBst, 0, 0.1, R1A, CONT_ON)
        ElseIf pCurrent <= 3 Then
            BstSourceSet(pBst, 0, 1, R3A, CONT_ON)
            Sleep(5)
            BstSourceSet(pBst, 0, 0.1, R1A, CONT_ON)
            Sleep(5)
        End If

        BstSourceSet(pBst, 0, 0.01, R100mA, CONT_ON)
        Sleep(5)

        BstDisable(pBst)
        Sleep(5)

        BstDisconnectAll(pBst)
        Sleep(5)

        BstClear(pBst)
        Sleep(5)

    End Sub

End Module
