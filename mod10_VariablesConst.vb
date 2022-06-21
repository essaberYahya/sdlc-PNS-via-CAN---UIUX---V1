Module Variables

    ' *** Enviroments Variables

    Public rCurrentDirectory As String

    ' ******************************************************************
    ' *** Unit Under Test Variables  
    ' ******************************************************************
    Public rpPercorsoFile As String
    Public rpCurrentBay As Integer
    Public rpCurrentSite As Integer

    Public rpIctResult As Integer

    ' --- Program Flow Control
    Public rpDrawRef As String
    Public rpNsite As Integer
    Public rpNtest As Integer
    Public rpSpecParagraph As String
    Public rpTestName As String
    Public rpDiagRemark As String
    Public rpMeasValue As Double
    Public rpTh As Double
    Public rpTl As Double
    Public rpUnitMeasure As String
    Public rpTestResult As Long

    ' ******************************************************************
    ' *** Configured By INI File (StartupPath & "Configuration\TPGM.ini"
    ' ******************************************************************

    ' *** [INFO]
    Public rpDebugMode As Long
    Public rpBayToBeDebug As Long
    Public rpDebugPrint As Long
    Public rpPrintInterTestResult As Long
    Public rpAddPrintEnable As Long
    Public rpPrintErrorOnFile As Long
    ' *** [ICT]
    Public rpForcePASS As Long
    ' *** [OBP]
    Public rpOBP_Printdebug As Long
    ' *** [FCT]
    Public rpFCT_TimeOutInizializzazione As Long
    Public rpFCT_TimeOutEeprom As Long
    Public rpFCT_HwVersion As String


    ' *** [CAN]
    Public rpCanPortHandle As Integer
    Public rpCanPortHandle1 As Integer
    Public rpCanPortHandle2 As Integer

    Public rpCanPort As ULong
    Public rpCanPort1 As ULong
    Public rpCanPort2 As ULong
    Public rpCanSpeed As UInteger
    Public rpCanSpeed1 As UInteger
    Public rpCanSpeed2 As UInteger

End Module


Module Constants

    ' *** RESULT Value
    Public Const PASS As Long = 0
    Public Const FAIL As Long = 1
    Public Const NONE As Long = -1

    ' *** STATUS Value
    Public Const kp_ON As Long = 0
    Public Const kp_OFF As Long = 1
    Public Const kp_DISABLED As Long = 0
    Public Const kp_ENABLED As Long = 1

    ' *** OUTPUT TYPE Value
    Public Const Typ_Hexdec As Long = 0   ' *** HEX ARRAY STRING
    Public Const Typ_Binary As Long = 1   ' *** BIN ARRAY STRING
    Public Const Typ_Decimal As Long = 2  ' *** DEC ARRAY STRING

    ' *** COMMON Value
    Public Const Kp_CanDelay As Integer = 30
    Public Const Kp_CanDelay_Slow As Integer = 120
End Module
