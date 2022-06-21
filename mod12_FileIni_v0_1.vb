Imports System.Windows.Forms

Module i_FileIni

  Private Declare Auto Function GetPrivateProfileString Lib "kernel32.dll" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

  Private Declare Auto Function WritePrivateProfileString Lib "kernel32.dll" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
     ByVal lpFileName As String) As Integer

  Public Function fGetINIFileName(Optional ByVal rIniFileName As String = "") As String

        If rIniFileName.Contains(":\") Or rIniFileName.Contains("\\") Then

            fGetINIFileName = rIniFileName

        Else

            If rIniFileName <> "" Then
                fGetINIFileName = Application.StartupPath & "\" & rIniFileName
            Else
                fGetINIFileName = Left(Application.ExecutablePath, Len(Application.ExecutablePath) - 3) & "INI"
            End If

        End If

    End Function

  Public Sub sSetStringToINIFile(ByVal rSectionName As String, ByVal rKeyName As String, ByVal rValue As String, Optional ByVal rIniFileName As String = "")

    WritePrivateProfileString(rSectionName, rKeyName, rValue, fGetINIFileName(rIniFileName))

  End Sub

  Public Function fGetStringFromINIFile(ByVal rSectionName As String, ByVal rKeyName As String, Optional ByVal rIniFileName As String = "") As String

    Dim Buffer As String
    Dim BufferLength As Integer

    Buffer = Space$(1024)
    BufferLength = GetPrivateProfileString(rSectionName, rKeyName, "", Buffer, 1024, fGetINIFileName(rIniFileName))
    Buffer = Left$(Buffer, BufferLength)
    BufferLength = InStr(Buffer, ";")
    If BufferLength > 0 Then
      fGetStringFromINIFile = RTrim(Left$(Buffer, BufferLength - 1))
    Else
      fGetStringFromINIFile = Buffer
    End If

  End Function

End Module
