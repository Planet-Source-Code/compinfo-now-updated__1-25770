Attribute VB_Name = "IniOps"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function GetSysIni(Section As String, key As String) As String
Dim RetVal As String * 128
lret = GetPrivateProfileString(Section, key, "None", RetVal, 128, "system.ini")
If lret = 0 Then
    GetSysIni = "Unknown"
Else
    GetSysIni = Left$(RetVal, lret)
End If
End Function
Public Function GetWinIni(Section As String, key As String) As String
Dim RetVal As String * 255
lret = GetPrivateProfileString(Section, key, "None", RetVal, 255, "win.ini")
If lret = 0 Then
    GetWinIni = "Unknown"
Else
    GetWinIni = Left$(RetVal, lret)
End If
End Function
Public Function GetMsdosSys(Section As String, key As String) As String
Dim RetVal As String * 255
WinEnv = String(145, Chr(0))
WinEnv = Left$(WinEnv, GetEnvironmentVariable("WINBOOTDIR", WinEnv, 145))
WinPath = Left(WinPath, 2)
lret = GetPrivateProfileString(Section, key, "None", RetVal, 255, "\msdos.sys")
If lret = 0 Then
    GetMsdosSys = "Unknown"
Else
    GetMsdosSys = Left$(RetVal, lret)
End If
End Function
