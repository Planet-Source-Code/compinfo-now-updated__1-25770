Attribute VB_Name = "DirsInfo"
Public Enum SHFolders
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_ALTSTARTUP = &H1D '// DBCS
    CSIDL_COMMON_ALTSTARTUP = &H1E '// DBCS
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
End Enum
Private Type ITEMIDLIST
    mkid As Long
End Type

Private Declare Function SHGetSpecialFolderLocation _
        Lib "shell32.dll" _
        (ByVal hwndOwner As Long, ByVal nFolder As SHFolders, _
        ppidl As ITEMIDLIST) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
        (ByVal pv As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
        ByVal pszPath As String) As Long
'Windows direktorijum
Public Sub WinDir()
WinPath = String(145, Chr(0))
WinPath = Left(WinPath, GetWindowsDirectory(WinPath, 145))
FrmMain.lbl110.Caption = "Windows folder:  " + WinPath
End Sub
'Windows\system direktorijum
Public Sub SysDir()
SysPath = String(145, Chr(0))
SysPath = Left(SysPath, GetSystemDirectory(SysPath, 145))
FrmMain.lbl111.Caption = "System folder:  " + SysPath
End Sub
Public Sub TempDir()
Mod2.WinEnv = String(145, Chr(0))
Mod2.WinEnv = Left$(WinEnv, GetEnvironmentVariable("temp", Mod2.WinEnv, 145))
FrmMain.lbl112.Caption = "Temporary folder:  " & Mod2.WinEnv
End Sub
Public Function GetBootDir() As String
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion\Setup"
        .ValueKey = "BootDir"
        .ValueType = REG_SZ
        GetBootDir = Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Setup"
        .ValueKey = "BootDir"
        .ValueType = REG_SZ
        GetBootDir = Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Function
Public Sub WinBootDir()
WinEnv = String(145, Chr(0))
WinEnv = Left$(WinEnv, GetEnvironmentVariable("WINBOOTDIR", WinEnv, 145))
FrmMain.lbl113.Caption = "Windows Boot folder:  " + WinEnv
End Sub
Public Sub GetConfigPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion\Setup"
        .ValueKey = "ConfigPath"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl115.Caption = "Config path:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Setup"
        .ValueKey = "ConfigPath"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl115.Caption = "Config path:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Sub GetICMPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion\Setup"
        .ValueKey = "ICMPath"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl116.Caption = "ICM path:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Setup"
        .ValueKey = "ICMPath"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl116.Caption = "ICM path:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Sub GetMediaPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion\Setup"
        .ValueKey = "MediaPath"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl117.Caption = "Media path:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Setup"
        .ValueKey = "MediaPath"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl117.Caption = "Media path:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub

Public Sub GetDevicePath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion"
        .ValueKey = "DevicePath"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl118.Caption = "Device path:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
        .ValueKey = "DevicePath"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl118.Caption = "Device path:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Sub GetOtherDevicePath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion"
        .ValueKey = "OtherDevicePath"
        .ValueType = REG_SZ
        If .Value = "" Then
           .Value = "Unknown or None"
        End If
        FrmMain.lbl119.Caption = "Other device path:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
        .ValueKey = "OtherDevicePath"
        .ValueType = REG_SZ
        If .Value = "" Then
           .Value = "Unknown or None"
        End If
        FrmMain.lbl119.Caption = "Other device path:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Sub GetCommonFilesPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion"
        .ValueKey = "CommonFilesDir"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11B.Caption = "Common files path:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
        .ValueKey = "CommonFilesDir"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11B.Caption = "Common files path:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Sub GetProgramFilesPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
        .ValueKey = "ProgramFilesDir"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11A.Caption = "Program files path:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
        .ValueKey = "ProgramFilesDir"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11A.Caption = "Program files path:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Sub GetWallPaperPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion"
        .ValueKey = "WallPaperDir"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11C.Caption = "WallPaper path:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
        .ValueKey = "WallPaperDir"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11C.Caption = "WallPaper path:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Sub GetPersonalPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion\explorer\Shell Folders"
        .ValueKey = "Personal"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11D.Caption = "Personal folder:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders"
        .ValueKey = "Personal"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11D.Caption = "Personal folder:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Sub GetCommonAppDataPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion\explorer\Shell Folders"
        .ValueKey = "Common AppData"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11E.Caption = "Common App Data folder:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders"
        .ValueKey = "Common AppData"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11E.Caption = "Common App Data folder:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Sub GetCommonDesktopPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion\explorer\Shell Folders"
        .ValueKey = "Common Desktop"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11F.Caption = "Common Desktop folder:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders"
        .ValueKey = "Common Desktop"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl11F.Caption = "Common Desktop folder:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub

Public Sub GetCommonStartupPath()
Dim oReg As New cRegistry
Dim sys As Object
Set sys = New OS
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion\explorer\Shell Folders"
        .ValueKey = "Common Startup"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl120.Caption = "Common Startup folder:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders"
        .ValueKey = "Common Startup"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "Unknown or None"
        End If
        FrmMain.lbl120.Caption = "Common Startup folder:  " + Trim(.Value)
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub
Public Function FolderLocation(lFolder As SHFolders, hwnd As Long) As String

    Dim lp As ITEMIDLIST
    Dim tmpStr As String
    SHGetSpecialFolderLocation hwnd, lFolder, lp
    tmpStr = Space$(255)
    SHGetPathFromIDList lp.mkid, tmpStr
    If InStr(tmpStr, Chr$(0)) > 0 Then
        tmpStr = Left$(tmpStr, InStr(tmpStr, Chr$(0)) - 1)
    End If
    CoTaskMemFree lp.mkid
    If tmpStr = "" Then tmpStr = "None Or Unknown"
    FolderLocation = tmpStr

End Function
Public Sub AddFolders()
tmpStr = FolderLocation(CSIDL_TEMPLATES, FrmMain.hwnd)
FrmMain.lbl121.Caption = "Shell New folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_STARTUP, FrmMain.hwnd)
FrmMain.lbl122.Caption = "StartUp folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_STARTMENU, FrmMain.hwnd)
FrmMain.lbl123.Caption = "Start menu folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_SENDTO, FrmMain.hwnd)
FrmMain.lbl123.Caption = "SendTo folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_RECENT, FrmMain.hwnd)
FrmMain.lbl124.Caption = "Recent folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_PRINTHOOD, FrmMain.hwnd)
FrmMain.lbl125.Caption = "PrinterHood folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_PRINTERS, FrmMain.hwnd)
FrmMain.lbl126.Caption = "Printers folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_NETWORK, FrmMain.hwnd)
FrmMain.lbl127.Caption = "Network folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_NETHOOD, FrmMain.hwnd)
FrmMain.lbl128.Caption = "NetworkHood folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_INTERNET_CACHE, FrmMain.hwnd)
FrmMain.lbl129.Caption = "Internet cache folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_INTERNET, FrmMain.hwnd)
FrmMain.lbl12A.Caption = "Internet folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_HISTORY, FrmMain.hwnd)
FrmMain.lbl12B.Caption = "History folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_FONTS, FrmMain.hwnd)
FrmMain.lbl12C.Caption = "Fonts folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_FAVORITES, FrmMain.hwnd)
FrmMain.lbl12D.Caption = "Favorites folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_DESKTOP, FrmMain.hwnd)
FrmMain.lbl12E.Caption = "Desktop folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_COOKIES, FrmMain.hwnd)
FrmMain.lbl12F.Caption = "Cookies folder:  " + tmpStr
tmpStr = FolderLocation(CSIDL_COMMON_DESKTOPDIRECTORY, FrmMain.hwnd)
FrmMain.lbl114.Caption = "Common Desktop folder:  " + tmpStr
End Sub
