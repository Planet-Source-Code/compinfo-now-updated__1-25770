Attribute VB_Name = "Mod1"
'Mem info
Public Sub MemDisp()
Mem.dwLength = Len(Mem)
GlobalMemoryStatus Mem
TotMem = (Mem.dwTotalPhys / 1024 + Mem.dwTotalPageFile / 1024) / 1024
TotMemAvail = (Mem.dwAvailPhys / 1024 + Mem.dwAvailPageFile / 1024) / 1024
TotalAvailPercent = TotMemAvail / (TotMem / 100)
FrmMain.PB1.Min = 0
FrmMain.PB1.Max = TotMem
FrmMain.PB1.Value = TotMemAvail
FrmMain.PB2.Min = 0
FrmMain.PB2.Max = Mem.dwTotalPhys
FrmMain.PB2.Value = Mem.dwAvailPhys
FrmMain.PB3.Min = 0
FrmMain.PB3.Max = Mem.dwTotalPageFile
FrmMain.PB3.Value = Mem.dwAvailPageFile
FrmMain.Lbl20.Caption = "Total system memory : " + CStr(TotMem) + " MB"
FrmMain.Lbl21.Caption = "Free system memory : " + CStr(TotMemAvail) + " MB," + CStr(TotalAvailPercent) + " %"
FrmMain.Lbl22.Caption = "Total physical memory : " & CStr(Mem.dwTotalPhys / 1024) & " KB"
FrmMain.Lbl23.Caption = "Free physical memory : " + CStr(Mem.dwAvailPhys / 1024) & " KB," + CStr(CInt(Mem.dwAvailPhys / (Mem.dwTotalPhys / 100))) & " %"
FrmMain.Lbl24.Caption = "Maximum swap file : " & CStr(Mem.dwTotalPageFile / 1024) & " KB"
FrmMain.Lbl25.Caption = "Current swap file : " + CStr(CLng((Mem.dwTotalPageFile / 1024) - (Mem.dwAvailPageFile / 1024))) & " KB"
FrmMain.Lbl26.Caption = "Free swap file : " + CStr(Mem.dwAvailPageFile / 1024) & " KB," + CStr(CInt(Mem.dwAvailPageFile / (Mem.dwTotalPageFile / 100))) & " %"
FrmMain.Lbl27.Caption = "Memory load index : " + CStr(Mem.dwMemoryLoad) & " %"
End Sub
'Windows info
Public Sub WinInfoDisp()
Dim oReg As New cRegistry
WinInfo.dwOSVersionInfoSize = Len(WinInfo)    'treba da je 148
lRet = GetVersionEx(WinInfo)
Select Case WinInfo.dwPlatformID
       Case VER_PLATFORM_WIN32_NT
           If WinInfo.dwMajorVersion = 5 And WinInfo.dwMinorVersion = 0 Then
               StrTmp = "Microsoft Windows 2000"
           ElseIf WinInfo.dwMajorVersion = 5 And WinInfo.dwMinorVersion = 1 Then
               StrTmp = "Microsoft Windows XP"
           Else
               StrTmp = "Microsoft Windows NT"
           End If
           If WinInfo.dwMinorVersion = 0 Then
               FrmMain.Lbl2.Caption = "Version:  " + CStr(WinInfo.dwMajorVersion) & "." & CStr(WinInfo.dwMinorVersion) & "0" & Left$(WinInfo.szCSDVersion, 2) & " Build " & CStr(NumMan.LoWord(WinInfo.dwBuildNumber)) + " Service Pack  " + CStr(WinInfo.wServicePackMajor) + "." + CStr(WinInfo.wServicePackMinor)
           Else
               FrmMain.Lbl2.Caption = "Version:  " + CStr(WinInfo.dwMajorVersion) & "." & CStr(WinInfo.dwMinorVersion) & Left$(WinInfo.szCSDVersion, 2) & " Build " & CStr(NumMan.LoWord(WinInfo.dwBuildNumber)) + " Service Pack  " + CStr(WinInfo.wServicePackMajor) + "." + CStr(WinInfo.wServicePackMinor)
           End If
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion"
               .ValueKey = "Plus! VersionNumber"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl3.Caption = "Enhacement pack version:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion"
               .ValueKey = "RegisteredOrganization"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl6.Caption = "Registered organization:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion"
               .ValueKey = "RegisteredOwner"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl7.Caption = "Registered owner:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion"
               .ValueKey = "ProductKey"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl8.Caption = "Product key:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows NT\CurrentVersion"
               .ValueKey = "ProductId"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.lbl9.Caption = "Windows ID:  " + Trim(.Value)
           End With
       Case VER_PLATFORM_WIN32_WINDOWS
           If WinInfo.dwMajorVersion = 4 And WinInfo.dwMinorVersion = 10 Then
               StrTmp = "Microsoft Windows 98"
               If Left$(WinInfo.szCSDVersion, 2) = " A" Then StrTmp = StrTmp + " Second Edition"
           ElseIf WinInfo.dwMajorVersion = 4 And WinInfo.dwMinorVersion = 90 Then
               StrTmp = "Microsoft Windows Millenium Edition"
           Else
               StrTmp = "Microsoft Windows 95"
               If Left$(WinInfo.szCSDVersion, 2) = " C" Then StrTmp = StrTmp + " OEM Service Release 2"
           End If
           If WinInfo.dwMinorVersion = 0 Then
               FrmMain.Lbl2.Caption = "Version:  " + CStr(WinInfo.dwMajorVersion) & "." & CStr(WinInfo.dwMinorVersion) & "0" & Left$(WinInfo.szCSDVersion, 2) & " Build " & CStr(NumMan.LoWord(WinInfo.dwBuildNumber))
           Else
               FrmMain.Lbl2.Caption = "Version:  " + CStr(WinInfo.dwMajorVersion) & "." & CStr(WinInfo.dwMinorVersion) & Left$(WinInfo.szCSDVersion, 2) & " Build " & CStr(NumMan.LoWord(WinInfo.dwBuildNumber))
           End If
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "Plus! VersionNumber"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl3.Caption = "Enhacement pack version:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "RegisteredOrganization"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl6.Caption = "Registered organization:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "RegisteredOwner"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl7.Caption = "Registered owner:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "ProductKey"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl8.Caption = "Product key:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "ProductId"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.lbl9.Caption = "Windows ID:  " + Trim(.Value)
           End With
       Case Else
           StrTmp = "Windows"
           If WinInfo.dwMinorVersion = 0 Then
               FrmMain.Lbl2.Caption = "Version:  " + CStr(WinInfo.dwMajorVersion) & "." & CStr(WinInfo.dwMinorVersion) & "0" & Left$(WinInfo.szCSDVersion, 2) & " Build " & CStr(NumMan.LoWord(WinInfo.dwBuildNumber))
           Else
               FrmMain.Lbl2.Caption = "Version:  " + CStr(WinInfo.dwMajorVersion) & "." & CStr(WinInfo.dwMinorVersion) & Left$(WinInfo.szCSDVersion, 2) & " Build " & CStr(NumMan.LoWord(WinInfo.dwBuildNumber))
           End If
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "Plus! VersionNumber"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl3.Caption = "Enhacement pack version:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "RegisteredOrganization"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl6.Caption = "Registered organization:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "RegisteredOwner"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl7.Caption = "Registered owner:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "ProductKey"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.Lbl8.Caption = "Product key:  " + Trim(.Value)
           End With
           With oReg
               .ClassKey = HKEY_LOCAL_MACHINE
               .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
               .ValueKey = "ProductId"
               .ValueType = REG_SZ
               If .Value = "" Then
                   .Value = "Unknown or None"
               End If
               FrmMain.lbl9.Caption = "Windows ID:  " + Trim(.Value)
           End With
End Select
If WinInfo.wProductType = VER_NT_WORKSTATION Then StrTmp = StrTmp + " Professional"
If WinInfo.wProductType = VER_NT_DOMAIN_CONTROLLER Then StrTmp = StrTmp + " domain controller"
If WinInfo.wProductType = VER_NT_SERVER Then StrTmp = StrTmp + " Server"
If WinInfo.wSuiteMask = VER_SUITE_BACKOFFICE Then StrTmp = StrTmp + vbCrLf + " BackOffice"
If WinInfo.wSuiteMask = VER_SUITE_DATACENTER Then StrTmp = StrTmp + vbCrLf + " DataCenter Server"
If WinInfo.wSuiteMask = VER_SUITE_ENTERPRISE Then StrTmp = StrTmp + vbCrLf + " Advanced Server"
If WinInfo.wSuiteMask = VER_SUITE_SMALLBUSINESS Then StrTmp = StrTmp + vbCrLf + " Small Business Server"
If WinInfo.wSuiteMask = VER_SUITE_SMALLBUSINESS_RESTRICTED Then StrTmp = StrTmp + vbCrLf + " Small Business Server with the restrictive client license in force"
If WinInfo.wSuiteMask = VER_SUITE_TERMINAL Then StrTmp = StrTmp + vbCrLf + " Terminal Services"
FrmMain.Lbl1.Caption = "Operating System:  " + StrTmp
If strtemp = "Microsoft Windows 95" Then
    If SystemParamsInt(SPI_GETWINDOWSEXTENSION, 1, Tempint, 0) Then FrmMain.Lbl1.Caption = "Operating System:  " + StrTmp + " with Windows Plus!"
End If
Set oReg = Nothing
End Sub
'Cpu i sys info
Public Sub CpuInfo()
GetSystemInfo CpuInformation
FrmMain.Lbl16.Caption = "Number of processors :  " + CStr(CpuInformation.dwNumberOfProcessors)
FrmMain.Lbl11.Caption = "Processor:  " + CpuVersion
FrmMain.Lbl12.Caption = "Class:  " + CpuClass
FrmMain.Lbl10.Caption = "Active processor mask :  " + CStr(CpuInformation.dwActiveProcessorMask)
FrmMain.Lbl13.Caption = "Processor Revision/stepping:  " + CStr(HiByte(CpuInformation.wProcessorRevision)) + " / " + CStr(LoByte(CpuInformation.wProcessorRevision))
FrmMain.Lbl14.Caption = "Block size for memory addressing :  " + CStr(CpuInformation.dwAllocationGranularity)
FrmMain.Lbl15.Caption = "Lowest load address :  " + CStr(CpuInformation.lpMinimumApplicationAdress)
FrmMain.lbl17.Caption = "Highest load address :  " + CStr(CpuInformation.lpMaximumApplicationAdress)
FrmMain.Lbl27.Caption = "Page size for virtual allocation :  " + CStr(CpuInformation.dwPageSize) + " Bytes"
If GetSystemMetrics(SM_SLOWMACHINE) Then
    FrmMain.lbl1E.Caption = "Low-end (slow) processor present."
Else
    FrmMain.lbl1E.Caption = "Mid- or Hi-end processor present."
End If
End Sub
Public Sub WinEnvironment()
WinEnv = String(145, Chr(0))
WinEnv = Left$(WinEnv, GetEnvironmentVariable("PATH", WinEnv, 145))
FrmMain.lblA.Caption = "Programs path:  " + WinEnv
WinEnv = String(145, Chr(0))
WinEnv = Left$(WinEnv, GetEnvironmentVariable("COMSPEC", WinEnv, 145))
FrmMain.lblB.Caption = "Command interpreter:  " + WinEnv
WinEnv = String(145, Chr(0))
WinEnv = Left$(WinEnv, GetEnvironmentVariable("prompt", WinEnv, 145))
FrmMain.lblC.Caption = "Prompt:  " + WinEnv
WinEnv = String(145, Chr(0))
WinEnv = GetSysIni("386enh", "pagingdrive")
FrmMain.Lbl28.Caption = "Paging Drive:  " + Left(WinEnv, 2)
End Sub
Public Sub GetSysPower()
Dim oReg As cRegistry
Set oReg = New cRegistry
lRet = GetSystemPowerStatus(SysPower)
Select Case SysPower.ACLineStatus
Case 0
    FrmMain.Lbl30.Caption = "AC Power: Off"
Case 1
    FrmMain.Lbl30.Caption = "AC Power: On"
Case 255
    FrmMain.Lbl30.Caption = "AC Power: Unknown"
End Select
Select Case SysPower.BatteryFlag
Case 1
    FrmMain.Lbl31.Caption = "Battery charge is high."
Case 2
    FrmMain.Lbl31.Caption = "Battery charge is low."
Case 4
    FrmMain.Lbl31.Caption = "Battery charge is critical."
Case 8
    FrmMain.Lbl31.Caption = "Battery is charging."
Case 128
    FrmMain.Lbl31.Caption = "The system has no battery."
Case 255
    FrmMain.Lbl31.Caption = "Battery charge status is unknown."
End Select
If SysPower.BatteryFullLifeTime = &HFFFFFFFF Then
    FrmMain.Lbl32.Caption = "Cannot determine battery full lifetime."
Else
    FrmMain.Lbl32.Caption = "Battery full life time: " + CStr(SysPower.BatteryFullLifeTime) + " seconds"
End If
If SysPower.BatteryLifePercent = 255 Then
    FrmMain.Lbl33.Caption = "Battery charge status is unknown."
    Else
    FrmMain.Lbl33.Caption = "Remaining left lifetime: " + CStr(SysPower.BatteryLifePercent) + " %"
End If
If SysPower.BatteryLifeTime = &HFFFFFFFF Then
    FrmMain.Lbl34.Caption = "Cannot determine battery left lifetime."
Else
    FrmMain.Lbl34.Caption = "Battery left life time: " + CStr(SysPower.BatteryFullLifeTime) + " seconds"
End If
ret = SystemParamsBool(SPI_GETSCREENSAVEACTIVE, 0, TempBool, 0)
If TempBool Then
    FrmMain.lbl35.Caption = "Screen Saving is enabled."
    ret = SystemParamsInt(SPI_GETSCREENSAVETIMEOUT, 0, StrInt, 0)
    FrmMain.lbl36.Caption = "Screen saving timeout:  " + CStr(StrInt / 60) + "  min(s)"
Else
    FrmMain.lbl35.Caption = "Screen Saving is not enabled."
    FrmMain.lbl36.Caption = "Screen saving timeout:  Unknown"
End If
With oReg
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Control Panel\Desktop"
    .ValueKey = "ScreenSaveLowPowerActive"
    .ValueType = REG_SZ
    If .Value = "0" Then
        FrmMain.lbl36.Caption = "Screen save low power active:  No"
    ElseIf .Value = "1" Then
        FrmMain.lbl36.Caption = "Screen save low power active:  Yes"
    Else
        FrmMain.lbl36.Caption = "Screen save low power active:  Unknown"
    End If
End With
With oReg
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Control Panel\Desktop"
    .ValueKey = "ScreenSavePowerOffActive"
    .ValueType = REG_SZ
    If .Value = "0" Then
        FrmMain.lbl37.Caption = "Screen save power off active:  No"
    ElseIf .Value = "1" Then
        FrmMain.lbl37.Caption = "Screen save power off active:  Yes"
    Else
        FrmMain.lbl37.Caption = "Screen save power off active:  Unknown"
    End If
End With
With oReg
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Control Panel\Desktop"
    .ValueKey = "ScreenSaveUsePassword"
    .ValueType = REG_SZ
    If .Value = "0" Then
        FrmMain.lbl38.Caption = "Screen save use password:  No"
    ElseIf .Value = "1" Then
        FrmMain.lbl38.Caption = "Screen save use password:  Yes"
    Else
        FrmMain.lbl38.Caption = "Screen save use password:  Unknown"
    End If
End With
With oReg
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Control Panel\PowerCfg"
    .ValueKey = "CurrentPowerPolicy"
    .ValueType = REG_SZ
    StrTmp = .Value
    FrmMain.lbl39.Caption = "Current power policy number:  " + .Value
End With
With oReg
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Control Panel\PowerCfg\PowerPolicies\" + StrTmp
    .ValueKey = "Name"
    .ValueType = REG_SZ
    FrmMain.lbl3A.Caption = "Current power policy name:  " + .Value
End With
With oReg
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Control Panel\PowerCfg\PowerPolicies\" + StrTmp
    .ValueKey = "Description"
    .ValueType = REG_SZ
    FrmMain.lbl3B.Caption = "Current power policy description:  " + .Value
End With
Set oReg = Nothing
End Sub
Public Sub MouseInfo()
Dim sys As Object
Set sys = New OS
FrmMain.List5.AddItem "General info:"
If GetSystemMetrics(SM_MOUSEPRESENT) Then
    FrmMain.List5.AddItem "     Mouse: Present"
Else
    FrmMain.List5.AddItem "     Mouse: Not present"
End If
FrmMain.List5.AddItem "     Model:  " + GetSysIni("boot.description", "mouse.drv")
Dim oReg As New cRegistry
If sys.IsWinNTless5 Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "SYSTEM\CurrentControlset\Control\class\{4D36E96F-E325-11CE-BFC1-08002BE10318}\0000"
        .ValueKey = "MouseType"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        FrmMain.List5.AddItem "     Mouse type:  " + Trim(.Value)
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "SYSTEM\Currentcontrolset\Services\class\mouse\0000"
        .ValueKey = "MouseType"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        FrmMain.List5.AddItem "     Mouse type:  " + Trim(.Value)
    End With
End If
Set oReg = Nothing
FrmMain.List5.AddItem "     Number of mouse buttons: " + CStr(GetSystemMetrics(SM_CMOUSEBUTTONS))
If GetSystemMetrics(SM_SWAPBUTTON) Then
    FrmMain.List5.AddItem "     Button configuration: Right-handed"
Else
    FrmMain.List5.AddItem "     Button configuration: Left-handed"
End If
If GetSystemMetrics(SM_MOUSEWHEELPRESENT) Then
    FrmMain.List5.AddItem "     Mouse wheel is present."
Else
    FrmMain.List5.AddItem "     Mouse wheel is not present."
End If
FrmMain.List5.AddItem ""
FrmMain.List5.AddItem "Mouse settings:"
If sys.IsWinNT Or sys.IsWin2000 Then
        FrmMain.List5.AddItem "     Pointer trails:  Unsupported"
Else
    lRet = SystemParamsInt(SPI_GETMOUSETRAILS, 0, Tempint, 0)
    If Tempint = 0 Or Tempint = 1 Then
        FrmMain.List5.AddItem "     Pointer trails:  No"
    Else
        FrmMain.List5.AddItem "     Pointer trails:  " + CStr(Tempint)
    End If
End If
lRet = SystemParamsInt(SPI_GETMOUSESPEED, 0, Tempint, 0)
FrmMain.List5.AddItem "     Mouse speed:  " + CStr(Tempint)
FrmMain.List5.AddItem "     Double click speed:  " + CStr(GetDoubleClickTime) + " ms"
lRet = SystemParamsInt(SPI_GETWHEELSCROLLLINES, 0, Tempint, 0)
FrmMain.List5.AddItem "     Mouse wheel scroll lines:  " + CStr(Tempint)
If sys.IsWin95 Then
    FrmMain.List5.AddItem "     Hot tracking is not supported."
Else
    lRet = SystemParamsBool(SPI_GETHOTTRACKING, 0, TempBool, 0)
    If TempBool = True Then
        FrmMain.List5.AddItem "     Hot tracking is enabled."
    Else
        FrmMain.List5.AddItem "     Hot tracking is not enabled."
    End If
End If
If sys.IsWin5 Then
    lRet = SystemParamsBool(SPI_GETCURSORSHADOW, 0, TempBool, 0)
    If TempBool Then
        FrmMain.List5.AddItem "     Cursor shadow:  Yes"
    Else
        FrmMain.List5.AddItem "     Cursor shadow:  No"
    End If
Else
    FrmMain.List5.AddItem "     Cursor shadow:  Unsupported"
End If
If sys.IsWin95 Then
    FrmMain.List5.AddItem "     Snap to default button:  Unsupported"
Else
    lRet = SystemParamsBool(SPI_GETSNAPTODEFBUTTON, 0, TempBool, 0)
    If TempBool Then
        FrmMain.List5.AddItem "     Snap to default button:  Yes"
    Else
        FrmMain.List5.AddItem "     Snap to default button:  No"
    End If
End If
If sys.IsWin95 Then
    FrmMain.List5.AddItem "     Mouse hover width:  Unsupported"
Else
    lRet = SystemParamsInt(SPI_GETMOUSEHOVERWIDTH, 0, Tempint, 0)
    FrmMain.List5.AddItem "     Mouse hover width:  " + CStr(Tempint)
End If
If sys.IsWin95 Then
    FrmMain.List5.AddItem "     Mouse hover height:  Unsupported"
Else
    lRet = SystemParamsInt(SPI_GETMOUSEHOVERHEIGHT, 0, Tempint, 0)
    FrmMain.List5.AddItem "     Mouse hover height:  " + CStr(Tempint)
End If
If sys.IsWin95 Then
    FrmMain.List5.AddItem "     Mouse hover time:  Unsupported"
Else
    lRet = SystemParamsInt(SPI_GETMOUSEHOVERTIME, 0, Tempint, 0)
    FrmMain.List5.AddItem "     Mouse hover time:  " + CStr(Tempint)
End If
FrmMain.List5.AddItem "     Drag tolerance rectangle width:  " + CStr(GetSystemMetrics(SM_CXDRAG))
FrmMain.List5.AddItem "     Drag tolerance rectangle height:  " + CStr(GetSystemMetrics(SM_CYDRAG))
FrmMain.List5.AddItem "     Width of the rectangle double click area:  " + CStr(GetSystemMetrics(SM_CXDOUBLECLK))
FrmMain.List5.AddItem "     Height of the rectangle double click area:  " + CStr(GetSystemMetrics(SM_CYDOUBLECLK))
Dim sv As Object
Set sv = New FileVersion
sv.GetVersionClassic (GetSysIni("boot", "mouse.drv"))
sv.GetFileInfo (GetSysIni("boot", "mouse.drv"))
FrmMain.List5.AddItem ""
FrmMain.List5.AddItem "Driver information:"
FrmMain.List5.AddItem "     Driver:  " + GetSysIni("boot", "mouse.drv")
FrmMain.List5.AddItem "     Driver description:  " + sv.ProductName
FrmMain.List5.AddItem "     Driver version:  " + CStr(sv.MajorVersion) + "." + CStr(sv.MinorVersion) + "." + CStr(sv.QFEVersion) + "." + CStr(sv.BuildNumber)
FrmMain.List5.AddItem "     Driver provider:  " + sv.CompanyName
FrmMain.List5.AddItem "     Driver date:  " + sv.FileDate
FrmMain.List5.AddItem "     Driver time:  " + sv.FileTime
Set sys = Nothing
End Sub
Public Sub KeyInfo()
Dim ret As Integer
Dim sv As Object
Dim lan As Object
Set sv = New FileVersion
Set lan = New OS
FrmMain.List6.AddItem "General info:"
FrmMain.List6.AddItem "     Model:  " + GetSysIni("boot.description", "keyboard.typ")
FrmMain.List6.AddItem "     Keyboard Type/SubType:  " + CStr(GetKeyboardType(0)) + "/" + CStr(GetKeyboardType(1))
StrInt = GetOEMCP
FrmMain.List6.AddItem "     OEM Code Page:  " + CStr(StrInt)
Select Case StrInt
Case 437
     FrmMain.List6.AddItem "     OEM Code Page name:  MS-DOS United States"
Case 708
     FrmMain.List6.AddItem "     OEM Code Page name:  Arabic (ASMO 708)"
Case 709
     FrmMain.List6.AddItem "     OEM Code Page name:  Arabic (ASMO 449+, BCON V4)"
Case 710
     FrmMain.List6.AddItem "     OEM Code Page name:  Arabic (Transparent Arabic)"
Case 720
     FrmMain.List6.AddItem "     OEM Code Page name:  Arabic (Transparent ASMO)"
Case 737
     FrmMain.List6.AddItem "     OEM Code Page name:  Greek (formerly 437G)"
Case 775
     FrmMain.List6.AddItem "     OEM Code Page name:  Baltic"
Case 850
     FrmMain.List6.AddItem "     OEM Code Page name:  MS-DOS Multilingual (Latin I)"
Case 852
     FrmMain.List6.AddItem "     OEM Code Page name:  MS-DOS Slavic (Latin II)"
Case 855
     FrmMain.List6.AddItem "     OEM Code Page name:  IBM Cyrillic (primarily Russian)"
Case 857
     FrmMain.List6.AddItem "     OEM Code Page name:  IBM Turkish"
Case 860
     FrmMain.List6.AddItem "     OEM Code Page name:  MS-DOS Portuguese"
Case 861
     FrmMain.List6.AddItem "     OEM Code Page name:  MS-DOS Icelandic"
Case 862
     FrmMain.List6.AddItem "     OEM Code Page name:  Hebrew"
Case 863
     FrmMain.List6.AddItem "     OEM Code Page name:  MS-DOS Canadian-French"
Case 864
     FrmMain.List6.AddItem "     OEM Code Page name:  Arabic"
Case 865
     FrmMain.List6.AddItem "     OEM Code Page name:  MS-DOS Nordic"
Case 866
     FrmMain.List6.AddItem "     OEM Code Page name:  MS-DOS Russian"
Case 869
     FrmMain.List6.AddItem "     OEM Code Page name:  IBM Modern Greek"
Case 874
     FrmMain.List6.AddItem "     OEM Code Page name:  Thai"
Case 932
     FrmMain.List6.AddItem "     OEM Code Page name:  Japan"
Case 936
     FrmMain.List6.AddItem "     OEM Code Page name:  Chinese (PRC, Singapore)"
Case 949
     FrmMain.List6.AddItem "     OEM Code Page name:  Korean"
Case 950
     FrmMain.List6.AddItem "     OEM Code Page name:  Chinese (Taiwan; Hong Kong SAR, PRC)"
Case 1361
     FrmMain.List6.AddItem "     OEM Code Page name:  Korean (Johab)"
Case Else
     FrmMain.List6.AddItem "     OEM Code Page name:  Unknown"
End Select
StrInt = GetACP
FrmMain.List6.AddItem "     ANSI Code Page:  " + CStr(StrInt)
Select Case StrInt
Case 874
     FrmMain.List6.AddItem "     ANSI Code Page name:  Thai"
Case 932
     FrmMain.List6.AddItem "     ANSI Code Page name:  Japan"
Case 936
     FrmMain.List6.AddItem "     ANSI Code Page name:  Chinese (PRC, Singapore)"
Case 949
     FrmMain.List6.AddItem "     ANSI Code Page name:  Korean"
Case 950
     FrmMain.List6.AddItem "     ANSI Code Page name:  Chinese (Taiwan; Hong Kong SAR, PRC)"
Case 1200
     FrmMain.List6.AddItem "     ANSI Code Page name:  Unicode (BMP of ISO 10646)"
Case 1250
     FrmMain.List6.AddItem "     ANSI Code Page name:  Windows 3.1 Eastern European"
Case 1251
     FrmMain.List6.AddItem "     ANSI Code Page name:  Windows 3.1 Cyrillic"
Case 1252
     FrmMain.List6.AddItem "     ANSI Code Page name:  Windows 3.1 Latin 1 (US, Western Europe)"
Case 1253
     FrmMain.List6.AddItem "     ANSI Code Page name:  Windows 3.1 Greek"
Case 1254
     FrmMain.List6.AddItem "     ANSI Code Page name:  Windows 3.1 Turkish"
Case 1255
     FrmMain.List6.AddItem "     ANSI Code Page name:  Hebrew"
Case 1256
     FrmMain.List6.AddItem "     ANSI Code Page name:  Arabic"
Case 1257
     FrmMain.List6.AddItem "     ANSI Code Page name:  Baltic"
Case Else
     FrmMain.List6.AddItem "     ANSI Code Page name:  Unknown"
End Select
Dim ret1 As Integer
SystemParamsInt SPI_GETKEYBOARDDELAY, 0, ret1, 0
Dim ret2 As Long
SystemParamsLong SPI_GETKEYBOARDSPEED, 0, ret2, 0
FrmMain.List6.AddItem ""
FrmMain.List6.AddItem "Keyboard settings:"
Dim sRet  As String
Dim retu
sRet = Space$(32)
GetKeyboardLayoutName sRet
retu = InStr(sRet, Chr$(0))
If retu > 0 Then FrmMain.List6.AddItem "     Layout ID:  " + Left$(sRet, retu - 1)
FrmMain.List6.AddItem "     Layout name:  " + lan.GetLanguageName(Left$(sRet, retu - 1))
FrmMain.List6.AddItem "     Repeat delay:  " + CStr(ret1)
FrmMain.List6.AddItem "     Repeat speed:  " + CStr(ret2)
FrmMain.List6.AddItem "     Number of function keys:  " + CStr(GetKeyboardType(2))
FrmMain.List6.AddItem "     Caret flash speed:  " + CStr(Mod2.GetCaretBlinkTime) + " ms"
If lan.IsWin95 Then
    lRet = SystemParamsLong(SPI_GETCARETWIDTH, 0, Mod2.ret, 0)
    FrmMain.List6.AddItem "     Caret width:  " + CStr(ret) + " pixels"
Else
    FrmMain.List6.AddItem "     Caret width:  Not Supported"
End If
sv.GetVersionClassic (GetSysIni("boot", "keyboard.drv"))
sv.GetFileInfo (GetSysIni("boot", "keyboard.drv"))
FrmMain.List6.AddItem ""
FrmMain.List6.AddItem "Driver Information:"
FrmMain.List6.AddItem "     Driver:  " + GetSysIni("boot", "keyboard.drv")
FrmMain.List6.AddItem "     Driver description:  " + sv.FileDescription
FrmMain.List6.AddItem "     Driver version:  " + CStr(sv.MajorVersion) + "." + CStr(sv.MinorVersion) + "." + CStr(sv.BuildNumber) + "." + CStr(sv.QFEVersion)
FrmMain.List6.AddItem "     Driver provider:  " + sv.CompanyName
FrmMain.List6.AddItem "     Driver date:  " + sv.FileDate
FrmMain.List6.AddItem "     Driver date:  " + sv.FileTime
Set sv = Nothing
End Sub
Public Sub GetCompName()
CompName = String(32, Chr(0))
ret = GetComputerName(CompName, 32)
If ret = 0 Then
FrmMain.Lbl3.Caption = "Cannot get computer name"
Exit Sub
Else
CompName = Left(CompName, 31)
FrmMain.Lbl4.Caption = "Computer name:  " + CompName
End If
End Sub
Public Sub GetUserName()
UserName = String(256, Chr(0))
ret = Mod2.GetUserName(UserName, 256)
If ret = 0 Then
FrmMain.Lbl4.Caption = "Cannot get User Name"
Exit Sub
Else
UserName = Left(UserName, 255)
FrmMain.Lbl5.Caption = "User Name:  " + UserName
End If
End Sub
Public Sub WinInfoII()
If GetSystemMetrics(SM_DEBUG) Then
    FrmMain.lbl80.Caption = "Debugging version of USER.EXE is installed."
Else
    FrmMain.lbl80.Caption = "Debugging version of USER.EXE is not installed."
End If
If GetSystemMetrics(SM_DBCSENABLED) Then
    FrmMain.lbl81.Caption = "DBCS version of USER.EXE is installed"
Else
    FrmMain.lbl81.Caption = "DBCS version of USER.EXE is not installed."
End If
If GetSystemMetrics(SM_CLEANBOOT) = 0 Then
    FrmMain.lbl82.Caption = "Windows had Normal boot."
ElseIf 1 Then
    FrmMain.lbl82.Caption = "Windows had Fail-Safe boot."
ElseIf 2 Then
    FrmMain.lbl82.Caption = "Windows had Fail-Safe with network boot."
End If
If GetSystemMetrics(SM_MIDEASTENABLED) = True Then
    FrmMain.lbl83.Caption = "Hebrew and Arabic languages are enabled on this system."
Else
    FrmMain.lbl83.Caption = "Hebrew and Arabic languages are not enabled on this system."
End If
If GetSystemMetrics(SM_PENWINDOWS) = True Then
    FrmMain.lbl84.Caption = "Microsoft Windows for Pen computing extensions are installed."
Else
    FrmMain.lbl84.Caption = "Microsoft Windows for Pen computing extensions are not installed."
End If
If GetSystemMetrics(SM_SECURE) = True Then
    FrmMain.lbl85.Caption = "Security is present."
Else
    FrmMain.lbl85.Caption = "Security is not present."
End If
ret = SystemParamsBool(SPI_GETSCREENREADER, 1, TempBool, 0)
If TempBool Then
    FrmMain.lbl86.Caption = "Screen reviewer utility is running."
Else
    FrmMain.lbl86.Caption = "Screen reviewer utility is not running."
End If
If GetSystemMetrics(SM_SHOWSOUNDS) Then
    FrmMain.lbl87.Caption = "Show sounds is enabled."
Else
    FrmMain.lbl87.Caption = "Show sounds is not enabled."
End If
If GetSystemMetrics(SM_MENUDROPALIGNMENT) Then
    FrmMain.lbl88.Caption = "Drop-down menus are right-aligned."
Else
    FrmMain.lbl88.Caption = "Drop-down menus are left-aligned."
End If
Dim oReg As New cRegistry
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "System\CurrentControlSet\Control\Shutdown"
        .ValueKey = "FastReboot"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        If Trim(.Value) = "0" Then
            FrmMain.lbl89.Caption = "Fast Reboot enabled: No"
        ElseIf Trim(.Value) = "1" Then
            FrmMain.lbl89.Caption = "Fast Reboot enabled: Yes"
        Else
            FrmMain.lbl89.Caption = "Fast Reboot enabled: Unknown"
        End If
    End With
Set oReg = Nothing
If (GetSystemMetrics(SM_NETWORK) And &H1) = &H1 Then
    FrmMain.lbl8A.Caption = "Network connection is installed."
Else
    FrmMain.lbl8A.Caption = "Network connection is not installed."
End If
StrTmp = GetMsdosSys("Options", "BootGUI")
If StrTmp = "1" Then
    FrmMain.lbl8B.Caption = "Boot in to GUI:  True"
Else
    FrmMain.lbl8B.Caption = "Boot in to GUI:  False"
End If
StrTmp = GetMsdosSys("Options", "AutoScan")
If StrTmp = "0" Then
    FrmMain.lbl8C.Caption = "AutoScan after bad shutdown:  False"
Else
    FrmMain.lbl8C.Caption = "AutoScan after bad shutdown:  True"
End If
StrTmp = GetMsdosSys("Options", "Logo")
If StrTmp = "0" Then
    FrmMain.lbl8D.Caption = "Show logo during start up:  False"
Else
    FrmMain.lbl8D.Caption = "Show logo during start up:  True"
End If
StrTmp = GetMsdosSys("Options", "DrvSpace")
If StrTmp = "0" Then
    FrmMain.lbl8E.Caption = "Load DriveSpace driver:  False"
Else
    FrmMain.lbl8E.Caption = "Load DriveSpace driver:  True"
End If
StrTmp = GetMsdosSys("Options", "DblSpace")
If StrTmp = "0" Then
    FrmMain.lbl8F.Caption = "Load DoubleSpace driver:  False"
Else
    FrmMain.lbl8F.Caption = "Load DoubleSpace driver:  True"
End If
End Sub
Public Function CpuVersion() As String
Dim sys As Object
Dim oReg As New cRegistry
Set sys = New OS
With oReg
    .ClassKey = HKEY_LOCAL_MACHINE
    .SectionKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
    .ValueKey = "VendorIdentifier"
    .ValueType = REG_SZ
    If .Value = "" Then
        .Value = "None"
    End If
    StrTmp = Trim(.Value)
End With
If sys.IsWinNT Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
        .ValueKey = "~MHz"
        .ValueType = REG_NONE
        CpuVersion = StrTmp + " " + .Value + " MHz"
    End With
Else
    CpuVersion = StrTmp
End If
Set sys = Nothing
Set oReg = Nothing
End Function
Public Function CpuClass() As String
    Dim oReg As New cRegistry
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
        .ValueKey = "Identifier"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        CpuClass = Trim(.Value)
    End With
Set oReg = Nothing
End Function
Public Sub BIOSInfo()
Dim oReg As New cRegistry
Dim sys As New OS
If sys.IsWin95 Or sys.IsWin98 Then
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Enum\Root\*PNP0C01\0000"
        .ValueKey = "BIOSName"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        FrmMain.lbl18.Caption = "Manufacturer:  " + Trim(.Value)
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Enum\Root\*PNP0C01\0000"
        .ValueKey = "BIOSVersion"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        FrmMain.lbl19.Caption = "Version:  " + Trim(.Value)
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Enum\Root\*PNP0C01\0000"
        .ValueKey = "BIOSDate"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        FrmMain.lbl1A.Caption = "BIOS Date:  " + Trim(.Value)
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Enum\Root\*PNP0C01\0000"
        .ValueKey = "Model"
        .ValueType = REG_BINARY
        FrmMain.lbl1B.Caption = "BIOS Model:  " + CStr(Hex(Asc(CStr(.Value))))
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Enum\Root\*PNP0C01\0000"
        .ValueKey = "SubModel"
        .ValueType = REG_BINARY
        StrTmp = Hex$(Asc(CStr(.Value)))
        If Len(StrTmp) = 1 Then StrTmp = "0" + StrTmp
        FrmMain.lbl1C.Caption = "BIOS SubModel:  " + StrTmp
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Enum\Root\*PNP0C01\0000"
        .ValueKey = "Revision"
        .ValueType = REG_BINARY
        StrTmp = Hex$(Asc(CStr(.Value)))
        If Len(StrTmp) = 1 Then StrTmp = "0" + StrTmp
        FrmMain.lbl1D.Caption = "BIOS Revision:  " + StrTmp
    End With
Else
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "HARDWARE\DESCRIPTION\System"
        .ValueKey = "BIOSName"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        FrmMain.lbl18.Caption = "Manufacturer:  " + Trim(.Value)
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "HARDWARE\DESCRIPTION\System"
        .ValueKey = "BIOSVersion"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        FrmMain.lbl19.Caption = "Version:  " + Trim(.Value)
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "HARDWARE\DESCRIPTION\System"
        .ValueKey = "BIOSDate"
        .ValueType = REG_SZ
        If .Value = "" Then
            .Value = "None"
        End If
        FrmMain.lbl1A.Caption = "BIOS Date:  " + Trim(.Value)
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "HARDWARE\DESCRIPTION\System"
        .ValueKey = "Model"
        .ValueType = REG_BINARY
        FrmMain.lbl1B.Caption = "BIOS Model:  " + CStr(Hex(Asc(CStr(.Value))))
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "HARDWARE\DESCRIPTION\System"
        .ValueKey = "SubModel"
        .ValueType = REG_BINARY
        StrTmp = Hex$(Asc(CStr(.Value)))
        If Len(StrTmp) = 1 Then StrTmp = "0" + StrTmp
        FrmMain.lbl1C.Caption = "BIOS SubModel:  " + StrTmp
    End With
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "HARDWARE\DESCRIPTION\System"
        .ValueKey = "Revision"
        .ValueType = REG_BINARY
        StrTmp = Hex$(Asc(CStr(.Value)))
        If Len(StrTmp) = 1 Then StrTmp = "0" + StrTmp
        FrmMain.lbl1D.Caption = "BIOS Revision:  " + StrTmp
    End With
End If
Set sys = Nothing
Set oReg = Nothing
End Sub

