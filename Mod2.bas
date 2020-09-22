Attribute VB_Name = "Mod2"
Option Explicit
'Konstante za product type
Public Const VER_NT_WORKSTATION = 1
Public Const VER_NT_DOMAIN_CONTROLLER = 2
Public Const VER_NT_SERVER = 3
'Konstante za suite type
Public Const VER_SUITE_DATACENTER = 128
Public Const VER_SUITE_ENTERPRISE = 2
Public Const VER_SUITE_SMALLBUSINESS = 1
Public Const VER_SUITE_BACKOFFICE = 4
Public Const VER_SUITE_TERMINAL = 16
Public Const VER_SUITE_SMALLBUSINESS_RESTRICTED = 32
Public Const VER_SUITE_COMMUNICATIONS = 8
Public Const VER_SUITE_EMBEDDEDNT = 64
Public Const VER_SUITE_SINGLEUSERTS = 256
'   Win 16s
Public Const VER_PLATFORM_WIN32s = 0
'   Win 95
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
'   Win NT
Public Const VER_PLATFORM_WIN32_NT = 2
'Konstante za CPU
Public Const SM_SLOWMACHINE = 73
'Konstante za misa
Public Const SM_MOUSEPRESENT = 19
Public Const SM_SWAPBUTTON = 23
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_MOUSEWHEELPRESENT = 75
Public Const SPI_GETMOUSETRAILS = 94
Public Const SPI_GETMOUSESPEED = 112
Public Const SPI_GETWHEELSCROLLLINES = 104
Public Const SPI_GETHOTTRACKING = 4110
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CXDRAG = 68
Public Const SM_CYDRAG = 69
Public Const SPI_GETMOUSEHOVERWIDTH = 98
Public Const SPI_GETMOUSEHOVERHEIGHT = 100
Public Const SPI_GETMOUSEHOVERTIME = 102
Public Const SPI_GETCURSORSHADOW = &H101A
Public Const SPI_GETSNAPTODEFBUTTON = 95
'Konstante za tastaturu
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_GETKEYBOARDDELAY = 22
Public Const SPI_GETCARETWIDTH = &H2006
'WinInfo konstante
Public Const SM_DEBUG = 22
Public Const SM_DBCSENABLED = 42
Public Const SM_CLEANBOOT = 67
Public Const SM_SECURE = 44
Public Const SPI_GETSCREENREADER = 70
Public Const SM_SHOWSOUNDS = 70
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_NETWORK = 63
'stanje memorije
Global Mem As MEMORYSTATUS
'konstante za GetTickCount
Public Const MPDay = 86400000
Public Const MPHour = 3600000
Public Const MPMinute = 60000
'konstante za screen saver
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const SPI_GETSCREENSAVETIMEOUT = 14
'Informacije o windowsu
Global WinInfo As OSVERSIONINFOEX
'Cpu info
Global CpuInformation As SYSTEMINFO
'Privremena promenljiva
Global StrInt As Integer
'Privremena promenljiva
Global StrTmp As String
'windows enviroment variables(programs path, temporary folder)
Global WinEnv As String
'Privremena promenljiva
Global ret As Long
'Privremena promenljiva
Global lRet As Long
'Privremena promenljiva
Global TempBool As Boolean
'Zbir fizicke memorije i swap fajla
Global TotMem As Long
'Kolicina slobodne memorije u zbiru fizicke memorije i swap fajla
Global TotMemAvail As Long
'Isto kao i gore samo u procentima
Global TotalAvailPercent As Integer
'windows direktorijum
Global WinPath As String
'Windows/system direktorijum
Global SysPath As String
'Snaga i jos neki parametri sys baterija
Global SysPower As SYSTEM_POWER_STATUS
'Ime Kompa
Global CompName As String
'Username ulogovanog korisnika
Global UserName As String
'Tip za syspower koji se salje GetSystemPowerStatus
Public Type SYSTEM_POWER_STATUS
        ACLineStatus As Byte
        BatteryFlag As Byte
        BatteryLifePercent As Byte
        Reserved1 As Byte
        BatteryLifeTime As Long
        BatteryFullLifeTime As Long
End Type


'
Public Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type



'Tip za mem koji se salje GlobalMemoryStatus
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

'Tip za verziju OSa koji se salje GetVersionEx
Type OSVERSIONINFOEX
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformID As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
        wServicePackMajor As Integer
        wServicePackMinor As Integer
        wSuiteMask As Integer
        wProductType As Byte
        wReserved As Byte
End Type
'Tip za procesor koji se salje GetSystemInfo
Type SYSTEMINFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAdress As Long
    lpMaximumApplicationAdress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type
'Deklaracije funkcija
'
Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
'Win direktorijum
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Info o sysu i procesoru
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEMINFO)
'Info o memu
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'Info o windowsu
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long
'Pomaze u raznim informacijama kao na primer da li je mis
'prisutan ili nije
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Windows\system direktorijum
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Vraca vrednosti sistemski varijabli kao npr. path iz autoexec.bat-a
Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
'vracaju keyboard delay i speed
Declare Function SystemParamsInt Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    lpvParam As Integer, ByVal fuWinIni As Long) As Long
Declare Function SystemParamsBool Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    lpvParam As Boolean, ByVal fuWinIni As Long) As Long
Declare Function SystemParamsLong Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    lpvParam As Long, ByVal fuWinIni As Long) As Long
'Vraca Layout tastature
Declare Function GetKeyboardLayoutName Lib "user32" Alias _
    "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'Vraca trenutni jezik koji se koristi
Public Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
'Vraca ime kompjutera
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Vraca User Name
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Vraca OEM indentifikator kodne stranice
Declare Function GetOEMCP Lib "kernel32" () As Long
'Vraca ANSI indentifikator kodne stranice
Declare Function GetACP Lib "kernel32" () As Long
'vraca informacije o tastauri
Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
'vraca vreme kojom kursor blinkuje
Declare Function GetCaretBlinkTime Lib "user32" () As Long
'Vraca Vreme za double click
Declare Function GetDoubleClickTime Lib "user32" () As Long
'Vraca vreme koje je proslo od startovanja windowsa
Declare Function GetTickCount Lib "kernel32.dll" () As Long
'copies memory
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source _
    As Any, ByVal Length As Long)
