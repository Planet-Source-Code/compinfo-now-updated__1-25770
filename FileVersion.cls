VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   Module          : cShellVersion
'   Description     : Returns the Shell DLL Versions
'   Author          : C. Eswar Santhosh
'   Last Updated    : 14th February, 2000.
'   Notes           : Many Shell DLLs export the function DllGetVersion. These include Shell32.dll, ShlWapi.dll, Comctl32.dll
'                     ShDocVw.dll etc.,
'
'   Copyright Info  :
'
'   This Class module is provided AS-IS. This Class module can be used as a part of a compiled
'   executable whether freeware or not. This Class module may not be posted to any web site
'   or BBS or any redistributable media like CD-ROM without the consent of the author.
'
'   Web Site : http://eswar_santhosh.tripod.com
'
'   e-mail   : eswar_santhosh@yahoo.com
'
'   Revision History :
'
'
'-------------------------------------------------------------------------------------------------------------

Option Explicit
                                                    
'
' API Constants
'
Private Const NOERROR As Long = 0                       ' Result

' Format Message Constants
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

' For Version Information
Private Const MAX_PATH As Long = 260
Private Const VOS_NT As Long = &H40000

'
' API Types
'
Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

' Supported by Version 5.0 or higher of Shlwapi.dll
Private Type DLLVERSIONINFO2
    Info1 As DLLVERSIONINFO
    dwFlags As Long             ' Currently, reserved and must be set to zero.
    ullVersion As Currency      ' Encoding of the 4 Bytes : [Major] [Minor] [Build] [QFE]
End Type

' For Older guys, this is the one that works
Private Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFileType As Long             '  e.g. VFT_DRIVER
        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
End Type

'
' API Declarations (This direct declarations avoid using the unreliable CreateThread to call the function)
'
Private Declare Function FormatMessage Lib "kernel32" Alias _
"FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, _
ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
ByVal lpBuffer As String, ByVal nSize As Long, _
Arguments As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

' For the Classic approach
Private Declare Function GetSystemDirectory Lib "kernel32" _
Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) _
As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) _
As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, _
ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long

' The win32api.txt declaration is wrong.
Private Declare Function VerQueryValue Lib "version.dll" _
Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
lplpBuffer As Any, _
puLen As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (dest As Any, ByVal Source As Long, ByVal Length As Long)

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
   (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

'
' Local Variables
'
Dim mCompanyName As String         '
Dim mFileDescription As String     '
Dim mFileVersion As String
Dim mInternalName As String        '
Dim mLegalCopyright As String      '
Dim mOriginalFileName As String    '
Dim mProductName As String         '
Dim mProductVersion As String      '
Dim md As Date
Dim mt As Date

Dim arrInfo() As Byte
Dim strLang As String
Dim lInfosize As Long
Dim lpInfoBlock As Long

Dim mresult

Dim m_Major As Long             ' Major Version
Dim m_Minor As Long             ' Minor Version
Dim m_BuildNumber As Long       ' Build Number
Dim m_QFEVersion As Long        ' QFE Number
Dim m_PlatformID As String ' Platform for which the DLL was written
Public Property Get FileDate() As String
    FileDate = md
End Property
Public Property Get FileTime() As String
    FileTime = mt
End Property
Public Property Get OriginalFileName() As String
    OriginalFileName = mOriginalFileName
End Property
Public Property Get FileDescription() As String
    FileDescription = mFileDescription
End Property
Public Property Get CompanyName() As String
    CompanyName = mCompanyName
End Property
Public Property Get InternalName() As String
    InternalName = mInternalName
End Property
Public Property Get LegalCopyright() As String
    LegalCopyright = mLegalCopyright
End Property
Public Property Get ProductName() As String
    ProductName = mProductName
End Property
Public Property Get ProductVersion() As String
    ProductVersion = mProductVersion
End Property
Public Property Get FileVersion() As String
    FileVersion = mFileVersion
End Property
Public Property Get MajorVersion() As Long
Attribute MajorVersion.VB_Description = "Returns the Major Version of the Shell DLL"
    MajorVersion = m_Major
End Property
Public Property Get MinorVersion() As Long
Attribute MinorVersion.VB_Description = "Returns the Minor Version of the Shell DLL"
    MinorVersion = m_Minor
End Property
Public Property Get BuildNumber() As Long
Attribute BuildNumber.VB_Description = "Build Number of the Shell DLL"
    BuildNumber = m_BuildNumber
End Property
Public Property Get QFEVersion() As Long
Attribute QFEVersion.VB_Description = "QFE Version of the Shell DLL if supported."
    QFEVersion = m_QFEVersion
End Property
Public Property Get Platform() As String
    Platform = m_PlatformID
End Property
Public Sub GetVersionClassic(ByVal LibName As String)
'
' In this case, these DLLs have only one folder where they reside, the System Folder
'
Dim mBuffer As String, mVerSize As Long, Dummy As Long, mVerInfo() As Byte
Dim mFixedFileInfo As VS_FIXEDFILEINFO, ptrBufferAddress As Long, lenBlock As Long

mBuffer = Space$(MAX_PATH)

GetSystemDirectory mBuffer, Len(mBuffer)

mBuffer = Left$(mBuffer, InStr(mBuffer, Chr$(0)) - 1)
mBuffer = mBuffer & "\" & LibName

If Dir(mBuffer) = "" Then
    Err.Raise 28001, App.Title & "File Version", "File was not found : " & LibName
    Exit Sub
End If

md = DateValue(FileDateTime(mBuffer))
mt = TimeValue(FileDateTime(mBuffer))

' Dummy will be set to zero by the function
mVerSize = GetFileVersionInfoSize(mBuffer, Dummy)

If mVerSize = 0 Then      ' This error will *never* happen
    Err.Raise 28002, App.Title & "File Version", "Version Information Resource not found"
End If

ReDim mVerInfo(mVerSize)

GetFileVersionInfo mBuffer, 0&, mVerSize, mVerInfo(0)

' Request the Root Block which returns a VS_FIXED_FILE_INFO Structure
VerQueryValue mVerInfo(0), "\", ptrBufferAddress, lenBlock

' Now ptrBufferAddress contains the Address of the VS_FIXED_FILE_INFO Block.
CopyMemory mFixedFileInfo, ByVal ptrBufferAddress, LenB(mFixedFileInfo)

' We are interested in the File Version members only.
' Get the File Version from them
CopyMemory m_Major, ByVal VarPtr(mFixedFileInfo.dwFileVersionMS) + 2, 2&
CopyMemory m_Minor, ByVal VarPtr(mFixedFileInfo.dwFileVersionMS), 2&
CopyMemory m_BuildNumber, ByVal VarPtr(mFixedFileInfo.dwFileVersionLS) + 2, 2&
CopyMemory m_QFEVersion, ByVal VarPtr(mFixedFileInfo.dwFileVersionLS), 2&
' The other thing is to get the Platform ID.
If mFixedFileInfo.dwFileOS = VER_PLATFORM_WIN32_WINDOWS Then
    m_PlatformID = "Windows 9x"
ElseIf mFixedFileInfo.dwFileOS = VER_PLATFORM_WIN32_NT Then
    m_PlatformID = "Windows NT"
Else
    m_PlatformID = "Windows 3.xx"
End If


End Sub
Private Function APIError(ByVal ErrNumber As Long) As String
'
' Returns the Error Message for the API Error.
'
Dim strMessage As String, lenMessage As Long

strMessage = Space$(1024)
lenMessage = FormatMessage(FORMAT_MESSAGE_IGNORE_INSERTS Or FORMAT_MESSAGE_FROM_SYSTEM, _
vbNull, ErrNumber, 0, strMessage, Len(strMessage), vbNull)

If lenMessage > 0 Then APIError = Left$(strMessage, lenMessage)
End Function

Public Function GetFileInfo(ByRef inFileSpec As String) As Boolean
    On Error Resume Next
    
    GetFileInfo = False                        ' Assume first
    
    Dim lpHandle As Long
    lpHandle = 0
    lInfosize = GetFileVersionInfoSize(inFileSpec, lpHandle)
    If lInfosize = 0 Then
        Exit Function
    End If

    ReDim arrInfo(lInfosize)
    mresult = GetFileVersionInfo(inFileSpec, 0&, lInfosize, arrInfo(0))
    If mresult = 0 Then
        Exit Function
    End If

      ' Call VerQueryValue to return selected version info pieces
      ' Note the 2nd argument is for pointing a zero-terminated string such as
      ' "\VarFileInfo\Translation" and "\StringFileInfo\lang-codepage\string-name"
      ' where lang-codepage has to be established from the first four bytes in the
      ' strFileInfoStringed info and string-name is one of predefined string names.
      ' The 3rd argument is a pointer to the requested info block.
      '----------------------------------------------------------------
      'lang-codepage data
      '  from first 4 bytes of "\VarFileInfo\Translation" piece of data
      '----------------------------------------------------------------
    mresult = VerQueryValue(arrInfo(0), "\VarFileInfo\Translation", lpInfoBlock, lInfosize)
    If mresult = 0 Then
        Exit Function
    End If

      ' To access the block info, we first move the obtained block info from
      ' strFileInfoString pointed by lpVerBlocke into an array.
    Dim arrTemp(4) As Byte
    Dim lngTemp As Long
    MoveMemory arrTemp(0), lpInfoBlock, lInfosize
      ' Word of language + word of codepage in the lang structure. High-low.
    lngTemp = arrTemp(2) + arrTemp(3) * &H100 + _
            arrTemp(0) * &H10000 + arrTemp(1) * &H1000000
    strLang = Hex(lngTemp)
      ' Length of lang-codepage is 8; pad zeros in front in case length < 8
      ' e.g. "04094E4" becomes "004094E4"
    Do While Len(strLang) < 8
        strLang = "0" & strLang
    Loop
      
     '----------------------------------------------------------------
     ' "\StringFileInfo\lang-codepage\string-name" pieces of data:
     '----------------------------------------------------------------
    mCompanyName = GetInfoBlock("CompanyName")
    mFileDescription = GetInfoBlock("FileDescription")
    mFileVersion = GetInfoBlock("FileVersion")
    mInternalName = GetInfoBlock("InternalName")
    mLegalCopyright = GetInfoBlock("LegalCopyright")
    mOriginalFileName = GetInfoBlock("OriginalFileName")
    mProductName = GetInfoBlock("ProductName")
    mProductVersion = GetInfoBlock("ProductVersion")
    If Trim(mProductVersion) <> "" Then
         GetFileInfo = True
    End If
End Function
Private Function GetInfoBlock(ByVal inStringName As String) As String
    GetInfoBlock = ""
     ' We don't know how long, fill the string with some char 0 first
    Dim strFileInfoString As String
    strFileInfoString = String(128, 0)
    mresult = VerQueryValue(arrInfo(0), "\StringFileInfo\" & strLang & _
        "\" & inStringName, lpInfoBlock, lInfosize)
    If mresult > 0 Then
        lstrcpy strFileInfoString, lpInfoBlock
        GetInfoBlock = mid(strFileInfoString, 1, InStr(strFileInfoString, vbNullChar) - 1)
    End If
End Function
