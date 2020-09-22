Attribute VB_Name = "HardInfo"
'Screen Resolution
Public Const CURVECAPS = 28     '  /* Curve capabilities                       */
Public Const LINECAPS = 30     '   /* Line capabilities                        */
Public Const POLYGONALCAPS = 32   '/* Polygonal capabilities                   */
Public Const TEXTCAPS = 34      '  /* Text capabilities                        */
Public Const CLIPCAPS = 36     '   /* Clipping capabilities                    */
Public Const RASTERCAPS = 38       '/* Bitblt capabilities
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DRIVERVERSION = 0
Public Const HORZRES = 8
Public Const VERTRES = 10
Public Const BITSPIXEL = 12
Public Const HORZSIZE = 4
Public Const VERTSIZE = 6

'/* Clipping Capabilities */
Public Const CP_RECTANGLE = 1         '/* Output clipped to rects          */
Public Const CP_REGION = 2         '   /* obsolete                         */


'/* Curve Capabilities */
Public Const CC_CIRCLES = 1  '   /* Can do circles                   */
Public Const CC_PIE = 2             '  /* Can do pie wedges                */
Public Const CC_CHORD = 4            ' /* Can do chord arcs                */
Public Const CC_ELLIPSES = 8        '  /* Can do ellipese                  */
Public Const CC_WIDE = 16      ' /* Can do wide lines                */
Public Const CC_STYLED = 32          ' /* Can do styled lines              */
Public Const CC_WIDESTYLED = 64      ' /* Can do wide styled lines         */
Public Const CC_INTERIORS = 128      ' /* Can do interiors                 */
Public Const CC_ROUNDRECT = 256        '/*                                  */

'/* Line Capabilities */
Public Const LC_POLYLINE = 2      '   /* Can do polylines                 */
Public Const LC_MARKER = 4        '   /* Can do markers                   */
Public Const LC_POLYMARKER = 8    '   /* Can do polymarkers               */
Public Const LC_WIDE = 16         '  /* Can do wide lines                */
Public Const LC_STYLED = 32       '  /* Can do styled lines              */
Public Const LC_WIDESTYLED = 64   '  /* Can do wide styled lines         */
Public Const LC_INTERIORS = 128    ' /* Can do interiors                 */

'/* Polygonal Capabilities */
Public Const PC_POLYGON = 1        '   /* Can do polygons                  */
Public Const PC_RECTANGLE = 2      '   /* Can do rectangles                */
Public Const PC_WINDPOLYGON = 4    '   /* Can do winding polygons          */
Public Const PC_SCANLINE = 8       '   /* Can do scanlines                 */
Public Const PC_WIDE = 16           '  /* Can do wide borders              */
Public Const PC_STYLED = 32         '  /* Can do styled borders            */
Public Const PC_WIDESTYLED = 64     '  /* Can do wide styled borders       */
Public Const PC_INTERIORS = 128      ' /* Can do interiors                 */
Public Const PC_POLYPOLYGON = 256    ' /* Can do polypolygons              */
Public Const PC_PATHS = 512          ' /* Can do paths                     */

'/* Text Capabilities */
Public Const TC_OP_CHARACTER = 1           '  /* Can do OutputPrecision   CHARACTER      */
Public Const TC_OP_STROKE = 2              '  /* Can do OutputPrecision   STROKE         */
Public Const TC_CP_STROKE = 4              '  /* Can do ClipPrecision     STROKE         */
Public Const TC_CR_90 = 8                  '  /* Can do CharRotAbility    90             */
Public Const TC_CR_ANY = 16                '  /* Can do CharRotAbility    ANY            */
Public Const TC_SF_X_YINDEP = 32           '  /* Can do ScaleFreedom      X_YINDEPENDENT */
Public Const TC_SA_DOUBLE = 64             '  /* Can do ScaleAbility      DOUBLE         */
Public Const TC_SA_INTEGER = 128            '  /* Can do ScaleAbility      INTEGER        */
Public Const TC_SA_CONTIN = 256            '  /* Can do ScaleAbility      CONTINUOUS     */
Public Const TC_EA_DOUBLE = 512            '  /* Can do EmboldenAbility   DOUBLE         */
Public Const TC_IA_ABLE = 1024              '  /* Can do ItalisizeAbility  ABLE           */
Public Const TC_UA_ABLE = 2048              '  /* Can do UnderlineAbility  ABLE           */
Public Const TC_SO_ABLE = 4096             '  /* Can do StrikeOutAbility  ABLE           */
Public Const TC_RA_ABLE = 8192             '  /* Can do RasterFontAble    ABLE           */
Public Const TC_VA_ABLE = 16384             '  /* Can do VectorFontAble    ABLE           */
Public Const TC_SCROLLBLT = 65536           '  /* Don't do text scroll with blt           */

'/* Raster Capabilities */
Public Const RC_BITBLT = 1         '   /* Can do standard BLT.
Public Const RC_BANDING = 2        '       /* Device requires banding support  */
Public Const RC_SCALING = 4        '       /* Device requires scaling support  */
Public Const RC_BITMAP64 = 8       '       /* Device can support >64K bitmap   */
Public Const RC_GDI20_OUTPUT = 16   '      /* has 2.0 output calls         */
'Public Const RC_GDI20_STATE = 32    '/*Unknown*/
Public Const RC_SAVEBITMAP = 64     '    /*saves bitmaps locally*/
Public Const RC_DI_BITMAP = 128       '     /* supports DIB to memory       */
Public Const RC_PALETTE = 256         '      /* supports a palette           */
Public Const RC_DIBTODEV = 512        '      /* supports DIBitsToDevice      */
Public Const RC_BIGFONT = 1024         '      /* supports >64K fonts          */
Public Const RC_STRETCHBLT = 2048       '      /* supports StretchBlt          */
Public Const RC_FLOODFILL = 4096       '      /* supports FloodFill           */
Public Const RC_STRETCHDIB = 8192      '      /* supports StretchDIBits       */
'Public Const RC_OP_DX_OUTPUT = 16384    '/*Unknown*/
Public Const RC_DEVBITS = 32768         '/*Supports device bitmaps*/


Private Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type

Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Sub GCard()
Dim tDevMode As DEVMODE, Id As Long, hDC As Long
FrmMain.lbl50.Caption = "Graphic card:  " + GetSysIni("boot.description", "display.drv")
hDC = GetDC(GetDesktopWindow())
FrmMain.lbl51.Caption = "Current screen resolution:  " + CStr(GetDeviceCaps(hDC, HORZRES)) + " x " + CStr(GetDeviceCaps(hDC, VERTRES))
FrmMain.lbl52.Caption = "Current color depth:  " + CStr(GetDeviceCaps(hDC, BITSPIXEL)) + " bits per pixel"
ret = GetDeviceCaps(hDC, HORZSIZE)
FrmMain.lbl54.Caption = "Physical screen width:  " + CStr(ret) + " mm, " + CStr(CLng(ret / 25)) + " in"
lret = GetDeviceCaps(hDC, VERTSIZE)
FrmMain.lbl55.Caption = "Physical screen height:  " + CStr(lret) + " mm, " + CStr(CLng(lret / 25)) + " in"
FrmMain.lbl56.Caption = "Recommended monitor size:  " + CStr(lret + ret) + " mm, " + CStr(CLng(ret / 25) + CLng(lret / 25)) + " in"
FrmMain.lbl53.Caption = "Supported Video Modes:"
Id = 0
ID1 = 50
Do
    ret = EnumDisplaySettings(0&, Id, tDevMode)
    If ret = 0 Then Exit Do
    With tDevMode
        FrmMain.List2.AddItem .dmPelsWidth & " pixels by " & .dmPelsHeight & " pixels, Color depth " & .dmBitsPerPel & " bits per pixel"
    End With
    Id = Id + 1
Loop
FrmMain.lbl57.Caption = "Video driver capabilities:"
FrmMain.List3.AddItem "Clipping Capabilities:"
If (GetDeviceCaps(hDC, CLIPCAPS) And CP_RECTANGLE) = CP_RECTANGLE Then
    FrmMain.List3.AddItem "     Can clip output to rectangle:  Yes"
Else
    FrmMain.List3.AddItem "     Can clip output to rectangle:  Yes"
End If
If (GetDeviceCaps(hDC, CLIPCAPS) And CP_REGION) = CP_REGION Then
    FrmMain.List3.AddItem "     Can clip output to region:  Yes"
Else
    FrmMain.List3.AddItem "     Can clip output to region:  No"
End If
FrmMain.List3.AddItem ""
FrmMain.List3.AddItem "Curve Capabilities:"
If (GetDeviceCaps(hDC, CURVECAPS) And CC_CIRCLES) = CC_CIRCLES Then
    FrmMain.List3.AddItem "     Can draw circles:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw circles:  No"
End If
If (GetDeviceCaps(hDC, CURVECAPS) And CC_PIE) = CC_PIE Then
    FrmMain.List3.AddItem "     Can draw pie wedges:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw pie wedges:  No"
End If
If (GetDeviceCaps(hDC, CURVECAPS) And CC_CHORD) = CC_CHORD Then
    FrmMain.List3.AddItem "     Can draw chord arcs:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw chord arcs:  No"
End If
If (GetDeviceCaps(hDC, CURVECAPS) And CC_ELLIPSES) = CC_ELLIPSES Then
    FrmMain.List3.AddItem "     Can draw ellipses:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw ellipses:  No"
End If
If (GetDeviceCaps(hDC, CURVECAPS) And CC_WIDE) = CC_WIDE Then
    FrmMain.List3.AddItem "     Can draw wide borders:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw wide borders:  No"
End If
If (GetDeviceCaps(hDC, CURVECAPS) And CC_STYLED) = CC_STYLED Then
    FrmMain.List3.AddItem "     Can draw styled borders:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw styled borders:  No"
End If
If (GetDeviceCaps(hDC, CURVECAPS) And CC_WIDESTYLED) = CC_WIDESTYLED Then
    FrmMain.List3.AddItem "     Can draw wide, styled borders:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw wide, styled borders:  No"
End If
If (GetDeviceCaps(hDC, CURVECAPS) And CC_INTERIORS) = CC_INTERIORS Then
    FrmMain.List3.AddItem "     Can draw interiors:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw interiors:  No"
End If
If (GetDeviceCaps(hDC, CURVECAPS) And CC_ROUNDRECT) = CC_ROUNDRECT Then
    FrmMain.List3.AddItem "     Can draw rounded rectangles:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw rounded rectangles:  No"
End If
FrmMain.List3.AddItem ""
FrmMain.List3.AddItem "Line Capabilities:"
If (GetDeviceCaps(hDC, LINECAPS) And LC_POLYLINE) = LC_POLYLINE Then
    FrmMain.List3.AddItem "     Can draw polylines:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw polylines:  No"
End If
If (GetDeviceCaps(hDC, LINECAPS) And LC_MARKER) = LC_MARKER Then
    FrmMain.List3.AddItem "     Can draw markers:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw markers:  No"
End If
If (GetDeviceCaps(hDC, LINECAPS) And LC_POLYMARKER) = LC_POLYMARKER Then
    FrmMain.List3.AddItem "     Can draw polymarkers:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw polymarkers:  No"
End If
If (GetDeviceCaps(hDC, LINECAPS) And LC_WIDE) = LC_WIDE Then
    FrmMain.List3.AddItem "     Can draw wide lines:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw wide lines:  No"
End If
If (GetDeviceCaps(hDC, LINECAPS) And LC_STYLED) = LC_STYLED Then
    FrmMain.List3.AddItem "     Can draw styled lines:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw styled lines:  No"
End If
If (GetDeviceCaps(hDC, LINECAPS) And LC_WIDESTYLED) = LC_WIDESTYLED Then
    FrmMain.List3.AddItem "     Can draw wide, styled lines:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw wide, styled lines:  No"
End If
If (GetDeviceCaps(hDC, LINECAPS) And LC_INTERIORS) = LC_INTERIORS Then
    FrmMain.List3.AddItem "     Can draw interiors:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw interiors:  No"
End If
FrmMain.List3.AddItem ""
FrmMain.List3.AddItem "Polygonal Capabilities:"
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_POLYGON) = PC_POLYGON Then
    FrmMain.List3.AddItem "     Can draw alternate-fill polygons:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw alternate-fill polygons:  No"
End If
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_RECTANGLE) = PC_RECTANGLE Then
    FrmMain.List3.AddItem "     Can draw rectangles:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw rectangles:  No"
End If
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_WINDPOLYGON) = PC_WINDPOLYGON Then
    FrmMain.List3.AddItem "     Can draw winding-fill polygons:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw winding-fill polygons:  No"
End If
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_SCANLINE) = PC_SCANLINE Then
    FrmMain.List3.AddItem "     Can draw a scanlines:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw a scanlines:  No"
End If
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_WIDE) = PC_WIDE Then
    FrmMain.List3.AddItem "     Can draw wide borders:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw wide borders:  No"
End If
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_STYLED) = PC_STYLED Then
    FrmMain.List3.AddItem "     Can draw styled borders:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw styled borders:  No"
End If
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_WIDESTYLED) = PC_WIDESTYLED Then
    FrmMain.List3.AddItem "     Can draw wide, styled borders:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw wide, styled borders:  No"
End If
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_INTERIORS) = PC_INTERIORS Then
    FrmMain.List3.AddItem "     Can draw interiors:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw interiors:  No"
End If
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_POLYPOLYGON) = PC_POLYPOLYGON Then
    FrmMain.List3.AddItem "     Can draw polypolygons:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw polypolygons:  No"
End If
If (GetDeviceCaps(hDC, POLYGONALCAPS) And PC_PATHS) = PC_PATHS Then
    FrmMain.List3.AddItem "     Can draw paths:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw paths:  No"
End If
FrmMain.List3.AddItem ""
FrmMain.List3.AddItem "Text Capabilities:"
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_OP_CHARACTER) = TC_OP_CHARACTER Then
    FrmMain.List3.AddItem "     Capable of character output precision:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of character output precision:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_OP_STROKE) = TC_OP_STROKE Then
    FrmMain.List3.AddItem "     Capable of stroke output precision:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of stroke output precision:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_CP_STROKE) = TC_CP_STROKE Then
    FrmMain.List3.AddItem "     Capable of stroke clip precision:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of stroke clip precision:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_CR_90) = TC_CR_90 Then
    FrmMain.List3.AddItem "     Capable of 90º character rotation:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of 90º character rotation:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_CR_ANY) = TC_CR_ANY Then
    FrmMain.List3.AddItem "     Capable of any angle character rotation:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of any angle character rotation:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_SF_X_YINDEP) = TC_SF_X_YINDEP Then
    FrmMain.List3.AddItem "     Capable of independent X-Y scaling:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of independent X-Y scaling:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_SA_DOUBLE) = TC_SA_DOUBLE Then
    FrmMain.List3.AddItem "     Capable of doubled character for scaling:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of doubled character for scaling:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_SA_INTEGER) = TC_SA_INTEGER Then
    FrmMain.List3.AddItem "     Capable of integer multiples character scaling:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of integer multiples character scaling:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_SA_CONTIN) = TC_SA_CONTIN Then
    FrmMain.List3.AddItem "     Any multiples for exact character scaling:  Yes"
Else
    FrmMain.List3.AddItem "     Any multiples for exact character scaling:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_EA_DOUBLE) = TC_EA_DOUBLE Then
    FrmMain.List3.AddItem "     Can draw double weighted characters:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw double weighted characters:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_IA_ABLE) = TC_IA_ABLE Then
    FrmMain.List3.AddItem "     Can italicize:  Yes"
Else
    FrmMain.List3.AddItem "     Can italicize:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_UA_ABLE) = TC_UA_ABLE Then
    FrmMain.List3.AddItem "     Can underline:  Yes"
Else
    FrmMain.List3.AddItem "     Can underline:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_SO_ABLE) = TC_SO_ABLE Then
    FrmMain.List3.AddItem "     Can draw strikeouts:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw strikeouts:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_RA_ABLE) = TC_RA_ABLE Then
    FrmMain.List3.AddItem "     Can draw raster fonts:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw raster fonts:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_VA_ABLE) = TC_VA_ABLE Then
    FrmMain.List3.AddItem "     Can draw vector fonts:  Yes"
Else
    FrmMain.List3.AddItem "     Can draw vector fonts:  No"
End If
If (GetDeviceCaps(hDC, TEXTCAPS) And TC_SCROLLBLT) = TC_SCROLLBLT Then
    FrmMain.List3.AddItem "     Cannot scroll using BitBlt:  Yes"
Else
    FrmMain.List3.AddItem "     Cannot scroll using BitBlt:  No"
End If
FrmMain.List3.AddItem ""
FrmMain.List3.AddItem "Raster Capabilities:"
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_BITBLT) = RC_BITBLT Then
    FrmMain.List3.AddItem "     Capable of transferring bitmaps:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of transferring bitmaps:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_BANDING) = RC_BANDING Then
    FrmMain.List3.AddItem "     Supports banding:  Yes"
Else
    FrmMain.List3.AddItem "     Supports banding:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_SCALING) = RC_SCALING Then
    FrmMain.List3.AddItem "     Supports scaling:  Yes"
Else
    FrmMain.List3.AddItem "     Supports scaling:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_BITMAP64) = RC_BITMAP64 Then
    FrmMain.List3.AddItem "     Supports bitmaps larger than 64 KB:  Yes"
Else
    FrmMain.List3.AddItem "     Supports bitmaps larger than 64 KB:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_GDI20_OUTPUT) = RC_GDI20_OUTPUT Then
    FrmMain.List3.AddItem "     Supports Windows 2.00:  Yes"
Else
    FrmMain.List3.AddItem "     Supports Windows 2.00:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_DI_BITMAP) = RC_DI_BITMAP Then
    FrmMain.List3.AddItem "     Supports DIBs:  Yes"
Else
    FrmMain.List3.AddItem "     Supports DIBs:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_PALETTE) = RC_PALETTE Then
    FrmMain.List3.AddItem "     Palette-based device:  Yes"
Else
    FrmMain.List3.AddItem "     Palette-based device:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_DIBTODEV) = RC_DIBTODEV Then
    FrmMain.List3.AddItem "     DIBs on device surface:  Yes"
Else
    FrmMain.List3.AddItem "     DIBs on device surface:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_BIGFONT) = RC_BIGFONT Then
    FrmMain.List3.AddItem "     Supports fonts larger than 64 KB:  Yes"
Else
    FrmMain.List3.AddItem "     Supports fonts larger than 64 KB:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_STRETCHBLT) = RC_STRETCHBLT Then
    FrmMain.List3.AddItem "     Stretch/Compress bitmaps:  Yes"
Else
    FrmMain.List3.AddItem "     Stretch/Compress bitmaps:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_FLOODFILL) = RC_FLOODFILL Then
    FrmMain.List3.AddItem "     Capable of performing flood fills:  Yes"
Else
    FrmMain.List3.AddItem "     Capable of performing flood fills:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_STRETCHDIB) = RC_STRETCHDIB Then
    FrmMain.List3.AddItem "     Stretch/Compress DIBs:  Yes"
Else
    FrmMain.List3.AddItem "     Stretch/Compress DIBs:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_SAVEBITMAP) = RC_SAVEBITMAP Then
    FrmMain.List3.AddItem "     Save bitmap locally:  Yes"
Else
    FrmMain.List3.AddItem "     Save bitmap locally:  No"
End If
If (GetDeviceCaps(hDC, RASTERCAPS) And RC_DEVBITS) = RC_DEVBITS Then
    FrmMain.List3.AddItem "     Supports device bitmaps:  Yes"
Else
    FrmMain.List3.AddItem "     Supports device bitmaps:  No"
End If
ReleaseDC GetDesktopWindow(), hDC
End Sub
Public Sub HardInfo95()

End Sub
Public Sub HardInfo98()

End Sub
Public Sub HardInfo2000()

End Sub
Public Sub HardInfoNT()

End Sub
