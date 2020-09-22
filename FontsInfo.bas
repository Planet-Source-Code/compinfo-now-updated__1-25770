Attribute VB_Name = "FontsInfo"
Public Const SPI_GETFONTSMOOTHING = 74
Public Sub FontsSmooth()
SystemParamsLong SPI_GETFONTSMOOTHING, 0, ret, 0
If ret Then
    FrmMain.lbl60.Caption = "Fonts smoothing: Enabled"
Else
    FrmMain.lbl60.Caption = "Fonts smoothing: Disabled"
End If
FrmMain.lbl61.Caption = "Number of fonts installed:  " + CStr(Screen.FontCount)
FrmMain.lbl62.Caption = "Fonts installed:  "
For ret = 0 To Screen.FontCount - 1
    FrmMain.List1.AddItem Screen.Fonts(ret) + " font", ret
Next
End Sub
