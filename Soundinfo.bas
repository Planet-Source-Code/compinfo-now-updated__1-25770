Attribute VB_Name = "SoundInfo"
'device capabilities
Public Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
'Sound card determination
Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Public Type WAVEOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * 32
    dwFormats As Long
    wChannels As Integer
    dwSupport As Long
End Type
'Sound card determination
Public Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
'device capabilities
Public Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Public Type WAVEINCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * 32
    dwFormats As Long
    wChannels As Integer
End Type

'Supports 11.025 kHz, 8-bit, mono playback.
Public Const WAVE_FORMAT_1M08 = &H1
'Supports 11.025 kHz, 16-bit, mono playback.
Public Const WAVE_FORMAT_1M16 = &H4
'Supports 11.025 kHz, 8-bit, stereo playback.
Public Const WAVE_FORMAT_1S08 = &H2
'Supports 11.025 kHz, 16-bit, stereo playback.
Public Const WAVE_FORMAT_1S16 = &H8
'Supports 22.05 kHz, 8-bit, mono playback.
Public Const WAVE_FORMAT_2M08 = &H10
'Supports 22.05 kHz, 16-bit, mono playback.
Public Const WAVE_FORMAT_2M16 = &H40
'Supports 22.05 kHz, 8-bit, stereo playback.
Public Const WAVE_FORMAT_2S08 = &H20
'Supports 22.05 kHz, 16-bit, stereo playback.
Public Const WAVE_FORMAT_2S16 = &H80
'Supports 44.1 kHz, 8-bit, mono playback.
Public Const WAVE_FORMAT_4M08 = &H100
'Supports 44.1 kHz, 16-bit, mono playback.
Public Const WAVE_FORMAT_4M16 = &H400
'Supports 44.1 kHz, 8-bit, stereo playback.
Public Const WAVE_FORMAT_4S08 = &H200
'Supports 44.1 kHz, 16-bit, stereo playback.
Public Const WAVE_FORMAT_4S16 = &H800

'Supports separate left and right channel volumes.
Public Const WAVECAPS_LRVOLUME = 8
'Supports pitch control.
Public Const WAVECAPS_PITCH = 1
'Supports playback rate control.
Public Const WAVECAPS_PLAYBACKRATE = 2
'Supports returning of sample-accurate position information.
Public Const WAVECAPS_SAMPLEACCURATE = 32
'Supports synchronous playback -- i.e., it will block while playing buffered audio.
Public Const WAVECAPS_SYNC = 16
'Supports volume control.
Public Const WAVECAPS_VOLUME = 4
Public Sub SoundCard()
Dim b As Boolean
lRet = waveInGetNumDevs
ret = waveOutGetNumDevs()
If ret > 0 Then
    FrmMain.lbl70.Caption = "Sound Card: Present"
Else
    FrmMain.lbl70.Caption = "Sound Card: Not Present"
    Exit Sub
End If
FrmMain.List4.AddItem "Sound Card settings:"
WinEnv = String(145, Chr(0))
WinEnv = Left$(WinEnv, GetEnvironmentVariable("BLASTER", WinEnv, 145))
ret = InStr(1, WinEnv, "A", vbBinaryCompare)
StrTmp = mid$(WinEnv, ret + 1, Len(WinEnv) - ret - 1)
lRet = InStr(1, StrTmp, " ", vbBinaryCompare)
StrTmp = mid$(WinEnv, ret + 1, lRet - 1)
FrmMain.List4.AddItem "     IO Port:  " + StrTmp
ret = InStr(1, WinEnv, "I", vbBinaryCompare)
StrTmp = mid$(WinEnv, ret + 1, Len(WinEnv) - ret - 1)
lRet = InStr(1, StrTmp, " ", vbBinaryCompare)
StrTmp = mid$(WinEnv, ret + 1, lRet - 1)
FrmMain.List4.AddItem "     IRQ number:  " + StrTmp
ret = InStr(1, WinEnv, "D", vbBinaryCompare)
StrTmp = mid$(WinEnv, ret + 1, Len(WinEnv) - ret - 1)
lRet = InStr(1, StrTmp, " ", vbBinaryCompare)
StrTmp = mid$(WinEnv, ret + 1, lRet - 1)
FrmMain.List4.AddItem "     DMA number:  " + StrTmp
FrmMain.List4.AddItem ""
lRet = -1
ret = waveInGetNumDevs()
Dim win As WAVEINCAPS
b = False
For lRet = -1 To ret - 1
    If b = False Then
        b = True
    Else
        FrmMain.List4.AddItem ""
    End If
    FrmMain.List4.AddItem "Wave Input Device " + CStr(lRet + 1)
    ret = waveInGetDevCaps(lRet, win, Len(win))
    FrmMain.List4.AddItem "     General info:"
    FrmMain.List4.AddItem "          Device Name:  " + win.szPname
    FrmMain.List4.AddItem "          Driver version:  " + CStr(NumMan.HiByte(win.vDriverVersion)) + "." + CStr(NumMan.LoByte(win.vDriverVersion))
    If win.wChannels = 1 Then FrmMain.List4.AddItem "          Number of audio chanels:  " + CStr(win.wChannels) + " (mono device)"
    If win.wChannels = 2 Then FrmMain.List4.AddItem "          Number of audio chanels:  " + CStr(win.wChannels) + " (stereo device)"
    If win.wChannels > 2 Then FrmMain.List4.AddItem "          Number of audio chanels:  " + CStr(win.wChannels)
    FrmMain.List4.AddItem "          Manufacturer ID:  " + CStr(win.wMid)
    FrmMain.List4.AddItem "          Product ID:  " + CStr(win.wPid)
    FrmMain.List4.AddItem ""
    FrmMain.List4.AddItem "     Specific wave formats supports:"
    If (win.dwFormats And WAVE_FORMAT_1M08) = WAVE_FORMAT_1M08 Then
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 8-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 8-bit, mono playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_1M16) = WAVE_FORMAT_1M16 Then
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 16-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 16-bit, mono playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_1S08) = WAVE_FORMAT_1S08 Then
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 8-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 8-bit, stereo playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_1S16) = WAVE_FORMAT_1S16 Then
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 16-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 16-bit, stereo playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_2M08) = WAVE_FORMAT_2M08 Then
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 8-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 8-bit, mono playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_2M16) = WAVE_FORMAT_2M16 Then
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 16-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 16-bit, mono playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_2S08) = WAVE_FORMAT_2S08 Then
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 8-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 8-bit, stereo playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_2S16) = WAVE_FORMAT_2S16 Then
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 16-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 16-bit, stereo playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_4M08) = WAVE_FORMAT_4M08 Then
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 8-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 8-bit, mono playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_4M16) = WAVE_FORMAT_4M16 Then
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 16-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 16-bit, mono playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_4S08) = WAVE_FORMAT_4S08 Then
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 8-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 8-bit, stereo playback:  No"
    End If
    If (win.dwFormats And WAVE_FORMAT_4S16) = WAVE_FORMAT_4S16 Then
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 16-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 16-bit, stereo playback:  No"
    End If
Next
lRet = -1
ret = waveOutGetNumDevs()
Dim w As WAVEOUTCAPS
For lRet = -1 To ret - 1
    FrmMain.List4.AddItem ""
    FrmMain.List4.AddItem "Wave Output Device " + CStr(lRet + 1)
    ret = waveOutGetDevCaps(lRet, w, Len(w))
    FrmMain.List4.AddItem "     General info:"
    FrmMain.List4.AddItem "          Device Name:  " + w.szPname
    FrmMain.List4.AddItem "          Driver version:  " + CStr((w.vDriverVersion And &HFF00) \ 256) + "." + CStr(w.vDriverVersion And &HFF)
    If w.wChannels = 1 Then FrmMain.List4.AddItem "          Number of audio chanels:  " + CStr(w.wChannels) + " (mono device)"
    If w.wChannels = 2 Then FrmMain.List4.AddItem "          Number of audio chanels:  " + CStr(w.wChannels) + " (stereo device)"
    If w.wChannels > 2 Then FrmMain.List4.AddItem "          Number of audio chanels:  " + CStr(w.wChannels)
    FrmMain.List4.AddItem "          Manufacturer ID:  " + CStr(w.wMid)
    FrmMain.List4.AddItem "          Product ID:  " + CStr(w.wPid)
    FrmMain.List4.AddItem ""
    FrmMain.List4.AddItem "     Specific wave formats supports:"
    If (w.dwFormats And WAVE_FORMAT_1M08) = WAVE_FORMAT_1M08 Then
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 8-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 8-bit, mono playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_1M16) = WAVE_FORMAT_1M16 Then
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 16-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 16-bit, mono playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_1S08) = WAVE_FORMAT_1S08 Then
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 8-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 8-bit, stereo playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_1S16) = WAVE_FORMAT_1S16 Then
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 16-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 11.025 kHz, 16-bit, stereo playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_2M08) = WAVE_FORMAT_2M08 Then
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 8-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 8-bit, mono playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_2M16) = WAVE_FORMAT_2M16 Then
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 16-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 16-bit, mono playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_2S08) = WAVE_FORMAT_2S08 Then
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 8-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 8-bit, stereo playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_2S16) = WAVE_FORMAT_2S16 Then
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 16-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 22.05 kHz, 16-bit, stereo playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_4M08) = WAVE_FORMAT_4M08 Then
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 8-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 8-bit, mono playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_4M16) = WAVE_FORMAT_4M16 Then
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 16-bit, mono playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 16-bit, mono playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_4S08) = WAVE_FORMAT_4S08 Then
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 8-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 8-bit, stereo playback:  No"
    End If
    If (w.dwFormats And WAVE_FORMAT_4S16) = WAVE_FORMAT_4S16 Then
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 16-bit, stereo playback:  Yes"
    Else
        FrmMain.List4.AddItem "          Supports 44.1 kHz, 16-bit, stereo playback:  No"
    End If
    FrmMain.List4.AddItem "     "
    FrmMain.List4.AddItem "     Device capabilities:"
    If w.dwSupport And WAVECAPS_PITCH Then
        FrmMain.List4.AddItem "          Pitch control:  Yes" + CStr(Chr(WAVECAPS_SAMPLEACCURATE))
    Else
        FrmMain.List4.AddItem "          Pitch control:  No"
    End If
    If (w.dwSupport And WAVECAPS_VOLUME) = WAVECAPS_VOLUME Then
        FrmMain.List4.AddItem "          Volume control:  Yes"
    Else
        FrmMain.List4.AddItem "          Volume control:  No"
    End If
    If (w.dwSupport And WAVECAPS_SYNC) = WAVECAPS_SYNC Then
        FrmMain.List4.AddItem "          Synchronous operation:  Yes"
    Else
        FrmMain.List4.AddItem "          Synchronous operation:  No"
    End If
    If w.dwSupport And WAVECAPS_PLAYBACKRATE Then
        FrmMain.List4.AddItem "          Playback rate control:  Yes"
    Else
        FrmMain.List4.AddItem "          Playback rate control:  No"
    End If
    If (w.dwSupport And WAVECAPS_SAMPLEACCURATE) = WAVECAPS_SAMPLEACCURATE Then
        FrmMain.List4.AddItem "          Sample-accurate position information:  Yes"
    Else
        FrmMain.List4.AddItem "          Sample-accurate position information:  No"
    End If
    If w.wChannels = 2 Then
        If (w.dwSupport And WAVECAPS_LRVOLUME) = WAVECAPS_LRVOLUME Then
            FrmMain.List4.AddItem "          Separate left and right channel volumes:  Yes"
        Else
            FrmMain.List4.AddItem "          Separate left and right channel volumes:  No"
        End If
    Else
    End If
Next
End Sub
