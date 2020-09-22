Attribute VB_Name = "NumMan"
Public Function HiByte(ByVal w) As Byte
   If w And &H8000 Then
      HiByte = &H80 Or ((w And &H7FFF) \ &HFF)
   Else
      HiByte = w \ 256
    End If
End Function
Public Function LoByte(ByVal w) As Byte
 LoByte = w And &HFF
End Function
Public Function LoWord(dw As Long) As Integer
  If dw And &H8000& Then
      LoWord = &H8000 Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
   End If
End Function
Public Function HiWord(dw As Long) As Integer
 If dw And &H80000000 Then
      HiWord = (dw \ 65535) - 1
 Else
    HiWord = dw \ 65535
 End If
End Function
Public Function MakeInt(ByVal LoByte As Byte, ByVal HiByte As Byte) As Integer
MakeInt = ((HiByte * &H100) + LoByte)
End Function
Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
MakeLong = ((HiWord * &H10000) + LoWord)
End Function
Public Function ULarge2Dbl(ul As LARGE_INTEGER) As Double
    Dim ld As Double
    Dim hd As Double
    
    If ul.LowPart < 0 Then
        ld = 2147483648#
        ul.LowPart = Abs(ul.LowPart)
        ld = ld + 2147483648# - CDbl(ul.LowPart)
    Else
        ld = CDbl(ul.LowPart)
    End If
    
    If ul.HighPart < 0 Then
        hd = 2147483648#
        ul.HighPart = Abs(ul.HighPart)
        hd = hd + 2147483648# - CDbl(ul.HighPart)
    Else
        hd = CDbl(ul.HighPart)
    End If
    ULarge2Dbl = ld + (hd * 4294967296#)
End Function
