Attribute VB_Name = "Module1"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public sx As Integer, sy As Integer, ex As Integer
Public ey As Integer, volt As Single, curr As Single
Public plot As Integer, temp As String
Public DacOut As Integer
Public mk_ch As VbMsgBoxResult
Public Function ChrToVal(Str)
Dim Hb, Lb As Integer
    If Str <> "" Then
        'Hb = Asc(Mid$(Str, 2, 1))
        'Lb = Asc(Mid$(Str, 3, 1))
        'ChrToVal = (Hb * 256) + Lb
        ChrToVal = CInt(Mid$(Str, 2, 4))
    End If
End Function


