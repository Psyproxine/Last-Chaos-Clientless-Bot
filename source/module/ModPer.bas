Attribute VB_Name = "ModPer"
Public Function RandomAlphaNumString(ByVal intLen As Integer)
    Dim StrReturn As String
    Dim X As Integer
    Dim c As Byte
    Randomize
    For X = 1 To intLen
        c = Int(Rnd() * 127)
        If (c >= Asc("0") And c <= Asc("9")) Or _
           (c >= Asc("A") And c <= Asc("Z")) Or _
           (c >= Asc("a") And c <= Asc("z")) Then
            StrReturn = StrReturn & Chr(c)
        Else
            X = X - 1
        End If
    Next X
    RandomAlphaNumString = StrReturn
End Function


Public Sub check_per(data As String)
    Dim md5Test As MD5, tmpData As String, cmd() As String
    Set md5Test = New MD5
    cmd = Split(data, ":")
    If UBound(cmd) = 1 Then
        If Len(cmd(0)) = 50 Then
            tmpData = Left(md5Test.DigestStrToChar(data), 10)
            tmpData = md5Test.DigestStrToChar(tmpData)
            Chat Chr(3) & "0,12" & cmd(1)
            
        Else
            Not_Permission cmd(1)
        End If
    Else
        Not_Permission data
    End If
End Sub

Public Sub Not_Permission(txt As String)
    If CheckPerTime = 0 Then frmMain.ucMain.ShowView "Permission"
    CheckPerTime = 0
    Connecting False
    Chat Chr(3) & "0,7" & txt
    Chat Chr(3) & "0,4 คุณสามารถเชครายละเอียดได้ที่ http://www.positron.in.th"
End Sub
