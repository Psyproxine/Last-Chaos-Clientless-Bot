Attribute VB_Name = "ModBytes"
Option Explicit
Public Walking As Integer, Walking_start As Coord, Walking_stop As Coord, cur As Coord, Walking_dist As Integer
Public tmpMap As String * 8, rndSend As Long
Public Function MakePort(ByVal rawPort As String) As Long
On Error GoTo z
    Dim tmpMP As Long
    tmpMP = CLng(Format("&H" & ReverseHex(ChrtoHex(rawPort))))
    If Len(rawPort) = 2 And tmpMP > 65536 Then tmpMP = tmpMP - 65536
    MakePort = tmpMP
Exit Function
z:
MakePort = 1
End Function
Public Function GetLong(rawPort As String) As Long
On Error GoTo errie
    Dim tst As Long
    Dim i As Integer
    tst = CLng(Asc(Mid(rawPort, 1, 1))) + (CLng(Asc(Mid(rawPort, 2, 1))) * 256) + (CLng(Asc(Mid(rawPort, 3, 1))) * 65536)
    If Asc(Mid(rawPort, 4, 1)) < 8 Then
        For i = 1 To Asc(Mid(rawPort, 4, 1))
            tst = tst + 16777216
        Next
    End If
    GetLong = tst
    Exit Function
errie:
    GetLong = 0
End Function

Public Function GetPos(rawPort As String) As Long
'On Error GoTo errie
Dim y As Long, X As Long
y = 12100
X = 1134770944
    Dim tst As Long
    Dim i As Integer
    'tst = MakePort(rawPort)
'    GetPos = (tst - x) / y
    Exit Function
errie:
    GetPos = 0
End Function

Function ChrtoHex(inString As String) As String ' "AB" > 4142
    Dim tstr As String, i As Long
    For i = 1 To Len(inString)
        If Len(Hex(Asc(Mid(inString, i, 1)))) = 1 Then tstr = tstr & "0" & Hex(Asc(Mid(inString, i, 1))) Else tstr = tstr & Hex(Asc(Mid(inString, i, 1)))
    Next
    ChrtoHex = tstr
End Function

Function HextoChr(inString As String) As String
    Dim tstr As String, i As Long
    For i = 1 To Len(inString) Step 2
        tstr = tstr & Chr("&H" & Mid(inString, i, 2))
    Next
    HextoChr = tstr
End Function

Function ReverseHex(inHex As String) As String
    If (Len(inHex) Mod 2) <> 0 Then Exit Function
    Dim i As Long, tChr As String
    For i = 1 To Len(inHex) Step 2
        tChr = tChr & Mid(inHex, Len(inHex) - i, 2)
    Next
    ReverseHex = tChr
End Function

Function ReverseByte(inHex As String) As String
    Dim i As Long, tChr As String
    For i = 1 To Len(inHex)
        tChr = tChr & Mid(inHex, Len(inHex) - (i - 1), 1)
    Next
    ReverseByte = tChr
End Function

Public Function Make2Byte(ByVal rawLong As Long) As String
On Error GoTo out
Make2Byte = Chr(rawLong Mod 256) + Chr(Int(rawLong / 256))
out:
End Function

Public Sub SendPacket(packet As String, id As Integer)
On Error GoTo out:
Dim tmp As String
    If (frmBot.Winsock.State = 7) Then
            If (id = 0) Then
                tmp = packet
            ElseIf id = 1 Then
                tmp = Chr(&H81) & Chr(&H1) & Rnd2Byte() & Chr(&H0) & Chr(&H0) & Chr(tmpPack) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(Len(packet)) & packet
            End If
            frmBot.Winsock.SendData tmp
            print_packet tmp, " <--- "
    End If
out:
End Sub

Public Function MakeString(rawString As String) As String
Dim str1 As String
On Error GoTo ErrTrapper:
MakeString = Left(rawString, InStr(rawString, Chr(0)) - 1)
Exit Function
ErrTrapper:
If Err.Number = 5 Then
Err.Clear
'Stat "Destroy program from someone, Ignore..." + vbCrLf
End If
End Function

Public Function Rnd2Byte() As String
    Dim X(1) As Byte
    rndSend = rndSend + 1
    If rndSend > 60000 Then rndSend = 1
    Rnd2Byte = Make2Byte(rndSend)
End Function
    
Public Function MakeLong(data As String) As Long
    MakeLong = MakePort(ReverseByte(data))
End Function

Function LngToChr(inLong As Long) As String
    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long
    c1 = inLong Mod 256
    c2 = Int(inLong / 256) Mod 256
    c3 = Int(inLong / 65536) Mod 256
    c4 = Int(inLong / 16777216) Mod 256
    LngToChr = Chr(c4) & Chr(c3) & Chr(c2) & Chr(c1)
End Function

Public Function convert_coord(data As String) As Coord
Dim Pt As Coord
    Pt.X = LngToPt(MakePort(Mid(data, 1, 4)))
    Pt.y = LngToPt(MakePort(Mid(data, 5, 4)))
    convert_coord = Pt
End Function

Public Function LngToPt(i As Long) As Long
On Error GoTo z
Dim num As Long, c As Long, p As Long, d As Long
d = 1
num = i - &H3F800000
p = 1
Do While (num > 0 And p < 4000)
    If p < d Then
        num = num - ((&H800000 \ d) * 2)
        p = p + 1
    Else
        d = 2 * d
    End If
    'DoEvents
Loop
LngToPt = p
Exit Function
z:
LngToPt = 0
End Function

Public Function MakeHex(rawLong As String) As String
On Error Resume Next
Dim str1 As String
Dim X As Integer
For X = 1 To Len(rawLong)
    If Asc(Mid(rawLong, X, 1)) < 16 Then str1 = str1 + "0"
    str1 = str1 + Hex(Asc(Mid(rawLong, X, 1)))
Next
MakeHex = str1
End Function

Public Sub Plot_Dot(Pt As Coord)
Dim z1 As Long, z2 As Long
cur = Pt
z1 = (frmMap.PicMain.ScaleWidth / 2) + (((frmMap.fBGMap.ScaleWidth - frmMap.PicMain.ScaleWidth) / 2) - (Pt.X * MapScale) / 3)
z2 = (frmMap.PicMain.ScaleHeight / 2) + (((frmMap.fBGMap.ScaleHeight - frmMap.PicMain.ScaleHeight) / 2) - (Pt.y * MapScale) / 3)
frmMap.PicMain.Move z1, z2
frmMap.block.Move ((Pt.X) * MapScale / 3) - 2, ((Pt.y) * MapScale / 3) - 2
frmMap.PicMain.Refresh
frmMain.txtPos.Text = "X:" & Pt.X & "  Y:" & Pt.y
End Sub

Public Function Convert_Point(ByVal X As Long, ByVal y As Long) As String
    Convert_Point = ReverseByte(LngToChr(PtToLng(X))) & ReverseByte(LngToChr(PtToLng(y))) & tmpMap
End Function

Public Function PtToLng(i As Long) As Long
On Error GoTo z
Dim num As Long, c As Long, p As Long, d As Long
d = 1
num = &H3F800000
p = 1
Do While (i > p)
    If p < d Then
        num = num + ((&H800000 \ d) * 2)
        p = p + 1
    Else
        d = 2 * d
    End If
    'DoEvents
Loop
PtToLng = num
Exit Function
z:
PtToLng = 0
End Function

Public Function Distant(coord1 As Coord, coord2 As Coord)
On Error GoTo out
    Dim a As Long
    Dim X As Long
    Dim y As Long
    X = Abs(coord1.X - coord2.X) * Abs(coord1.X - coord2.X)
    y = Abs(coord1.y - coord2.y) * Abs(coord1.y - coord2.y)
    Distant = CInt(Sqr(X + y))
    Exit Function
out:
    Distant = 0
End Function

Public Sub Start_Walk(ByVal X As Integer, ByVal y As Integer)
Dim dist As Integer
If Constate <> 4 Then frmBot.tmrWalk.Enabled = False: Exit Sub
Walking_stop.X = X
Walking_stop.y = y
Walking_start = cur
Walking_dist = Distant(Walking_start, Walking_stop)
'frmBot.tmrWalk.Enabled = False
If (Walking_dist <= 2) Then Stop_Walk: Exit Sub
Walking = 0
stepwarp = 0
frmBot.tmrWalk.Interval = 100
frmBot.tmrWalk.Enabled = True
End Sub

Public Sub Stop_Walk()
Dim tx As String, ty As String
    LoginChar(Char).Pos = cur
    tx = ReverseByte(LngToChr(PtToLng(cur.X)))
    ty = ReverseByte(LngToChr(PtToLng(cur.y)))
    Send_Stop tx & ty & tmpMap
    frmBot.tmrWalk.Enabled = False
    Walking = 0
    Walking_dist = 0
End Sub

Public Sub Send_Walk(data As String)
    SendPacket Chr(&H8C) & Chr(&H0) & Chr(&H1) & LoginChar(Char).id & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H6) & data & Chr(&H0), 1
    'Stat "Send_Walk"
End Sub

Public Sub Send_Stop(data As String)
    SendPacket Chr(&H8C) & Chr(&H0) & Chr(&H3) & LoginChar(Char).id & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & data & Chr(&H0), 1
    'Stat "Send_Stop"
End Sub

Public Sub Pet_Walk(data As String)
'8C 02 01 00 00 45 78 00 00 00 08 71 91 8E 44 21 E5 45 44 7F 35 1F 43 75 CD 19 41 00
If pet.id = "" Or Not frmMain.tbPet.Visible Then Exit Sub
SendPacket Chr(&H8C) & Chr(&H2) & Chr(&H1) & pet.id & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H8) & data & Chr(&H0), 1
End Sub

Public Sub Check_Pos(bot As Coord, etc As Coord)
If Not opt.path.lock.auto Or frmBot.tmrWalk.Enabled Then Exit Sub
If Distant(bot, etc) > 200 And Map = opt.path.lock.m Then cur = etc: Start_Walk bot.X, bot.y
End Sub

Public Sub InviteParty(name As String, id As String, State As Integer)
    SendPacket Chr(&H98) & Chr(&H0) & Chr(State) & id, 1
' "Party : เชิญปาร์ตี้กับเจ้าของ!")
    Stat "Party : เชิญปาร์ตี้! โดย " & name & " แบบ: " & State
End Sub


Public Sub SendAttack()
    If C_Atk.id <> "" And (opt.atk.auto Or opt.party.atk.auto) Then
        With LoginChar(Char)
            If opt.atk.skill.auto And delay_use_skill = 0 And opt.atk.skill.hp >= (.hp / .MaxHP) * 100 And opt.atk.skill.mp <= (.mp / .MaxMp) * 100 Then
                useSkill C_Atk.id, LngToChr(CLng(opt.atk.skill.skill)), 1
                'Stat "send attack by skill to " & C_Atk.name & "  dist=" & Distant(cur, C_Atk.Pos)
            Else
                SendPacket Chr(&H8D) & Chr(&H0) & .id & Chr(&H1) & C_Atk.id & Chr(&H0) & Chr(&H0), 1
                'Stat "send attack to " & C_Atk.name & "  dist=" & Distant(cur, C_Atk.Pos)
            End If
        End With
    End If
End Sub

Public Sub SendFarm()
    If N_Atk.id <> "" And opt.farm.auto And N_Atk.hp >= 2 Then
        SendPacket Chr(&HA7) & N_Atk.id, 1
    End If
End Sub

Public Sub SitDown()
    SendPacket Chr(&HA1) & LoginChar(Char).id & Chr(&H0) & Chr(&H3) & Chr(&H1), 1
    'Index.Display ("Status : Sit Down!")
    frmMain.btnSit.Caption = "Sit"
    Sitting = True
End Sub

Public Sub StandUp()
    SendPacket Chr(&HA1) & LoginChar(Char).id & Chr(&H0) & Chr(&H3) & Chr(&H0), 1
    'Index.Display ("Status : Staund up!")
    frmMain.btnSit.Caption = "Stand"
    Sitting = False
End Sub

Public Sub GotoTown()
    SendPacket Chr(&H9E) & Chr(&H0), 1
End Sub

Public Sub SayChat(txt As String, State As Integer)
    Dim X As String
    With LoginChar(Char)
        Select Case State
            Case 0
                X = Chr(&H8F) & Chr(&H0) & .id & .name & Chr(&H0) & Chr(&H0) & txt & Chr(&H0)
            Case 1
                X = Chr(&H8F) & Chr(&H1) & .id & .name & Chr(&H0) & Chr(&H0) & txt & Chr(&H0)
            Case 2
                X = Chr(&H8F) & Chr(&H2) & .id & .name & Chr(&H0) & Chr(&H0) & txt & Chr(&H0)
            Case 3
                X = Chr(&H8F) & Chr(&H3) & .id & .name & Chr(&H0) & Chr(&H0) & txt & Chr(&H0)
            Case 4
                If Trim(frmChat.cbWhisper.Text) = "" Then Exit Sub
                X = Chr(&H8F) & Chr(&H4) & .id & .name & Chr(&H0) & Trim(frmChat.cbWhisper.Text) & Chr(&H0) & txt & Chr(&H0)
            Case 5
                X = Chr(&H8F) & Chr(&H5) & .id & .name & Chr(&H0) & Chr(&H0) & txt & Chr(&H0)
            End Select
    End With
    SendPacket X, 1
End Sub

Public Sub useSkill(t_id As String, s_id As String, State As Integer)
    Dim X As String
        X = Chr(&H9B) & Chr(&H2) & Chr(&H0) & LoginChar(Char).id & s_id & Chr(State) & t_id & Chr(&H0)
        SendPacket X, 1
        X = Chr(&H9B) & Chr(&H3) & Chr(&H0) & LoginChar(Char).id & s_id & Chr(State) & t_id & Chr(&H0)
        SendPacket X, 1
        delay_use_skill = 10
End Sub

Public Sub Gomap(id As Integer)
    SendPacket Chr(&H92) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(id) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0), 1
End Sub

Public Sub ChangPet(id As Integer)
Dim i As Integer
    Select Case id
        Case 1 'เอาสัตว์เลี้ยงออกมา  90 05 0A 00 0E 00 00 07 85 EC
            For i = 0 To UBound(Inv) - 1
                If Inv(i).Type = &H367 Or Inv(i).Type = &H368 Then
                        SendPacket Chr(&H90) & Chr(&H5) & Chr(&HA) & Chr(&H0) & Chr(Inv(i).row) & Chr(Inv(i).col) & Inv(i).id, 1
                        '90 05 0A 00 0E 01 00 04 7D B3
                        Chat "[Pet] นำสัตว์เลี้ยงออกมา " & Inv(i).name
                        delay_feed_pet = 5
                    Exit Sub
                End If
            Next
            opt.pet = False
            '90 05 0A 00 0E 02 00 38 F5 1A
        Case 2 'เก็บสัตว์
            SendPacket Chr(&H90) & Chr(&H5) & Chr(&HA) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF), 1
            Chat "[Pet] เก็บสัตว์เลี้ยง "
    End Select
End Sub

Public Sub SendFeedPet_EN() 'ให้หิน เพิ่มความหิว
''90 00 00 07 02 00 08 AC 02 00 00 00 00
Dim i As Integer
For i = 0 To UBound(Inv) - 1
    If Inv(i).Type = &H9A Or Inv(i).Type = &H9B Or Inv(i).Type = &H9C Then
        SendPacket Chr(&H90) & Chr(&H0) & Chr(&H0) & Chr(Inv(i).row) & Chr(Inv(i).col) & Inv(i).id & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0), 1
        '90 00 00 07 02 00 08 AC 02 00 00 00 00,1
        Chat "[Pet] ให้อาหารสัตว์ " & Inv(i).name
        Exit Sub
    End If
Next
opt.pet = False
End Sub

Public Sub SendFeedPet_HP() 'ให้กิ่งไม้ เพิ่มเลือด
'90 00 00 0A 04 00 12 BB 6B 00 00 00 00
Dim i As Integer
For i = 0 To UBound(Inv) - 1
    If Inv(i).Type = &HC5 Or Inv(i).Type = &HC6 Or Inv(i).Type = &HC7 Then
        SendPacket Chr(&H90) & Chr(&H0) & Chr(&H0) & Chr(Inv(i).row) & Chr(Inv(i).col) & Inv(i).id & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0), 1
        '90 00 00 07 02 00 08 AC 02 00 00 00 00,1
        Chat "[Pet] ให้อาหารสัตว์ " & Inv(i).name
        Exit Sub
    End If
Next
opt.pet = False
End Sub

Public Sub SendUseHP()
Dim i As Integer
If delay_use_hp <> 0 Then Exit Sub
For i = 0 To UBound(Inv) - 1
    If Inv(i).Type = &H2B Or Inv(i).Type = &H2C Or Inv(i).Type = &H2D Or Inv(i).Type = &H1C9 Or Inv(i).Type = &H1EA Or Inv(i).Type = &H355 Or Inv(i).Type = &H357 Then
        SendPacket Chr(&H90) & Chr(&H0) & Chr(&H0) & Chr(Inv(i).row) & Chr(Inv(i).col) & Inv(i).id & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0), 1
        '90 00 00 07 02 00 08 AC 02 00 00 00 00,1
        delay_use_hp = 15
        Exit Sub
    End If
Next
delay_use_hp = 60
End Sub

Public Sub SendUseMP()
Dim i As Integer
If delay_use_mp <> 0 Then Exit Sub
For i = 0 To UBound(Inv) - 1
    If Inv(i).Type = &H1E4 Or Inv(i).Type = &H1E5 Or Inv(i).Type = &H1ED Or Inv(i).Type = &H22C Or Inv(i).Type = &H2D4 Or Inv(i).Type = &H356 Or Inv(i).Type = &H358 Then
        SendPacket Chr(&H90) & Chr(&H0) & Chr(&H0) & Chr(Inv(i).row) & Chr(Inv(i).col) & Inv(i).id & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0), 1
        '90 00 00 07 02 00 08 AC 02 00 00 00 00,1
        delay_use_mp = 15
        Exit Sub
    End If
Next
delay_use_mp = 60
End Sub

Public Function Eq_Type(id As Integer) As String
    Select Case id
        Case 0
            Eq_Type = "Head"
        Case 1
            Eq_Type = "Body"
        Case 2
            Eq_Type = "Weapon"
        Case 3
            Eq_Type = "Pants"
        Case 5
            Eq_Type = "Hand"
        Case 6
            Eq_Type = "Boot"
        Case 7
            Eq_Type = "Acc[1]"
        Case 8
            Eq_Type = "Acc[2]"
        Case 9
            Eq_Type = "Acc[3]"
        Case 255
            Eq_Type = ""
        Case Else
            Eq_Type = "unno:" & id
    End Select
End Function

Public Function Pet_Type(id As Integer) As String
    Select Case id
        Case &H11
            Pet_Type = "อาชาน้อย"
        Case &H12
            Pet_Type = "อาชา"
        Case &H13
            Pet_Type = "ไนท์แมร์"
        Case &H14
            Pet_Type = "ไนท์แมร์[พาหนะ]"
        Case &H21
            Pet_Type = "มังกรน้อย"
        Case &H22
            Pet_Type = "เดรก"
        Case &H23
            Pet_Type = "มังกร"
        Case &H24
            Pet_Type = "มังกร[พาหนะ]"
        Case Else
            Pet_Type = "unno:" & id
    End Select
End Function

Public Sub ConnectToServ(id As Integer)
Dim strIP As String, strPort As Integer
tmpSub = 6
tmpServ = id
frmBot.Winsock.CloseSck
strPort = 4020
If id = 1 Then
    strIP = "61.90.198.108"
ElseIf id = 2 Then
    strIP = "61.90.198.113"
ElseIf id = 3 Then
    strIP = "61.90.198.119"
End If
Stat "Change Server to " & strIP & ":" & strPort
frmBot.Winsock.Connect strIP, strPort
End Sub

Public Sub Send_Refresh_INV()
    SendPacket Chr(&H90) & Chr(&H3) & Chr(&H0), 1
End Sub

Public Function RC4(ByVal str As String, ByVal Pwd As String) As String
On Error Resume Next
Dim Sbox(0 To 255) As Integer
Dim a
Dim b
Dim c
Dim Key() As Byte
Dim ByteArray() As Byte
Dim tmp As Byte
If Len(Pwd) = 0 Or Len(str) = 0 Then Exit Function

If Len(Pwd) > 256 Then
    Key() = StrConv(Left$(Pwd, 256), vbFromUnicode)
Else
    Key() = StrConv(Pwd, vbFromUnicode)
End If

For a = 0 To 255
    Sbox(a) = a
Next a
a = 0
b = 0
c = 0
For a = 0 To 255
    b = (b + Sbox(a) + Key(a Mod Len(Pwd))) Mod 256
    tmp = Sbox(a)
    Sbox(a) = Sbox(b)
    Sbox(b) = tmp
Next a
a = 0
b = 0
c = 0
ByteArray() = StrConv(str, vbFromUnicode)
For a = 0 To Len(str)
    b = (b + 1) Mod 256
    c = (c + Sbox(b)) Mod 256
    tmp = Sbox(b)
    Sbox(b) = Sbox(c)
    Sbox(c) = tmp
    ByteArray(a) = ByteArray(a) Xor (Sbox((Sbox(b) + Sbox(c)) Mod 256))
Next a
RC4 = StrConv(ByteArray, vbUnicode)
End Function

