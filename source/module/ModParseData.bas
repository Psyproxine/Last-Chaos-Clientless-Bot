Attribute VB_Name = "ModParseData"
Option Explicit
Public Chan_List() As Channel_Info, send80 As Boolean, sendConfirmLoading As Boolean
Public tmppacket() As String, tmpCharSelect As String, serv_index As Integer, tmrRep As Long
Public tmpPack As Integer, tmpEquip As Integer
Public Sub print_packet(packet As String, Text As String)
    If Not opt.debug.show Then Exit Sub
    Dim tstr As String
    Dim X As Integer
    tmppacket(UBound(tmppacket)) = packet
    ReDim Preserve tmppacket(UBound(tmppacket) + 1)
    tstr = ""
    For X = 1 To IIf(Len(packet) >= 8, 8, Len(packet))
       If Asc(Mid(packet, X, 1)) < 16 Then tstr = tstr + "0"
       tstr = tstr + Hex(Asc(Mid(packet, X, 1))) + " "
       'If x Mod 19 = 0 Then tstr = tstr & vbCrLf
    Next
        tstr = Left(tstr, Len(tstr) - 1)
        frmDebug.lstpacket.AddItem "[" & time & "]" & Text & " " & tstr
        'DoColor frmDebug.rtbDebug, Chr(3) & "14[" & Chr(3) & "15" & Time & Chr(3) & "4 " & Text & Chr(3) & "14] " & tstr
End Sub

Public Sub PrintP(head As String, packet As String)
    Dim tstr As String, sstr As String
    Dim X As Integer
    tstr = ""
    For X = 1 To Len(packet)
       If Asc(Mid(packet, X, 1)) < 16 Then tstr = tstr + "0"
       tstr = tstr + Hex(Asc(Mid(packet, X, 1))) + " "
       sstr = sstr + hex2string(Trim(Mid$(packet, X, 1))) + " "
       If X Mod 16 = 0 Then
            tstr = tstr & vbCrLf
            sstr = sstr & vbCrLf
        End If
    Next
       tstr = Left(tstr, Len(tstr) - 1)
       Open App.path & "\log\" & ChrtoHex(head) & ".txt" For Append As #1
       Print #1, " == " & ChrtoHex(head) & "                  " & Len(packet) & " Bytes == "
       Print #1, tstr
       Print #1, sstr
       Close 1
End Sub


Public Sub ParseData()
Dim ChopNumber As Long
Dim tmpData As String
Dim tData As String
Dim NoJump As Boolean, i As Integer, j As Integer, z As Integer
  'On Error GoTo runtime5

restart:


If Len(RecvData) < 4 Then Exit Sub
ChopNumber = MakeLong(Mid(RecvData, 11, 2))
If Len(RecvData) < ChopNumber + 12 Then Exit Sub
tmpData = Mid(RecvData, 13, ChopNumber)

'61.90.198.104
Chat ChrtoHex(RC4(Mid(tmpData, 3), "LASTCHAOS"))

Connecting False
RecvData = ""
Select Case Mid(tmpData, 1, 2)
Case Chr(&H22) & Chr(&H0)  'Server List
      Pack_22 tmpData
Case Chr(&H1) & Chr(&H1F)   'Login Error
        Stat Chr(&H3) & "0,4 Login Error"
        Connecting False
       'reconnect
Case Chr(&H1) & Chr(&H20)
        Stat Chr(&H3) & "0,4 ไอดีนี้อยู่ในระหว่างการใช้งาน"
        Connecting False
        reconnect
Case Chr(&H1) & Chr(&H2B)
        Stat Chr(&H3) & "0,4 ไอดีนี้ไม่สามารถเล่น24ชม"
        Connecting False
Case Chr(&H1) & Chr(&HD5)
        
Case Chr(&H2) & Chr(&H1) 'ยืนยันหลังเลือกตัว
        SendPacket Chr(&HA6) & LngToChr(tmrRep), 1
        frmBot.tmrLoad.Enabled = True
        Stat "Loading Data.."
Case Chr(&H2) & Chr(&H2) 'ลิสตัวละคร
        Unload frmServ
      Pack_202 tmpData
Case Chr(&H2) & Chr(&H3) 'ลิสตัวละคร
If Char > -1 And Char < frmChar.lstChar.ListCount Then frmChar.lstChar.ListIndex = Char: frmChar.lstChar_DblClick
    If Constate < 3 Then frmChar.show
Case Chr(&H6) & Chr(&H0) 'รายละเอียดตัวละคร
    Pack_600 tmpData
Case Chr(&H7) & Chr(&H0)          'Found Monster(42 byte) & People(>42 byte)
    Pack_700 tmpData
Case Chr(&H7) & Chr(&H1)          'Monster Born
    Pack_700 tmpData
Case Chr(&H8) & Chr(&H0)          'People  หายไป
    Pack_800 tmpData
Case Chr(&H8) & Chr(&H1)          'Monster  หายไป
    Pack_801 tmpData
Case Chr(&H8) & Chr(&H2)          'Pet หายไป
    Pack_802 tmpData
    'Chat ChrtoHex(tmpData)
Case Chr(&H9) & Chr(&H0)  'บอกตำแหน่ง ตอนเข้าเกมส์
    Pack_900 tmpData
    frmMain.Caption = LoginChar(Char).name & " - " & Version
Case Chr(&HA) & Chr(&H0) 'item
    Pack_A00 tmpData
Case Chr(&HA) & Chr(&H1)          'Item Ware House First
    Pack_A00 tmpData
Case Chr(&HA) & Chr(&H2)          'End item Ware house
    Pack_A00 tmpData
Case Chr(&HB) & Chr(&H0)
    'Chat ChrtoHex(tmpData)
Case Chr(&HC) & Chr(&H0)                        'People move
    Pack_C00 tmpData
Case Chr(&HC) & Chr(&H1)                        'Monster move
    Pack_C01 tmpData
Case Chr(&HC) & Chr(&H2)                        'Pet move
    Pack_C02 tmpData
Case Chr(&HE) & Chr(&H0)                   'Atk Monster Recv
    Pack_E00 tmpData
Case Chr(&HE) & Chr(&H1)                   'Monster Atk
    Pack_E01 tmpData
Case Chr(&HF) & Chr(&H0)          'Chat ธรรมดา
    Pack_F00 tmpData, 0
Case Chr(&HF) & Chr(&H1)          'Chat ปาร์ตี้
    Pack_F00 tmpData, 1
Case Chr(&HF) & Chr(&H2)          'Chat กิลด์
    Pack_F00 tmpData, 2
Case Chr(&HF) & Chr(&H3)          'Chat สังคม
    Pack_F00 tmpData, 3
Case Chr(&HF) & Chr(&H4)          'Chat กระซิบ
    Pack_F00 tmpData, 4
Case Chr(&HF) & Chr(&H5)          'Chat ตะโกน
    Pack_F00 tmpData, 5
Case Chr(&HF) & Chr(&H6)          'Chat GM แหกปาก
    Pack_F00 tmpData, 6
Case Chr(&H10) & Chr(&H0)                'ใช้ item
    Pack_1000 tmpData
Case Chr(&H10) & Chr(&H1)                'พยายามเก้บของ
Case Chr(&H10) & Chr(&H4)               'ของหมด
    Pack_1004 tmpData
Case Chr(&H10) & Chr(&H5)               'ถอด ใส่ อุปกรณ์
    Pack_1005 tmpData
Case Chr(&H10) & Chr(&H7)                  'ของเข้า ชนิดใหม่
    Pack_1007 tmpData
Case Chr(&H10) & Chr(&H8)                  'ของเข้า มีชนิดนั้น อยู่แล้ว
    Pack_1008 tmpData
Case Chr(&H10) & Chr(&H9)                   'Item Drop
    Pack_1009 tmpData
Case Chr(&H10) & Chr(&HA)                  'เจอ item หล่นอยู่
    Pack_100A tmpData
Case Chr(&H10) & Chr(&HB)                   'item หายปาย
    Pack_100B tmpData
Case Chr(&H10) & Chr(&H1C)                   '10 1C 4A 61 6E 65 6B 75 00 00 00 00 8D 00 00 00 00 00 00 00 01

Case Chr(&H11) & Chr(&H0)              'ไม่มีชื่อคนที่ซิบ
    Chat Chr(&H3) & "5[Whisper] ไม่มีบุคคลนี้อยู่"
Case Chr(&H11) & Chr(&H16)              'อุปกรณ์หมดความทนทาน
    Pack_1116 tmpData
Case Chr(&H11) & Chr(&H17)              'ใส่อุปกรณ์ไม่ตรง  ใส่แร่อยู่
    Pack_1116 tmpData
Case Chr(&H11) & Chr(&H18)              'ใส่อุปกรณ์ไม่ตรง  สมุนไพรอยู่
    Pack_1116 tmpData
Case Chr(&H11) & Chr(&H19)              'ใส่อุปกรณ์ไม่ตรง  ประจุอยู่
    Pack_1116 tmpData
Case Chr(&H11) & Chr(&H38)              'ใส่อุปกรณ์ไม่ตรง  ประจุอยู่
    Pack_1138 tmpData
Case Chr(&H13) & Chr(&H0)              'วาบไปเมืองอื่น
    Pack_1300 tmpData
Case Chr(&H16) & Chr(&H2)          'Server ทัก ตอบ ใน 35 วินาที หลังจากส่ง ไป 25 วิ server ทักใหม่ + timerz_reply ครั้งละ 600
   frmBot.tmrReply.Enabled = True
Case Chr(&H15) & Chr(&H0)                'บันทึกตำแหน่ง
     Pack_1500 tmpData
Case Chr(&H18) & Chr(&H0)                   'โดน ขอปาร์ตี้
    Pack_1800 tmpData
Case Chr(&H18) & Chr(&H3)                       ' ปาตี้หาย
    'Pack_1808 tmpData
Case Chr(&H18) & Chr(&H4)                       '   (มีแต่หัวแพคเกจ ไม่มีข้อมูล)
Case Chr(&H18) & Chr(&H5)                       'ลิสคนในปาตี้
    Pack_1805 tmpData
Case Chr(&H18) & Chr(&H6)                       '  คนออกตี้
    Pack_1806 tmpData
Case Chr(&H18) & Chr(&H8)                     ' ยุบปาตี้
    Pack_1808 tmpData
Case Chr(&H18) & Chr(&H9)                     ' สถานะคนในปาร์ตี้ เช่นเลือด
    Pack_1809 tmpData
Case Chr(&H18) & Chr(&HA)                     ' เชิญปาตี้ไม่ได้
    Stat "Party: เชิญเข้าปาร์ตี้ไม่ได้"
Case Chr(&H1B) & Chr(&H0)                'Skill
    Pack_1B00 tmpData
Case Chr(&H1B) & Chr(&H2)               'เริ่มใช้สกิล
Case Chr(&H1B) & Chr(&H3)               ' ใช้สกิลเสร็จแล้ว
    Pack_1B03 tmpData
Case Chr(&H1B) & Chr(&H4)               '
Case Chr(&H1B) & Chr(&H6)               '
Case Chr(&H1C) & Chr(&H0)                'สถานะคนรอบข้าง
    'Chat ChrtoHex(tmpData)
Case Chr(&H1C) & Chr(&H1)                'สถานะคนรอบข้าง
    'Chat ChrtoHex(tmpData)
Case Chr(&H1C) & Chr(&H2)                'สถานะคนรอบข้าง
    'Chat ChrtoHex(tmpData)
Case Chr(&H1D) & Chr(&H0)                 'Hp mp คนรอบข้าง
    'Chat ChrtoHex(tmpData)
Case Chr(&H1D) & Chr(&H1)                    ' HP ของมอน
    Pack_1D01 tmpData
Case Chr(&H1F) & Chr(&H0)               'ท่าทางคนรอบข้าง
    'Chat ChrtoHex(tmpData)
Case Chr(&H1F) & Chr(&H2)              ' เห็น

Case Chr(&H1F) & Chr(&H1)              ' เก็บเกี่ยว
    Pack_1F02 tmpData
Case Chr(&H1F) & Chr(&H3)              ' ใช้ item
Case Chr(&H20) & Chr(&H0)          'Exp Sp
    Pack_2000 tmpData
Case Chr(&H21) & Chr(&H0)                   'นั่ง ยืน
    Pack_2100 tmpData
Case Chr(&H24) & Chr(&H0)                   'ไม่รู้
Case Chr(&H2A) & Chr(&H2)          'ชื่อกิล
    Pack_2A02 tmpData
Case Chr(&H2A) & Chr(&H3)          'ลิสคนในกิล
    Pack_2A03 tmpData
Case Chr(&H2A) & Chr(&H4)          'คนในกิลเปลี่ยนสถานะ
    Pack_2A04 tmpData
Case Chr(&H2C) & Chr(&H2)          'ตั้งร้าน
    Pack_2C02 tmpData
Case Chr(&H37) & Chr(&H0)          'รายละเอียดสัตว์เลี้ยง
    Pack_3700 tmpData
Case Else
    Stat "UnKnown Packet: " & MakePort(Mid(RecvData, 13, 2)) & " (Chr(&H" & UCase(Hex(MakePort(Mid(RecvData, 13, 1)))) & ") " & "Chr(&H" & UCase(Hex(MakePort(Mid(RecvData, 14, 1)))) & "))   ความยาว>" & ChopNumber
    PrintP Mid(RecvData, 13, 2), tmpData
End Select
RecvData = Mid$(RecvData, ChopNumber + 13)
If Len(RecvData) > 12 Then GoTo restart
Exit Sub
runtime5:
RecvData = ""
End Sub

Private Sub Pack_22(data As String)
Dim i As Integer
Dim Ch_Head As String
Dim Ch_Heads() As String
Dim tmpData As String
Dim tmp As String
Dim Serv As String
Dim leng As Integer
Dim lserv As Integer
Dim lindex As Integer
Dim l As Integer
Dim sel1 As Integer
Dim sel2 As Integer
   ' frmServ.lstServ.Clear
    If Constate > 1 Or Logining Then Exit Sub
    Constate = 1
    serv_index = serv_index + 1
    data = Mid(data, 13, Len(data) - 12)
    lserv = MakePort(Left$(data, 2))
    If serv_index = lserv Then frmServ.show
    tmpData = Right$(data, Len(data) - 4) ' เอาตัวบอกจำนวนเซิฟออก
    lindex = 0
start:
    Serv = MakePort(Left$(tmpData, 2)) ' ลำดับของเซิฟ
    If (Serv > UBound(Chan_List)) Then ReDim Preserve Chan_List(Serv)
    Chan_List(Serv).id = Serv
    frmServ.lstServ.AddItem Serv & ". Serv. " & Serv
    tmpData = Right(tmpData, Len(tmpData) - 8) ' เอาตัวบอกลำดับเซิฟออก
    leng = MakePort(Left$(tmpData, 1)) 'ตัวบอกจำนวนเซิฟย่อย
    tmpData = Right$(tmpData, Len(tmpData) - 1) ' เอาตัวบอกจำนวนเซิฟย่อยออก
    ReDim Chan_List(Serv).ChannelX(0)
    For i = 1 To leng
        ReDim Preserve Chan_List(Serv).ChannelX(UBound(Chan_List(Serv).ChannelX) + 1)
        With Chan_List(Serv).ChannelX(UBound(Chan_List(Serv).ChannelX))
            .id = MakePort(Mid(tmpData, 4, 1)) ' ลำดับของเซิฟย่อย
            tmpData = Right$(tmpData, Len(tmpData) - 8) ' เอาลำดับเซิฟย่อยออก
            tmp = MakeString(Left$(tmpData, 16)) ' ชื่อของเซิฟย่อย
            .name = tmp
            tmpData = Right$(tmpData, Len(tmpData) - 16) ' เอาชื่อเซิฟย่อยออก
            tmp = MakePort(Mid$(tmpData, 2, 1) & Left$(tmpData, 1)) ' จำนวนคนของเซิฟย่อย
            .Connection = tmp
            tmpData = Right$(tmpData, Len(tmpData) - 2) ' เอาจำนวนคนเซิฟย่อยออก
            
            If LimitSERV Then
                'จำกัดเซิฟย่อยที่เล่น
                If Server > 0 And Chan_List(Serv).id = Server Then
                    Logining = True
                    ConnectToServ Server
                    Unload frmServ
                    Exit Sub
                End If
            Else
                'ไม่จำกัดเซิฟย่อยที่เล่น
                If (Server > 0 And World > 0) Then
                    If (Chan_List(Serv).id = Server And .id = World) Then
                        frmBot.Winsock.CloseSck
                        Stat "Change Server to " & .name & ":" & .Connection & " [" & Server & "/" & World & "]"
                        frmBot.Winsock.Connect .name, .Connection
                        Unload frmServ
                        Exit Sub
                    End If
                End If
            End If
        End With
    Next
    If Len(tmpData) > 0 Then GoTo start
End Sub

Public Sub Pack_202(data As String)
On Error Resume Next
Dim str() As String, i As Integer, c As Integer, a As String
    If Constate > 2 Then Exit Sub
    Constate = 2
        c = UBound(LoginChar)
        LoginChar(c).id = Mid(data, 3, 4)
        i = InStr(8, data, Chr(&H0), vbBinaryCompare)
        LoginChar(c).name = Mid$(data, 7, i - 7)
        data = Mid(data, i + 1)
        LoginChar(c).Class = MakePort(Left(data, 1))
        data = Mid(data, 8)
        LoginChar(c).level = MakePort(Left(data, 1))
        data = Right(data, Len(data) - 1)

            'Get Exp
            If MakePort(ReverseHex(Mid(data, 1, 8))) > MakeLong(Mid(data, 9, 8)) Then
                LoginChar(c).EXP = 0
            Else
                LoginChar(c).EXP = MakeLong(Mid(data, 1, 8))
            End If
            LoginChar(c).MaxEXP = MakeLong(Mid(data, 9, 8))
            data = Right(data, Len(data) - 16)
            'Get Unknow
          '  Characterz(p).Stat.Unknow = makelong(Mid(Data, i + 1, 4))
            data = Right(data, Len(data) - 4)
            'Get Hp
            LoginChar(c).hp = MakeLong(Left(data, 4))
            data = Right(data, Len(data) - 4)
            LoginChar(c).MaxHP = MakeLong(Left(data, 4))
            'Get Mp
            data = Right(data, Len(data) - 4)
            LoginChar(c).mp = MakeLong(Left(data, 4))
            data = Right(data, Len(data) - 4)
            LoginChar(c).MaxMp = MakeLong(Left(data, 4))
            'i_name = i_name + 8

            
            
        'LoginChar(c).Pos.x = MakePort(Mid(str(i), 5, 4))
        'LoginChar(c).Pos.y = MakePort(Mid(str(i), 9, 4))
        'Stat "Move to " & convert_coord(LoginChar(c).Pos.x, True) & " : " & convert_coord(LoginChar(c).Pos.y, False)
        frmChar.lstChar.AddItem c + 1 & ". " & LoginChar(c).name
        ReDim Preserve LoginChar(UBound(LoginChar) + 1)
     '  For c = 0 To frmChar.lstChar.ListCount - 1
     '     If (Left(frmChar.lstChar.List(c), InStr(frmChar.lstChar.List(c), ".") - 1) = CStr(Char)) Then
     '            frmChar.lstChar.ListIndex = c
     '            frmChar.lstChar_DblClick
     '            Exit For
     '      End If
     '      Next
    'End If
End Sub

 Private Sub Pack_600(data As String)    'Character Status
        'On Error Resume Next
        Dim Buff1, Buff2 As String
        'Dim BInt As Integer
        '01 81 00 00 01 D8 00 00 00 00 00 82 06 00
        '0  free
        '00 1D   Lv
        '00 00 00 00 00 4F EE B7      00 00 00 00 00 74 04 C4    Exp
        '00 00 00 CA                  00 00 02 A0                Hp
        '00 00 06 1B                  00 00 06 1B                Mp
        '00 00 00 00    'พละกำลัง เต็ม
        '00 00 00 01     'ว่องไว
        '00 00 00 35    'สติปัญญา
        '00 00 00 0E     'ร่างกาย
        '00 00 00 00    'ในวงเล็บ พละกำลัง
        '00 00 00 00    'ในวงเล็บ ว่องไว
        '00 00 00 24     'ในวงเล็บ สติปัญญา
        '00 00 00 04     'ในวงเล็บ ร่างกาย
        '00 00 00 23     'พลังโจมตี
        '00 00 01 48     'พลังโจมตีด้วยเวทย์มนต์
        '00 00 01 F4     'พลัง ป้องกัน
        '00 00 00 2E     'พลังป้องกันเวทย์มนต์
        '00 2E 92 39
        '00 00 0D A5
        '00 00 1B 30
        '00 00 C0 3F
        '00 00 C0 40
        '10 F8
        '00 00 00 00 00
        'FF FF FF FE      'pk
        '00 00 00 00
        '00 00 70 41
        '00 00 00 00 00 00
        
        '06
        '00 00 00 26
        '00 00 00 00 00 20 1A AF     00 00 00 00 02 42 B1 F0          EXP
        '00 00 05 66                             00 00 07 06                                  HP
        '00 00 00 3F                             00 00 05 B9                                 MP
        '00 00 00 24                                                                                    Str
        '00 00 00 13                                                                                    AGI
        '00 00 00 08                                                                                    INT
        '00 00 00 0C                                                                                    VIT
        '00 00 00 08
        '00 00 00 04
        '00 00 00 08
        '00 00 00 04
        '00 00 01 B0                                                                                    ATK
        '00 00 00 68                                                                                    MATK
        '00 00 02 A4                                                                                    DEF
        '00 00 00 49                                                                                    MDEF
        '00 87 64 A8                                                                                    SP
        '00 00 10 62
        '00 00 24 90
        '00 00 C0 3F
        '00 00 F0 40
        '0A 00 00 00
        '00 00 00 00
        '00 00 00 00
        '00 00 00 33
        '33 13 40 00
        '00 00 00 00
        '00 00 00 00 00 00 00
        With LoginChar(Char)
            Buff1 = Right(data, Len(data) - 3) 'Cut 06 00 00
            Buff2 = Left(Buff1, 2)
            .level = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 2)  'Cut Lv
            Buff2 = Left(Buff1, 8)
            .EXP = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 8)  'Cut exp
            Buff2 = Left(Buff1, 8)
            .MaxEXP = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 8)  'Cut maxexp
            Buff2 = Left(Buff1, 4)
            .hp = MakeLong(Buff2)
    
            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut hp
            Buff2 = Left(Buff1, 4)
            .MaxHP = MakeLong(Buff2)
            
            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut maxhp
            Buff2 = Left(Buff1, 4)
            .mp = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut mp
            Buff2 = Left(Buff1, 4)
            .MaxMp = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut maxmp
            Buff2 = Left(Buff1, 4)
            .str = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut str
            Buff2 = Left(Buff1, 4)
            .Agi = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut agi
            Buff2 = Left(Buff1, 4)
            .Int = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut inte
            Buff2 = Left(Buff1, 4)
            .Vit = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut vit
            Buff2 = Left(Buff1, 4)
            .Str0 = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut str0
            Buff2 = Left(Buff1, 4)
            .Agi0 = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut agi0
            Buff2 = Left(Buff1, 4)
            .Int0 = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut inte0
            Buff2 = Left(Buff1, 4)
            .Vit0 = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut vit0
            Buff2 = Left(Buff1, 4)
            .atk = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut attk
            Buff2 = Left(Buff1, 4)
            .Mtk = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut mgattk
            Buff2 = Left(Buff1, 4)
            .Def = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut def
            Buff2 = Left(Buff1, 4)
            .Mef = MakeLong(Buff2)

            Buff1 = Right(Buff1, Len(Buff1) - 4)  'Cut mgdef
            Buff2 = Left(Buff1, 4)
            .Spp = MakeLong(Buff2) / 10000


        End With
        Update_Stats
    End Sub 'Character Status

   Private Sub Pack_700(data As String)  'Found monster or people
        'On Error Resume Next
            Dim i As Integer, t As Integer
        If MakeLong(Mid(data, 3, 1)) = 2 Then
        '07 01 02 00 00 3E 59 00 02 A3 75 4A 61 6E 65 6B 75 00 21 08 F1 D5 44 A8 7D 11 44 1F 05 C8 42 6C AA B5 C2 00 00 00 00 64 00 00 00 64 0A
        '07 01
        '02
        '00 00 3E 59
        '00 02 A3 75
        '4A 61 6E 65 6B 75
        '00 21 '
        '3A E6 9C 44 59 4D 91 44 48 61 20 43 2F C0 D3 C2
        '00
        '00 00 00 64
        '00 00 00 64
        'FF
        
                With Apet(UBound(Apet))
                        .id = Mid(data, 4, 4)
                        i = InStr(12, data, Chr(&H0), vbBinaryCompare)
                        .name = Mid$(data, 12, i - 12)
                        .Type = Pet_Type(MakeLong(Mid(data, i, 2)))
                        .Coord = Mid(data, i + 2, 16)
                        .Pos = convert_coord(.Coord)
                        .hp = MakeLong(Mid(data, i + 19, 4))
                        .MaxHP = MakeLong(Mid(data, i + 23, 4))
                End With
                If Mid(data, 8, 4) = LoginChar(Char).id Then
                    pet = Apet(UBound(Apet))
                    Update_Pet
                End If
                ReDim Preserve Apet(UBound(Apet) + 1)
                Update_Apet
        ElseIf MakeLong(Mid(data, 3, 1)) = 1 Then
            If Mid(data, 42, 1) <> Chr(&H0) Then
                'This is a NPC
                With NPC(UBound(NPC))
                    .id = Mid(data, 4, 4)
                    .Type = MakeLong(Mid(data, 8, 4))
                    .name = Return_Monster_Name(.Type)
                    .Coord = Mid(data, 12, 16)
                    .hp = MakeLong(Mid(data, 29, 4))
                    .MaxHP = MakeLong(Mid(data, 33, 4))
                    .Agriculture = IIf(Mid(data, 42, 1) = Chr(&H14) Or Mid(data, 42, 1) = Chr(&H1E), True, False)
                End With
                ReDim Preserve NPC(UBound(NPC) + 1)
                Update_npc
                If opt.farm.auto And N_Atk.id <> "" And (Mid(data, 42, 1) = Chr(&H14) Or Mid(data, 42, 1) = Chr(&H1E)) And _
                N_Atk.MaxHP = 1000 And MakeLong(Mid(data, 33, 4)) = 300 Then
                    N_Atk.id = ""
                    SelectFarm
                End If
            Else
                'This is a monster
                With Monster(UBound(Monster))
                    .id = Mid(data, 4, 4)
                    .Type = MakeLong(Mid(data, 8, 4))
                    .name = Return_Monster_Name(.Type)
                    .Coord = Mid(data, 12, 16)
                    .Pos = convert_coord(.Coord)
                    .hp = MakeLong(Mid(data, 29, 4))
                    .MaxHP = MakeLong(Mid(data, 33, 4))
                End With
                ReDim Preserve Monster(UBound(Monster) + 1)
                'Update_Monster
            End If
        Else
            People(UBound(People)).id = Mid(data, 4, 4)
            i = InStr(9, data, Chr(&H0), vbBinaryCompare)
            People(UBound(People)).name = Mid$(data, 8, i - 8)
            People(UBound(People)).Class = MakeLong(Mid(data, i + 1, 1))
            People(UBound(People)).Coord = Mid(data, i + 5, 16)
            People(UBound(People)).Pos = convert_coord(People(UBound(People)).Coord)
            t = MakePort(Mid(data, i + 95, 1))
            If t > 0 Then
                People(UBound(People)).Guild = Mid(data, i + 96, t)
            End If
            ReDim Preserve People(UBound(People) + 1)
            'Update_People
            'MessageBox.Show(Peoples.Item(0).Name)
            'Any people
        End If

    End Sub 'Found monster or people
        
Private Sub Pack_800(data As String)
    'On Error Resume Next
    Dim i As Integer, found As Boolean
    For i = 0 To UBound(People) - 1
        If People(i).id = Mid(data, 3, 4) Then found = True
        If found Then
            People(i) = People(i + 1)
        End If
    Next
    If found Then ReDim Preserve People(UBound(People) - 1)
End Sub 'people  Disappear
    
Public Sub Pack_801(data As String)
    'On Error Resume Next
    Dim i As Integer, found As Boolean
    found = False
    For i = 0 To UBound(Monster) - 1
        If Monster(i).id = Mid(data, 3, 4) Or Monster(i).id = "" Then found = True
        If found Then Monster(i) = Monster(i + 1)
    Next
    If C_Atk.id = Mid(data, 3, 4) Then C_Atk.id = ""
    If found Then ReDim Preserve Monster(UBound(Monster) - 1): Exit Sub
    found = False
    For i = 0 To UBound(NPC) - 1
        If NPC(i).id = Mid(data, 3, 4) Or NPC(i).id = "" Then found = True
        If found Then NPC(i) = NPC(i + 1)
    Next
    If found Then ReDim Preserve NPC(UBound(NPC) - 1): Update_npc
End Sub 'MOnster Disappear
    
Public Sub Pack_802(data As String)
    Dim i As Integer, found As Boolean
    found = False
    For i = 0 To UBound(Apet) - 1
        If Apet(i).id = Mid(data, 3, 4) Or Apet(i).id = "" Then found = True
        If found Then Apet(i) = Apet(i + 1)
    Next
    If found Then ReDim Preserve Apet(UBound(Apet) - 1): Update_Apet
End Sub

Public Sub Pack_900(data As String)
'09
'00 00 55 DD     id
'B9 D9 EB B9 E9 CD C2 E2 B9 B5 C1           น ู ๋ น ้ อ ย โ น ต ม
'00 03 00 03           . . . . .
'02 00 00 00 00 00 00 00 00
'00 B0 94 44 00 60 6A 44 B8 DE 20 43 00 00 00 00
'00 00 04 FE 2A 00 00 00 00 00 0A        . . . .

'09
'00 02 83 45
'50 6f 73 69 74 72 6f  6e                                       Positro n
'00 01 00 02
'02 00 00 00 00 00 00 00 00
'00 60  8b 44 00 00 6e 44 48 6120 43 00 00 00 00
'00 00  04 fe 2a 00 00 00 00 01 0a
'00 05 07 40
Dim i As Integer, tmp As Integer
'Chat "PAck_900: " & ChrtoHex(data)
i = InStr(6, data, Chr(&H0), vbBinaryCompare)
Map = CInt(MakeLong(Mid(data, i + 8, 1)))
If opt.path.lock.auto Then delay_go_map = 2
data = Mid(data, i + 13)
tmpPack = (MakeLong(Right(data, 4)) - (MakeLong(LoginChar(Char).id) * 2)) + (LoginChar(Char).Class * 2)   '(MakeLong(Right(Data, 1)) + opt.debug.pack) Mod 256
LoginChar(Char).Coord = Left(data, 16)
LoginChar(Char).Pos = convert_coord(data)
tmpMap = Mid(LoginChar(Char).Coord, 9, 8)
Plot_Dot LoginChar(Char).Pos
Constate = 4
delay_feed_pet = 15
frmMap.slMap_Change
End Sub
    
Public Sub Pack_A00(data As String)
'0A000001
'002988AB0000020E060000000000000000FFFFFFFF000000000000000100
'002988AC00000013FF0000000000000000FFFFFFFF000000000000006800
'002988AD0000021C010000000000000000FFFFFFFF000000000000000100
'FFFFFFFFFFFFFFFF

'0A000200
'002988AE000002F8FF0000000000000000FFFFFFFF000000000000000100
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF

'0A000000
'0029936B00000228000000000000000000FFFFFFFF000000000000000100
'0029936C0000020CFF0000000000000000FFFFFFFF000000000000000100
'0029936D00000210020000000000000000FFFFFFFF000000000000000100
'0029936E0000020D030000000000000000FFFFFFFF000000000000000100
'0029936F0000020F050000000000000000FFFFFFFF000000000000000100
'0A000001
'002993700000020E060000000000000000FFFFFFFF000000000000000100
'0029937100000013FF0000000000000000FFFFFFFF000000000000006800
'002993720000021C010000000000000000FFFFFFFF000000000000000100
'FFFFFFFFFFFFFFFF
'0A000200
'00299373000002F8FF0000000000000000FFFFFFFF000000000000000100
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF


'0A000000

'00081712
'000000B9
'FF0000000000000080FFFFFFFF000000000000000100000817130000008DFF0000000000000000FFFFFFFF00000000000000250000081714000001F5FF000000000000000CFFFFFFFF000000000000000500000817150000002EFF0000000000000000FFFFFFFF00000000000000020000081716000001E4FF0000000000000000FFFFFFFF000000000000001100

'0A000001000817170000020CFF0000000000000000FFFFFFFF0000000000000001000008171800000013FF0000000000000000FFFFFFFF0000000000021D300000081719000000C1FF000000000000000000001C200000000000000001000008171A00000072FF0000000000000000FFFFFFFF0000000000000001000008171B0000007FFF0000000000000000FFFFFFFF000000000000000900

'0A0000020008171C0000007BFF0000000000000000FFFFFFFF000000000000000C000008171D00000077FF0000000000000000FFFFFFFF0000000000000004000008171E000000A0FF0000000000000000FFFFFFFF000000000000000C000008171F00000054FF0000000000000009FFFFFFFF0000000000000001000008172000000054FF0000000000000001FFFFFFFF000000000000000100

'0A000003

'00081721
'00000076
'FF
'00000000
'00000000
'FFFFFFFF
'00000000
'00000001
'00

'00081722
'0000007D
'FF
'00000000
'00000000
'FFFFFFFF
'00000000
'00000002
'00

'00081723
'0000010FFF0000000000000000FFFFFFFF00000000
'00000001
'00

'00081724000001FFFF0000000000000000FFFFFFFF000000000000000B000008172500000078FF0000000000000000FFFFFFFF000000000000000300

        On Error Resume Next
        'Chat ChrtoHex(Data)
        Dim T1, T2 As String 'T1 แบบที่ส่ง (A0 02) คือส่งหมดแล้ว,T2 ลำดับที่ส่ง
        T1 = Mid(data, 1, 2): T2 = Mid(data, 3, 2)
        If T1 = Chr(&HA) & Chr(&H1) Then ReDim Inv(0): frmInv.lvInv.ListItems.Clear
        Dim Qu, Refine As String, LV As String, Dura As String, State As Integer
        Dim Item_Id As String, Item_Type As String, Blog As String
        Dim DString As String, Op As Integer
        Dim row As Integer, col As Integer
        State = CInt(MakeLong(Mid(data, 3, 1)))
        row = CInt(MakeLong(Mid(data, 4, 1)))
        DString = Right(data, Len(data) - 4)  'ฝากไว้ก่อน
        col = 0
        Do While Len(DString) >= 30
            If Mid(DString, 1, 4) = Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) Then
                DString = Mid(DString, 5)
                GoTo a
            End If
            With Inv(UBound(Inv))
                .id = Mid(DString, 1, 4)
                .Type = MakeLong(Mid(DString, 5, 4))
                If .Type = &H367 Or .Type = &H368 Then
                    frmMain.mngetfeeden.Enabled = True
                    frmMain.mngetfeeden.Enabled = True
                    frmMain.mnPet(1).Enabled = True
                    frmMain.mnPet(2).Enabled = True
                End If
                .equip = CInt(MakeLong(Mid(DString, 9, 1)))
                .row = row
                .col = col
                .Refine = MakeLong(Mid(DString, 10, 4))
                .name = IIf(.Refine > 0, "+" & .Refine & " ", "") & Return_Item_Name(.Type)
                .LV = MakeLong(Mid(DString, 14, 4))
                .Durability = -1
                If Mid(DString, 18, 4) <> Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) Then .Durability = MakeLong(Mid(DString, 18, 4))
                .Amount = MakeLong(Mid(DString, 26, 4))
                'MessageBox.Show(makelong(Item_Id))
            End With
            Op = MakeLong(Mid(DString, 30, 1)) * 2
            If Len(DString) >= 30 Then DString = Right(DString, (Len(DString) - (30 + Op)))
            ReDim Preserve Inv(UBound(Inv) + 1)
a:
            col = col + 1
        Loop
        'If T1 = Chr(&HA) & Chr(&H2) Then frmItem.ManageItem()
        Update_Inv
End Sub 'Warehouse

Private Sub Pack_C00(data As String)
   '     On Error Resume Next
'        Dim Type As Byte
        Dim i As Integer
        Dim X As Integer, tmp() As String
        'If LoginChar(Char).ID = Mid(Data, 4, 4) Then Chat ChrtoHex(Data)
        For i = 0 To UBound(People) - 1
            If People(i).id = Mid(data, 4, 4) Then
                People(i).Coord = Mid(data, 12, 16)
                People(i).Pos = convert_coord(People(i).Coord)
                If Active And Not Sitting Then Check_Pos cur, People(i).Pos
                'Update_People
                Exit Sub
            End If
        Next
    End Sub 'test Owned walk and follow
    
Private Sub Pack_C01(data As String)
        'On Error Resume Next
        Dim id As String
        Dim Coord As String
        Dim i As Integer
        id = Mid(data, 4, 4)
        Coord = Mid(data, 12, 16)
        For i = 0 To UBound(Monster) - 1
            If Monster(i).id = Mid(data, 4, 4) Then
                Monster(i).Coord = Coord
                Monster(i).Pos = convert_coord(Coord)
                If Monster(i).id = C_Atk.id Then C_Atk = Monster(i): Update_Curmon
                'If C_Atk.ID <> "" Then Check_Pos cur, Monster(i).Pos
                If Active And Not Sitting Then Check_Pos cur, Monster(i).Pos
                'Update_Monster
                Exit For
            End If
        Next i
End Sub 'monster walk

Private Sub Pack_C02(data As String)
'Chat ChrtoHex(Data)
        Dim id As String
        Dim Coord As String
        Dim i As Integer
        id = Mid(data, 4, 4)
        Coord = Mid(data, 12, 16)
        For i = 0 To UBound(Apet) - 1
            If Apet(i).id = Mid(data, 4, 4) Then
                Apet(i).Coord = Coord
                Apet(i).Pos = convert_coord(Coord)
                If Apet(i).id = pet.id Then pet = Apet(i): Update_Pet
                Update_Apet
                Exit For
            End If
        Next i
End Sub

Private Sub Pack_E00(ByVal data As String)
    'On Error Resume Next
    Dim id As String
    Dim Monster_ID As String
    Dim i As Integer, X As Integer, z As Integer
    Dim hp, atk As Long, found As Boolean, tmp() As String

    id = Mid(data, 3, 4)
    Monster_ID = Mid(data, 13, 4)
    hp = MakeLong(Mid(data, 17, 4))
    atk = MakeLong(Mid(data, 25, 4))

    If (id = pet.id Or Monster_ID = pet.id) Then Chat ChrtoHex(data)
    
    i = Return_Monster(Monster_ID)
    If i > -1 Then
        Monster(i).hp = hp
        If Monster(i).MaxHP < hp Then Monster(i).MaxHP = hp
        If (Monster_ID = C_Atk.id) Then C_Atk = Monster(i)
        If C_Atk.MaxHP < hp Then C_Atk.MaxHP = hp
        If C_Atk.hp <= 0 Then Pack_801 Chr(&H0) & Chr(&H0) & C_Atk.id
        If hp <= 0 Then Pack_801 Chr(&H0) & Chr(&H0) & Monster_ID
        Update_Curmon
                
        If id = LoginChar(Char).id Then
            Stat Chr(3) & "3You Attack to " & Monster(i).name & ", " & IIf(atk <> 0, atk & " Damage", "Miss!")
            delay_Mon_Atk = 15
            'frmAtk.atkResPonse.Enabled = False
            Exit Sub
        End If
        
        For X = 0 To UBound(People) - 1
            If People(X).id = id Then
                Monster(i).IsAttack = True
                Stat People(X).name & " attack to " & Monster(i).name & ", " & IIf(atk <> 0, atk & " Damage", "Miss!") & " " & IIf(Monster(i).id = C_Atk.id, Chr(3) & "0,6[Jam!]" & Chr(3), "")
                Exit For
            End If
        Next
                
        If opt.party.atk.auto And id <> LoginChar(Char).id Then
            tmp = Split(opt.party.atk.name, ";")
            For z = 0 To UBound(tmp)
                For X = 0 To UBound(party) - 1
                    If party(X).id = id And party(X).name = Trim(tmp(z)) Then
                        Monster(i).Atk_me = True
                        C_Atk = Monster(i)
                        delay_Mon_Atk = 15
                        C_Atk.hp = hp
                        Update_Curmon
                        Exit Sub
                    End If
                Next
            Next
        End If
        Monster(i).IsAttack = True
    End If
End Sub

Private Sub Pack_E01(ByVal data As String)
    'On Error Resume Next
    Dim id As String, i As Integer, found As Boolean, z As Integer, X As Integer
    Dim Monster_ID As String, hp As Long, atk As String, tmp() As String
    id = Mid(data, 13, 4)
    Monster_ID = Mid(data, 3, 4)
    atk = MakeLong(Mid(data, 25, 4))
    'Pet
    '0E01
    '0014A6A1
    '00
    'FFFFFFFF
    '02
    '00004578
    '00000058
    '00000000
    '00000001
    '0D04
    
    i = Return_Monster(Monster_ID)
    If i > -1 And id = LoginChar(Char).id Then
        Stat Chr(3) & "5" & Monster(i).name & "  Attack you,  " & IIf(atk <> 0, atk & " Damage", "Miss!")
        Monster(i).Atk_me = True
        C_Atk.Atk_me = True
        If Sitting Then StandUp
        If Monster_ID = C_Atk.id Then delay_Mon_Atk = 15
        If C_Atk.id = "" And Not frmBot.tmrWalk.Enabled Then C_Atk = Monster(i)
        Update_Curmon
        Exit Sub
    ElseIf i > -1 And id = pet.id And frmMain.tbPet.Visible Then
        Stat Chr(3) & "7" & Monster(i).name & "  Attack your Pet,  " & IIf(atk <> 0, atk & " Damage", "Miss!")
        Monster(i).Atk_me = True
        If Sitting Then StandUp
        If Monster_ID = C_Atk.id Then delay_Mon_Atk = 15
        C_Atk = Monster(i)
        Update_Curmon
        Exit Sub
    ElseIf id = LoginChar(Char).id Then
        If Sitting Then StandUp
        Monster(UBound(Monster)).id = Monster_ID
        Monster(UBound(Monster)).name = "[Monster " & ChrtoHex(Monster_ID) & "]"
        Monster(UBound(Monster)).Coord = Convert_Point(cur.X, cur.y)
        Monster(UBound(Monster)).MaxHP = 1
        Monster(UBound(Monster)).Pos = cur
        Monster(UBound(Monster)).Atk_me = True
        Stat Chr(3) & "5Bug: " & Monster(UBound(Monster)).name & " Attack you, " & IIf(atk <> 0, atk & " Damage", "Miss!")
        If C_Atk.id = "" And Not frmBot.tmrWalk.Enabled Then C_Atk = Monster(UBound(Monster)):        delay_Mon_Atk = 15
        Update_Curmon
        ReDim Preserve Monster(UBound(Monster) + 1)
        Exit Sub
    ElseIf id = pet.id And frmMain.tbPet.Visible Then
        If Sitting Then StandUp
        Monster(UBound(Monster)).id = Monster_ID
        Monster(UBound(Monster)).name = "[Monster " & ChrtoHex(Monster_ID) & "]"
        Monster(UBound(Monster)).Coord = Convert_Point(pet.Pos.X, pet.Pos.y)
        Monster(UBound(Monster)).MaxHP = 1
        Monster(UBound(Monster)).Pos = pet.Pos
        Monster(UBound(Monster)).Atk_me = True
        Stat Chr(3) & "7Bug: " & Monster(UBound(Monster)).name & " Attack your Pet, " & IIf(atk <> 0, atk & " Damage", "Miss!")
        C_Atk = Monster(UBound(Monster))
        Update_Curmon
        ReDim Preserve Monster(UBound(Monster) + 1)
        Exit Sub
    ElseIf i > -1 Then
        For X = 0 To UBound(People) - 1
            If People(X).id = id Then
                Monster(i).IsAttack = True
                Stat Monster(i).name & " attack to " & People(X).name & ", " & IIf(atk <> 0, atk & " Damage", "Miss!")
                Exit For
            End If
        Next
    Else
        For X = 0 To UBound(People) - 1
            If People(X).id = id Then
                Monster(UBound(Monster)).id = Monster_ID
                Monster(UBound(Monster)).name = "[Monster " & ChrtoHex(Monster_ID) & "]"
                Monster(UBound(Monster)).Coord = People(X).Coord
                Monster(UBound(Monster)).MaxHP = 1
                Monster(UBound(Monster)).Pos = People(X).Pos
                Monster(UBound(Monster)).IsAttack = True
                Stat Monster(UBound(Monster)).name & " attack to " & People(X).name & ", " & IIf(atk <> 0, atk & " Damage", "Miss!")
                ReDim Preserve Monster(UBound(Monster) + 1)
                Exit For
            End If
        Next
    End If
        
    If opt.party.protect.auto Then
        tmp = Split(opt.party.protect.name, ";")
        For z = 0 To UBound(tmp)
            For X = 0 To UBound(party) - 1
                If party(X).id = id And party(X).name = Trim(tmp(z)) And Trim(tmp(z)) <> "" Then
                    If C_Atk.id <> Monster_ID Then
                        If Sitting Then StandUp
                        If i > -1 Then
                            C_Atk = Monster(i)
                        Else
                            C_Atk.id = Monster_ID
                            C_Atk.name = "[Monster " & ChrtoHex(Monster_ID) & "]"
                            C_Atk.Coord = party(X).Coord
                            C_Atk.Pos = party(X).Pos
                        End If
                        Start_Walk C_Atk.Pos.X, C_Atk.Pos.y
                    End If
                    Update_Curmon
                    Stat Chr(3) & "13" & C_Atk.name & " Attack " & party(X).name & " [Party], " & IIf(atk <> 0, atk & " Damage", "Miss!")
                    Exit Sub
                End If
            Next
        Next
    End If
    If i > -1 Then Monster(i).IsAttack = True
End Sub

Private Sub Pack_F00(data As String, State As Integer)
    'On Error Resume Next
    Dim name, id As String, i As Integer
    Dim Textz As String
    'Dim Sep As String() = {Chr(&H0)}
    Dim X() As String, n As String
    id = Mid(data, 3, 4)
    '0F 00 00 02 8C E1 B9 E9 CD A7 A1 D4           . . .  . . . . . แ น ้ อ ง ก ิ
   'BF BB D9 BB D9 EB 00 00 C1 D9 B9 E2 B5 B9 20 39           ฟ ป ู ป ู ๋ . . ม ู น โ ต น  9
   '30 30 30 00

    i = InStr(7, data, Chr(&H0), vbBinaryCompare)
    name = Mid$(data, 7, i - 7)
    Textz = Mid(data, i + 2)
    If State = 4 Then
    i = InStr(2, Textz, Chr(&H0), vbBinaryCompare)
    Textz = Mid$(Textz, i + 1, i)
    End If
    Textz = MakeString(Textz)
    i = Return_People(id)
    If i > -1 Then
        n = "[" & People(i).Pos.X & "," & People(i).Pos.y & "][" & Distant(cur, People(i).Pos) & "] "
    End If
    If i > -1 And Not frmBot.tmrChat.Enabled Then
        frmBot.tmrChat.Interval = RandNum(3500, 1500)
        frmBot.tmrChat.Enabled = True
    End If
    Select Case State
        Case 0
            Chat n & Chr(&H3) & "1" & name & ": " & Textz
        Case 1
            Chat n & Chr(&H3) & "8[Party] " & name & ": " & Textz
        Case 2
            Chat n & Chr(&H3) & "12[Guild] " & name & ": " & Textz
        Case 3
            Chat n & Chr(&H3) & "3[Trade] " & name & ": " & Textz
        Case 4
            Chat n & Chr(&H3) & "5[Whisper] " & name & ": " & Textz
        Case 5
            Chat n & Chr(&H3) & "13[Shout] " & name & ": " & Textz
        Case 6
            Chat n & Chr(&H3) & "7[GM] " & name & ": " & Textz
        Case Else
            Chat n & Chr(&H3) & "0,4[ELSE]" & name & ": " & Textz
        End Select
    'frmChat.ManageChat(Name, Textz, Ty, Make_FullHex(ID))
End Sub 'Chat Management

Private Sub Pack_1000(data As String)
'10 00 00 07 02 00 19 6F 3F 00 00 00 00
Dim i As Integer
For i = 0 To UBound(Inv) - 1
    If Inv(i).id = Mid(data, 6, 4) Then
        Stat "[Item] ใช้ item: " & Inv(i).name
        Exit Sub
    End If
Next
End Sub


Private Sub Pack_1004(data As String)
'10 00 00 07 02 00 19 6F 3F 00 00 00 00
Dim i As Integer, found As Boolean
found = False
For i = 0 To UBound(Inv) - 1
    If Inv(i).row = MakeLong(Mid(data, 4, 1)) And Inv(i).col = MakeLong(Mid(data, 5, 1)) Then found = True
    If found Then Inv(i) = Inv(i + 1)
Next
If found Then ReDim Preserve Inv(UBound(Inv) - 1): delay_refresh_inv = 3
End Sub

Private Sub Pack_1005(data As String)
'10 05 0A 00 00 00 FF FF FF FF 00 08 01 00 3C A4 81 ถอดออก
'10 05 0A 00 08 01 00 3C A4 81 00 00 00 FF FF FF FF  ใส่
Dim i As Integer, found As Boolean
If Mid(data, 7, 4) = Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) Then 'ถอด
    For i = 0 To UBound(Inv) - 1
         If Inv(i).id = Mid(data, 14, 4) Then
            Stat "Item: ถอดอุปกรณ์ " & Inv(i).name
            If Mid(data, 3, 1) = Chr(&HA) Then
                pet.id = ""
                frmMain.tbPet.Visible = False
                frmMain.Form_Resize
            End If
            delay_refresh_inv = 3
            Exit Sub
         End If
    Next
ElseIf Mid(data, 14, 4) = Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) Then 'ใส่
    For i = 0 To UBound(Inv) - 1
         If Inv(i).id = Mid(data, 7, 4) Then
            Stat "Item: ใส่อุปกรณ์ " & Inv(i).name
            delay_refresh_inv = 3
            Exit Sub
         End If
    Next
End If
End Sub

Private Sub Pack_1007(ByVal data As String)
    'On Error Resume Next
    Dim val As Long
    val = MakeLong(Mid(data, 30, 4))
    Stat "Item : " & IIf(val > 0, "ได้รับ", "ใช้") & " [" & Return_Item_Name(MakeLong(Mid(data, 10, 4))) & "][" & Abs(MakeLong(Mid(data, 30, 4))) & "]"
    'delay_refresh_inv = 3
    Inv(UBound(Inv)).row = MakeLong(Mid(data, 4, 1))
    Inv(UBound(Inv)).col = MakeLong(Mid(data, 5, 1))
    Inv(UBound(Inv)).id = Mid(data, 6, 1)
    Inv(UBound(Inv)).Amount = val
    Inv(UBound(Inv)).name = Return_Item_Name(MakeLong(Mid(data, 10, 4)))
    Inv(UBound(Inv)).Type = MakeLong(Mid(data, 10, 4))
    ReDim Preserve Inv(UBound(Inv) + 1)
    '1007
    '00
    '0504
    '004E1C72
    '0000009C
    'FF0000000000000000FFFFFFFF000000000000012C00

End Sub 'Recv Item

Private Sub Pack_1008(ByVal data As String)
    'On Error Resume Next
        Dim id, IdHex, Alli, val As Long, Dura As Long
        Dim i As Integer
        id = Mid(data, 6, 4)
        Alli = Mid(data, 22, 8)
        val = MakeLong(Mid(data, 30, 8))
            For i = 0 To UBound(Inv) - 1
                If Inv(i).id = Mid(data, 6, 4) Then
                    If Mid(data, 18, 4) <> (Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF)) Then
                        Stat "Item : [" & Inv(i).name & "] เหลือความทนทาน " & MakeLong(Mid(data, 18, 4))
                    Else
                        Stat "Item : " & IIf(val > 0, "ได้รับ", "ใช้") & " [" & Inv(i).name & "][" & MakeLong(Mid(data, 30, 8)) & "]"
                        Inv(i).Amount = MakeLong(Mid(data, 22, 8))
                        Update_Inv
                    End If
                    If Inv(i).Amount = 0 Then delay_refresh_inv = 3
                    Exit Sub
                End If
            Next
End Sub 'Recv item

Private Sub Pack_1009(ByVal data As String)
    On Error Resume Next
    '10 09
    '00 18 08 7C
    '00 00 00 13
    '00 00 00 00 00 00 00 44
    '4C A0 A3 44 EE E8 CF 43 33 73 30 43 FA 7E CA BF
    '00 01 00 0A 50 F3
    With Item(UBound(Item))
        .id = Mid(data, 3, 4)
        .Type = MakeLong(Mid(data, 7, 4))
        .name = Return_Item_Name(.Type)
        .Amount = MakeLong(Mid(data, 11, 8))
        .Coord = Mid(data, 19, 16)
        .Pos = convert_coord(.Coord)
        Stat "Item : Drop [" & .name & "][" & .Amount & "] on X:" & .Pos.X & " Y:" & .Pos.y & "[" & Distant(cur, .Pos) & "]"
        'If Active And Distant(.Pos, cur) <= 10 Then SendPacket Chr(&H90) & Chr(&H1) & Chr(&H0) & LoginChar(Char).ID & .ID, 1  'เก็บของ
        If Active And Not Sitting Then Check_Pos cur, .Pos   '90 01 00 00 02 A3 75 00 1E A6 29
    End With
    ReDim Preserve Item(UBound(Item) + 1)
    Pack_801 Chr(&H0) & Chr(&H0) & Mid(data, 37, 4)
End Sub

Private Sub Pack_100A(ByVal data As String)
    'On Error Resume Next
    'Chat "100A " & ChrtoHex(data)
    'Exit Sub
    '10 0A
    '00 18 07 4B
    '00 00 00 13
    '00 00 00 00 00 00 00 4E
    '42 65 A8 44 B4 F5 B3 43 AE 07 2E 43 AB 0B 2A 40
    '00
    With Item(UBound(Item))
        .id = Mid(data, 3, 4)
        .Type = MakeLong(Mid(data, 7, 4))
        .name = Return_Item_Name(.Type)
        .Amount = MakeLong(Mid(data, 11, 8))
        .Coord = Mid(data, 19, 16)
        .Pos = convert_coord(.Coord)
        Stat "Item : Drop [" & .name & "][" & .Amount & "] on X:" & .Pos.X & " Y:" & .Pos.y & " [" & Distant(cur, .Pos) & "]"
        If Active And Not Sitting Then Check_Pos cur, .Pos
        'SendPacket Chr(&H10) & Chr(&H1) SendPacket Chr(&H& LoginChar(Char).ID & .ID, 1   'เก็บของ
    End With
    ReDim Preserve Item(UBound(Item) + 1)
End Sub

Private Sub Pack_100B(ByVal data As String)
    'On Error Resume Next
    Dim i As Integer, found As Boolean
    For i = 0 To UBound(Item) - 1
        If Item(i).id = Mid(data, 3, 4) Then
            found = True
            Stat "Item : Disappear [" & Item(i).name & "] on X:" & Item(i).Pos.X & " Y:" & Item(i).Pos.y
        End If
        If found Then
            Item(i) = Item(i + 1)
        End If
    Next
    If found Then ReDim Preserve Item(UBound(Item) - 1)
End Sub 'Item disappeared

Private Sub Pack_1116(ByVal data As String)
Dim i As Integer
If delay_change_equip <> 0 Then Exit Sub
Stat "คุณสวมใส่อุปกรณ์ไม่ถูกต้อง"
Update_Curnpc
If N_Atk.id = "" Then Exit Sub
If tmpEquip >= UBound(Inv) - 1 Then tmpEquip = 0
For i = tmpEquip To UBound(Inv) - 1
    With Inv(i)
        tmpEquip = i + 1
        If .Durability > 10 Then
            Stat "พยายามใส่ " & .name
            SendPacket Chr(&H90) & Chr(&H5) & Chr(&H2) & Chr(&H0) & Chr(.row) & Chr(.col) & .id, 1
            '90 05 02 00 05 04 00 2B 22 61
            '90 05 02 00 06 02 00 0F C0 87
            delay_change_equip = 2
            Exit Sub
        End If
    End With
Next
End Sub

Private Sub Pack_1138(ByVal data As String)
opt.pet = False
Dim stime As Long
stime = MakeLong(Mid(data, 7, 4))
Chat "Pet: สัตว์เลี้ยงตาย! ท่านเหลือเวลา " & (stime \ 3600) & "ชั่วโมง " & (stime \ 60) Mod 60 & "นาที " & (stime Mod 60) & "วินาที เพื่อเปิดผลึกอีกครั้ง"
End Sub

Public Sub Pack_1300(data As String)
'13
'00 00 00 07    แมบ
'00 20 16 44 00 C0 90 43 A4 70 C7 42 00 00 00 00  ตำแหน่ง
'00
Dim i As Integer
Map = CInt(MakeLong(Mid(data, 5, 1)))
If opt.path.lock.auto Then delay_go_map = 5
LoginChar(Char).Coord = Mid(data, 6, 16)
LoginChar(Char).Pos = convert_coord(LoginChar(Char).Coord)
'Chat "Pack_1300: " & Map & " X:" & cur.x & " Y:" & cur.y
tmpMap = Mid(LoginChar(Char).Coord, 9, 8)
Plot_Dot LoginChar(Char).Pos
Constate = 4
delay_feed_pet = 15
frmBot.tmrLoad.Enabled = True
frmMap.slMap_Change
End Sub

Public Sub Pack_1500(data As String)
Dim i As Integer, d As String, Pt As Coord, cnt As Integer
cnt = MakeLong(Mid(data, 3, 1))
ReDim sPoint(0)
frmPoint.lvPoint.ListItems.Clear
If cnt = 0 Then Exit Sub
d = Mid(data, 4)
d = Left(d, Len(d) - 4)
Do While Len(d) > 0
    With sPoint(UBound(sPoint))
        .id = MakeLong(Left(d, 1))
        .Map = MakeLong(Mid(d, 2, 4))
        .Coord = Mid(d, 6, 8)
        .Pos = convert_coord(Mid(d, 6, 8))
        i = InStr(14, d, Chr(&H0), vbBinaryCompare)
        .name = Mid(d, 14, i - 14)
        cnt = frmPoint.lvPoint.ListItems.Add.Index
        frmPoint.lvPoint.ListItems(cnt).Text = .id
        frmPoint.lvPoint.ListItems(cnt).SubItems(1) = .name
        frmPoint.lvPoint.ListItems(cnt).SubItems(2) = .Map
        frmPoint.lvPoint.ListItems(cnt).SubItems(3) = .Pos.X & " : " & MakePort(Mid(d, 6, 4)) & " : " & ChrtoHex(Mid(d, 6, 4))
        frmPoint.lvPoint.ListItems(cnt).SubItems(4) = .Pos.y & " : " & MakePort(Mid(d, 10, 4)) & " : " & ChrtoHex(Mid(d, 10, 4))
        d = Mid(d, i + 1)
    End With
    ReDim Preserve sPoint(UBound(sPoint) + 1)
Loop
End Sub

Private Sub Pack_1800(data As String)
    'On Error Resume Next
    Dim ty As Byte
    Dim id As String
    Dim name As String
    Dim X As Integer, tmp() As String
    ty = MakePort(Mid(data, 3, 1))
    id = Mid(data, 4, 4)
    name = Trim(Mid(data, 8, Len(data) - 8))
    If id <> LoginChar(Char).id Then
        Stat Chr(&H3) & "9Party : Party Request!!! By : " & name
        If opt.party.resp.auto Then
            tmp = Split(opt.party.resp.name, ";")
            For X = 0 To UBound(tmp)
                If name = Trim(tmp(X)) And Trim(tmp(X)) <> "" Then
                    frmBot.tmrParty.Enabled = True
                    Stat "ตอบรับปาร์ตี้กับ : " & name
                    Exit Sub
                End If
            Next
        End If
        frmInvite.tmrdelay.Enabled = True
        frmInvite.txt.Caption = "ขอปาร์ตี้โดย : " & name & " แบบ : " & ty
        frmInvite.show
    Else
        Stat Chr(&H3) & "9Party : Party Request!!! for : " & name
    End If
End Sub 'Recv Party by anyone

Public Sub Pack_1805(data As String)
Dim sdata As String, i As Integer
'01 81 00 00 13 C4 00         ๘  C ๖ ว . C . . . . . . . ฤ .
'   00 00 00 00 2F 18 05 01 00 01 77 4F 4F 6F 4D 69           . . . . / . . . . . w O O o M i
'   6B 61 6F 4F 00 02 02 00 00 00 27 00 00 04 D4 00           k a o O . . . . . . ' . . . ิ .
'   00 04 D4 00 00 05 AA 00 00 05 AA 77 96 95 44 5D           . . ิ . . . ช . . . ช w . . D ]
'  39 74 44 00        9 t D .
sdata = Mid(data, 4)
Do While Len(sdata) > 0
    party(UBound(party)).id = Left(sdata, 4)
    i = InStr(5, sdata, Chr(&H0), vbBinaryCompare)
    party(UBound(party)).name = Mid(sdata, 5, i - 5)
    party(UBound(party)).Class = CInt(MakeLong(Mid(sdata, i + 1, 1)))
    party(UBound(party)).level = CInt(MakeLong(Mid(sdata, i + 6, 1)))
    party(UBound(party)).hp = MakeLong(Mid(sdata, i + 7, 4))
    party(UBound(party)).MaxHP = MakeLong(Mid(sdata, i + 11, 4))
    party(UBound(party)).mp = MakeLong(Mid(sdata, i + 15, 4))
    party(UBound(party)).MaxMp = MakeLong(Mid(sdata, i + 19, 4))
    party(UBound(party)).Coord = Mid(sdata, i + 23, 8)
    party(UBound(party)).Pos = convert_coord(party(UBound(party)).Coord)
    sdata = Mid(sdata, i + 32)
    ReDim Preserve party(UBound(party) + 1)
Loop
Update_Party
End Sub
    
Private Sub Pack_1806(data As String)
Dim i As Integer, F As Boolean
For i = 0 To UBound(party) - 1
    If party(i).id = Mid(data, 3, 4) Then F = True
    If F Then party(i) = party(i + 1)
Next
If F Then ReDim Preserve party(UBound(party) - 1)
Update_Party
End Sub
    
Private Sub Pack_1808(data As String)
ReDim party(0)
Update_Party
End Sub
    
Private Sub Pack_1809(data As String)
    'On Error Resume Next
    Dim nameid As String
    Dim hp As Long, MaxHP As Long, mp As Long, MaxMp As Long, tmp() As String
    Dim LV As Integer, Coord As String
    Dim X As Integer, i As Integer, z As Integer, ds As Long, tx As Long, ty As Long, to2 As Coord
    nameid = Mid(data, 3, 4)
    LV = MakeLong(Mid(data, 7, 4))
    hp = MakeLong(Mid(data, 11, 4))
    MaxHP = MakeLong(Mid(data, 15, 4))
    mp = MakeLong(Mid(data, 19, 4))
    MaxMp = MakeLong(Mid(data, 23, 4))
    Coord = Mid(data, 27, 8)
    '18 09 00 02 83 45 00 00 00 10 00 00 01 1A 00 00 02 C6 00 00 01 5A 00 00 01 5A 6F 8C B4 44 27 46 E3 43 00
    For i = 0 To UBound(party) - 1
        If party(i).id = nameid Then
            party(i).hp = hp
            party(i).MaxHP = MaxHP
            party(i).mp = mp
            party(i).MaxMp = MaxMp
            party(i).level = LV
            party(i).Coord = Coord
            party(i).Pos = convert_coord(Coord)
            Update_Party
            If opt.party.follow.auto And Active Then
                tmp = Split(opt.party.follow.name, ";")
                For X = 0 To UBound(tmp)
                    If party(i).name = tmp(X) And Trim(tmp(X)) <> "" Then
                        ds = Distant(cur, party(i).Pos)
                        If ds > 0 Then
                            tx = party(i).Pos.X - cur.X
                            ty = party(i).Pos.y - cur.y
                            to2.X = (cur.X + (tx * (ds - opt.party.follow.Min) / ds))
                            to2.y = (cur.y + (ty * (ds - opt.party.follow.Min) / ds))
                            If Distant(party(i).Pos, cur) > opt.party.follow.Max And Distant(to2, Walking_stop) > 1 Then
                                Stat "[Follow] Found: " & Trim(tmp(X)) & "   X:" & party(i).Pos.X & " Y:" & party(i).Pos.y
                                Start_Walk to2.X, to2.y
                                GoTo z
                            End If
                        End If
                    End If
                Next
            End If
z:
            If opt.party.hp1.auto And delay_use_skill = 0 Then
                tmp = Split(opt.party.hp1.name, ";")
                For X = 0 To UBound(tmp)
                    If party(i).id = nameid And party(i).name = tmp(X) And Trim(tmp(X)) <> "" Then
                        If CInt((hp / MaxHP) * 100) <= opt.party.hp1.hp And (LoginChar(Char).mp / LoginChar(Char).MaxMp) * 100 >= opt.party.hp1.mp Then
                            For z = 0 To UBound(skill) - 1
                                If MakeLong(skill(z).id) = CLng(opt.party.hp1.skill) Then
                                    useSkill nameid, skill(z).id, 0
                                    Exit Sub
                                End If
                            Next
                        End If
                    End If
                Next
            End If
            If opt.party.mp1.auto Then
                tmp = Split(opt.party.mp1.name, ";")
                For X = 0 To UBound(tmp)
                    If party(i).id = nameid And party(i).name = tmp(X) And Trim(tmp(X)) <> "" Then
                        If CInt((mp / MaxMp) * 100) <= opt.party.mp1.hp And (LoginChar(Char).mp / LoginChar(Char).MaxMp) * 100 >= opt.party.mp1.mp Then
                            For z = 0 To UBound(skill) - 1
                                If MakeLong(skill(z).id) = CLng(opt.party.mp1.skill) Then
                                    useSkill nameid, skill(z).id, 0
                                    Exit Sub
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
End Sub 'Party mp hp recv
    
Public Sub Pack_1B00(data As String)
Dim sdata As String, i As Integer, F As Boolean
ReDim skill(0)
sdata = Mid(data, 4)
Do While Len(sdata) > 0
    F = False
    skill(UBound(skill)).id = Left(sdata, 4)
    For i = 0 To UBound(tmpSkill) - 1
        If tmpSkill(i).Type = MakeLong(skill(UBound(skill)).id) Then
            skill(UBound(skill)).name = tmpSkill(i).name
            skill(UBound(skill)).Detail = tmpSkill(i).Detail
            F = True
            Exit For
        End If
    Next
    If Not F Then
        skill(UBound(skill)).name = "Skill: " & ChrtoHex(Left(sdata, 4))
        skill(UBound(skill)).Detail = ""
    End If
    skill(UBound(skill)).LV = CInt(MakePort(Mid(sdata, 5, 1)))
    sdata = Mid(sdata, 6)
    ReDim Preserve skill(UBound(skill) + 1)
Loop
Update_Skill
End Sub

Public Sub Pack_1B03(data As String)
Dim p1 As String, sk As Long, p2 As String
p1 = Mid(data, 4, 4)
sk = MakeLong(Mid(data, 8, 4))
p2 = Mid(data, 13, 4)
If (p1 = LoginChar(Char).id And p2 = LoginChar(Char).id) Then 'ใช้กับตัวเราเอง
    Stat "You using skill [" & Return_Skill_Name(sk) & "]"
ElseIf p1 = LoginChar(Char).id And Return_People(p2) > -1 Then 'เราใช้สกิลกับคนอื่น
    Stat "You using skill [" & Return_Skill_Name(sk) & "] on " & People(Return_People(p2)).name
ElseIf p1 = LoginChar(Char).id And Return_Monster(p2) > -1 Then 'เราใช้สกิลกับมอน
    Stat "You using skill [" & Return_Skill_Name(sk) & "] on " & Monster(Return_Monster(p2)).name
ElseIf Return_People(p1) > -1 And p2 = LoginChar(Char).id Then 'คนอื่นใช้สกิลกับเรา
    Stat People(Return_People(p1)).name & " using skill [" & Return_Skill_Name(sk) & "] on you"
ElseIf Return_Monster(p1) > -1 And p2 = LoginChar(Char).id Then 'มอนใช้สกิลกับเรา
    Stat Monster(Return_Monster(p1)).name & " using skill [" & Return_Skill_Name(sk) & "] on you"
ElseIf Return_People(p1) > -1 And Return_People(p2) > -1 Then 'คนอื่นใช้สกิลกับคนอื่น
    Stat People(Return_People(p1)).name & " using skill [" & Return_Skill_Name(sk) & "] on " & People(Return_People(p2)).name
ElseIf Return_People(p1) > -1 And Return_Monster(p2) > -1 Then  'คนอื่นใช้สกิลกับมอน
    Stat People(Return_People(p1)).name & " using skill [" & Return_Skill_Name(sk) & "] on " & Monster(Return_Monster(p2)).name
ElseIf Return_Monster(p1) > -1 And Return_People(p2) > -1 Then 'มอนใช้สกิลกับคนอื่น
    Stat Monster(Return_Monster(p1)).name & " using skill [" & Return_Skill_Name(sk) & "] on " & People(Return_People(p2)).name
ElseIf Return_Monster(p1) > -1 And Return_Monster(p2) > -1 Then 'มอนใช้สกิลกับมอน
    Stat Monster(Return_Monster(p1)).name & " using skill [" & Return_Skill_Name(sk) & "] on " & Monster(Return_Monster(p2)).name
ElseIf p1 = LoginChar(Char).id Then
    Stat "You using skill [" & Return_Skill_Name(sk) & "] on Unknown"
ElseIf p2 = LoginChar(Char).id Then
    Stat "Unknown using skill [" & Return_Skill_Name(sk) & "] on you"
End If
End Sub

Private Sub Pack_1D01(data As String)
        On Error Resume Next
        Dim id As String, Remain As String
        Dim i As Integer, X As Integer, found As Boolean
        '1D 01 00 00 0C E3 00 00 00 84 00 00 01 2C 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00
        id = Mid(data, 3, 4)
        Remain = Mid(data, 7, 4)
        If id = C_Atk.id Then
            C_Atk.hp = MakeLong(Remain)
            Update_Curmon
        End If
        For i = 0 To UBound(Monster) - 1
            If Monster(i).id = id Then
                Monster(i).hp = MakeLong(Remain)
                If Monster(i).hp <= 0 Then
                    Pack_801 Chr(&H0) & Chr(&H0) & Monster(i).id
                End If
                Exit Sub
            End If
        Next
            'ค่าที่เหลือ
        If opt.farm.auto Then
                For i = 0 To UBound(NPC) - 1
                    If NPC(i).id = id Then
                        NPC(i).hp = MakeLong(Remain)
                        If NPC(i).hp <= 1 Then
                            Pack_801 Chr(&H0) & Chr(&H0) & NPC(i).id
                            If N_Atk.id = id Then N_Atk.id = "": SelectFarm
                        End If
                        If N_Atk.id = id Then N_Atk = NPC(i): delay_select_npc = 15
                        Update_Curnpc
                        Update_npc
                        Exit Sub
                    End If
                Next
        End If
out:
    'Update_Monster
    End Sub 'Monster Hp

Public Sub Pack_1F02(data As String)
Dim a As Integer, b As Integer
a = Return_People(Mid(data, 5, 4))
b = Return_NPC(Mid(data, 10, 4))
If a > -1 And b > -1 Then
    Stat People(a).name & " กำลังเก็บเกี่ยว " & NPC(b).name & " [" & MakeLong(Mid(data, 14, 4)) & "]"
End If
End Sub

Private Sub Pack_2000(data As String)    'Exp And Sp
    'On Error Resume Next
    Dim EXP, SP As Integer
    Dim d As String
    '01 81 00 00 3E E5 00 00 00 00 00 0D  ->   20 00 00 00 00     00 00 06 F6     00 00 02 FF     'Exp   06 F6    1782   , 'SP 02 FF 767
    EXP = MakeLong(Mid(data, 6, 4))
    SP = MakeLong(Mid(data, 10, 4))
    sExp = sExp + EXP
    With LoginChar(Char)
        .EXP = .EXP + EXP
        .SP = .SP + SP
        Stat "earn " & EXP & "exp, " & SP & "sp"
    End With
    Update_Stats
End Sub 'Exp And Sp
    
Public Sub Pack_2100(data As String)
Dim i  As Integer, n As String
i = Return_People(Mid(data, 2, 4))
If i > -1 Then
    n = "[" & People(i).Pos.X & "," & People(i).Pos.y & "][" & Distant(cur, People(i).Pos) & "] "
    If Mid(data, 7, 1) = Chr(&H3) Then
        Stat n & People(i).name & ": " & IIf(Right(data, 1) = Chr(&H1), "Sitting", "Standup")
    ElseIf Mid(data, 7, 1) <> Chr(&H3) Then
        Chat n & People(i).name & ": Send Emotion [" & Emo(Mid(data, 7, 1)) & "]"
    End If
End If
End Sub

Private Sub Pack_2A02(data As String)
    Dim i As Integer, X As Integer
    i = MakePort(Mid(data, 6, 1))
    frmClan.lvClan.ListItems.Clear
    X = frmClan.lvClan.ListItems.Add.Index
    frmClan.lvClan.ListItems(X).Text = ""
    frmClan.lvClan.ListItems(X).SubItems(1) = "Guild: " & Mid(data, 7, i)
    frmClan.lvClan.ListItems(X).SubItems(2) = "Lv: " & MakePort(Mid(data, i + 11, 1))
    LoginChar(Char).Clan = Mid(data, 7, i)
    Update_Stats
End Sub
    
Private Sub Pack_2A03(data As String)
    Dim i As Integer, X As Integer, sdata As String
    sdata = Mid(data, 11)
    Do While Len(sdata) > 6
        X = frmClan.lvClan.ListItems.Add.Index
        frmClan.lvClan.ListItems(X).Text = ChrtoHex(Left(sdata, 4))
        i = InStr(5, sdata, Chr(&H0), vbBinaryCompare)
        frmClan.lvClan.ListItems(X).SubItems(1) = Mid$(sdata, 5, i - 4)
        frmClan.lvClan.ListItems(X).SubItems(2) = MakeHex(Mid(sdata, i, 6))
        sdata = Mid(sdata, i + 6)
    Loop
End Sub

Private Sub Pack_2A04(data As String)
    Dim i As Integer, X As Integer, sdata As String
'01 81 00 00 00 42        . . . . . . . . . . . . . . . B
 '  00 00 00 00 00 14 2A 04 00 00 00 06 00 01 77 4F           . . . . . . * . . . . . . . w O
 '  4F 6F 4D 69 6B 61 6F 4F 00 01
        i = InStr(11, data, Chr(&H0), vbBinaryCompare)
        Chat Chr(&H3) & "12[Guild] " & Mid$(data, 11, i - 11) & " เปลี่ยนสถานะ: " & IIf(Right(data, 1) = Chr(&H1), "Online", "Offine")
End Sub

Private Sub Pack_2C02(data As String)
Dim i As Integer
For i = 0 To UBound(People) - 1
    If People(i).id = Mid(data, 3, 4) Then
        People(i).Shop = Mid(data, 8, Len(data) - 8)
        Stat "[Shop] " & People(i).name & ": " & People(i).Shop
        Exit Sub
    End If
Next
End Sub

Private Sub Pack_3700(data As String)
'Chat ChrtoHex(Data)
If pet.id <> Mid(data, 6, 4) Or MakeLong(Mid(data, 5, 1)) <> 0 Then
    Exit Sub
End If
'37
'00000000
'00003E59
'21
'00000001
'00000000
'00000000
'00000000
'0000028E
'00000064
'00000064
'00000000000000340000006400000036000000640000000000000000000000000000000000FF
'[0:31:08]
'37
'00000004
'00003E59
'00000000

'[21:46:14] 37000000000000457822000000220000000000000DA20000000000002BE50000005D00000064000000010000005F0000006400000064000000640000000000000000000000000000000000FF00000000
'[21:46:14] 370000000400004578000000070000011C010000011D010000011E010000011F01000001200100000112170000011609
If Not frmMain.tbPet.Visible Then frmMain.tbPet.Visible = True:        frmMain.Form_Resize
pet.LV = MakeLong(Mid(data, 11, 4))
pet.EXP = MakeLong(Mid(data, 15, 8))
pet.MaxEXP = MakeLong(Mid(data, 23, 8))
pet.hp = MakeLong(Mid(data, 31, 4))
pet.MaxHP = MakeLong(Mid(data, 35, 4))
pet.SP = MakeLong(Mid(data, 39, 4))
pet.EN = MakeLong(Mid(data, 43, 4))
pet.MaxEN = MakeLong(Mid(data, 47, 4))
pet.FL = MakeLong(Mid(data, 51, 4))
pet.MaxFL = MakeLong(Mid(data, 55, 4))
Update_Pet
End Sub
