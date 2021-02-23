VERSION 5.00
Begin VB.Form frmBot 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   4455
   ClientTop       =   3915
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   5100
   Begin VB.Timer tmrParty 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1800
      Top             =   720
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   2250
   End
   Begin VB.Timer tmrChat 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2040
      Top             =   2310
   End
   Begin VB.Timer tmrSession2 
      Interval        =   850
      Left            =   1620
      Top             =   1470
   End
   Begin VB.Timer tmrRecon 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2670
      Top             =   2340
   End
   Begin VB.Timer tmrNoMon 
      Interval        =   500
      Left            =   1500
      Top             =   2040
   End
   Begin VB.Timer tmrWalk 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   2370
   End
   Begin VB.Timer tmrReply 
      Enabled         =   0   'False
      Interval        =   35000
      Left            =   2910
      Top             =   1470
   End
   Begin VB.Timer tmrSession 
      Interval        =   1000
      Left            =   840
      Top             =   1500
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Winsock As CSocketMaster
Attribute Winsock.VB_VarHelpID = -1
Private Session As Integer, DelayLoading As Integer
Private ProcessID As Long, index_update As Integer, timeout As Integer

Private Sub Form_Load()
   Set Winsock = New CSocketMaster
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Winsock = Nothing
End Sub

Private Sub tmrChat_Timer()
        frmBot.tmrChat.Enabled = False
End Sub

Private Sub tmrLoad_Timer()
    SendPacket Chr(&H85), 1
    tmrLoad.Enabled = False
End Sub

Private Sub tmrNoMon_Timer()
Dim i As Integer
If Constate < 4 Or C_Atk.id <> "" Or Not Active Or tmrWalk.Enabled Or Sitting Then Exit Sub
'If UBound(Item) > 0 Then SelectItem
If UBound(Item) > 0 Then
    For i = 0 To UBound(Item) - 1
        If Item(i).keep <= 5 Then
            If Distant(cur, Item(i).Pos) <= 2 Then
                SendPacket Chr(&H90) & Chr(&H1) & Chr(&H0) & LoginChar(Char).id & Item(i).id, 1
                Item(i).keep = Item(i).keep + 1
                Exit Sub
            ElseIf Distant(cur, Item(i).Pos) <= 16 Then
                Start_Walk Item(i).Pos.X, Item(i).Pos.Y
                Exit Sub
            End If
        End If
    Next
End If
If opt.farm.auto Then
    SelectFarm
ElseIf (opt.atk.auto Or opt.party.atk.auto) Then
    SelectMonster
End If
End Sub

Private Sub tmrParty_Timer()
tmrParty.Enabled = False
SendPacket Chr(&H98) & Chr(&H1), 1
End Sub

Private Sub tmrRecon_Timer()
If delayRecon = 0 Then
    frmMain.btnConnect.Caption = "Connect"
    frmMain.btnConnect_Click
    tmrRecon.Enabled = False
    Exit Sub
End If
If delayRecon > 0 Then delayRecon = delayRecon - 1
Stat Chr(3) & "14waiting for reconnect, remain " & delayRecon & "s"
End Sub

Private Sub tmrSession_Timer()
Dim ax As Long, ay As Long
    'Session = Session + 1
    SSec = SSec + 1
    If (SSec > 59) Then
         SMin = SMin + 1
         SSec = 0
         If (SMin > 59) Then
             SHour = SHour + 1
             SMin = 0
         End If
     End If
    Dim i As Integer, F As Boolean
    rndSend = rndSend + 16
    frmMain.sbMain.Panels(2).Text = "Time: " & MakeTime & " , Exp: " & Format(sExp, "##,##")
If Constate = 4 Then
    Update_Curmon
    If delay_send_Move > 0 Then delay_send_Move = delay_send_Move - 1
    If delay_use_skill > 0 Then delay_use_skill = delay_use_skill - 1
    If delay_Mon_Atk > 0 And timeout < 5 Then delay_Mon_Atk = delay_Mon_Atk - 1
    If delay_invite_party > 0 Then delay_invite_party = delay_invite_party - 1
    If delay_change_equip > 0 Then delay_change_equip = delay_change_equip - 1
    If delay_feed_pet > 0 Then delay_feed_pet = delay_feed_pet - 1
    If delay_select_npc > 0 Then delay_select_npc = delay_select_npc - 1
    If delay_use_hp > 0 Then delay_use_hp = delay_use_hp - 1
    If delay_use_mp > 0 Then delay_use_mp = delay_use_mp - 1
    If delay_refresh_inv > -1 Then delay_refresh_inv = delay_refresh_inv - 1
    
    'refresh inv
    If delay_refresh_inv = 0 Then Send_Refresh_INV
        
    'ขุดแร่ เก็บเกี่ยว
    If opt.farm.auto And delay_select_npc = 0 And N_Atk.id <> "" And Active And Not Sitting Then
        If N_Atk.hp < 5 Then Pack_801 Chr(&H0) & Chr(&H0) & N_Atk.id
        N_Atk.id = ""
        SelectFarm
    End If
    
    'ใช้ item hp
    If opt.hp.auto And Active And delay_use_hp = 0 And LoginChar(Char).MaxHP > 0 Then
        If (opt.hp.hp / 100) >= (LoginChar(Char).hp / LoginChar(Char).MaxHP) Then SendUseHP
    End If
    
    'ใช้ item mp
    If opt.mp.auto And Active And delay_use_mp = 0 And LoginChar(Char).MaxMp > 0 Then
        If opt.mp.mp / 100 >= LoginChar(Char).mp / LoginChar(Char).MaxMp Then SendUseMP
    End If
    
    If frmMap.cTool.value = 1 And Active Then ' And (.hp / .MaxHp) * 100 > opt.heal.sit_hp.max Then
        If delay_go_map > -1 Then delay_go_map = delay_go_map - 1
        If Map <> opt.path.lock.m And delay_go_map = 0 Then
            Gomap opt.path.lock.m
            Stat "Change map " & Map & " to " & opt.path.lock.m
            delay_go_map = 2
        ElseIf C_Atk.id = "" And Not tmrWalk.Enabled And Map = opt.path.lock.m And delay_go_map < 0 And (LoginChar(Char).hp / LoginChar(Char).MaxHP) * 100 >= opt.heal.sit_hp.Min Then
            If (cur.X < (frmMap.sPos.Left / MapScale) * 3) Or cur.X > ((frmMap.sPos.Left + frmMap.sPos.Width) / MapScale) * 3 Or _
            (cur.Y < (frmMap.sPos.Top / MapScale) * 3) Or cur.Y > ((frmMap.sPos.Top + frmMap.sPos.Height) / MapScale) * 3 Then
                Start_Walk ((frmMap.sPos.Left + (frmMap.sPos.Width / 2)) / MapScale) * 3, ((frmMap.sPos.Top + (frmMap.sPos.Height / 2)) / MapScale) * 3
            End If
        End If
    End If
    If delay_feed_pet = 0 And Active And opt.pet And Not frmMain.tbPet.Visible Then ChangPet 1
    
    
    If pet.MaxEN > 0 And pet.id <> "" And Active Then
        If delay_feed_pet = 0 And (pet.EN / pet.MaxEN) < 0.9 And opt.pet Then
            SendFeedPet_EN
            delay_feed_pet = 5
        ElseIf (pet.EN / pet.MaxEN) <= 0.5 And opt.pet And Not opt.pet Then
            ChangPet 2
        ElseIf (pet.EN / pet.MaxEN) <= 0.3 Then
            ChangPet 2
            opt.pet = False
        End If
    End If
    
    If pet.MaxHP > 0 And pet.id <> "" And Active Then
        If delay_feed_pet = 0 And (pet.hp / pet.MaxHP) < 0.9 And opt.pet Then
            SendFeedPet_HP
            delay_feed_pet = 5
        ElseIf (pet.hp / pet.MaxHP) <= 0.5 And opt.pet And Not opt.pet Then
            ChangPet 2
        ElseIf (pet.hp / pet.MaxHP) <= 0.3 Then
            ChangPet 2
            opt.pet = False
        End If
    End If
        
    If delay_Mon_Atk = 0 And C_Atk.id <> "" Then
        Pack_801 Chr(&H0) & Chr(&H0) & C_Atk.id
        For i = 0 To UBound(Monster) - 1
            If Trim(Monster(i).id) = "" Then F = True
            If F Then Monster(i) = Monster(i + 1)
        Next
        If F Then ReDim Preserve Monster(UBound(Monster) - 1)
    End If
    Update_Monster
    Update_People
    Update_Item
    If (index_update Mod 2) Then check_party
    index_update = index_update + 1
    If index_update >= 5 Then index_update = 0
    timeout = timeout + 1
    If (timeout > opt.basic.timeout) Then
        Stat Chr(&H3) & "0,4 TimeOut!, waiting for reconnect "
        timeout = 0
        reconnect
        ShowBalloonTip "TimeOut!, waiting for reconnect", NIIF_ERROR
    End If
End If
End Sub

Private Sub tmrSession2_Timer()
Dim i As Integer
If Constate = 4 And Active Then
    If Not Sitting Then
        If UBound(Item) > 0 Then
            For i = 0 To UBound(Item) - 1
                If Distant(cur, Item(i).Pos) <= 2 And Item(i).keep <= 5 Then
                    SendPacket Chr(&H90) & Chr(&H1) & Chr(&H0) & LoginChar(Char).id & Item(i).id, 1
                    Exit For
                End If
            Next
        End If
        'If C_Atk.ID <> "" And (Distant(cur, C_Atk.Pos) <= 2 Or C_Atk.Atk_me) Then
        If C_Atk.id <> "" And (Distant(cur, C_Atk.Pos) <= 2 Or ((LoginChar(Char).Class = 2 Or LoginChar(Char).Class = 3) And Distant(cur, C_Atk.Pos) < 15)) Then
            If frmBot.tmrWalk.Enabled Then
                If Walking_stop.X <> cur.X And Walking_stop.Y <> cur.Y Then Walking_stop = cur
                Stop_Walk
            End If
            SendAttack
        ElseIf C_Atk.id <> "" Then
            Start_Walk C_Atk.Pos.X, C_Atk.Pos.Y
            delay_Mon_Atk = 15
        ElseIf N_Atk.id <> "" And Distant(cur, N_Atk.Pos) <= 3 Then
            If frmBot.tmrWalk.Enabled Then
                If Walking_stop.X <> cur.X And Walking_stop.Y <> cur.Y Then Walking_stop = cur
                Stop_Walk
            End If
            SendFarm
        ElseIf N_Atk.id <> "" Then
            Start_Walk N_Atk.Pos.X, N_Atk.Pos.Y
        End If
    End If
    With LoginChar(Char)
        If opt.heal.recon.auto And (.hp / .MaxHP) * 100 <= opt.heal.recon.hp And C_Atk.id <> "" Then
            Connecting False
            reconnect
        End If
        If opt.heal.hp1.auto Then
            If ((.hp / .MaxHP) * 100 <= opt.heal.hp1.hp And (.mp / .MaxMp) * 100 >= opt.heal.hp1.mp) Then
                useSkill .id, LngToChr(CLng(opt.heal.hp1.skill)), 0
            End If
        End If
        If opt.heal.sit_hp.auto Then
            If ((.hp / .MaxHP) * 100 <= opt.heal.sit_hp.Min And C_Atk.id = "" And Not Sitting) Then
                'Sitting = True
                SitDown
            End If
        End If
        If opt.heal.sit_mp.auto Then
            If ((.mp / .MaxMp) * 100 <= opt.heal.sit_mp.Min And C_Atk.id = "" And Not Sitting) Then
                'Sitting = True
                SitDown
            End If
        End If
        If opt.heal.sit_hp.auto And opt.heal.sit_mp.auto Then
            If (((.hp / .MaxHP) * 100 >= opt.heal.sit_hp.Max And (.mp / .MaxMp) * 100 >= opt.heal.sit_mp.Max) Or C_Atk.id <> "") And Sitting Then
                'Sitting = False
                StandUp
            End If
        ElseIf opt.heal.sit_hp.auto Then
            If ((.hp / .MaxHP) * 100 >= opt.heal.sit_hp.Max Or C_Atk.id <> "") And Sitting Then
                'Sitting = False
                StandUp
            End If
        ElseIf opt.heal.sit_mp.auto Then
            If ((.mp / .MaxMp) * 100 >= opt.heal.sit_mp.Max Or C_Atk.id <> "") And Sitting Then
                'Sitting = False
                StandUp
            End If
        End If
        If .hp <= 0 Then ShowBalloonTip "You're Dead!", NIIF_WARNING: GotoTown
    End With
End If
End Sub

Private Sub tmrWalk_Timer()
If Constate <> 4 Or Sitting Then Exit Sub
Dim tx As Long, ty As Long, sx As String, sy As String, ds As Integer, Pt As Coord
cur.X = Walking_start.X + ((Walking_stop.X - Walking_start.X) * (Walking / Walking_dist))
cur.Y = Walking_start.Y + ((Walking_stop.Y - Walking_start.Y) * (Walking / Walking_dist))
Plot_Dot cur
If Distant(cur, Walking_stop) <= 2 Then
    Walking_stop = cur
    Stop_Walk
ElseIf C_Atk.id <> "" And ((LoginChar(Char).Class = 2 Or LoginChar(Char).Class = 3) And Distant(cur, Walking_stop) < 15) Then
    Walking_stop = cur
    Stop_Walk
ElseIf Distant(cur, Walking_stop) <= 10 Then
    tx = Walking_stop.X
    ty = Walking_stop.Y
    sx = ReverseByte(LngToChr(PtToLng(tx)))
    sy = ReverseByte(LngToChr(PtToLng(ty)))
    Send_Walk sx & sy & tmpMap
ElseIf Walking Mod 10 = 0 And Distant(cur, Walking_stop) <= 100 Then
    tx = Walking_start.X + ((Walking_stop.X - Walking_start.X) * ((Walking + 10) / Walking_dist))
    ty = Walking_start.Y + ((Walking_stop.Y - Walking_start.Y) * ((Walking + 10) / Walking_dist))
    sx = ReverseByte(LngToChr(PtToLng(tx)))
    sy = ReverseByte(LngToChr(PtToLng(ty)))
    Send_Walk sx & sy & tmpMap
    tx = Walking_start.X + ((Walking_stop.X - Walking_start.X) * ((Walking) / Walking_dist))
    ty = Walking_start.Y + ((Walking_stop.Y - Walking_start.Y) * ((Walking) / Walking_dist))
    sx = ReverseByte(LngToChr(PtToLng(tx)))
    sy = ReverseByte(LngToChr(PtToLng(ty)))
    Pet_Walk sx & sy & tmpMap
ElseIf Distant(cur, Walking_stop) > 100 Then
    If Walking Mod 10 = 0 Then
        tx = Walking_start.X + ((Walking_stop.X - Walking_start.X) * ((Walking + 100) / Walking_dist))
        ty = Walking_start.Y + ((Walking_stop.Y - Walking_start.Y) * ((Walking + 100) / Walking_dist))
    sx = ReverseByte(LngToChr(PtToLng(tx)))
    sy = ReverseByte(LngToChr(PtToLng(ty)))
        Send_Walk sx & sy & tmpMap
        Walking = Walking + 99
        Pet_Walk sx & sy & tmpMap
    End If
End If
Walking = Walking + 1
End Sub

Private Sub Winsock_CloseSck()
Stat "4Disconnected!"
Clear_Array
Connecting False
reconnect
End Sub

Private Sub Winsock_Connect()
Dim Sepe As String
RecvData = ""
Clear_Array
Stat Winsock.RemoteHostIP & ":" & Winsock.RemotePort & " 3Connected!"
Dim Info As String
'SendPacket Chr(&H83) & Chr(&H0) & Chr(&H0) & Chr(&H3) & Chr(&H87) & Chr(&H0) & _
                         User & Chr(&H0) & Pass & Chr(&H0), 1
SendPacket HextoChr("81D533EBE460D42615D5A918245E317DC1763F33882512142C3EDAD1607E"), 1
Chat ChrtoHex(Chr(&H83) & Chr(&H0) & Chr(&H0) & Chr(&H3) & Chr(&H87) & Chr(&H0) & _
                         User & Chr(&H0) & Pass & Chr(&H0))
serv_index = 0
Check_AP
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
'On Error GoTo errie
Dim tData As String, len2 As Integer
Winsock.GetData tData
print_packet tData, " ---> "
timeout = 0
start:
RecvData = RecvData & tData
    
If Len(RecvData) >= 2 Then ParseData
Exit Sub
errie:
RecvData = ""
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Stat "4Error! " & Description
    Winsock.CloseSck
    Clear_Array
    Connecting False
    reconnect
Err.Clear
End Sub

Private Sub tmrReply_Timer()
    If tmrRep + 600 > &HFFFFFF Then
        tmrRep = 600
    Else
        tmrRep = tmrRep + 600
    End If
    SendPacket Chr(&HA6) & LngToChr(tmrRep), 1
    tmrReply.Enabled = False
End Sub
