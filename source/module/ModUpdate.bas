Attribute VB_Name = "ModUpdate"
Option Explicit
Public Sub Update_People()
Dim i As Integer, x As Integer, d As Integer
If cur.x = 0 Then cur = LoginChar(Char).Pos
For x = 0 To UBound(People) - 1
    If x + 1 > frmPeople.lvPlayer.ListItems.Count Then
        i = frmPeople.lvPlayer.ListItems.Add.Index
        frmPeople.lvPlayer.ListItems(i).Text = ChrtoHex(People(x).id)
        frmPeople.lvPlayer.ListItems(i).SubItems(1) = People(x).name
        frmPeople.lvPlayer.ListItems(i).SubItems(2) = People(x).Guild
        frmPeople.lvPlayer.ListItems(i).SubItems(3) = Return_Class(People(x).Class)
        Load frmMap.blPeo(i): Load frmMap.lbPeo(i)
        People(x).Pos = convert_coord(People(x).Coord)
        d = Distant(cur, People(x).Pos)
        If d <= 16 Then tmpMap = Mid(People(x).Coord, 9, 8)
        frmPeople.lvPlayer.ListItems(i).SubItems(4) = "X:" & People(x).Pos.x & "  Y:" & People(x).Pos.y & "  [" & d & "]": People_Move i
    Else
        With frmPeople.lvPlayer.ListItems(x + 1)
            If .Text <> ChrtoHex(People(x).id) Then .Text = ChrtoHex(People(x).id)
            If .SubItems(1) <> People(x).name Then .SubItems(1) = People(x).name
            If .SubItems(2) <> People(x).Guild Then .SubItems(2) = People(x).Guild
            If .SubItems(3) <> Return_Class(People(x).Class) Then .SubItems(3) = Return_Class(People(x).Class)
            People(x).Pos = convert_coord(People(x).Coord)
            d = Distant(cur, People(x).Pos)
            If d <= 16 Then tmpMap = Mid(People(x).Coord, 9, 8)
            If .SubItems(4) <> "X:" & People(x).Pos.x & "  " & "Y: " & People(x).Pos.y & "  [" & d & "]" Then .SubItems(4) = "X:" & People(x).Pos.x & "  " & "Y: " & People(x).Pos.y & "  [" & d & "]": People_Move x + 1
        End With
    End If
Next
If UBound(People) < frmMap.blPeo.UBound Then
    For i = UBound(People) + 1 To frmMap.blPeo.UBound
        If i <> 0 Then
            Unload frmMap.blPeo(i)
            Unload frmMap.lbPeo(i)
        Else
            frmMap.blPeo(i).Visible = False
            frmMap.lbPeo(i).Visible = False
        End If
    Next
End If
If UBound(People) < frmPeople.lvPlayer.ListItems.Count Then
    For i = frmPeople.lvPlayer.ListItems.Count To UBound(People) + 1 Step -1
        If i <> 0 Then
            frmPeople.lvPlayer.ListItems.Remove i
        End If
    Next
End If
End Sub

Public Sub Update_Party()
Dim i As Integer, found As Boolean, x As Integer, d As Integer
If cur.x = 0 Then cur = LoginChar(Char).Pos
For x = 0 To UBound(party) - 1
    If x + 1 > frmParty.lvParty.ListItems.Count Then
        i = frmParty.lvParty.ListItems.Add.Index
        frmParty.lvParty.ListItems(i).Text = ChrtoHex(party(x).id)
        frmParty.lvParty.ListItems(i).SubItems(1) = party(x).name
        frmParty.lvParty.ListItems(i).SubItems(2) = party(x).level
        frmParty.lvParty.ListItems(i).SubItems(3) = party(x).hp & " / " & party(x).MaxHP
        frmParty.lvParty.ListItems(i).SubItems(4) = party(x).mp & " / " & party(x).MaxMp
        frmParty.lvParty.ListItems(i).SubItems(5) = Return_Class(party(x).Class)
        d = Distant(cur, party(x).Pos)
        If d <= 16 Then tmpMap = Mid(party(x).Coord, 9, 8)
        frmParty.lvParty.ListItems(i).SubItems(6) = "X:" & party(x).Pos.x & "  Y:" & party(x).Pos.y & "  [" & d & "]"
    Else
        With frmParty.lvParty.ListItems(x + 1)
            If .Text <> ChrtoHex(party(x).id) Then .Text = ChrtoHex(party(x).id)
            If .SubItems(1) <> party(x).name Then .SubItems(1) = party(x).name
            If .SubItems(2) <> party(x).level Then .SubItems(2) = party(x).level
            If .SubItems(3) <> party(x).hp & " / " & party(x).MaxHP Then .SubItems(3) = party(x).hp & " / " & party(x).MaxHP
            If .SubItems(4) <> party(x).mp & " / " & party(x).MaxMp Then .SubItems(4) = party(x).mp & " / " & party(x).MaxMp
            If .SubItems(5) <> Return_Class(party(x).Class) Then .SubItems(5) = Return_Class(party(x).Class)
            party(x).Pos = convert_coord(party(x).Coord)
            d = Distant(cur, party(x).Pos)
            If d <= 16 Then tmpMap = Mid(party(x).Coord, 9, 8)
            If .SubItems(6) <> "X:" & party(x).Pos.x & "  " & "Y: " & party(x).Pos.y & "  [" & d & "]" Then .SubItems(6) = "X:" & party(x).Pos.x & "  " & "Y: " & party(x).Pos.y & "  [" & d & "]"
        End With
    End If
Next
If UBound(party) < frmParty.lvParty.ListItems.Count Then
    For i = frmParty.lvParty.ListItems.Count To UBound(party) + 1 Step -1
        If i <> 0 Then
            frmParty.lvParty.ListItems.Remove i
        End If
    Next
End If
End Sub

Public Sub People_Move(x As Integer)
    frmMap.blPeo.Item(x).Visible = True
    frmMap.lbPeo.Item(x).Visible = IIf(frmMap.cPeople.Value = 1, True, False)
    frmMap.lbPeo.Item(x).Alignment = 2
    frmMap.lbPeo.Item(x).Caption = ""
    frmMap.lbPeo.Item(x).Move ((People(x - 1).Pos.x) * MapScale / 3) - 2, ((People(x - 1).Pos.y) * MapScale / 3) - 14
    frmMap.blPeo.Item(x).Move ((People(x - 1).Pos.x) * MapScale / 3) - 2, ((People(x - 1).Pos.y) * MapScale / 3) - 2
    frmMap.lbPeo.Item(x).Caption = People(x - 1).name
    'FrmField.PicMain.Refresh
End Sub

Public Sub Update_Monster()
Dim i As Integer, x As Integer, d As Integer
If cur.x = 0 Then cur = LoginChar(Char).Pos
For x = 0 To UBound(Monster) - 1
    If x + 1 > frmMon.lvMon.ListItems.Count Then
        i = frmMon.lvMon.ListItems.Add.Index
        With frmMon.lvMon.ListItems(i)
            .Text = ChrtoHex(Monster(x).id)
            .SubItems(1) = Monster(x).name
            .SubItems(2) = Monster(x).hp & "/" & Monster(x).MaxHP
            Monster(x).Pos = convert_coord(Monster(x).Coord)
            d = Distant(cur, Monster(x).Pos)
            If d <= 16 Then tmpMap = Mid(Monster(x).Coord, 9, 8)
            Load frmMap.blmons(i): Load frmMap.lbMons(i)
            .SubItems(3) = "X:" & Monster(x).Pos.x & "  " & "Y: " & Monster(x).Pos.y & "  [" & d & "]"
            Monster_Move i
        End With
    Else
        With frmMon.lvMon.ListItems(x + 1)
            If .Text <> ChrtoHex(Monster(x).id) Then .Text = ChrtoHex(Monster(x).id)
            If .SubItems(1) <> Monster(x).name Then .SubItems(1) = Monster(x).name
            If .SubItems(2) <> Monster(x).hp & "/" & Monster(x).MaxHP Then .SubItems(2) = Monster(x).hp & "/" & Monster(x).MaxHP
            Monster(x).Pos = convert_coord(Monster(x).Coord)
            d = Distant(cur, Monster(x).Pos)
            If d <= 16 Then tmpMap = Mid(Monster(x).Coord, 9, 8)
            If .SubItems(3) <> "X:" & Monster(x).Pos.x & "  " & "Y: " & Monster(x).Pos.y & "  [" & d & "]" Then .SubItems(3) = "X:" & Monster(x).Pos.x & "  " & "Y: " & Monster(x).Pos.y & "  [" & d & "]"
            Monster_Move x + 1
        End With
    End If
    If Monster(x).id = C_Atk.id Then
        frmMap.blmons(x + 1).BackColor = &HFFFFFF
        frmMap.lbMons(x + 1).BackColor = &HFFFFFF
    Else
        frmMap.blmons(x + 1).BackColor = &HFF
        frmMap.lbMons(x + 1).BackColor = &HFF
    End If
Next
If UBound(Monster) < frmMap.blmons.UBound Then
    For i = UBound(Monster) + 1 To frmMap.blmons.UBound
        If i <> 0 Then
            Unload frmMap.blmons(i)
            Unload frmMap.lbMons(i)
        Else
            frmMap.blmons(i).Visible = False
            frmMap.lbMons(i).Visible = False
        End If
    Next
End If
If UBound(Monster) < frmMon.lvMon.ListItems.Count Then
    For i = frmMon.lvMon.ListItems.Count To UBound(Monster) + 1 Step -1
        If i <> 0 Then
            frmMon.lvMon.ListItems.Remove i
        End If
    Next
End If
End Sub


Public Sub Update_Apet()
Dim i As Integer, x As Integer, d As Integer
For x = 0 To UBound(Apet) - 1
    If x + 1 > frmPet.lvPet.ListItems.Count Then
        i = frmPet.lvPet.ListItems.Add.Index
        With frmPet.lvPet.ListItems(i)
            .Text = ChrtoHex(Apet(x).id)
            .SubItems(1) = Apet(x).name & " - " & Apet(x).Type
            .SubItems(2) = Apet(x).hp & "/" & Apet(x).MaxHP
            Apet(x).Pos = convert_coord(Apet(x).Coord)
            d = Distant(cur, Apet(x).Pos)
            .SubItems(3) = "X:" & Apet(x).Pos.x & "  " & "Y: " & Apet(x).Pos.y & "  [" & d & "]"
        End With
    Else
        With frmPet.lvPet.ListItems(x + 1)
            If .Text <> ChrtoHex(Apet(x).id) Then .Text = ChrtoHex(Apet(x).id)
            If .SubItems(1) <> Apet(x).name & " - " & Apet(x).Type Then .SubItems(1) = Apet(x).name & " - " & Apet(x).Type
            If .SubItems(2) <> Apet(x).hp & "/" & Apet(x).MaxHP Then .SubItems(2) = Apet(x).hp & "/" & Apet(x).MaxHP
            Apet(x).Pos = convert_coord(Apet(x).Coord)
            
            d = Distant(cur, Apet(x).Pos)
            If .SubItems(3) <> "X:" & Apet(x).Pos.x & "  " & "Y: " & Apet(x).Pos.y & "  [" & d & "]" Then .SubItems(3) = "X:" & Apet(x).Pos.x & "  " & "Y: " & Apet(x).Pos.y & "  [" & d & "]"
        End With
    End If
Next
If UBound(Apet) < frmPet.lvPet.ListItems.Count Then
    For i = frmPet.lvPet.ListItems.Count To UBound(Apet) + 1 Step -1
        If i <> 0 Then
            frmPet.lvPet.ListItems.Remove i
        End If
    Next
End If
End Sub


Public Sub Monster_Move(x As Integer)
    frmMap.blmons.Item(x).Visible = True
    frmMap.lbMons.Item(x).Visible = IIf(frmMap.cMonster.Value = 1, True, False)
    frmMap.lbMons.Item(x).Alignment = 2
    frmMap.lbMons.Item(x).Caption = ""
    frmMap.lbMons.Item(x).Move ((Monster(x - 1).Pos.x) * MapScale / 3) - 2, ((Monster(x - 1).Pos.y) * MapScale / 3) - 14
    frmMap.blmons.Item(x).Move ((Monster(x - 1).Pos.x) * MapScale / 3) - 2, ((Monster(x - 1).Pos.y) * MapScale / 3) - 2
    frmMap.lbMons.Item(x).Caption = Monster(x - 1).name
    'FrmField.PicMain.Refresh
End Sub

Public Sub Update_Item()
Dim i As Integer, x As Integer, d As Integer
If cur.x = 0 Then cur = LoginChar(Char).Pos
For x = 0 To UBound(Item) - 1
    If x + 1 > frmItem.lvItem.ListItems.Count Then
        i = frmItem.lvItem.ListItems.Add.Index
        frmItem.lvItem.ListItems(i).Text = ChrtoHex(Item(x).id)
        frmItem.lvItem.ListItems(i).SubItems(1) = Item(x).name
        frmItem.lvItem.ListItems(i).SubItems(2) = Item(x).Amount
        'Load frmMap.blPeo(i): Load frmMap.lbPeo(i)
        Item(x).Pos = convert_coord(Item(x).Coord)
        d = Distant(cur, Item(x).Pos)
        If d <= 16 Then tmpMap = Mid(Item(x).Coord, 9, 8)
        frmItem.lvItem.ListItems(i).SubItems(3) = "X:" & Item(x).Pos.x & "  Y:" & Item(x).Pos.y & "  [" & d & "]"
    Else
        With frmItem.lvItem.ListItems(x + 1)
            If .Text <> ChrtoHex(Item(x).id) Then .Text = ChrtoHex(Item(x).id)
            If .SubItems(1) <> Item(x).name Then .SubItems(1) = Item(x).name
            If .SubItems(2) <> Item(x).Amount Then .SubItems(2) = Item(x).Amount
            Item(x).Pos = convert_coord(Item(x).Coord)
            d = Distant(cur, Item(x).Pos)
            If d <= 16 Then tmpMap = Mid(Item(x).Coord, 9, 8)
            If .SubItems(3) <> "X:" & Item(x).Pos.x & "  " & "Y: " & Item(x).Pos.y & "  [" & d & "]" Then .SubItems(3) = "X:" & Item(x).Pos.x & "  " & "Y: " & Item(x).Pos.y & "  [" & d & "]"
        End With
    End If
Next
'If UBound(Item) < frmMap.blPeo.UBound Then
'    For i = UBound(Item) + 1 To frmMap.blPeo.UBound
'        If i <> 0 Then
'            Unload frmMap.blPeo(i)
'            Unload frmMap.lbPeo(i)
'        Else
'            frmMap.blPeo(i).Visible = False
'            frmMap.lbPeo(i).Visible = False
'        End If
'    Next
'End If
If UBound(Item) + 1 <= frmItem.lvItem.ListItems.Count Then
    For i = frmItem.lvItem.ListItems.Count To UBound(Item) + 1 Step -1
        If i <> 0 Then
            frmItem.lvItem.ListItems.Remove i
        End If
    Next
End If
End Sub

Public Sub Update_npc()
Dim i As Integer, found As Boolean, x As Integer, d As Integer
If cur.x = 0 Then cur = LoginChar(Char).Pos


For x = 0 To UBound(NPC) - 1
    If x + 1 > frmNPC.lvNPC.ListItems.Count Then
        i = frmNPC.lvNPC.ListItems.Add.Index
        frmNPC.lvNPC.ListItems(i).Text = ChrtoHex(NPC(x).id)
        frmNPC.lvNPC.ListItems(i).SubItems(1) = NPC(x).name
        frmNPC.lvNPC.ListItems(i).SubItems(2) = NPC(x).hp & "/" & NPC(x).MaxHP
        frmNPC.lvNPC.ListItems(i).SubItems(3) = IIf(NPC(x).Agriculture, "X", "")
        NPC(x).Pos = convert_coord(NPC(x).Coord)
        d = Distant(cur, NPC(x).Pos)
        If d <= 16 Then tmpMap = Mid(NPC(x).Coord, 9, 8)
        frmNPC.lvNPC.ListItems(i).SubItems(4) = "X:" & NPC(x).Pos.x & "  " & "Y: " & NPC(x).Pos.y & "  [" & d & "]"
        NPC_Move i
    Else
        With frmNPC.lvNPC.ListItems(x + 1)
            If .Text <> ChrtoHex(NPC(x).id) Then .Text = ChrtoHex(NPC(x).id)
            If .SubItems(1) <> NPC(x).name Then .SubItems(1) = NPC(x).name
            If .SubItems(2) <> NPC(x).hp & "/" & NPC(x).MaxHP Then .SubItems(2) = NPC(x).hp & "/" & NPC(x).MaxHP
            If .SubItems(3) <> IIf(NPC(x).Agriculture, "X", "") Then .SubItems(3) = IIf(NPC(x).Agriculture, "X", "")
            NPC(x).Pos = convert_coord(NPC(x).Coord)
            d = Distant(cur, NPC(x).Pos)
            If .SubItems(4) <> "X:" & NPC(x).Pos.x & "  " & "Y: " & NPC(x).Pos.y & "  [" & d & "]" Then .SubItems(4) = "X:" & NPC(x).Pos.x & "  " & "Y: " & NPC(x).Pos.y & "  [" & d & "]"
            NPC_Move x + 1
        End With
    End If
Next

If UBound(NPC) < frmMap.blNPC.UBound Then
    For i = UBound(NPC) + 1 To frmMap.blNPC.UBound
        If i <> 0 Then
            Unload frmMap.blNPC(i)
            Unload frmMap.lbNPC(i)
        Else
            frmMap.blNPC(i).Visible = False
            frmMap.lbNPC(i).Visible = False
        End If
    Next
End If
If UBound(NPC) < frmNPC.lvNPC.ListItems.Count Then
    For i = frmNPC.lvNPC.ListItems.Count To UBound(NPC) + 1 Step -1
        If i <> 0 Then
            frmNPC.lvNPC.ListItems.Remove i
        End If
    Next
End If

End Sub

Public Sub NPC_Move(x As Integer)
    If x >= frmMap.blNPC.Count Then
        Load frmMap.blNPC(x)
        Load frmMap.lbNPC(x)
        'Load frmMap.lbShop(x)
    End If
    frmMap.blNPC.Item(x).Visible = True
    frmMap.lbNPC.Item(x).Visible = IIf(frmMap.cNPC.Value = 1, True, False)
    frmMap.lbNPC.Item(x).Alignment = 2
    frmMap.lbNPC.Item(x).Caption = ""
    frmMap.lbNPC.Item(x).Move ((NPC(x - 1).Pos.x) * MapScale / 3) - 2, ((NPC(x - 1).Pos.y) * MapScale / 3) - 14
    frmMap.blNPC.Item(x).Move ((NPC(x - 1).Pos.x) * MapScale / 3) - 2, ((NPC(x - 1).Pos.y) * MapScale / 3) - 2
    frmMap.lbNPC.Item(x).Caption = NPC(x - 1).name
    'FrmField.PicMain.Refresh
End Sub


Public Sub Update_Inv()
Dim i As Integer, found As Boolean, x As Integer
For x = 0 To UBound(Inv) - 1
    found = False
    For i = 1 To frmInv.lvInv.ListItems.Count
        If frmInv.lvInv.ListItems(i).Text = ChrtoHex(Inv(x).id) Then
            frmInv.lvInv.ListItems(i).SubItems(1) = Inv(x).name
            frmInv.lvInv.ListItems(i).SubItems(2) = Inv(x).Amount
            frmInv.lvInv.ListItems(i).SubItems(3) = Inv(x).row & "," & Inv(x).col 'IIf(Inv(x).block = 255, Inv(x).block, "X")
            frmInv.lvInv.ListItems(i).SubItems(4) = IIf(Inv(x).State > 0, "Quest", Eq_Type(Inv(x).equip)) '" Durability" & Inv(x).Durability & " Lv" & Inv(x).Lv & " Refine" & Inv(x).Refine
           GoTo nextX
        End If
    Next
    i = frmInv.lvInv.ListItems.Add.Index
    frmInv.lvInv.ListItems(i).Text = ChrtoHex(Inv(x).id)
    frmInv.lvInv.ListItems(i).SubItems(1) = Inv(x).name
    frmInv.lvInv.ListItems(i).SubItems(2) = Inv(x).Amount
    frmInv.lvInv.ListItems(i).SubItems(3) = Inv(x).row & "," & Inv(x).col 'IIf(Inv(x).block = 255, Inv(x).block, "X")
    frmInv.lvInv.ListItems(i).SubItems(4) = IIf(Inv(x).State > 0, "Quest", Eq_Type(Inv(x).equip))    '" Durability" & Inv(x).Durability & " Lv" & Inv(x).Lv & " Refine" & Inv(x).Refine
nextX:
Next
End Sub

Public Sub Update_Skill()
Dim i As Integer, found As Boolean, x As Integer
For x = 0 To UBound(skill) - 1
    For i = 1 To frmSkill.lvSkill.ListItems.Count
        If frmSkill.lvSkill.ListItems(i).Text = ChrtoHex(skill(x).id) Then
            frmSkill.lvSkill.ListItems(i).SubItems(2) = skill(x).LV
            GoTo nextX
        End If
    Next
    i = frmSkill.lvSkill.ListItems.Add.Index
    frmSkill.lvSkill.ListItems(i).Text = ChrtoHex(skill(x).id)
    frmSkill.lvSkill.ListItems(i).SubItems(1) = skill(x).name
    frmSkill.lvSkill.ListItems(i).SubItems(2) = skill(x).LV
    frmSkill.lvSkill.ListItems(i).SubItems(3) = skill(x).Detail
nextX:
Next
End Sub

Public Sub Update_Stats()
On Error Resume Next
    Dim Percent As Double, tstr As String
    With LoginChar(Char)
        frmMain.txtChar.Text = "Char: " & LoginChar(Char).name
        frmInfo.lvInfo.ListItems(1).SubItems(1) = .name
        frmInfo.lvInfo.ListItems(2).SubItems(1) = .Clan
        frmInfo.lvInfo.ListItems(2).SubItems(3) = Return_Class(.Class)
        frmMain.txtLv.Text = "Lv: " & LoginChar(Char).level
        frmInfo.lvInfo.ListItems(3).SubItems(1) = .level
        frmInfo.lvInfo.ListItems(3).SubItems(3) = Format(.Kill, "##,##")
        Percent = Format(.hp / .MaxHP, "##.####")
        frmMain.pbCHP.Value = Percent * 100
        frmMain.pbCHP.Text = .hp & " / " & .MaxHP
        frmMain.pbCHP.ToolTipText = Percent * 100 & "%"
        frmInfo.lvInfo.ListItems(5).SubItems(1) = .hp & " / " & .MaxHP & " (" & Percent * 100 & "%)"
        Percent = Format(.mp / .MaxMp, "##.####")
        frmMain.pbCMP.Value = Percent * 100
        frmMain.pbCMP.Text = .mp & " / " & .MaxMp
        frmMain.pbCHP.ToolTipText = Percent * 100 & "%"
        frmInfo.lvInfo.ListItems(5).SubItems(3) = .mp & " / " & .MaxMp & " (" & Percent * 100 & "%)"
        Percent = Format(.EXP / .MaxEXP, "##.####")
        frmMain.pbCEXP.Value = Percent * 100
        frmMain.pbCEXP.Text = .EXP & " / " & .MaxEXP
        frmMain.pbCEXP.ToolTipText = Percent * 100 & "%"
        frmInfo.lvInfo.ListItems(6).SubItems(1) = .Spp
        frmInfo.lvInfo.ListItems(6).SubItems(3) = .EXP & " / " & .MaxEXP & " (" & Percent * 100 & "%)"
        frmInfo.lvInfo.ListItems(8).SubItems(1) = .str & " (" & .Str0 & ")"
        frmInfo.lvInfo.ListItems(8).SubItems(3) = .atk
        frmInfo.lvInfo.ListItems(9).SubItems(1) = .Agi & " (" & .Agi0 & ")"
        frmInfo.lvInfo.ListItems(9).SubItems(3) = .Mtk
        frmInfo.lvInfo.ListItems(10).SubItems(1) = .Int & " (" & .Int0 & ")"
        frmInfo.lvInfo.ListItems(10).SubItems(3) = .Def
        frmInfo.lvInfo.ListItems(11).SubItems(1) = .Vit & " (" & .Vit0 & ")"
        frmInfo.lvInfo.ListItems(11).SubItems(3) = .Mef
        frmInfo.lvInfo.ListItems(14).SubItems(1) = Format(.Money, "##,##")
        If ((.str + .Agi + .Int + .Vit) - (.Str0 + .Agi0 + .Int0 + .Vit0) > LimitLEVEL) Or .level > LimitLEVEL Then
            Connecting False
            Chat "คุณไม่สามารถเปิดบอทตัวละครนี้ได้ เนื่องจากมีเลเวลมากกว่า " & LimitLEVEL
        End If
    End With
End Sub

Public Sub Clear_Array()
Dim i As Integer
RecvData = ""
ReDim People(0)
ReDim LoginChar(0)
ReDim Monster(0)
ReDim Item(0)
ReDim Chan_List(0)
ReDim Inv(0)
ReDim party(0)
ReDim NPC(0)
ReDim skill(0)
ReDim Apet(0)
Constate = 0
Logining = False
'InActive True
frmMain.tbPet.Visible = False
frmMain.Form_Resize
pet.id = ""
frmPeople.lvPlayer.ListItems.Clear
frmParty.lvParty.ListItems.Clear
frmMon.lvMon.ListItems.Clear
frmItem.lvItem.ListItems.Clear
frmInv.lvInv.ListItems.Clear
frmClan.lvClan.ListItems.Clear
frmNPC.lvNPC.ListItems.Clear
frmSkill.lvSkill.ListItems.Clear
frmPet.lvPet.ListItems.Clear
For i = 1 To frmMap.blPeo.UBound
    Unload frmMap.blPeo(i): Unload frmMap.lbPeo(i)
Next
For i = 1 To frmMap.blmons.UBound
    Unload frmMap.blmons(i): Unload frmMap.lbMons(i)
Next
frmBot.tmrWalk.Enabled = False
frmMain.mngetfeeden.Enabled = False
frmMain.mngetfeeden.Enabled = False
frmMain.mnPet(1).Enabled = False
frmMain.mnPet(2).Enabled = False
C_Atk.id = ""
N_Atk.id = ""
delay_go_map = -1
delay_use_hp = 15
delay_use_mp = 15
End Sub

Public Sub InActive(t As Boolean)
Active = Not t
frmMain.btnActive.Caption = IIf(t, "InActive", "Active")
End Sub

Public Sub check_party()
If delay_invite_party > 0 Or Not opt.party.Invite.auto Then Exit Sub
Dim i As Integer, x As Integer, z As Integer, tmp() As String, F As Boolean
tmp = Split(opt.party.Invite.name, ";")
If UBound(tmp) < UBound(party) Then Exit Sub
For i = 0 To UBound(tmp)
    If Trim(tmp(i)) <> "" Then
        For x = 0 To UBound(People) - 1
            If People(x).name = Trim(tmp(i)) Then
                F = False
                For z = 0 To UBound(party) - 1
                    If party(z).id = People(x).id Then F = True: Exit For
                Next
                If Not F Then
                     InviteParty People(x).name, People(x).id, opt.party.Invite.State
                     delay_invite_party = 10
                End If
            End If
        Next
    End If
Next
End Sub

Public Sub Update_Curmon()
    If C_Atk.id <> "" Then
        frmMain.pbCMon.Text = C_Atk.hp & "/" & C_Atk.MaxHP
        If C_Atk.MaxHP > 0 Then frmMain.pbCMon.Value = (C_Atk.hp / C_Atk.MaxHP) * 100
        frmMain.txtTarget.Text = "Target: " & C_Atk.name & "  X:" & C_Atk.Pos.x & " Y:" & C_Atk.Pos.y & " [" & Distant(cur, C_Atk.Pos) & "]"
    ElseIf N_Atk.id = "" Then
        frmMain.pbCMon.Text = ""
        frmMain.pbCMon.Value = 0
        frmMain.txtTarget.Text = "Target: -none-"
    End If
    frmMain.txtWalk.Visible = IIf(frmBot.tmrWalk.Enabled, True, False)
End Sub

Public Sub Update_Curnpc()
    If N_Atk.id <> "" Then
        frmMain.pbCMon.Text = N_Atk.hp & "/" & N_Atk.MaxHP
        If N_Atk.MaxHP > 0 Then frmMain.pbCMon.Value = (N_Atk.hp / N_Atk.MaxHP) * 100
        frmMain.txtTarget.Text = "Target: " & N_Atk.name & "  X:" & N_Atk.Pos.x & " Y:" & N_Atk.Pos.y & " [" & Distant(cur, N_Atk.Pos) & "]"
    Else
        frmMain.pbCMon.Text = ""
        frmMain.pbCMon.Value = 0
        frmMain.txtTarget.Text = "Target: -none-"
    End If
    frmMain.txtWalk.Visible = IIf(frmBot.tmrWalk.Enabled, True, False)
End Sub

Public Sub Update_Pet()
    With frmMain
        If Not .tbPet.Visible Then .tbPet.Visible = True:        frmMain.Form_Resize
        .pPetHP.Text = pet.hp & "/" & pet.MaxHP
        If pet.MaxHP > 0 Then .pPetHP.Value = (pet.hp / pet.MaxHP) * 100
        .pPetEN.Text = pet.EN & "/" & pet.MaxEN
        If pet.MaxEN > 0 Then .pPetEN.Value = (pet.EN / pet.MaxEN) * 100
        .pPetEXP.Text = pet.EXP & "/" & pet.MaxEXP
        If pet.MaxEXP > 0 Then .pPetEXP.Value = (pet.EXP / pet.MaxEXP) * 100
        .pPetFL.Text = pet.FL & "/" & pet.MaxFL
        If pet.MaxFL > 0 Then .pPetFL.Value = (pet.FL / pet.MaxFL) * 100
        .txtPet.Text = "Pet:  Lv." & pet.LV & "  X:" & pet.Pos.x & ", Y:" & pet.Pos.y & " [" & Distant(cur, pet.Pos) & "] SP:" & pet.SP
    End With
End Sub
