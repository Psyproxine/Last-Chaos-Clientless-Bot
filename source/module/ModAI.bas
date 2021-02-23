Attribute VB_Name = "ModAI"
Option Explicit
Public Sub SelectMonster()
Dim i As Integer, Cur_Distance As Integer, Sel_distance As Integer, Bestmonster As Integer, found As Boolean
Dim CurMonsterName As String, tmp() As String, z As Integer, x As Integer
    If UBound(Monster) < 1 Or Constate <> 4 Or Not Active Or C_Atk.id <> "" Or Sitting Or Not opt.atk.auto Then Exit Sub
    With LoginChar(Char)
        If (.hp / .MaxHP < 0.5 And Not Sitting And C_Atk.id = "") Then
            'Sitting = True
            SitDown
            Exit Sub
        End If
    End With
    Sel_distance = 200
    For i = 0 To UBound(Monster) - 1
        Monster(i).Pos = convert_coord(Monster(i).Coord)
        If Trim(Monster(i).id) = "" Or Trim(Monster(i).name) = "" Then GoTo endloop
        If Monster(i).Atk_me Then Bestmonster = i: GoTo tokill
        If Monster(i).NoAttack Then GoTo endloop
        If Monster(i).IsPet Then GoTo endloop
        If Monster(i).IsAttack Then GoTo endloop
        If frmMap.cTool.Value = 1 Then
            If Monster(i).Pos.x < (frmMap.sPos.Left / MapScale) * 3 Or Monster(i).Pos.x > ((frmMap.sPos.Left + frmMap.sPos.Width) / MapScale) * 3 Then GoTo endloop
            If Monster(i).Pos.y < (frmMap.sPos.Top / MapScale) * 3 Or Monster(i).Pos.y > ((frmMap.sPos.Top + frmMap.sPos.Height) / MapScale) * 3 Then GoTo endloop
        End If
        Cur_Distance = Distant(Monster(i).Pos, cur)
        If Cur_Distance < Sel_distance And Cur_Distance < 200 Then
            Sel_distance = Cur_Distance
            Bestmonster = i
            found = True
        End If
endloop:
    Next
If (Not Sitting) And found Then
tokill:
    CurMonsterName = Return_Monster_Name(Monster(Bestmonster).Type)
    Stat "10Select [3" + CurMonsterName + "10] as a Target, Locking..."
    delay_Mon_Atk = 15
    N_Atk.id = ""
    C_Atk = Monster(Bestmonster)
    C_Atk.hp = C_Atk.MaxHP
    Start_Walk C_Atk.Pos.x, C_Atk.Pos.y
End If
Exit Sub
errie:
Err.Clear
End Sub

Public Sub SelectFarm()
Dim i As Integer, Cur_Distance As Integer, Sel_distance As Integer, Bestmonster As Integer, found As Boolean
Dim CurMonsterName As String, tmp() As String, z As Integer, x As Integer
    If UBound(NPC) < 1 Or Constate <> 4 Or Not Active Or C_Atk.id <> "" Or N_Atk.id <> "" Or Sitting Or Not opt.farm.auto Then Exit Sub
    Sel_distance = 200
    Bestmonster = 0
    For i = 0 To UBound(NPC) - 1
        NPC(i).Pos = convert_coord(NPC(i).Coord)
        If Trim(NPC(i).id) = "" Or Trim(NPC(i).name) = "" Or NPC(i).hp <= 1 Then GoTo endloop
        If Not NPC(i).Agriculture Then GoTo endloop
        'If NPC(i).HP < 100 Then GoTo endloop
        If NPC(i).hp = 300 Then Bestmonster = i: GoTo tokill
        If frmMap.cTool.Value = 1 Then
            If NPC(i).Pos.x < (frmMap.sPos.Left / MapScale) * 3 Or NPC(i).Pos.x > ((frmMap.sPos.Left + frmMap.sPos.Width) / MapScale) * 3 Then GoTo endloop
            If NPC(i).Pos.y < (frmMap.sPos.Top / MapScale) * 3 Or NPC(i).Pos.y > ((frmMap.sPos.Top + frmMap.sPos.Height) / MapScale) * 3 Then GoTo endloop
        End If
        Cur_Distance = Distant(NPC(i).Pos, cur)
        If (NPC(i).MaxHP <= NPC(Bestmonster).MaxHP) Then
            Sel_distance = Cur_Distance
            Bestmonster = i
            found = True
        End If
endloop:
    Next
If (Not Sitting) And found Then
tokill:
    CurMonsterName = NPC(Bestmonster).name
    Stat "10Select [3" + CurMonsterName + "10] as a Target, Locking..."
    delay_Mon_Atk = 15
    delay_select_npc = 15
    C_Atk.id = ""
    N_Atk = NPC(Bestmonster)
    N_Atk.hp = N_Atk.MaxHP
    Start_Walk N_Atk.Pos.x, N_Atk.Pos.y
End If
Exit Sub
errie:
Err.Clear
End Sub

Public Sub Item2Pet()
'90 00 00 0C 00 00 37 F8 A6 00 00 00 00
End Sub
