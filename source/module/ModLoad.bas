Attribute VB_Name = "ModLoad"
Option Explicit
Option Compare Text
#If Win16 Then
        Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal FileName As String) As Integer
        Private Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal Default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal FileName As String) As Integer
#Else
        Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
        Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Public Sub wini(head As String, Key As String, Default As String, File As String)
WritePrivateProfileString head, Key, Default, App.path & "\" & File & ".ini"
End Sub

Public Function rini(head As String, Key As String, Default As String, File As String) As String
On Error GoTo z:
    Dim sRet As String
    sRet = String(255, Chr(0))
    rini = Trim(Left(sRet, GetPrivateProfileString(head, ByVal Key, Default, sRet, 255, App.path & "\" & File & ".ini")))
    If LenB(rini) = 0 Or Left$(rini, 6) = "Error " Then rini = Trim(Default)
    Exit Function
z:
    rini = Trim(Default)
End Function

Public Sub Load_Option()
Dim tmp As String, cmd() As String
ReDim cmd(0)
'  Attack
tmp = rini("attack", "auto", 0, "config\options")
opt.atk.auto = False
opt.farm.auto = False
If tmp = 1 Then
    opt.atk.auto = True
ElseIf tmp = 2 Then
    opt.farm.auto = True
End If

tmp = rini("attack", "autoskill", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.atk.skill.auto = True
    opt.atk.skill.skill = CInt("&h" & cmd(1))
    opt.atk.skill.hp = CInt(cmd(2))
    opt.atk.skill.mp = CInt(cmd(3))
Else
    opt.atk.skill.auto = False
End If

frmBot.tmrSession2.Enabled = False
frmBot.tmrSession2.Interval = rini("attack", "speed", 1200, "config\options")
frmBot.tmrSession2.Enabled = True

' Path
tmp = rini("path", "lock", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.path.lock.auto = True
    opt.path.lock.X = CInt(cmd(1))
    opt.path.lock.y = CInt(cmd(2))
    opt.path.lock.w = CInt(cmd(3))
    opt.path.lock.h = CInt(cmd(4))
    opt.path.lock.m = CInt("&H" & cmd(5))
    frmMap.sPos.Width = CInt(opt.path.lock.w * MapScale) \ 3
    frmMap.sPos.Height = CInt(opt.path.lock.h * MapScale) \ 3
    frmMap.sPos.Left = CInt(((opt.path.lock.X) - (opt.path.lock.w / 2)) * MapScale) \ 3
    frmMap.sPos.Top = CInt(((opt.path.lock.y) - (opt.path.lock.h / 2)) * MapScale) \ 3
    frmMap.cTool.value = 1
    frmMap.cTool_Click
Else
    opt.path.lock.auto = False
    frmMap.cTool.value = 0
    frmMap.cTool_Click
End If

'  Party
tmp = rini("party", "invite", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.party.Invite.auto = True
    opt.party.Invite.name = cmd(1)
    opt.party.Invite.State = Int(cmd(2))
Else
    opt.party.Invite.auto = False
    opt.party.Invite.name = ""
End If
tmp = rini("party", "response", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.party.resp.auto = True
    opt.party.resp.name = cmd(1)
Else
    opt.party.resp.auto = False
    opt.party.resp.name = ""
End If
tmp = rini("party", "follow", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.party.follow.auto = True
    opt.party.follow.name = cmd(1)
    opt.party.follow.Max = CInt(cmd(2))
    opt.party.follow.Min = CInt(cmd(3))
Else
    opt.party.follow.auto = False
    opt.party.follow.name = ""
End If
tmp = rini("party", "hp1", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.party.hp1.auto = True
    opt.party.hp1.name = cmd(1)
    opt.party.hp1.skill = CInt("&H" & cmd(2))
    opt.party.hp1.hp = cmd(3)
    opt.party.hp1.mp = cmd(4)
Else
    opt.party.hp1.auto = False
    opt.party.hp1.name = ""
End If
tmp = rini("party", "protect", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.party.protect.auto = True
    opt.party.protect.name = cmd(1)
Else
    opt.party.protect.auto = False
    opt.party.protect.name = ""
End If
tmp = rini("party", "follow_atk", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.party.atk.auto = True
    opt.party.atk.name = cmd(1)
Else
    opt.party.atk.auto = False
    opt.party.atk.name = ""
End If
'Heal
tmp = rini("heal", "sit_hp", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.heal.sit_hp.auto = True
    opt.heal.sit_hp.Min = cmd(1)
    opt.heal.sit_hp.Max = cmd(2)
Else
    opt.heal.sit_hp.auto = False
    opt.heal.sit_hp.Min = cmd(1)
    opt.heal.sit_hp.Max = cmd(2)
End If
tmp = rini("heal", "sit_mp", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.heal.sit_mp.auto = True
    opt.heal.sit_mp.Min = cmd(1)
    opt.heal.sit_mp.Max = cmd(2)
Else
    opt.heal.sit_mp.auto = False
    opt.heal.sit_mp.Min = cmd(1)
    opt.heal.sit_mp.Max = cmd(2)
End If
tmp = rini("heal", "hp_recon", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.heal.recon.auto = True
    opt.heal.recon.hp = CInt(cmd(1))
Else
    opt.heal.recon.auto = False
End If
tmp = rini("heal", "hp1", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.heal.hp1.auto = True
    opt.heal.hp1.skill = CInt("&H" & cmd(1))
    opt.heal.hp1.hp = cmd(2)
    opt.heal.hp1.mp = cmd(3)
Else
    opt.heal.hp1.auto = False
End If

'Basic
opt.basic.relogin = rini("basic", "relogin", 10, "config\options")
opt.basic.timeout = rini("basic", "timeout", 60, "config\options")
opt.debug.show = rini("debug", "show", 0, "config\options")

'style
opt.style.theme = rini("style", "theme", 2, "config\options")
frmMain.mnStyle_Click opt.style.theme

'pet
opt.pet = rini("pet", "auto", 0, "config\options")

'Auto HP/MP
tmp = rini("useitem", "hp", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.hp.auto = True
    opt.hp.hp = cmd(1)
Else
    opt.hp.auto = False
End If
tmp = rini("useitem", "mp", "#", "config\options")
cmd = Split(tmp, "#")
If cmd(0) = "1" Then
    opt.mp.auto = True
    opt.mp.mp = cmd(1)
Else
    opt.mp.auto = False
End If
End Sub

Public Sub Load_Monster()
Open App.path & "\data\monster.txt" For Input As #1
Dim tstr As String, tmp() As String
ReDim tmpMon(0)
Do While Not EOF(1)
    Line Input #1, tstr
    tmp = Split(tstr, ":")
    If UBound(tmp) > 0 Then
        tmpMon(UBound(tmpMon)).Type = CInt("&H" & tmp(0))
        tmpMon(UBound(tmpMon)).name = tmp(1)
        ReDim Preserve tmpMon(UBound(tmpMon) + 1)
    End If
Loop
Close #1
End Sub

Public Function Return_Item_Name(id As Long) As String
Dim i As Integer
For i = 0 To UBound(tmpItem)
If tmpItem(i).Type = id Then
Return_Item_Name = tmpItem(i).name
Exit Function
End If
Next
Return_Item_Name = "Unknown: " & Hex(id) & " : " & UBound(tmpItem)
End Function

Public Sub Load_Item()
Open App.path & "\data\item.txt" For Input As #1
Dim tstr As String, tmp() As String
ReDim tmpItem(0)
Do While Not EOF(1)
    Line Input #1, tstr
    tmp = Split(tstr, ":")
    If UBound(tmp) > 0 Then
        tmpItem(UBound(tmpItem)).Type = "&H" & tmp(0)
        tmpItem(UBound(tmpItem)).name = tmp(1)
        tmpItem(UBound(tmpItem)).Detail = tmp(2)
        ReDim Preserve tmpItem(UBound(tmpItem) + 1)
    End If
Loop
Close #1
End Sub

Public Function Return_Monster_Name(id As Long) As String
Dim i As Integer
For i = 0 To UBound(tmpMon) - 1
If tmpMon(i).Type = id Then
Return_Monster_Name = tmpMon(i).name
Exit Function
End If
Next
Return_Monster_Name = "Unknown: " & Hex(id) & " : " & UBound(tmpMon)
End Function

Public Sub Load_Skill()
Open App.path & "\data\skill.txt" For Input As #1
Dim tstr As String, tmp() As String
ReDim tmpSkill(0)
Do While Not EOF(1)
    Line Input #1, tstr
    tmp = Split(tstr, ":")
    If UBound(tmp) > 0 Then
        tmpSkill(UBound(tmpSkill)).Type = CLng("&H" & tmp(0))
        tmpSkill(UBound(tmpSkill)).name = tmp(1)
        tmpSkill(UBound(tmpSkill)).Detail = tmp(2)
        ReDim Preserve tmpSkill(UBound(tmpSkill) + 1)
    End If
Loop
Close #1
End Sub

Public Function Return_Skill_Name(id As Long) As String
Dim i As Integer
For i = 0 To UBound(tmpSkill) - 1
If tmpSkill(i).Type = id Then
Return_Skill_Name = tmpSkill(i).name
Exit Function
End If
Next
Return_Skill_Name = "Unknown: " & Hex(id)
End Function

Public Function Return_People(id As String) As Integer
Dim i As Integer
For i = 0 To UBound(People) - 1
If People(i).id = id Then
Return_People = i
Exit Function
End If
Next
Return_People = -1
End Function

Public Function Return_Monster(id As String) As String
Dim i As Integer
For i = 0 To UBound(Monster) - 1
If Monster(i).id = id Then
Return_Monster = i
Exit Function
End If
Next
Return_Monster = -1
End Function

Public Function Return_NPC(id As String) As Integer
Dim i As Integer
For i = 0 To UBound(NPC) - 1
If NPC(i).id = id Then
Return_NPC = i
Exit Function
End If
Next
Return_NPC = -1
End Function
