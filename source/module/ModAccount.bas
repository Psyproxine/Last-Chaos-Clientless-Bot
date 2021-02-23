Attribute VB_Name = "ModAccount"
Option Explicit
Private Type location_detail
    name As String
    IP As String
    Port As String
    Serv As Integer
End Type
Public Type Channels
    id As Byte
    val As Byte
    name As String
    Connection As String
End Type

Public Type Channel_Info
    id As Integer
    Channel_Name As String
    ChannelX() As Channels
End Type

Public Type Character
    id As String * 4
    name As String
    level As Long
    Class As Integer
    Clan As String
    Kill As Integer
    hp As Long
    MaxHP As Long
    mp As Long
    MaxMp As Long
    EXP As Long
    MaxEXP As Long
    str As Long
    Int As Long
    Agi As Long
    Vit As Long
    Str0 As Long
    Agi0 As Long
    Int0 As Long
    Vit0 As Long
    Spp As Long
    SP As Long
    
    
    atk As Long
    Mtk As Long
    Def As Long
    Mef As Long
    
    Money As Long
    
    Pos As Coord
    Coord As String
End Type
Public Location() As location_detail
Public tmpServ As Integer, tmpSub As Integer
Public LoginChar() As Character
Public party() As Character
Public delayRecon As Integer
Public CheckPerTime As Date
Public EncryptData As String

Public Sub Load_Account()
ReDim Location(0)
Dim All As String, ac() As String, i As Integer, X As Integer
'User = rini("Account", "user", "", "account")
User = rini("Account", "user", "", "account")
Pass = rini("Account", "pass", "", "account")
Loc = rini("Account", "location", "", "account")
Server = rini("Account", "server", 0, "account")
World = rini("Account", "world", 0, "account")
Char = rini("Account", "char", 0, "account")

frmPermission.txtPer_User.Text = rini("Web", "username", "", "account")
frmPermission.txtPer_Pass.Text = rini("Web", "password", "", "account")

frmMain.txtUser.Text = User
frmMain.txtPass.Text = Pass

frmMain.cbChar.AddItem "Manual"
frmMain.cbChar.ListIndex = 0
For i = 1 To 5
    frmMain.cbChar.AddItem CStr(i)
    If Char = i Then frmMain.cbChar.ListIndex = i
Next


All = rini("All_Serv", "Serv", "", "account")
If All <> "" Then
    ac = Split(All, ";")
    For i = 0 To UBound(ac)
        If Trim(ac(i)) <> "" Then
            frmMain.cbLocation.AddItem Trim(ac(i))
            ReDim Preserve Location(X)
            Location(X).name = Trim(ac(i))
            Location(X).IP = rini(ac(i), "ip", "", "account")
            Location(X).Port = rini(ac(i), "port", "", "account")
            Location(X).Serv = rini(ac(i), "serv", 0, "account")
            If Loc = Location(X).name Then frmMain.cbLocation.ListIndex = i
            X = X + 1
        End If
    Next
End If
End Sub

Public Sub AddAccount()
    User = Trim(frmMain.txtUser.Text)
    Pass = Trim(frmMain.txtPass.Text)
    Loc = Trim(frmMain.cbLocation.Text)
    If (Trim(frmMain.cbServer.Text) <> "" And Trim(frmMain.cbServer.Text) <> "Manual") Then
        Server = CInt(Trim(frmMain.cbServer.Text))
    Else
        Server = 0
    End If
    If (Trim(frmMain.cbWorld.Text) <> "" And Trim(frmMain.cbWorld.Text) <> "Manual") Then
        World = CInt(Trim(frmMain.cbWorld.Text))
    Else
        World = 0
    End If
    If (Trim(frmMain.cbChar.Text) <> "" And Trim(frmMain.cbChar.Text) <> "Manual") Then
        Char = CInt(Trim(frmMain.cbChar.Text)) - 1
    Else
        Char = -1
    End If
End Sub

Public Sub SaveAccount()
    wini "Account", "name", Trim(frmMain.txtUser.Text), "account"
    wini "Account", "pass", Trim(frmMain.txtPass.Text), "account"
    wini "Account", "location", Trim(frmMain.cbLocation.Text), "account"
    wini "Account", "server", Trim(frmMain.cbServer.Text), "account"
    wini "Account", "world", Trim(frmMain.cbWorld.Text), "account"
    wini "Account", "char", Trim(frmMain.cbChar.Text), "account"
End Sub

Public Sub Bot_Connect()
Dim i As Integer
'If CheckPerTime = 0 Or DateDiff("h", CheckPerTime, time) >= 12 Then
'    EncryptData = RandomAlphaNumString(100)
'    frmPermission.WebSock.CloseSck
'    frmPermission.WebSock.Connect "positron.in.th", 80
'Else
    For i = 0 To UBound(Location)
        If Location(i).name = Loc Then
            frmBot.Winsock.CloseSck
            frmBot.Winsock.Connect Location(i).IP, Location(i).Port
            Stat "Connecting to " & Location(i).IP & ":" & Location(i).Port
            Exit Sub
        End If
    Next
'End If
End Sub


Public Function Return_Class(id As Integer) As String
    Select Case id
        Case 0: Return_Class = "Titan"
        Case 1: Return_Class = "Knight"
        Case 2: Return_Class = "Healer"
        Case 3: Return_Class = "Mage"
        Case 4: Return_Class = "Rouge"
        Case 5: Return_Class = "Sorceror"
        Case Else: Return_Class = "UnKnown Class: " & id
    End Select
End Function

Public Sub Connecting(b As Boolean)
    With frmMain
        If b Then
            .btnConnect.Caption = "Disconnect"
        Else
            .btnConnect.Caption = "Connect"
            frmBot.Winsock.CloseSck
        End If
        '.txtUser.Enabled = Not B
        .txtUser.Enabled = Not b
        .txtPass.Enabled = Not b
        .cbLocation.Enabled = Not b
        .cbServer.Enabled = Not b
        .cbWorld.Enabled = Not b
        .cbChar.Enabled = Not b
    End With
    Constate = 0
    frmMain.btnSit.Caption = "Stand"
End Sub

Public Function hex2string(str As String) As String
Dim tmp As String
tmp = "1234567890-=!@#$%^&*()_+qwertyuiop[]QWERTYUIOP{}asdfghjkl;'ASDFGHJKL:""zxcvbnm,./ZXCVBNM<>?å/-À¶ØÖ¤µ¨¢ª+ñòóôÙßõö÷øùæäÓ¾ÐÑÕÃ¹ÂºÅð""®±¸íê³Ï­°,¿Ë¡´àéèÒÊÇ§Ä¦¯â¬çëÉÈ«.¼»áÍÔ×·Áã½()©ÎÚì?²ÌÆ"
hex2string = IIf(InStr(tmp, str) > 0, str, ".")
End Function

Public Sub reconnect()
delayRecon = opt.basic.relogin
frmBot.tmrRecon.Enabled = True
ShowBalloonTip "Waiting for reconnect in " & delayRecon & "sec", NIIF_ERROR
End Sub

Public Function Version()
    Version = "Aggressive Powered - v " & App.Major & "." & App.Minor & "." & Right("0000" & App.Revision, 4) & " - killzone"
End Function

Public Function MakeTime() As String
On Error GoTo z
    If (SHour < 10) Then
        MakeTime = "0" + CStr(SHour) + ":"
    Else
        MakeTime = CStr(SHour) + ":"
    End If
    If (SMin < 10) Then
        MakeTime = MakeTime + "0" + CStr(SMin) + ":"
    Else
        MakeTime = MakeTime + CStr(SMin) + ":"
    End If
    If (SSec < 10) Then
        MakeTime = MakeTime + "0" + CStr(SSec)
    Else
        MakeTime = MakeTime + CStr(SSec)
    End If
z:
End Function

Public Function Emo(id As String) As String
    Dim X As String
    Select Case id
        Case Chr(&HA)
            Emo = "·Ñ¡·ÒÂ" '       21 00 02 BA C6 00 0A 00
        Case Chr(&HB)
            Emo = "ÃèÒàÃÔ§"     'ÃèÒàÃÔ§          21 00 02 BA C6 00 0B 00
        Case Chr(&HD)
            Emo = "àÊÕÂã¨"     'àÊÕÂã¨          21 00 02 BA C6 00 0D 00
        Case Chr(&HE)
            Emo = "ÂÍ´àÂÕèÂÁ"     'ÂÍ´àÂÕèÂÁ       21 00 02 BA C6 00 0E 00
        Case Chr(&HF)
            Emo = "»ÃºÁ×Í"         '»ÃºÁ×Í         21 00 02 BA C6 00 0F 00
        Case Chr(&H10)
            Emo = "»®ÔàÊ¸"      '»®ÔàÊ¸          21 00 02 BA C6 00 10 00
        Case Chr(&H11)
            Emo = "âÍéÍÇ´"     'âÍéÍÇ´          21 00 02 BA C6 00 11 00
        Case Chr(&H13)
            Emo = "µÓË¹Ô"     'µÓË¹Ô            21 00 02 BA C6 00 13 00
        Case Chr(&H14)
            Emo = "àªÕÂÃì"  'àªÕÂÃì             21 00 02 BA C6 00 14 00
        Case Chr(&H15)
            Emo = "·éÒ·ÒÂ"   '·éÒ·ÒÂ          21 00 02 BA C6 00 15 00
        Case Chr(&H16)
            Emo = "à¤ÒÃ¾"   'à¤ÒÃ¾
        Case Else
            Emo = "Unknown Case: " & ChrtoHex(id)
    End Select
End Function

Public Function RandNum(Upper As Integer, Lower As Integer) As Long
  On Error GoTo LocalError
  Randomize
  RandNum = ((Upper - Lower) * Rnd + Lower)
  Exit Function
LocalError:
  RandNum = 1
End Function

Public Sub Play(File As String)
    If FileExists(File) Then
        File = """" & File & """"
        'mciSendString "open " & file & " type MPEGVideo", 0, 0, 0
        'mciSendString "play " & file, 0, 0, 0
    End If
End Sub

Function FileExists(ByVal strPathName As String) As Boolean
On Error GoTo errHandle
Open strPathName For Input As #1: Close 1
FileExists = True: Exit Function
errHandle:
FileExists = False
End Function
