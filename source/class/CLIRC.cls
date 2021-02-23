VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLIRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Code by Michael Leaney _
 _
Please ask for authors permission before _
using code for any purpose what so ever _
many long hard hours of work have been _
put into this code, so please pay respect _
 _
 leahcimic@hotmail.com _
 _
 ThankyouOption Explicit
 
Public CLUser As CLIRCUser
Public CLChan As CLIRCChans

Private cSocketHandle As Object
Private cErrDescription As String
Private cErrNumber As Integer
Private Params(100) As String
Private tmpUser As String
Private tmpNick As String
Private tmpHost As String
Private tmpTarget As String
Private tmpMsg As String
Private tmpChan As String
Private privConnected As Boolean

Private inBuffer As String

Public Server As String

Public Event Connected()
Public Event PrivateMessage(nick As String, User As String, Host As String, Target As String, Message As String)
Public Event Action(nick As String, User As String, Host As String, Target As String, Message As String)
Public Event Notice(nick As String, User As String, Host As String, Target As String, Message As String)
Public Event Topic(nick As String, User As String, Host As String, Chan As String, Topic As String)
Public Event Join(nick As String, User As String, Host As String, Channel As String)
Public Event Part(nick As String, User As String, Host As String, Channel As String, Message As String)
Public Event Quit(nick As String, User As String, Host As String, Message As String)
Public Event OP(srcNick As String, srcUser As String, srcHost As String, Chan As String, TargetNick As String)
Public Event DEOP(srcNick As String, srcUser As String, srcHost As String, Chan As String, Target As String)
Public Event BAN(srcNick As String, srcUser As String, srcHost As String, Chan As String, tarNick As String, TarUser As String, TarHost As String)
Public Event UNBAN(srcNick As String, srcUser As String, srcHost As String, Chan As String, tarNick As String, TarUser As String, TarHost As String)
Public Event VOICE(srcNick As String, srcUser As String, srcHost As String, Chan As String, Target As String)
Public Event DEVOICE(srcNick As String, srcUser As String, srcHost As String, Chan As String, Target As String)
Public Event chanOnModerated(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOffModerated(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOnInviteOnly(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOffInviteOnly(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOnSecret(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOffSecret(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOnPrivate(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOffPrivate(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOnOpsChangeTopics(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOffOpsChangeTopics(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOnKey(srcNick As String, srcUser As String, srcHost As String, Chan As String, Key As String)
Public Event chanOffKey(srcNick As String, srcUser As String, srcHost As String, Chan As String, Key As String)
Public Event chanOnLimit(srcNick As String, srcUser As String, srcHost As String, Chan As String, Limit As String)
Public Event chanOffLimit(srcNick As String, srcUser As String, srcHost As String, Chan As String, Limit As String)
Public Event chanOnNoExternal(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event chanOffNoExternal(srcNick As String, srcUser As String, srcHost As String, Chan As String)
Public Event Kick(srcNick As String, srcUser As String, srcHost As String, Chan As String, Target As String, Message As String)
Public Event nick(srcNick As String, srcUser As String, srcHost As String, newNick As String)
Public Event errNickInUse(strNick As String)

Private User As privUser

Private Type privUser
    nick As String
    Fullname As String
    UserName As String
    HostName As String
    ServerName As String
End Type


Sub DriveCTCP(srcNick As String, srcUser As String, srcHost As String, srcTarget As String, srcMessage As String)

    Dim srcNum As Integer
    
    If InStr(srcMessage, " ") Then srcNum = InStr(srcMessage, " ") Else srcNum = Len(srcMessage) + 1
    
    Select Case Mid(srcMessage, 1, srcNum - 1)
        Case "PING"
            Notice srcNick, "" + srcMessage + ""
        Case "VERSION"
            Notice srcNick, "VERSION 8,1clIRC Visual Basic IRC - Michael Leaney"
        Case "TIME"
            Notice srcNick, "TIME " + Time$ + ""
        Case "FINGER"
            Notice srcNick, "FINGER Please stop fingering me!"
        Case "ACTION"
            RaiseEvent Action(srcNick, srcUser, srcHost, srcTarget, Mid(srcMessage, 8))
    End Select
End Sub

Public Property Get isConnected() As Boolean
    isConnected = privConnected
End Property


Function MODELIST(Mode As String, Chan As String, List As String)

'510 char limit

Dim tmpNicks() As String
Dim tmpNum As Integer
Dim Prefix As String
Dim tmpMsg As String
Dim tmpCnt As Integer

tmpNicks = Split(List + " #EOL#", " ")

Prefix = "MODE " + Chan + " " + Mode


For tmpNum = 0 To 10000
    If tmpNicks(tmpNum) = "#EOL#" Then Exit For
    If CLChan.NotOpChar(tmpNicks(tmpNum)) = User.nick Then GoTo SkipMe
    If Len(Prefix + String(tmpCnt, Right(Mode, 1)) + " " + tmpMsg + " " + CLChan.NotOpChar(tmpNicks(tmpNum))) > 511 Or tmpCnt = 6 Then
        SendData Prefix + String(tmpCnt - 1, Right(Mode, 1)) + " " + tmpMsg: tmpMsg = CLChan.NotOpChar(tmpNicks(tmpNum) + " "): tmpCnt = 1
        Debug.Print Prefix + String(tmpCnt - 1, Right(Mode, 1)) + " " + tmpMsg
    Else
        tmpCnt = tmpCnt + 1
        tmpMsg = tmpMsg + CLChan.NotOpChar(tmpNicks(tmpNum)) + " "
    End If
SkipMe:
Next tmpNum

SendData Prefix + String(tmpCnt - 1, Right(Mode, 1)) + " " + tmpMsg
Debug.Print Prefix + String(tmpCnt - 1, Right(Mode, 1)) + " " + tmpMsg
End Function
Function KICKLIST(Chan As String, List As String, Comment As String)

Dim tmpNicks() As String
Dim tmpNum As Integer

tmpNicks = Split(List + " #EOL#", " ")

For tmpNum = 0 To 10000
    If tmpNicks(tmpNum) = "#EOL#" Then Exit For
    If CLChan.NotOpChar(tmpNicks(tmpNum)) = User.nick Then GoTo SkipMe
        Debug.Print Chan, tmpNicks(tmpNum), Comment
        Kick Chan, CLChan.NotOpChar(tmpNicks(tmpNum)), Comment
    
SkipMe:
Next tmpNum

End Function

Function ConvNToMask(List As String, Mode As Integer)

Dim tmpNicks() As String
Dim tmpNum As Integer
Dim tmpMsg As String
tmpNicks = Split(List + " #EOL#", " ")

For tmpNum = 0 To 10000
    If tmpNicks(tmpNum) = "#EOL#" Then Exit For
    If CLChan.NotOpChar(tmpNicks(tmpNum)) = User.nick Then GoTo SkipMe
    If tmpNicks(tmpNum) = "" Then GoTo SkipMe
    Select Case Mode
        Case 1
            tmpMsg = tmpMsg + tmpNicks(tmpNum) + "!*@*" + " "
        Case 2
            tmpMsg = tmpMsg + "*!" + CLUser.UserFromNick(tmpNicks(tmpNum)) + "@*" + " "
        Case 3
            tmpMsg = tmpMsg + "*!" + CLUser.UserFromNick(tmpNicks(tmpNum)) + "@" + CLUser.HostFromNick(tmpNicks(tmpNum)) + " "
        Case 4
            tmpMsg = tmpMsg + "*!*@" + CLUser.HostFromNick(tmpNicks(tmpNum)) + " "
    End Select
SkipMe:
Next tmpNum

ConvNToMask = tmpMsg

End Function

Public Function Connect(ServName As String, Optional ServPort As String = "6667") As Boolean
'Code by Michael Leaney _
 _
Please ask for authors permission before _
using code for any purpose what so ever _
many long hard hours of work have been _
put into this code, so please pay respect _
 _
 leahcimic@hotmail.com _
 _
 Thankyou
    With User
        If .Fullname = "" Then GoTo ReqVarNotSet
        If .HostName = "" Then GoTo ReqVarNotSet
        If .ServerName = "" Then GoTo ReqVarNotSet
        If .UserName = "" Then GoTo ReqVarNotSet
        If .nick = "" Then GoTo ReqVarNotSet
    End With
    
    cSocketHandle.Connect ServName, ServPort
    Connect = True
    Exit Function
    
ReqVarNotSet:
    RaiseError 11
    Connect = False
  
End Function



Public Sub Connected()
'Code by Michael Leaney _
 _
Please ask for authors permission before _
using code for any purpose what so ever _
many long hard hours of work have been _
put into this code, so please pay respect _
 _
 leahcimic@hotmail.com _
 _
 Thankyou
    SendData "NICK " + User.nick
    SendData "USER " + User.UserName + " " + User.HostName + " " + User.ServerName + " :" + User.Fullname
    
    'RaiseEvent Connected
    
End Sub

Private Function DriveParams()
'Code by Michael Leaney _
 _
Please ask for authors permission before _
using code for any purpose what so ever _
many long hard hours of work have been _
put into this code, so please pay respect _
 _
 leahcimic@hotmail.com _
 _
 ThankyouDim tmpNicks() As String
Dim Host As String
Dim tmpMode As String
Dim tmpNum As Integer
Dim tmpNicks() As String
    
    TrimParams
    
    'Ping? Pong!
    If Params(0) = "PING" Then SendData "PONG " + Params(1)
    'PRIVMSG

    Select Case Params(1)
        
        Case "PRIVMSG"
            PutUserInfoToTmp Params(0)
            tmpTarget = Params(2)
            tmpMsg = Mid(Params(3), 2)
            If tmpMsg Like "*" Then
                DriveCTCP tmpNick, tmpUser, tmpHost, tmpTarget, Mid(tmpMsg, 2, Len(tmpMsg) - 2): Exit Function
            End If
            If tmpMsg = "USEHOST" And tmpNick = User.nick And tmpTarget = User.nick Then
                User.HostName = Host: Exit Function
            End If
            RaiseEvent PrivateMessage(tmpNick, tmpUser, tmpHost, tmpTarget, tmpMsg)
        Case "JOIN"
            PutUserInfoToTmp Params(0)
                'CLUser DATA - This manages the UserDB and keeps it up to date
                CLChan.StoreInfo Mid(Params(2), 2), ""
                CLChan.AddNick tmpNick, Mid(Params(2), 2)
                CLUser.StoreNUHInfo tmpNick, tmpUser, tmpHost
                CLUser.AddChannel tmpNick, Mid(Params(2), 2)
                
                'END CLUUser
            RaiseEvent Join(tmpNick, tmpUser, tmpHost, Mid(Params(2), 2))
        Case "PART"
            PutUserInfoToTmp Params(0)
            'CLUser DATA - This manages the UserDB and keeps it up to date
            CLChan.RemoveNick Params(2), tmpNick
            CLUser.RemoveChannel tmpNick, Params(2)
            
            'END CLUser
            RaiseEvent Part(tmpNick, tmpUser, tmpHost, Params(2), Mid(Params(3), 2))
        Case "NOTICE"
            PutUserInfoToTmp Params(0)
            tmpTarget = Params(2)
            tmpMsg = Mid(Params(3), 2)
             If tmpMsg Like "*" Then
                DriveCTCP tmpNick, tmpUser, tmpHost, tmpTarget, Mid(tmpMsg, 2, Len(tmpMsg) - 2): Exit Function
            End If
            RaiseEvent Notice(tmpNick, tmpUser, tmpHost, tmpTarget, tmpMsg)
        Case "QUIT"
            PutUserInfoToTmp Params(0)
                'CLUser DATA - This manages the UserDB and keeps it up to date
                 CLChan.RemoveAllNick tmpNick
                 
                 CLUser.RemoveNick tmpNick
                
                'END CLUser
            tmpMsg = Mid(Params(2), 2)
            RaiseEvent Quit(tmpNick, tmpUser, tmpHost, tmpMsg)
        Case "NICK"
            PutUserInfoToTmp Params(0)
                If tmpNick = User.nick Then
                    User.nick = TrimIt(Params(2))
                End If

                'CLUser DATA - This manages the UserDB and keeps it up to date
    
                CLChan.NickChange tmpNick, TrimIt(Params(2))
                CLUser.ChangeNick tmpNick, TrimIt(Params(2))
                
                'END CLUser
            RaiseEvent nick(tmpNick, tmpUser, tmpHost, TrimIt(Params(2)))
        Case "KICK"
            PutUserInfoToTmp Params(0)
                'CLUser DATA - This manages the UserDB and keeps it up to date
                CLChan.RemoveNick Params(3), tmpNick
                CLUser.RemoveChannel tmpNick, Params(3)
                'END CLUser
            RaiseEvent Kick(tmpNick, tmpUser, tmpHost, Params(2), Params(3), Mid(Params(4), 2))
        Case "TOPIC"
            PutUserInfoToTmp Params(0)
                'CLChan DATA -
                CLChan.StoreInfo Params(2), TrimIt(Params(3))
                CLChan.StoreWhoSetTopic Params(2), tmpNick
                'END
            RaiseEvent Topic(tmpNick, tmpUser, tmpHost, Params(2), TrimIt(Params(3)))
        Case "MODE"
            ':leahciM!phil@vw10587.esc.net.au MODE #new +b *!*@vw10587.esc.net.au
            PutUserInfoToTmp Params(0)
            If IsChan(Params(2)) Then
                'Channel Mode, Lots of work needs to be done HERE!
                ChanMode
           Else
                'Person Mode, Finish me one day
                tmpTarget = Params(2)
                tmpMode = Params(3)
                If InStr(1, tmpMode, ":") Then
                End If
                    
            End If
        Case "001"
            PrivMSG UserNick, "USEHOST"
            RaiseEvent Connected
            privConnected = True
        Case "002"
        Case "003"
        Case "004"
        Case "251"
        Case "252"
        Case "253"
        Case "254"
            'Channels Formed
        Case "255"
            'Clients
        Case "332"
            CLChan.StoreInfo Params(3), TrimIt(Params(4))
        Case "333"
            CLChan.StoreWhoSetTopic Params(3), Params(4)
        Case "353"
            tmpChan = Params(4)
            tmpNicks() = Split(TrimIt(Params(5)) + " #EOL#", " ")
            For tmpNum = 0 To UBound(tmpNicks)
                If tmpNicks(tmpNum) = "#EOL#" Then Exit Function
                CLUser.StoreNUHInfo CLChan.NotOpChar(tmpNicks(tmpNum)), "", ""
                CLUser.AddChannel CLChan.NotOpChar(tmpNicks(tmpNum)), tmpChan
                CLChan.AddNick tmpNicks(tmpNum), tmpChan
            Next tmpNum
        Case "375"
            'Start Of MOTD
        Case "378"
            'MOTD
        Case "376"
            'end of MOTD
        Case "377"
        Case "433"
            '433 temp29687 leahciM :Nickname is already in use. Please use another Nickname
            RaiseEvent errNickInUse(Params(3))
        End Select
End Function

Public Function UsrIsOnChan(nick As String, Chan As String) As Boolean
    IsOnChan = CLUser.IsOnChan(nick, Chan)
End Function
Public Function UsrChanList(nick As String) As String
UsrChanList = CLUser.ChanList(nick)
End Function
Private Function TrimIt(srcString As String) As String
    If Left(srcString, 1) = ":" Then TrimIt = Trim(Right(srcString, Len(srcString) - 1)) Else TrimIt = Trim(srcString)
End Function

Private Sub ChanMode()
Dim tmpScan As Integer
Dim tmpAction As Boolean
Dim tmpNum As Integer
Dim srcNick As String
Dim srcUser As String
Dim srcHost As String

srcNick = tmpNick
srcUser = tmpUser
srcHost = tmpHost

tmpChan = Params(2)
tmpMode = Params(3)
tmpNum = 0

For tmpScan = 1 To Len(tmpMode)
    If Mid(tmpMode, tmpScan, 1) <> "+" And Mid(tmpMode, tmpScan, 1) <> "-" Then tmpNum = tmpNum + 1
    Select Case LCase(Mid(tmpMode, tmpScan, 1))
        Case "+"
            tmpAction = True
        Case "-"
            tmpAction = False
        Case "o"
            tmpTarget = Params(3 + tmpNum)
            
            If tmpAction = True Then
                'CLUser STUFF
                CLChan.AddNick "@" + tmpTarget, tmpChan
                CLUser.AddChannel tmpTarget, "@" + tmpChan
                'END
                RaiseEvent OP(srcNick, srcUser, srcHost, tmpChan, tmpTarget)
            Else
                'CLUser STUFF
                CLChan.AddNick tmpTarget, tmpChan
                CLUser.AddChannel tmpTarget, tmpChan
                'END
                RaiseEvent DEOP(srcNick, srcUser, srcHost, tmpChan, tmpTarget)
            End If
        Case "b"
            tmpTarget = Params(3 + tmpNum)
            PutUserInfoToTmp tmpTarget
            
            If tmpAction = True Then
                RaiseEvent BAN(srcNick, srcUser, srcHost, tmpChan, tmpNick, tmpUser, tmpHost)
            Else
                RaiseEvent UNBAN(srcNick, srcUser, srcHost, tmpChan, tmpNick, tmpUser, tmpHost)
            End If
        Case "v"
            tmpTarget = Params(3 + tmpNum)
            If tmpAction = True Then
                'CLUser STUFF
                CLChan.AddNick "+" + tmpTarget, tmpChan
                CLUser.AddChannel tmpTarget, "+" + tmpChan
                'END
                RaiseEvent VOICE(srcNick, srcUser, srcHost, tmpChan, tmpTarget)
            Else
                'CLUser STUFF
                CLChan.AddNick tmpTarget, tmpChan
                CLUser.AddChannel tmpTarget, "@" + tmpChan
                'END
                RaiseEvent DEVOICE(srcNick, srcUser, srcHost, tmpChan, tmpTarget)
            End If
        Case "m"
            'moderated
            If tmpAction = True Then
                RaiseEvent chanOnModerated(srcNick, srcUser, srcHost, tmpChan)
            Else
                RaiseEvent chanOffModerated(srcNick, srcUser, srcHost, tmpChan)
            End If
            tmpNum = tmpNum - 1
        Case "t"
            'ops set topic
            If tmpAction = True Then
                RaiseEvent chanOnOpsChangeTopics(srcNick, srcUser, srcHost, tmpChan)
            Else
                RaiseEvent chanOffOpsChangeTopics(srcNick, srcUser, srcHost, tmpChan)
            End If
            
            tmpNum = tmpNum - 1
        Case "n"
            'no external
            If tmpAction = True Then
                RaiseEvent chanOnNoExternal(srcNick, srcUser, srcHost, tmpChan)
            Else
                RaiseEvent chanOffNoExternal(srcNick, srcUser, srcHost, tmpChan)
            End If
            tmpNum = tmpNum - 1
        Case "i"
            'invite only
            If tmpAction = True Then
                RaiseEvent chanOnInviteOnly(srcNick, srcUser, srcHost, tmpChan)
            Else
                RaiseEvent chanOffInviteOnly(srcNick, srcUser, srcHost, tmpChan)
            End If
            tmpNum = tmpNum - 1
        Case "k"
            If tmpAction = True Then
                RaiseEvent chanOnKey(srcNick, srcUser, srcHost, tmpChan, Params(3 + tmpNum))
            Else
                RaiseEvent chanOffKey(srcNick, srcUser, srcHost, tmpChan, Params(3 + tmpNum))
            End If
            'key
        Case "l"
            If tmpAction = True Then
                RaiseEvent chanOnLimit(srcNick, srcUser, srcHost, tmpChan, Params(3 + tmpNum))
            Else
                RaiseEvent chanOffLimit(srcNick, srcUser, srcHost, tmpChan, Params(3 + tmpNum))
            End If
            'limit users
        Case "p"
            If tmpAction = True Then
                RaiseEvent chanOnPrivate(srcNick, srcUser, srcHost, tmpChan)
            Else
                RaiseEvent chanOffPrivate(srcNick, srcUser, srcHost, tmpChan)
            End If
            tmpNum = tmpNum - 1
            'Private
        Case "s"
            'secret
            If tmpAction = True Then
                RaiseEvent chanOnSecret(srcNick, srcUser, srcHost, tmpChan)
            Else
                RaiseEvent chanOffSecret(srcNick, srcUser, srcHost, tmpChan)
            End If
            tmpNum = tmpNum - 1
    End Select
    
Next tmpScan


End Sub

Public Function IsChan(Chan As String) As Boolean
    If Left(Chan$, 1) = "#" Then
        IsChan = True
    Else
        IsChan = False
    End If
End Function
Public Property Get errDescription() As String
 errDescription = cErrDescription
End Property
Public Property Get errNumber() As Integer
    errNumber = cErrNumber
End Property

Public Function Join(newChan As String, Optional newKey As String = "") As Boolean

    If SendData("JOIN " + newChan + " " + newKey) = False Then GoTo ErrHandler
    Join = True
    Exit Function
    
ErrHandler:
    RaiseError 21
    Join = False

End Function

Public Function Action(newTarget As String, newMessage As String) As Boolean

    If PrivMSG(newTarget, "ACTION " + newMessage + "") = False Then GoTo ErrHandler
    Action = True
    Exit Function
    
ErrHandler:
    RaiseError 21
    Action = False

End Function


Sub killParams()
Dim tmpNum As Integer

For tmpNum = 0 To UBound(Params)
        Params(tmpNum) = ""
Next tmpNum



End Sub

Public Function nick(newNick As String) As Boolean
    
    If SendData("NICK " + newNick) = False Then GoTo ErrHandler
    nick = True
    Exit Function
    
ErrHandler:
    nick = False
    RaiseError 21
    
End Function
Public Function OP(Chan As String, nick As String) As Boolean

    If Me.Mode(Chan, "+o", nick) = True Then
        OP = True
        Exit Function
    Else
        OP = False
    End If
    
End Function

Public Function DEOP(Chan As String, nick As String) As Boolean

    If Me.Mode(Chan, "-o", nick) = True Then
        DEOP = True
        Exit Function
    Else
        DEOP = False
    End If
    
End Function
Public Function BAN(Chan As String, Mask As String) As Boolean

    If Me.Mode(Chan, "+b", Mask) = True Then
        BAN = True
        Exit Function
    Else
        BAN = False
    End If
    
End Function
Public Function UNBAN(Chan As String, Mask As String) As Boolean

    If Me.Mode(Chan, "-b", Mask) = True Then
        UNBAN = True
        Exit Function
    Else
        UNBAN = False
    End If
    
End Function


Public Function PassParams(vData As String) As Boolean

Dim tmpNum As Integer
Dim CountParams As Integer
Dim FirstParam As String
Dim LastParam As String

    killParams
    CountParams = 1
    
    FirstParam = Mid(vData, 1, InStr(1, vData, " "))
    vData = Mid(vData, InStr(1, vData, " ") + 1)
    
    If InStr(1, vData, ":") > 0 Then
        LastParam = Mid(vData, InStr(1, vData, ":"))
        vData = Mid(vData, 1, InStr(1, vData, ":") - 1)
    End If
    
    vData = Trim(vData)
    
    'CountParams
    If vData = "" Then tmpNum = 1: GoTo SkipProcess
        
    For tmpNum = 1 To Len(vData)
        If Mid(vData, tmpNum, 1) = " " Then CountParams = CountParams + 1
    Next tmpNum
    'End Count
    
    'Put Params in Array
    For tmpNum = 1 To CountParams - 1
        Params(tmpNum) = Mid(vData, 1, InStr(1, vData, " "))
        vData = Mid(vData, InStr(1, vData, " ") + 1)
    Next tmpNum
     
    Params(tmpNum) = vData

    tmpNum = tmpNum + 1

SkipProcess:
    Params(tmpNum) = LastParam
    Params(0) = FirstParam
    
    PassParams = True
    DriveParams
    
    Exit Function
    
ErrHandler:
    PassParams = False
    RaiseError 22
    
End Function

Private Function PutUserInfoToTmp(vData)
         If Left(vData, 1) = ":" Then vData = Right(vData, Len(vData) - 1)
        
        If InStr(1, vData, "!") = 0 Or InStr(1, vData, "@") = 0 Then
            tmpNick = vData: Exit Function
        End If
        
        tmpNick = Mid(vData, 1, InStr(1, vData, "!") - 1)
        tmpUser = Mid(vData, InStr(1, vData, "!") + 1, (InStr(1, vData, "@") - InStr(1, vData, "!")) - 1)
        tmpHost = Mid(vData, InStr(1, vData, "@") + 1)
End Function

Public Function Quit(QuitMessage As String) As Boolean

    If SendData("QUIT :" + QuitMessage) = False Then GoTo ErrHandler
    Quit = True
    Exit Function
    
ErrHandler:
    Quit = False
    RaiseError 21
    
End Function

Public Function Part(vData As String) As Boolean

    If SendData("PART " + vData) = False Then GoTo ErrHandler
    Part = True
    Exit Function
    
ErrHandler:
    Part = False
    RaiseError 21
    
End Function

Public Function Mode(Channel As String, tmpMode As String, Optional Options As String = "") As Boolean

    If SendData("MODE " + Channel + " " + tmpMode + " " + Options) = False Then GoTo ErrHandler
    Mode = True
    Exit Function
    
ErrHandler:
    Mode = False
    RaiseError 21
    
End Function

Public Sub RecvData(vData As String)
Dim CountParams As Integer
Dim tmpNum As Integer

    CountParams = 0
    
    'Read Buffer
    inBuffer = inBuffer + vData
    
    'Place Buffer in vData
    vData = inBuffer
    
    For tmpNum = 1 To Len(vData)
        If Mid(vData, tmpNum, 1) = Chr$(10) Then CountParams = CountParams + 1
    Next tmpNum
    
    For tmpNum = 1 To CountParams
          Debug.Print Mid(vData, 1, InStr(1, vData, Chr$(10)) - 1)
        PassParams Mid(vData, 1, InStr(1, vData, Chr$(10)) - 1):
        vData = Mid(vData, InStr(1, vData, Chr$(10)) + 1)
    Next tmpNum

    inBuffer = vData

End Sub

Public Function Topic(Channel As String, newTopic As String) As Boolean

    If SendData("TOPIC " + Channel + " :" + newTopic) = False Then GoTo ErrHandler
    Topic = True
    Exit Function
    
ErrHandler:
    Topic = False
    RaiseError 21
    
End Function

Public Function Kick(Channel As String, nick As String, Optional Comment As String = "") As Boolean

    If SendData("KICK " + Channel + " " + nick + " :" + Comment) = False Then GoTo ErrHandler
    Kick = True
    Exit Function
    
ErrHandler:
    Kick = False
    RaiseError 21
    
End Function


Public Function PrivMSG(Destination As String, Message As String) As Boolean
    Dim Chars As String
    Dim LengthAllowed As String
    Chars = ":" + User.nick + "!" + User.UserName + "@" + User.HostName + " PRIVMSG " + Destination + " :"
    LengthAllowed = 510 - Len(Chars)
    Debug.Print LengthAllowed, Format(Len(Message) / LengthAllowed, "##.##")
Repeat:
    
    If Len(Message) > LengthAllowed Then
        If SendData("PRIVMSG " + Destination + " :" + Left(Message, LengthAllowed)) = False Then GoTo ErrHandler
        Message = Mid(Message, LengthAllowed)
    End If
    
    If Len(Message) > LengthAllowed Then
        GoTo Repeat
    Else
        If SendData("PRIVMSG " + Destination + " :" + Message) = False Then GoTo ErrHandler
    End If
    
    'If SendData("PRIVMSG " + Destination + " :" + Message) = False Then GoTo ErrHandler
    PrivMSG = True
    Exit Function
    
ErrHandler:
    PrivMSG = False
    RaiseError 21
    
End Function

Public Function Notice(Destination As String, Message As String) As Boolean

    If SendData("NOTICE " + Destination + " :" + Message) = False Then GoTo ErrHandler
    Notice = True
    Exit Function
    
ErrHandler:
    Notice = False
    RaiseError 21
    
End Function

Sub TrimParams()
    Dim tmpNum As Integer
    
    For tmpNum = 0 To 100
        Params(tmpNum) = Trim(Params(tmpNum))
        Params(tmpNum) = Replace(Params(tmpNum), Chr(13), "")
    Next tmpNum


End Sub

Public Function WhoIs(Target As String) As Boolean

    If SendData("WHOIS " + Target) = False Then GoTo ErrHandler
    WhoIs = True
    Exit Function
    
ErrHandler:
    WhoIs = False
    RaiseError 21
    
End Function


Public Function WhoWas(Target As String) As Boolean

    If SendData("WHOWAS " + Target) = False Then GoTo ErrHandler
    WhoWas = True
    Exit Function
    
ErrHandler:
    WhoWas = False
    RaiseError 21
    
End Function



Public Function Invite(nick As String, Channel As String) As Boolean

    If SendData("INVITE " + nick + " " + Channel) = False Then GoTo ErrHandler
    Invite = True
    Exit Function
    
ErrHandler:
    Invite = False
    RaiseError 21
    
End Function

Private Sub RaiseError(vError As Integer)
    Select Case vError
        Case 11
            cErrDescription = "UserFullName, UserHostName, UserServName, UserName and UserNick are required to be set!"
        Case 21
            cErrDescription = "The socket is in an invalid state to send data"
        Case 22
            cErrDescription = "Error in PassParams please contact software manafacture"
    End Select
End Sub



Public Function SendData(vData As String) As Boolean
    
    On Error GoTo ErrHandler
    
    cSocketHandle.SendData vData + vbCrLf
    SendData = True
    Exit Function
    
ErrHandler:
    RaiseError 21
    SendData = False
    
    
End Function

Public Property Get SocketHandle() As Object
    SocketHandle = cSocketHandle
End Property
Public Property Let SocketHandle(vLink As Object)
    Set cSocketHandle = vLink
End Property


Public Property Let UserNick(vData As String)
    User.nick = vData
End Property
Public Property Let UserServName(vData As String)
    User.ServerName = vData
End Property
Public Property Let UserHostName(vData As String)
    User.HostName = vData
End Property

Public Property Let UserFullName(vData As String)
    User.Fullname = vData
End Property
Public Property Get UserFullName() As String
    UserFullName = User.Fullname
End Property

Public Property Get UserNick() As String
    UserNick = User.nick
End Property

Public Property Get UserServName() As String
    UserServName = User.ServerName
End Property

Public Property Get UserName() As String
    UserName = User.UserName
End Property
Public Property Get UserHostName() As String
    UserHostName = User.HostName
End Property

Public Property Let UserName(vData As String)
    User.UserName = vData
End Property

Private Sub Class_Initialize()
    'Code by Michael Leaney _
 _
Please ask for authors permission before _
using code for any purpose what so ever _
many long hard hours of work have been _
put into this code, so please pay respect _
 _
 leahcimic@hotmail.com _
 _
 Thankyou
    
    Set CLUser = New CLIRCUser
    Set CLChan = New CLIRCChans
End Sub

Private Sub Class_Terminate()
'Code by Michael Leaney _
 _
Please ask for authors permission before _
using code for any purpose what so ever _
many long hard hours of work have been _
put into this code, so please pay respect _
 _
 leahcimic@hotmail.com _
 _
 Thankyou
        
End Sub


