VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLIRCUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Option Explicit
Option Compare Text
Public Event farout()
Private Type privUser
    
    nick As String
    User As String
    Host As String
    FName As String
    
    Channel As String
   
    Modes As String
    
    Server As String
    IdleTime As Integer
    SignOnTime As Double
    InUse As Boolean
    
    Tag As String
    Custom1 As String
    Custom2 As String
    Custom3 As String
    Custom4 As String
    Custom5 As String
    Custom6 As String
    Custom7 As String
    Custom8 As String
    Custom9 As String
    Custom10 As String
    
    Custom11 As Boolean
    Custom12 As Boolean
    Custom13 As Boolean
    Custom14 As Boolean
    Custom15 As Boolean
    
    Custom16 As Integer
    Custom17 As Integer
    Custom18 As Integer
    Custom19 As Integer
    Custom20 As Integer
    'Add any tag's you'd like to keep tabs on in the User
    'Module. This includes any extra's that the class does
    'not support

End Type

Private User(6000) As privUser

Public Function ChanList(nick As String) As String
    On Error GoTo ErrHandler
    
    ChanList = User(FindUser(nick)).Channel
    Exit Function
    
ErrHandler:
    ChanList = "Chan Not Found"
    Exit Function
End Function


Public Function IsOnChan(nick As String, Chan As String) As Boolean
    If Left(Chan, 1) = "#" Then Chan = Mid(Chan, 2)
    If InStr(User(FindUser(nick)).Channel + " ", Chan + " ") Then
        IsOnChan = True
    Else
        IsOnChan = False
    End If
End Function

Public Function AddChannel(strNick As String, strChan As String)
On Error GoTo ErrHandler
    Dim tmpNum As Integer
    Dim tmpChan() As String
    
    tmpChan = Split(User(FindUser(strNick)).Channel + " [EOL]", " ")
    
    For tmpNum = 0 To 10000
        If tmpChan(tmpNum) = "[EOL]" Then Exit For
        If tmpChan(tmpNum) = strChan Then Exit Function
        If NotOpChar(tmpChan(tmpNum)) = NotOpChar(strChan) Then
            tmpChan(tmpNum) = strChan
            User(FindUser(strNick)).Channel = Replace(Join(tmpChan, " "), " [EOL]", "")
            Exit Function
        End If
    Next tmpNum
    
    User(FindUser(strNick)).Channel = Replace(Join(tmpChan, " "), " [EOL]", "") + " " + strChan
    Exit Function
ErrHandler:
End Function

Public Function ChangeNick(srcNick As String, tarNick As String) As Boolean
    Dim tmpNick As Integer
    
    tmpNick = FindUser(srcNick)
    If tmpNick = -1 Then ChangeNick = False: Exit Function
    
    User(tmpNick).nick = tarNick
    ChangeNick = True
End Function
Public Sub Debugs(strUser As String)
    Debug.Print User(FindUser(strUser)).Channel
    
End Sub

Public Function RemoveChannel(strUser As String, strChan As String) As Boolean
    Dim tmpNum As Integer
    Dim tmpChan() As String
    On Error GoTo ErrHandler
    
    tmpChan = Split(User(FindUser(strUser)).Channel + " [EOL]", " ")
    
    For tmpNum = 0 To 10000
        If tmpChan(tmpNum) = "[EOL]" Then Exit For
        If tmpChan(tmpNum) = strChan Then
            tmpChan(tmpNum) = "[REMOVED]"
        End If
        If NotOpChar(tmpChan(tmpNum)) = NotOpChar(strChan) Then
            tmpChan(tmpNum) = "[REMOVED]"
        End If
    Next tmpNum
    
    User(FindUser(strUser)).Channel = Replace(Replace(Join(tmpChan, " "), " [EOL]", ""), " [REMOVED]", "")

    Exit Function
ErrHandler:
    RemoveChannel = False
    
End Function

Private Function OpChar(Chan As String) As String
    If Left(Chan, 1) = "#" Then OpChar = "": Exit Function
    OpChar = Left(Chan, 1)
End Function

Private Function NotOpChar(Chan As String) As String
    If Left(Chan, 1) = "#" Then NotOpChar = Chan: Exit Function
    NotOpChar = Mid(Chan, 2)
End Function

Private Function FindNewUser() As Integer
    Dim tmpNum As Integer
    
    For tmpNum = 1 To 6000
        If User(tmpNum).InUse = False Then
            FindNewUser = tmpNum: Exit Function
        End If
    Next tmpNum
    
    FindNewUser = -1
End Function

Public Function FindUser(nick As String) As Integer
    Dim tmpNum As Integer
    
    For tmpNum = 1 To 6000
        If User(tmpNum).nick = nick Then
            FindUser = tmpNum: Exit Function
        ElseIf User(tmpNum).nick = "" Then
            FindUser = -1: Exit Function
        End If
    Next tmpNum

End Function

Public Function RemoveNick(srcNick As String) As Boolean
    Dim tmpNick As Integer
    
    tmpNick = FindUser(srcNick)
    If tmpNick = -1 Then RemoveNick = False: Exit Function
        
    User(tmpNick).InUse = False
    
End Function

Public Function StoreNUHInfo(strNick As String, strUser As String, strHost As String) As Boolean
    Dim tmpNum As Integer
    
    tmpNum = FindUser(strNick)
    If tmpNum = -1 Then tmpNum = FindNewUser: User(tmpNum).InUse = True
    
    User(tmpNum).nick = strNick
    User(tmpNum).User = strUser
    User(tmpNum).Host = strHost
    
    StoreNUHInfo = True
    
    Exit Function
    
End Function


Property Get UserFromNick(nick As String) As String
    UserFromNick = User(FindUser(nick)).User
End Property


Property Get HostFromNick(nick As String) As String
    HostFromNick = User(FindUser(nick)).Host
End Property

