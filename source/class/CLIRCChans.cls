VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLIRCChans"
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

Private Type privChan
    
    Nicks As String
    Chan As String
    Topic As String
    Modes As String
    Bans As String
    TopicWhoSet As String
    
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

Private Chan(6000) As privChan

Public Function NickList(strChan As String) As String
    On Error GoTo ErrHandler
    NickList = Chan(FindChan(strChan)).Nicks
    Exit Function

ErrHandler:
    
End Function


Public Function NickNOOPList(strChan As String) As String
    On Error GoTo ErrHandler
    
    Dim tmpNick() As String
    Dim tmpNum As Integer
    Dim tmpNicks As String
    
    
    tmpNick = Split(Chan(FindChan(strChan)).Nicks + " #EOL#", " ")
    
    For tmpNum = 0 To 6000
        If tmpNick(tmpNum) = "#EOL#" Then Exit For
        If Left$(tmpNick(tmpNum), 1) <> "@" Then tmpNicks = tmpNicks + tmpNick(tmpNum) + " "
    Next tmpNum

    NickNOOPList = tmpNicks
    Exit Function

ErrHandler:
    
End Function

Public Function NickListing(strChan As String, isLike As String) As String
    On Error GoTo ErrHandler
    
    Dim tmpNick() As String
    Dim tmpNum As Integer
    Dim tmpNicks As String
   
    
    tmpNick = Split(Chan(FindChan(strChan)).Nicks + " #EOL#", " ")
    
    For tmpNum = 0 To 6000
        If tmpNick(tmpNum) = "#EOL#" Then Exit For
        If tmpNick(tmpNum) Like isLike Then tmpNicks = tmpNicks + tmpNick(tmpNum) + " "
    Next tmpNum

    NickListing = tmpNicks
    Exit Function

ErrHandler:
    
End Function




Public Function NickOPList(strChan As String) As String
    On Error GoTo ErrHandler
    
    Dim tmpNick() As String
    Dim tmpNum As Integer
    Dim tmpNicks As String
    
    
    tmpNick = Split(Chan(FindChan(strChan)).Nicks + " #EOL#", " ")
    
    For tmpNum = 0 To 6000
        If tmpNick(tmpNum) = "#EOL#" Then Exit For
        If Left$(tmpNick(tmpNum), 1) = "@" Then tmpNicks = tmpNicks + tmpNick(tmpNum) + " "
    Next tmpNum

    NickOPList = tmpNicks

    Exit Function

ErrHandler:
    
End Function




Public Function IsOnChan(nick As String, strChan As String) As Boolean
    If InStr(Chan(FindChan(strChan)).Nicks + " ", nick + " ") Then
        IsOnChan = True
    Else
        IsOnChan = False
    End If
End Function

Public Function AddNick(strNick As String, strChan As String)
    
    Dim tmpNum As Integer
    Dim tmpNick() As String
    
    tmpNick = Split(Chan(FindChan(strChan)).Nicks + " [EOL]", " ")
    
    For tmpNum = 0 To 10000
        If tmpNick(tmpNum) = "[EOL]" Then Exit For
        If tmpNick(tmpNum) = strNick Then Exit Function
        If NotOpChar(tmpNick(tmpNum)) = NotOpChar(strNick) Then
            tmpNick(tmpNum) = strNick: Chan(FindChan(strChan)).Nicks = Replace(Join(tmpNick, " "), " [EOL]", ""): Exit Function
        End If
    Next tmpNum
    
    Chan(FindChan(strChan)).Nicks = Replace(Join(tmpNick, " "), " [EOL]", "") + " " + strNick
    

End Function
Public Function RemoveAllNick(strNick As String)
    Dim tmpNum As Double
    strNick = Trim(strNick)
    
    For tmpNum = 1 To 10000
        If Chan(tmpNum).Chan = "" Then Exit Function
        Me.RemoveNick Chan(tmpNum).Chan, strNick
    Next tmpNum
    
    
End Function


Public Function NickChange(strNick As String, newNick As String)
    Dim tmpNum As Double
    
    For tmpNum = 1 To 10000
        If Chan(tmpNum).Chan = "" Then Exit For
        Chan(tmpNum).Nicks = Replace(Chan(tmpNum).Nicks + " ", strNick + " ", newNick + " ")
        
    Next tmpNum
    
    
    
End Function



Public Function RemoveNick(strChan As String, strNick As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim tmpNum As Integer
    Dim tmpNick() As String
    
    tmpNick = Split(Chan(FindChan(strChan)).Nicks + " [EOL]", " ")
    
    For tmpNum = 0 To 10000
        If tmpNick(tmpNum) = "[EOL]" Then Exit For
        If tmpNick(tmpNum) = strNick Then
            tmpNick(tmpNum) = "[REMOVED]"
        End If
        If NotOpChar(tmpNick(tmpNum)) = NotOpChar(strNick) Then
            tmpNick(tmpNum) = "[REMOVED]"
        End If
    Next tmpNum
    
    Chan(FindChan(strChan)).Nicks = Replace(Replace(Join(tmpNick, " "), " [EOL]", ""), " [REMOVED]", "")
    
    Exit Function
    
ErrHandler:

End Function

Private Function OpChar(Chan As String) As String
    
    Select Case Left(Chan, 1)
        Case "@", "+", "%", "$"
            OpChar = Mid(Chan, 1, 1)
        Case Else
            OpChar = ""
    End Select
    
End Function

Public Function NotOpChar(Chan As String) As String
    
    Select Case Left(Chan, 1)
        Case "@", "+", "%", "$"
            NotOpChar = Mid(Chan, 2)
        Case Else
            NotOpChar = Chan
    End Select
    
End Function

Private Function FindNewChan() As Integer
    Dim tmpNum As Integer
    
    For tmpNum = 1 To 6000
        If Chan(tmpNum).InUse = False Then
            FindNewChan = tmpNum: Exit Function
        End If
    Next tmpNum
    
    FindNewChan = -1
End Function

Public Function FindChan(strChan As String) As Integer
    Dim tmpNum As Integer
    
    For tmpNum = 1 To 6000
        If Chan(tmpNum).Chan = strChan Then
            FindChan = tmpNum: Exit Function
        ElseIf Chan(tmpNum).Chan = "" Then
            FindChan = -1: Exit Function
        End If
    Next tmpNum
    FindChan = -1

End Function

Public Function RemoveChan(srcChan As String) As Boolean
    Dim tmpChan As Integer
    
    tmpChan = FindChan(srcChan)
    If tmpChan = -1 Then RemoveChan = False
        
    Chan(tmpChan).InUse = False
    
End Function

Public Function StoreInfo(strChan As String, strTopic As String) As Boolean
    Dim tmpNum As Integer
    
    tmpNum = FindChan(strChan)
    If tmpNum = -1 Then tmpNum = FindNewChan: Chan(tmpNum).InUse = True
    
    Chan(tmpNum).Chan = strChan
    Chan(tmpNum).Topic = strTopic
    
    StoreInfo = True
    
    Exit Function
    
End Function






Function WhoSetTopic(strChan As String) As String
    WhoSetTopic = Chan(FindChan(strChan)).TopicWhoSet
End Function

Sub StoreWhoSetTopic(strChan As String, strNick As String)
    Chan(FindChan(strChan)).TopicWhoSet = strNick
End Sub
