Attribute VB_Name = "ModRTB"
Option Explicit
' Changed this module around a little. Got rid of those old ugly Global declarations
' removed a couple things that are unused.


Public Type RGB
    r As Integer
    g As Integer
    B As Integer
End Type

' These constants here are the characters that IRC uses for each respective type of text.
Public Declare Function GetTextCharset Lib "gdi32" (ByVal hdc As Long) As Long

Public Const strBold = ""
Public Const strUnderline = ""
Public Const strReverse = ""
Public Const strColor = ""
Public Const intIndent = 0
Public Const intFontSize = 8
Public Const Cancel = 15

'variables
' Depending on what colors your rtf are, set these accordingly. you might even want to make
' them into variables so that they can be changed.
Public Const lngForeColor& = vbBlack
Public Const lngBackColor& = &H80000000

Public Colors(99) As RGB
Public DefinedColors As Integer


Public Sub msIRC(Stri As String)
    'DoColor frmChat.rtbChat, Chr(3) & "14[" & Chr(3) & "15" & Time & Chr(3) & "14] " & Stri
    PutText frmIRC.rtbirc, Chr(3) & "14[" & Chr(3) & "15" & time & Chr(3) & "14] " & Stri & vbCrLf
End Sub

Public Sub Chat(Stri As String)
    'DoColor frmChat.rtbChat, Chr(3) & "14[" & Chr(3) & "15" & Time & Chr(3) & "14] " & Stri
    PutText frmChat.rtbChat, Chr(3) & "14[" & Chr(3) & "15" & time & Chr(3) & "14] " & Stri & vbCrLf
End Sub

Public Sub Stat(Stri As String)
If Not frmStat.rtbStat.Visible Then
    frmStat.rtbStat.Visible = True
    frmStat.wbLoad.Visible = False
End If
    PutText frmStat.rtbStat, Chr(3) & "14[" & Chr(3) & "15" & time & Chr(3) & "14] " & Stri & vbCrLf
    'DoColor frmStat.rtbStat, Chr(3) & "14[" & Chr(3) & "15" & Time & Chr(3) & "14] " & Stri
End Sub

Public Function ColorTable() As String
    Dim i As Integer, strTable As String
    Dim r As Integer, B As Integer, g As Integer
    strTable = "{\colortbl ;"
    'MsgBox DefinedColors
    For i = 0 To 15
        r = Colors(i).r
        g = Colors(i).g
        B = Colors(i).B
        strTable = strTable & "\red" & r & "\green" & g & "\blue" & B & ";"
    Next i
    strTable = strTable & "}"
    ColorTable = strTable
End Function

Public Sub LoadColors()
    Dim i As Integer, strFile As String
    strFile = App.path & "\colors.inf"
    
    On Error GoTo errorHandler
    DefinedColors = 0
    Open strFile For Input As #1
        Do
            Input #1, Colors(DefinedColors).r
            Input #1, Colors(DefinedColors).g
            Input #1, Colors(DefinedColors).B
            DefinedColors = DefinedColors + 1
        Loop Until EOF(1) Or DefinedColors >= 99
    Close #1
    Exit Sub
    
errorHandler:
    Select Case Err
        Case 53:
            MsgBox "(ERROR 53) The color information file does not exist, and therefore could not be loaded.  Please check the manual on how to fix this problem.  The program will now exit.", vbCritical
            End
        Case 76:
            MsgBox "(ERROR 76) The color information file does not exist, and therefore could not be loaded.  Please check the manual on how to fix this problem.  The program will now exit.", vbCritical
            End
        Case 62:
            MsgBox "(ERROR 62) The color information file is not complete.  Please check the manual for more information.", vbCritical
            End
        Case Else:
            MsgBox "An unknown error has occured.  The following information has been obtained, but is not documented in the manual." & vbCrLf & vbCrLf & "Error #" & Err & " : " & Error, vbCritical
    End Select

End Sub

Public Function RAnsiColor(lngColor As Long) As Integer

    Dim i As Integer
    For i = 0 To 15
        If RGB(Colors(i).r, Colors(i).g, Colors(i).B) = lngColor Then
            RAnsiColor = i
            Exit Function
        End If
    Next i
    
    If lngColor = lngForeColor Then
        RAnsiColor = 1
    ElseIf lngColor = lngBackColor Then
        RAnsiColor = 0
    Else
        RAnsiColor = 99
    End If
    
End Function


Public Function Red(ByVal Color As Long)
    Red = Color Mod 256
End Function


Public Function Green(ByVal Color As Long)
    Green = (Color / 256) Mod 256
End Function


Public Function Blue(ByVal Color As Long)
    Blue = Color / 65536
End Function

Public Sub SetRGB(lngColor As Long, ByRef r As Integer, ByRef g As Integer, ByRef B As Integer)
    
    r = Red(lngColor)
    g = Green(lngColor)
    B = Blue(lngColor)
    Exit Sub
End Sub


' This module holds the required functions for everything other than modColors.
' I gathered all of these subs and functions from VCV's sIRC project that had
' them. I assume they are the same since the PutText function was the same in his
' sIRC project (to the letter and comment). I've gone through the code and commented
' as much as I can but I don't have a knowledge of RTF Formatting and don't care to
' learn the ins and outs of this very dead format. (But not dead for us VB6
' Programmers RIGHT?!) Either way this works like he says.
'  If you do not wish to keep colors.inf you can hardcode the variables yourself
' into the LoadColors sub.
' After I went through and commented the coding I ran it through Morgan Haueisen's
' Code Formatter (Simple) v1.4.4 project to pretty up what I or VCV did.


Function ValidColorCode(strCode As String) As Boolean
 ' checks strCode to see if it is a valid IRC color code.
 ' It returns true if it is, false if not.
    If strCode = "" Then ValidColorCode = True: Exit Function
    Dim c1 As Integer, c2 As Integer
    If strCode Like "" Or _
       strCode Like "#" Or _
       strCode Like "##" Or _
       strCode Like "#,#" Or _
       strCode Like "##,#" Or _
       strCode Like "#,##" Or _
       strCode Like "#," Or _
       strCode Like "##," Or _
       strCode Like "##,##" Or _
       strCode Like ",#" Or _
       strCode Like ",##" Then
        Dim strCol() As String
        strCol = Split(strCode, ",")
        '
        If UBound(strCol) = -1 Then
            ValidColorCode = True
        ElseIf UBound(strCol) = 0 Then
            If strCol(0) = "" Then strCol(0) = 0
            If Int(strCol(0)) >= 0 And Int(strCol(0)) <= 99 Then
                ValidColorCode = True
                Exit Function
            Else
                ValidColorCode = False
                Exit Function
            End If
        Else
            If strCol(0) = "" Then strCol(0) = lngForeColor
            If strCol(1) = "" Then strCol(1) = 0
            c1 = Int(strCol(0))
            c2 = Int(strCol(1))
            If Int(c2) < 0 Or Int(c2) > 99 Then
                ValidColorCode = False
                Exit Function
            Else
                ValidColorCode = True
                Exit Function
            End If
        End If
        ValidColorCode = True
        Exit Function
    Else
        ValidColorCode = False
        Exit Function
    End If
End Function

Function LeftOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    
    intPos = InStr(strData, strDelim)
    If intPos Then
        LeftOf = Left$(strData, intPos - 1)
    Else
        LeftOf = strData
    End If
End Function

Function LeftR(strData As String, intMin As Integer)
    
    On Error Resume Next
    LeftR = Left$(strData, Len(strData) - intMin)
End Function

Function RightOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    intPos = InStr(strData, strDelim)
    
    If intPos Then
        RightOf = Mid$(strData, intPos + 1, Len(strData) - intPos)
    Else
        RightOf = strData
    End If
End Function

Function RightR(strData As String, intMin As Integer)
    On Error Resume Next
    RightR = Right$(strData, Len(strData) - intMin)
End Function

Public Sub PutText(rtf As RichTextBox, strData As String)
On Error Resume Next
    
    If strData = "" Then Exit Sub ' Obviously if there is no data we don't
                                  ' want to waste time crunching nothing!
    
    ' Variable declarations local to this method
    Dim i As Long, Length As Integer, strChar As String, strBuffer As String
    Dim clr As Integer, bclr As Integer, dftclr As Integer, strRTFBuff As String
    Dim bbbold As Boolean, bbunderline As Boolean, bbreverse As Boolean, strTmp As String
    Dim lngFC As String, lngBC As String, lngStart As Long, lngLength As Long, strPlaceHolder As String
    
    lngStart = rtf.SelStart   'Starting position to use
    lngLength = rtf.SelLength 'Length of text already in the box
    
    
    '* if not inialized, set font, intialiaze
    Dim btCharSet As Long
    Dim strRTF As String
    
'/*   Uncomment this block if you want the first line to be a size or two bigger than intFontSize.
     ' I couldn't figure out for the life of my why this block of code makes the first line bigger... it's
     ' the same.
   ' If rtf.Tag <> "init'd" Then
   '     rtf.Tag = "init'd" ' tag the boxes to reflect the inintialization.
   '     strFontName = rtf.Font.Name ' sets strFontName to the boxes font
   '     'rtf.Parent.FontName = strFontName ' sets the parent font name (I'm clueless to this one)
   '     btCharSet = GetTextCharset(rtf.Parent.hdc) ' Retrieve the character set from the box
   '     strRTF = "" ' Make sure this variable is empty
   '     ' so far as I can tell this next section defines the font, available colors, and the font size.
   '     strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
   '     strRTF = strRTF & ColorTable & vbCrLf
   '     strRTF = strRTF & "{\stylesheet{ Normal;}{\s1 heading 1;}{\s2 heading 2;}}" & vbCrLf
   '     strRTF = strRTF & "\viewkind4\uc1\pard\cf0\f0\fs" & CInt(intFontSize * 2) & "\fi-" & intIndent & "\li" & intIndent & "\ffprot1 "
   '     strPlaceHolder = " \n"
   '     For i = 0 To DefinedColors - 1
   '         strRTF = strRTF & "\cf" & i & strPlaceHolder
   '     Next
   '     strRTF = strRTF & "}" ' closing bracket for the RTF
   '     rtf.TextRTF = strRTF  ' set the RTF Text in the box to our temp string.
   ' End If ' repeat of the above, with a few less things
'*/
        btCharSet = GetTextCharset(rtf.Parent.hdc)
        strRTF = ""
        strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & frmChat.rtbChat.Font.name & ";}}" & vbCrLf
        strRTF = strRTF & ColorTable & vbCrLf
        strRTF = strRTF & "{\stylesheet{ Normal;}{\s1 heading 1;}{\s2 heading 2;}}" & vbCrLf
        strRTF = strRTF & "\viewkind4\uc1\pard\cf0\f0\fs" & CInt(intFontSize * 2) & "\fi-" & intIndent & "\li" & intIndent & "\ffprot1 "
    
    strRTFBuff = "\b0\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1 & "\i0\ulnone "
  
    
    dftclr = RAnsiColor(lngForeColor)
    'add a custom timestamp here. Make sure you have a variable named bTimestamp and
    ' it is set to true. This one adds "[00:00]" Personally I like "[mm:dd][hh:mm:ss]"
        
    Length = Len(strData) ' gets the length of data input
    i = 1 ' sets i to 1
    
    Do ' starts the loop
        strChar = Mid$(strData, i, 1) ' Picks out a single character based on the int i.
        Select Case strChar ' Select Case to see what strChar is
            Case Chr(Cancel)    'cancel code
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                lngFC = CStr(RAnsiColor(lngForeColor))
                lngBC = CStr(RAnsiColor(lngBackColor))
                strRTFBuff = strRTFBuff & strBuffer & "\b0\ul0\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1
                strBuffer = ""
                i = i + 1
            Case strBold  ' bold code
                bbbold = Not bbbold
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer & "\b"
                If bbbold = False Then strRTFBuff = strRTFBuff & "0"
                strBuffer = ""
                i = i + 1
            Case strUnderline ' Underline code
                bbunderline = Not bbunderline
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer & "\ul"
                If bbunderline = False Then strRTFBuff = strRTFBuff & "none"
                strBuffer = ""
                i = i + 1
            Case strReverse  ' reverse code
                bbreverse = Not bbreverse
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " " ' & strBuffer & "\"
                If bbreverse = False Then
                    If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                    strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1
                Else
                    If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                    strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngBackColor) + 1 & "\highlight" & RAnsiColor(lngForeColor) + 1
                End If
                
                strBuffer = ""
                i = i + 1
            Case strColor  ' color code
                
                strTmp = "" ' makes sure the tmp var is empty
                i = i + 1

                Do Until Not ValidColorCode(strTmp) Or i > Length  ' This goes through and gets individual characters
                    strTmp = strTmp & Mid$(strData, i, 1)          ' and adds them together until it gets a valid color
                    i = i + 1                                      ' code
                Loop
                
                strTmp = LeftR(strTmp, 1) ' Everytime it comes up with a valid code there is the first letter of whatever
                If strTmp = "" Then       ' was after it, so we remove it.
                    lngFC = CStr(RAnsiColor(lngForeColor)) ' This executes if there are no color codes
                    lngBC = CStr(RAnsiColor(lngBackColor))
                Else ' there must be color codes! YaY!
                    lngFC = LeftOf(strTmp, ",") ' sets the forecolor to the first number
                    lngFC = CStr(CInt(lngFC))   ' Makes sure its all proper
                    If InStr(strTmp, ",") Then  ' if "," is in the string
                        lngBC = RightOf(strTmp, ",") ' then set the background color to what is right of ,
                        If lngBC <> "" Then lngBC = CStr(CInt(lngBC)) Else lngBC = CStr(RAnsiColor(lngBackColor))
                        ' ^ if there is no lngBC code then set it to the normal background color
                    Else ' There is no , or back color code
                        lngBC = ""
                    End If
                End If
                
                If lngFC = "" Then lngFC = CStr(lngForeColor) ' if there wasn't a foreground color above, set it to normal
                lngFC = Int(lngFC) + 1 ' Adds one number to whatever number is set (gotta love the 0 index bull MS pulled on us)
                If lngBC <> "" Then lngBC = Int(lngBC) + 1 ' if BC isn't empty, then increment it's number also
                                
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " " ' adds a space to the end of strRTFBuff
                strRTFBuff = strRTFBuff & strBuffer ' combines strRTFBuffer and our plain inputed text
                strRTFBuff = strRTFBuff & "\cf" & lngFC ' Adds the RTF color code and our forecolor
                If lngBC <> "" Then strRTFBuff = strRTFBuff & "\highlight" & lngBC ' if the bc isn't empty then add the part to
                                                                                   ' change the background color
                
                i = i - 1 ' decrement i
                strBuffer = "" ' empty the buffer
                If i >= Length Then GoTo TheEnd ' if it reached the end of the input, then don't loop and go to the end
                
            Case Else  ' If strChar doesn't match up to the special codes, it adds it to the buffer and goes on.
                Select Case strChar
                Case "}", "{", "\" ' Special characters used by RTF need to be escaped before plain text use.
                    strBuffer = strBuffer & "\" & strChar ' escapes the char
                Case Else ' if it doesn't match the special chars just add it to the buffer and go on.
                    strBuffer = strBuffer & strChar
                End Select
                i = i + 1 ' increment i
        End Select
        
    Loop Until i > Length
    
   
TheEnd:
    If strBuffer <> "" Then
        strRTFBuff = strRTFBuff & " " & strBuffer  ' Combines the last part of the buffer with the RTF text
    End If
    
    'Clipboard.SetText rtf.TextRTF & vbCrLf & vbCrLf & vbCrLf & strRTF & strRTFBuff & vbCrLf & "}", 1
    
    strRTFBuff = strRTFBuff & vbCrLf ' adds a end line
    rtf.SelStart = Len(rtf.text)     ' selects the last character as the starting point to add text
    rtf.SelLength = 0                ' makes sure nothing is selected
    'If rtf.Text = "" Then
    '    rtf.SelRTF = strRTF & strRTFBuff & vbCrLf & " }" & vbCrLf  ' if it's empty we don't need \par
    'Else
        rtf.SelRTF = strRTF & strRTFBuff & "\par" & "}" & vbCrLf  ' if there is text already we need to add
    'End If
    If Len(rtf.text) > 20000 Then
        rtf.SelStart = 0
        rtf.SelLength = 15000
        rtf.SelText = ""
    End If
    rtf.SelStart = Len(rtf.text)
    rtf.SelLength = 0
End Sub

