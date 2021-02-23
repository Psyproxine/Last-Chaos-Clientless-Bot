Attribute VB_Name = "ModColorRTB"
Option Explicit
'back color text
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_USER = &H400
Public Const SCF_SELECTION = &H1&
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Type POINTL
    x As Long
    y As Long
End Type
'Stop painting
'Modified Sendmessage to calculate current character under mouse
Public Declare Function SendMessageP Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As POINTL) As Long
'Hide the Caret
Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Const EM_SCROLL = &HB5
Private Const EM_GETLINECOUNT = &HBA
Public Const EM_CHARFROMPOS = &HD7
Dim ColorColl As Collection
Public charf As CHARFORMAT2
Public Const LF_FACESIZE = 32
Public Const CFM_BACKCOLOR = &H4000000
Public Type CHARFORMAT2
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60
    
    ' Additional stuff supported by RICHEDIT20
    wWeight As Integer            ' /* Font weight (LOGFONT value)      */
    sSpacing As Integer           ' /* Amount to space between letters  */
    crBackColor As Long        ' /* Background color                 */
    lLCID As Long               ' /* Locale ID                        */
    dwReserved As Long         ' /* Reserved. Must be 0              */
    sStyle As Integer            ' /* Style handle                     */
    wKerning As Integer            ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte     ' /* Underline type                   */
    bAnimation As Byte         ' /* Animated text like marching ants */
    bRevAuthor As Byte         ' /* Revision author index            */
    bReserved1 As Byte
End Type

Public Color As UserColor
Public Type UserColor
    normal As Integer
    bgText As Integer
    ctcp As Integer
    notice As Integer
    Action As Integer
    invite As Integer
    join As Integer
    kick As Integer
    mode As Integer
    nick As Integer
    notify As Integer
    part As Integer
    quit As Integer
    topic As Integer
    whois As Integer
    Server As Integer
End Type

Public Function HighLightWord(mForm As Form, mRTF As RichTextBox, mWord As String, mHighLightColor As Long, Optional HighlightAll As Boolean = False, Optional DontLock As Boolean) As Long
    Dim TempRTF As String
    Dim z As Long
    Dim st As Long
    Dim found As Long
    Dim HLNum As Long
    Dim RepairCtbl As Boolean
    Dim OldCol As Long
    Dim Oldst As Long
    Dim curvl As Long
    If InStr(1, mRTF.text, mWord) = 0 Then Exit Function 'If there is no item found then bail out
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    If Not DontLock Then LockWindowUpdate mForm.hwnd
    UnHighLight mForm, mRTF, , True, True 'remove existing highlighting
    GetColorTable mRTF 'Read the color table in the RTF code
    For z = 1 To ColorColl.Count
        If ColorColl(z) = mHighLightColor Then
            HLNum = z - 1 'The highlight color is in the Colortable
            Exit For
        End If
    Next
    If HLNum = 0 Then 'If the highlight color is not in the Colortable
        z = mRTF.Find(mWord, 0)
        If z < 0 Then 'If there is no item found then bail out
            If Not DontLock Then LockWindowUpdate 0
            Exit Function
        End If
        'Select the first letter of the first found item
        mRTF.SelLength = 0
        mRTF.SelStart = z
        mRTF.SelLength = 1
        Oldst = z
        OldCol = mRTF.SelColor 'Remember it's color so we can return it later
        mRTF.SelColor = mHighLightColor 'Change it's color to the Highlight color
        'This will cause the color to appear in the Colortable in the correct order
        GetColorTable mRTF 'Read the modified Colortable
        For z = 1 To ColorColl.Count 'Locate it's index in the Colortable
            If ColorColl(z) = mHighLightColor Then
                HLNum = z - 1 'Found it!
                Exit For
            End If
        Next
        RepairCtbl = True 'Fix the altered characters SelColor later
    End If
    Screen.MousePointer = 11
    If HighlightAll Then mRTF.SelStart = 0 'If we're doing all then start at the beginning
    st = mRTF.SelStart
    Do Until found = -1 'Now hunt for the chosen word
        'Place markers in the ".Text" property that we can
        'later locate in the ".textRTF" property
        found = mRTF.Find(mWord, st)
        If found = -1 Then Exit Do
        mRTF.SelStart = found + mRTF.SelLength
        mRTF.SelText = "%%%ZENDBB%%%" 'Mark the start position
        mRTF.SelStart = found
        mRTF.SelText = "%%%ZSTART%%%" 'Mark the end position
        st = mRTF.SelStart + Len(mWord) + 12
        If Not HighlightAll Then Exit Do
    Loop
    'Apply highlighting to the locations we have marked
    TempRTF = mRTF.TextRTF
    TempRTF = Replace(TempRTF, "%%%ZENDBB%%%", "\highlight0 ")
    TempRTF = Replace(TempRTF, "%%%ZSTART%%%", "\highlight" & HLNum & " ")
    mRTF.TextRTF = TempRTF
    If RepairCtbl Then 'Fix up any colors we may have changed
        mRTF.SelStart = Oldst
        mRTF.SelLength = 1
        mRTF.SelColor = OldCol
        mRTF.SelStart = 0
    End If
    mRTF.SelStart = 0
    Screen.MousePointer = 0
    SetScrollPos mForm, mRTF, curvl, True 'return the scroll position to what it was
    If Not DontLock Then LockWindowUpdate 0
End Function
Public Sub HighLightSelection(mForm As Form, mRTF As RichTextBox, mHighLightColor As Long, Optional DontLock As Boolean)
    'This is trickier than the other Highlight functions because
    'we have to allow for existing highlighting in various colors
    
    Dim TempRTF As String
    Dim SelStart As Long
    Dim SelEnd As Long
    Dim SelectedText As String
    Dim BeforeHL As String
    Dim AfterHL As String
    Dim FirstSelHL As String
    Dim LastSelHL As String
    Dim StartReplaceHL As String
    Dim EndReplaceHL As String
    Dim TempNum As String
    Dim z As Long
    Dim st As Long
    Dim found As Long
    Dim HLNum As Long
    Dim RepairCtbl As Boolean
    Dim OldCol As Long
    If mRTF.SelLength = 0 Then Exit Sub
    st = mRTF.SelStart
    found = mRTF.SelLength
    If Not DontLock Then LockWindowUpdate mForm.hwnd
    'Locate the chosen color in the Colortable
    GetColorTable mRTF
    For z = 1 To ColorColl.Count
        If ColorColl(z) = mHighLightColor Then
            HLNum = z - 1
            Exit For
        End If
    Next
    'If we didn't find it then modify the content
    'to place the color in the Colortable
    If HLNum = 0 Then
        mRTF.SelStart = st
        mRTF.SelLength = 1
        OldCol = mRTF.SelColor
        mRTF.SelColor = mHighLightColor
        GetColorTable mRTF
        For z = 1 To ColorColl.Count
            If ColorColl(z) = mHighLightColor Then
                HLNum = z - 1
                Exit For
            End If
        Next
        RepairCtbl = True
    End If
    mRTF.SelStart = st
    mRTF.SelLength = 0
    'Place markers around the selection
    mRTF.SelText = "%%%ZSTART%%%"
    mRTF.SelStart = st + found + 12
    mRTF.SelText = "%%%ZENDBB%%%"
    TempRTF = mRTF.TextRTF
    SelStart = InStr(1, TempRTF, "%%%ZSTART%%%")
    SelEnd = InStr(1, TempRTF, "%%%ZENDBB%%%") + 12
    'Place the selected text RTF code in a variable
    SelectedText = Mid(TempRTF, SelStart, SelEnd - SelStart)
    
    'inspect the preceding RTF code for any highlighting
    z = InStrRev(TempRTF, "\highlight", SelStart)
    'If there's highlighting, record its number(color index)
    If z <> 0 Then BeforeHL = Mid(TempRTF, z + 10, 1)
    
    'inspect the RTF code after the selection for any highlighting
    z = InStr(SelEnd, TempRTF, "\highlight")
    'If there's highlighting, record its number(color index)
    If z <> 0 Then AfterHL = Mid(TempRTF, z + 10, 1)
    
    'inspect the RTF code of the selection for any highlighting
    'find the first highlighting entry in the selection
    z = InStr(1, SelectedText, "\highlight")
    'If there's highlighting, record the first highlighting entry's number(color index)
    If z <> 0 Then FirstSelHL = Mid(SelectedText, z + 10, 1)
    'find the last highlighting entry in the selection
    z = InStrRev(SelectedText, "\highlight")
    If z <> 0 Then
        'if found record it's number(color index)
        LastSelHL = Mid(SelectedText, z + 10, 1)
        'Ok, we've got all the selections highlighting recorded
        'now we remove ALL highlighting from the selection
        Do
            TempNum = Mid(SelectedText, z + 10, 1)
            SelectedText = Replace(SelectedText, "\highlight" & TempNum & " ", "", , 1)
            z = InStr(1, SelectedText, "\highlight")
            If z = 0 Then Exit Do
        Loop
        'retuen the altered seleted RTF code back to the entire RTF code
        TempRTF = Left(TempRTF, SelStart - 1) & SelectedText & Right(TempRTF, Len(TempRTF) - SelEnd + 1)
    Else
        'If there was no highlighting in the selection then
        'use any highlighting data from BEFORE the selection
        If BeforeHL <> "" And BeforeHL <> "0" Then
            LastSelHL = BeforeHL
        End If
    End If
    
    'Now to replace our markers with the appropriate RTF tags according to
    'the highlighting tags found before/in/after the selection
    
    'Prepare the replacement strings
    StartReplaceHL = IIf(BeforeHL = "0" Or BeforeHL = "", "\highlight" & HLNum & " ", "\highlight0 " & "\highlight" & HLNum & " ")
    EndReplaceHL = IIf(LastSelHL = "0" Or LastSelHL = "", "\highlight0 ", "\highlight0 " & "\highlight" & LastSelHL & " ")
    'Do the replacing
    TempRTF = Replace(TempRTF, "%%%ZSTART%%%", StartReplaceHL)
    TempRTF = Replace(TempRTF, "%%%ZENDBB%%%", EndReplaceHL)
    'return the RTF code to the richtextbox
    mRTF.TextRTF = TempRTF
    'Return any adjustments back to what it was
    If RepairCtbl Then
        mRTF.SelStart = st
        mRTF.SelLength = 1
        mRTF.SelColor = OldCol
        mRTF.SelStart = 0
    End If
    mRTF.SelStart = st
    mRTF.Refresh
    If Not DontLock Then LockWindowUpdate 0
End Sub
Public Sub UnHighLight(mForm As Form, mRTF As RichTextBox, Optional mHighLightColor As Long, Optional AllHighlighting As Boolean, Optional DontLock As Boolean)
    Dim tmpRTF As String, z As Long
    Dim curvl As Long, HLNum As Long
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    If InStr(1, mRTF.TextRTF, "\highlight") = 0 Then Exit Sub
    If Not DontLock Then LockWindowUpdate mForm.hwnd
    If Not AllHighlighting Then
        'find the color index in the Colortable
        'of the desired color
        GetColorTable mRTF
        For z = 1 To ColorColl.Count
            If ColorColl(z) = mHighLightColor Then
                HLNum = z - 1
                Exit For
            End If
        Next
        'remove such entries from the RTF code
        tmpRTF = mRTF.TextRTF
        tmpRTF = Replace(tmpRTF, "\highlight" & HLNum & " ", "")
    Else
        'color doesn't matter - just remove any highlighting
        tmpRTF = Replace(mRTF.TextRTF, "\highlight0 ", "")
        tmpRTF = Replace(tmpRTF, "\highlight0", "")
        z = 1
        If InStr(1, tmpRTF, "\highlight") <> 0 Then
            Do
                If InStr(1, tmpRTF, "\highlight" & z) <> 0 Then
                    tmpRTF = Replace(tmpRTF, "\highlight" & z & " ", "")
                    tmpRTF = Replace(tmpRTF, "\highlight" & z & "", "")
                End If
                If InStr(1, tmpRTF, "\highlight") = 0 Then Exit Do
                z = z + 1
            Loop
        End If
    End If
    'return the adjusted RTF code to the richtextbox
    mRTF.TextRTF = tmpRTF
    SetScrollPos mForm, mRTF, curvl, True
    If Not DontLock Then LockWindowUpdate 0
End Sub
Private Sub GetColorTable(mRTF As RichTextBox)
    Dim z As Long, z1 As Long, temp As String, tmp() As String, tmpCol() As String
    Set ColorColl = New Collection
    ColorColl.Add 0
    'Parse the RTF code to extract the Colortable
    z = InStr(1, mRTF.TextRTF, "{\colortbl")
    If z = 0 Then
        Exit Sub
    Else
        'Parse the Colortable to extract the colors used
        z1 = InStr(z, mRTF.TextRTF, "}")
        If z1 = 0 Then
            Exit Sub
        Else
            temp = Mid(mRTF.TextRTF, z, z1 - z + 1)
            tmp = Split(temp, ";")
            For z = 1 To UBound(tmp) - 1
                If tmp(z) <> "" Then
                    If Left(tmp(z), 1) = "\" Then tmp(z) = Right(tmp(z), Len(tmp(z)) - 1)
                    tmpCol = Split(tmp(z), "\")
                    'Dump the colors found into a collection
                    ColorColl.Add RGB(Val(Right(tmpCol(0), Len(tmpCol(0)) - 3)), Val(Right(tmpCol(1), Len(tmpCol(1)) - 5)), Val(Right(tmpCol(2), Len(tmpCol(2)) - 4)))
                End If
            Next
        End If
    End If
End Sub
Public Function HighLightComments(mForm As Form, mRTF As RichTextBox, mHighLightColor As Long, mCommentCharacter As String, Optional DontLock As Boolean)
    Dim temp As String
    Dim tmp() As String
    Dim tmp2() As String
    Dim z As Long
    Dim zStart As Long
    Dim HLNum As Long
    Dim RepairCtbl As Boolean
    Dim HasChanged As Boolean
    Dim curvl As Long
    'Any comments
    z = InStr(1, mRTF.text, mCommentCharacter)
    'None found - so bail out here
    If z = 0 Then Exit Function
    'Remember the scroll position
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    If Not DontLock Then LockWindowUpdate mForm.hwnd
    'find the highlight color in the Colortable
    GetColorTable mRTF
    For z = 1 To ColorColl.Count
        If ColorColl(z) = mHighLightColor Then
            HLNum = z - 1
            Exit For
        End If
    Next
    'If the highlight color was absent from the Colortable
    'adjust the content so it will appear there
    If HLNum = 0 Then
        z = mRTF.Find(mCommentCharacter, 0)
        If z < 0 Then
            If Not DontLock Then LockWindowUpdate 0
            Exit Function
        End If
        mRTF.SelLength = 0
        mRTF.SelStart = z
        mRTF.SelLength = 1
        mRTF.SelColor = mHighLightColor
        GetColorTable mRTF
        For z = 1 To ColorColl.Count
            If ColorColl(z) = mHighLightColor Then
                HLNum = z - 1
                Exit For
            End If
        Next
        RepairCtbl = True
    End If
    Screen.MousePointer = 11
    'split the code into lines
    tmp = Split(mRTF.TextRTF, vbCrLf)
    For z = 0 To UBound(tmp)
        'try exclude some RTF characters
        If Left(tmp(z), 9) <> "\par \'af" Then
            'find the first occurence of the comment character in each line
            zStart = InStr(1, tmp(z), mCommentCharacter)
            If zStart <> 0 Then
                'if we found any
                'are there any quotes in the line?
                If InStr(1, tmp(z), Chr(34)) Then
                    temp = Left(tmp(z), zStart - 1)
                    'is the quote BEFORE the comment character?
                    If InStr(1, temp, Chr(34)) Then
                        'if it is, how many quotes are there?
                        tmp2 = Split(temp, Chr(34))
                        'if there's an even number of quotes
                        'then our character is valid
                        If UBound(tmp2) Mod 2 = 0 Then
                            tmp(z) = Left(tmp(z), zStart - 1) & "\highlight" & HLNum & " " & Right(tmp(z), Len(tmp(z)) - zStart + 1) & "\highlight0 "
                        End If
                    Else 'if no quotes then our character is valid
                        tmp(z) = Left(tmp(z), zStart - 1) & "\highlight" & HLNum & " " & Right(tmp(z), Len(tmp(z)) - zStart + 1) & "\highlight0 "
                    End If
                Else 'if no quotes then our character is valid
                    tmp(z) = Left(tmp(z), zStart - 1) & "\highlight" & HLNum & " " & Right(tmp(z), Len(tmp(z)) - zStart + 1) & "\highlight0 "
                End If
            End If
        End If
    Next
    'put the code back together
    mRTF.TextRTF = join(tmp, vbCrLf)
    mRTF.Refresh
    'repair ant adjustments made
    If RepairCtbl Then
        Do
            z = mRTF.Find(mCommentCharacter, z)
            If z < 0 Then Exit Do
            If mRTF.SelColor = mHighLightColor Then
                mRTF.SelColor = vbBlack
                Exit Do
            End If
            mRTF.SelColor = vbBlack
            z = z + 1
        Loop
    End If
    mRTF.SelStart = 0
    Screen.MousePointer = 0
    'return the scrollbars back to there original position
    SetScrollPos mForm, mRTF, curvl, True
    If Not DontLock Then LockWindowUpdate 0
End Function
Public Sub SetScrollPos(mForm As Form, mRTF As RichTextBox, mPos As Long, Optional DontLock As Boolean)
    Dim CurLineCount As Long, curvl As Long, lastvl As Long
    'how many lines?
    CurLineCount = SendMessage(mRTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
    'what's the current top line?
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    If Not DontLock Then LockWindowUpdate mForm.hwnd
    'use a "PageUp" or "PageDown" for a few big jumps to get close to our target line
    If mPos < curvl Then
        Do Until curvl < mPos
            SendMessage mRTF.hwnd, EM_SCROLL, 2, 0
            curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
            If curvl = 0 Or curvl = CurLineCount Then Exit Do
        Loop
    Else
        Do Until curvl > mPos
            SendMessage mRTF.hwnd, EM_SCROLL, 3, 0
            curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
            If curvl = 0 Or curvl = CurLineCount Or lastvl = curvl Then Exit Do
            lastvl = curvl
        Loop
    End If
    'Now do some fine adjustment line by line to get
    'it exactly right
    Do Until curvl = mPos
        If mPos < curvl Then
            SendMessage mRTF.hwnd, EM_SCROLL, 0, 0
        Else
            SendMessage mRTF.hwnd, EM_SCROLL, 1, 0
        End If
        curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
        If curvl = 0 Or curvl = CurLineCount Or lastvl = curvl Then
            If curvl = 0 Then SendMessage mRTF.hwnd, EM_SCROLL, 0, 0
            Exit Do
        End If
        lastvl = curvl
    Loop
    If Not DontLock Then LockWindowUpdate 0
End Sub
Public Sub SelectAll(mForm As Form, mRTF As RichTextBox, Optional DontLock As Boolean)
    'Standard select all, but adjust the scrollbars
    'so we dont end up at the bottom of the page
    Dim curvl As Long, st As Long
    If Not DontLock Then LockWindowUpdate mForm.hwnd
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    mRTF.SelStart = 0
    mRTF.SelLength = Len(mRTF.text)
    SetScrollPos mForm, mRTF, curvl, True
    mRTF.SetFocus
    If Not DontLock Then LockWindowUpdate 0
End Sub
Public Sub SelectAbove(mForm As Form, mRTF As RichTextBox, Optional DontLock As Boolean)
    Dim curvl As Long, st As Long
    If Not DontLock Then LockWindowUpdate mForm.hwnd
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    st = mRTF.SelStart
    mRTF.SelStart = 0
    mRTF.SelLength = st
    SetScrollPos mForm, mRTF, curvl, True
    mRTF.SetFocus
    If Not DontLock Then LockWindowUpdate 0
End Sub
Public Sub SelectBelow(mForm As Form, mRTF As RichTextBox, Optional DontLock As Boolean)
    Dim curvl As Long, st As Long
    If Not DontLock Then LockWindowUpdate mForm.hwnd
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    st = mRTF.SelStart + mRTF.SelLength
    mRTF.SelStart = st
    mRTF.SelLength = Len(mRTF.text) - st
    SetScrollPos mForm, mRTF, curvl, True
    mRTF.SetFocus
    If Not DontLock Then LockWindowUpdate 0
End Sub
Public Sub StringColor(mForm As Form, mRTF As RichTextBox, mColor As Long, Optional DontLock As Boolean)
    Dim st As Long, sl As Long, FT As Long
    Dim curvl As Long, OldStart As Long
    'Use standard ".selcolor" to change the colors
    'It would be quicker to use Regular Expressions to
    'edit the actual RTF code, but this is pretty quick
    'for this demo
    If Not DontLock Then LockWindowUpdate mForm.hwnd
    OldStart = mRTF.SelStart
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    With mRTF
        If Len(.text) = 0 Then Exit Sub
        FT = .Find(Chr(34), 0)
        If FT <> -1 Then
            Do Until .SelStart + .SelLength >= Len(.text)
                DoEvents
                .SelStart = .SelStart + 1
                .Span Chr(34), True, True
                st = .SelStart
                sl = .SelLength
                If InStr(1, .SelText, vbCrLf) = 0 Then
                    .SelStart = st - 1
                    .SelLength = 1
                    .SelLength = sl + 2
                    .SelColor = mColor
                End If
                .SelLength = 0
                .SelStart = st + sl + 2
                FT = .Find(Chr(34), .SelStart)
                If FT = -1 Then Exit Do
            Loop
        End If
        .SelStart = 0
    End With
    mRTF.SelStart = OldStart
    SetScrollPos mForm, mRTF, curvl, True
    If Not DontLock Then LockWindowUpdate 0
End Sub

Public Function GetVBKeyWords() As String()
    GetVBKeyWords = Split("#Const|#Else|#ElseIf|#End|#If|Alias|Alias|And|As|Base|Binary|Boolean|Byte|ByVal|Call|Case|CBool|CByte|CCur|CDate|CDbl|CDec|CInt|CLng|Close|Compare|Const|CSng|CStr|Currency|CVar|CVErr|Decimal|Declare|DefBool|DefByte|DefCur|DefDate|DefDbl|DefDec|DefInt|DefLng|DefObj|DefSng|DefStr|DefVar|Dim|Do|Double|Each|Else|ElseIf|End|Enum|Eqv|Erase|Error|Exit|Explicit|False|For|Function|Get|Global|GoSub|GoTo|If|Imp|In|Input|Input|Integer|Is|LBound|Let|Lib|Like|Line|Lock|Long|Loop|LSet|Name|New|Next|Not|Object|On|Open|Option|Or|Output|Print|Private|Property|Public|Put|Random|Read|ReDim|Resume|Return|RSet|Seek|Select|Set|Single|Spc|Static|String|Stop|Sub|Tab|Then|Then|True|Type|UBound|Unlock|Variant|Wend|While|With|Xor|Nothing|To", "|")
End Function

Public Function GetCurrentPosition(mRTF As RichTextBox) As Long
    GetCurrentPosition = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
End Function

