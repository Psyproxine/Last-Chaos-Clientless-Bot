VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDebug 
   Caption         =   "Debug"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   9060
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3330
      TabIndex        =   6
      Top             =   600
      Width           =   4755
   End
   Begin VB.OptionButton Op 
      Caption         =   "ส่งข้อมูลโดยเพิ่ม เปิด/ความยาว/ปิด โดยอัตโนมัติ"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   4
      Top             =   30
      Width           =   3915
   End
   Begin VB.OptionButton Op 
      Caption         =   "ส่งข้อมูลตามที่กรอก"
      Height          =   255
      Index           =   0
      Left            =   3420
      TabIndex        =   3
      Top             =   30
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.TextBox txtPacket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3330
      TabIndex        =   2
      Text            =   "AA 55 06 00 08 00 10 00 00 00 55 AA"
      Top             =   300
      Width           =   4755
   End
   Begin VB.ListBox lstpacket 
      Appearance      =   0  'Flat
      Height          =   3150
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   3315
   End
   Begin RichTextLib.RichTextBox rtbDebug 
      Height          =   2055
      Left            =   3600
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3625
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmDebug.frx":0000
   End
   Begin Aggressive.chameleonButton chbSend 
      Height          =   285
      Left            =   8100
      TabIndex        =   5
      Top             =   300
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   503
      BTYPE           =   5
      TX              =   "Send"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   0
      FOCUSR          =   0   'False
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   0
      MPTR            =   1
      MICON           =   "frmDebug.frx":0092
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Aggressive.chameleonButton chSearch 
      Height          =   285
      Left            =   8100
      TabIndex        =   7
      Top             =   600
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   503
      BTYPE           =   5
      TX              =   "Search"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   0
      FOCUSR          =   0   'False
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   0
      MPTR            =   1
      MICON           =   "frmDebug.frx":00AE
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chbSend_Click()
If Trim(txtPacket.Text) = "" Then Exit Sub
If Op(0).Value = True Then
SendPacket h2s(Trim(txtPacket.Text)), 0
Else
SendPacket h2s(Trim(txtPacket.Text)), 1
End If
End Sub

Private Sub chSearch_Click()
If Trim(txtSearch) <> "" And Trim(rtbDebug.Text) <> "" Then HighLightWord Me, rtbDebug, txtSearch, &HFF, True, False
End Sub

Private Sub Form_Resize()
On Error Resume Next
lstpacket.Move 0, 0, lstpacket.Width, Me.ScaleHeight
rtbDebug.Move lstpacket.Width + 30, txtSearch.Top + txtSearch.Height + 30, Me.ScaleWidth - lstpacket.Width - 30, Me.ScaleHeight - txtSearch.Top - txtSearch.Height - 30
End Sub

Public Function h2s(Data As String) As String
Dim t() As String, s As String, x As Integer
Data = Replace(Data, vbCrLf, "")
    t = Split(Data, " ")
    s = ""
    For x = 0 To UBound(t)
      If Trim(t(x)) <> "" Then s = s & CStr(Chr("&H" & t(x)))
    Next
h2s = s
End Function

Private Sub Form_Unload(Cancel As Integer)
ReDim tmppacket(0)
End Sub

Private Sub lstpacket_Click()
Dim packet As String, tstr As String, sstr As String, x As Integer, ts As Boolean, ss As Boolean
tstr = " "
sstr = " "
rtbDebug.Text = ""
packet = tmppacket(lstpacket.ListIndex)
    For x = 1 To Len(packet)
       If Asc(Mid(packet, x, 1)) < 16 Then tstr = tstr + "0"
       tstr = tstr + Hex(Asc(Mid(packet, x, 1))) + " "
       sstr = sstr + hex2string(Trim(Mid$(packet, x, 1))) + " "
       
       If x < Len(packet) And x Mod 16 = 0 And Hex(Asc(Mid(packet, x, 1))) = "AA" And x > 1 Then
            If Hex(Asc(Mid(packet, x + 1, 1))) = "55" And Hex(Asc(Mid(packet, x - 1, 1))) <> "55" Then
                tstr = Replace(tstr, "AA", Chr(3) & "4,01AA" & Chr(3))
                ts = True
            End If
       End If
       
       If x < Len(packet) And x Mod 16 = 0 And Hex(Asc(Mid(packet, x, 1))) = "55" And x > 1 Then
            If Hex(Asc(Mid(packet, x + 1, 1))) = "AA" And Hex(Asc(Mid(packet, x - 1, 1))) <> "AA" Then
                tstr = Replace(tstr, "55", Chr(3) & "4,0155" & Chr(3))
            End If
       End If
       
       If x < Len(packet) And x Mod 16 = 1 And x > 16 And Hex(Asc(Mid(packet, x, 1))) = "AA" Then
            If Hex(Asc(Mid(packet, x - 1, 1))) = "55" And Hex(Asc(Mid(packet, x + 1, 1))) <> "55" Then
                tstr = Replace(tstr, "AA", Chr(3) & "4,01AA" & Chr(3))
            End If
       End If
       
       If x < Len(packet) And x Mod 16 = 0 And Hex(Asc(Mid(packet, x, 1))) = "55" And x > 1 Then
            If Hex(Asc(Mid(packet, x - 1, 1))) = "AA" And Hex(Asc(Mid(packet, x + 1, 1))) = "AA" Then
                tstr = Replace(tstr, "55", Chr(3) & "4,0155" & Chr(3))
            End If
       End If
       
       If x = Len(packet) And x Mod 16 = 1 And x > 16 And Hex(Asc(Mid(packet, x, 1))) = "AA" Then
            If Hex(Asc(Mid(packet, x - 1, 1))) = "55" Then
                tstr = Replace(tstr, "AA", Chr(3) & "4,01AA" & Chr(3))
            End If
       End If
       If x = Len(packet) And x Mod 16 = 1 And x > 16 And Hex(Asc(Mid(packet, x, 1))) = "55" Then
            If Hex(Asc(Mid(packet, x - 1, 1))) = "AA" Then
                tstr = Replace(tstr, "55", Chr(3) & "4,0155" & Chr(3))
            End If
       End If
    
       
       If x Mod 16 = 0 Then
            tstr = tstr & "     " & Chr(9) & sstr
            tstr = Replace(tstr, "AA 55", Chr(3) & "4,01AA 55" & Chr(3))
            tstr = Replace(tstr, "55 AA", Chr(3) & "4,0155 AA" & Chr(3))
            PutText rtbDebug, " " & tstr
            tstr = " "
            sstr = " "
        End If
    Next
    If x Mod 16 <> 0 Then
            tstr = tstr & "     " & Chr(9) & sstr
            tstr = Replace(tstr, "AA 55", Chr(3) & "4,01AA 55" & Chr(3))
            tstr = Replace(tstr, "55 AA", Chr(3) & "4,0155 AA" & Chr(3))
            PutText rtbDebug, tstr
    End If
        rtbDebug.SelStart = 0
End Sub
