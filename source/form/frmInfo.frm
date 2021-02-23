VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfo 
   Caption         =   "Info"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvInfo 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8421504
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   6174
      EndProperty
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
lvInfo.ListItems.Add 1, "Name", "Name"
lvInfo.ListItems.Add 2, "Guild", "Guild"
lvInfo.ListItems(2).SubItems(2) = "Class"
lvInfo.ListItems.Add 3, "Level", "Level"
'lvInfo.ListItems(3).SubItems(2) = "Killed"
lvInfo.ListItems.Add 4, "", ""
lvInfo.ListItems.Add 5, "HP", "HP"
lvInfo.ListItems(5).SubItems(2) = "MP"
lvInfo.ListItems.Add 6, "HATE", "SP"
lvInfo.ListItems(6).SubItems(2) = "EXP"
lvInfo.ListItems.Add 7, "", ""
lvInfo.ListItems.Add 8, "STR", "STR"
lvInfo.ListItems(8).SubItems(2) = "ATK"
lvInfo.ListItems.Add 9, "AGI", "AGI"
lvInfo.ListItems(9).SubItems(2) = "MTK"
lvInfo.ListItems.Add 10, "INT", "INT"
lvInfo.ListItems(10).SubItems(2) = "DEF"
lvInfo.ListItems.Add 11, "VIT", "VIT"
lvInfo.ListItems(11).SubItems(2) = "MEF"
lvInfo.ListItems.Add 12, "", ""
lvInfo.ListItems.Add 13, "Money", "Money"
Dim i As Integer
For i = 1 To 13
      '  frmInfo.lvInfo.ListItems(i).Bold = True
        frmInfo.lvInfo.ListItems(i).ForeColor = &HFF6600
        If i <> 1 And i <> 3 And i <> 4 And i <> 7 And i <> 12 And i <> 13 Then
       '     frmInfo.lvInfo.ListItems(i).ListSubItems(2).Bold = True
            frmInfo.lvInfo.ListItems(i).ListSubItems(2).ForeColor = &HFF6600
        End If
Next
End Sub

Private Sub Form_Resize()
lvInfo.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
