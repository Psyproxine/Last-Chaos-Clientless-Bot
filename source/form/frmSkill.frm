VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSkill 
   Caption         =   "Skill"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   Icon            =   "frmSkill.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvSkill 
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
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Lv"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Detail"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmSkill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
lvSkill.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
