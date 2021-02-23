VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPermission 
   Caption         =   "Permission"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "Check User Permission"
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4965
      Begin MSComctlLib.ListView lvReport 
         Height          =   2385
         Left            =   60
         TabIndex        =   5
         Top             =   855
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   4207
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.TextBox txtPer_Pass 
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   990
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   540
         Width           =   1755
      End
      Begin VB.TextBox txtPer_User 
         Height          =   285
         Left            =   990
         TabIndex        =   1
         Top             =   210
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   165
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         Height          =   165
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents WebSock  As CSocketMaster
Attribute WebSock.VB_VarHelpID = -1
Private strData As String

Private Sub Form_Load()
Set WebSock = New CSocketMaster
Me.Icon = frmMain.Icon
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
lvReport.Width = Me.ScaleWidth - (lvReport.Left * 2)
lvReport.Height = Me.ScaleHeight - lvReport.Top - 60
End Sub



Private Sub WebSock_CloseSck()
Dim i As Integer, str() As String
str = Split(strData, vbCrLf & vbCrLf)
check_per Trim(str(1))
    'lvReport.ListItems.Clear
    'i = lvReport.ListItems.Add.Index
    'lvReport.ListItems(i).Text = "zz"
    'lvReport.ListItems(i).SubItems(1) = str(1)
End Sub

Private Sub WebSock_Connect()
Dim strBody As String, strHttp As String
    strData = ""
    strBoundary = RandomAlphaNumString(32)
    strBody = "--" & strBoundary & vbCrLf
    ' Username
    strBody = strBody & "Content-Disposition: form-data; name=""per_user""" & vbCrLf & vbCrLf & _
                   txtPer_User.Text & vbCrLf
    strBody = strBody & "--" & strBoundary & vbCrLf
    ' Password
    strBody = strBody & "Content-Disposition: form-data; name=""per_pass""" & vbCrLf & vbCrLf & _
                   txtPer_Pass.Text & vbCrLf
    strBody = strBody & "--" & strBoundary & vbCrLf
    ' EncryptData
    strBody = strBody & "Content-Disposition: form-data; name=""per_key""" & vbCrLf & vbCrLf & _
                   EncryptData & vbCrLf
    strBody = strBody & "--" & strBoundary & vbCrLf
    ' ID Name
    strBody = strBody & "Content-Disposition: form-data; name=""per_id""" & vbCrLf & vbCrLf & _
                    frmMain.txtUser.Text & vbCrLf
    strBody = strBody & "--" & strBoundary & "--"
    
    lngLength = Len(strBody)
    
    strHttp = "POST /ap/permission.php HTTP/1.0" & vbCrLf
    strHttp = strHttp & "Host: www.positron.in.th" & vbCrLf
    strHttp = strHttp & "Content-Type: multipart/form-data, boundary=" & strBoundary & vbCrLf
    strHttp = strHttp & "Content-Length: " & lngLength & vbCrLf & vbCrLf
    strHttp = strHttp & strBody
    
    WebSock.SendData strHttp
End Sub

Private Sub WebSock_DataArrival(ByVal bytesTotal As Long)
    Dim tstr As String
    WebSock.GetData tstr
    strData = strData & tstr '& vbCrLf
End Sub

Private Sub WebSock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Chat "[Permission]" & Description
Err.Clear
End Sub

