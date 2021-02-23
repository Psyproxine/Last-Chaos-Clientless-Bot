Attribute VB_Name = "ModStyle"
Option Explicit
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub


Public Sub CreateStyle()
With frmMain.ucMain
      
      ' You can set a custom color scheme
   '   .SetScheme New SchemeOffice2003

      ' Set the window handle of the main window
      .MainHwnd = frmMain.hwnd
            
      ' Publish all views
      .AddView "Stat", frmStat
      .AddView "Chat", frmChat
      .AddView "IRC", frmIRC
      .AddView "Info", frmInfo
      .AddView "Skill", frmSkill
      .AddView "Inv", frmInv
      .AddView "Point", frmPoint
      .AddView "Permission", frmPermission
      .AddView "Clan", frmClan
     ' .AddView "VMap", frmVMap
      .AddView "People", frmPeople
      .AddView "Mons", frmMon
      .AddView "Pets", frmPet
      .AddView "NPC", frmNPC
      .AddView "Item", frmItem
      .AddView "Map", frmMap
      .AddView "Party", frmParty
      .AddView "Shop", frmShop
      
      With .AddPerspective("Main")
         With .AddFolder("Left_Folder", vbRelRight, 0.45, .ID_EDITOR_AREA)
            .AddView "Stat"
            .AddView "Permission"
            .ActiveViewId = "Stat"
         End With
         With .AddFolder("Right_Folder", vbRelRight, 0.55, "Left_Folder")
            .AddView "Info"
            .AddView "Skill"
            .AddView "Map"
           ' .AddView "Buf"
            .AddView "Shop"
            .AddView "Clan"
            .AddView "Point"
            .ActiveViewId = "Info"
         End With
         
         With .AddFolder("Left_Bottom_Folder", vbRelBottom, 0.5, "Left_Folder")
            .AddView "Chat"
            .AddView "IRC"
            .ActiveViewId = "Chat"
         End With
         
         With .AddFolder("Right_Bottom_Folder", vbRelBottom, 0.6, "Right_Folder")
            .AddView "People"
            .AddView "Mons"
            .AddView "NPC"
            .AddView "Inv"
            .AddView "Item"
            .AddView "Pets"
            .AddView "Party"
            .ActiveViewId = "People"
         End With
         
         .ActiveViewId = "Stat"
      End With
      .ShowPerspective "Main"
      .Refresh
   End With
End Sub

Public Sub CreateChat()
frmChat.lstChatMenu.AddItem "Public"
frmChat.lstChatMenu.AddItem "Party"
frmChat.lstChatMenu.AddItem "Guild"
frmChat.lstChatMenu.AddItem "Trade"
frmChat.lstChatMenu.AddItem "Whisper"
frmChat.lstChatMenu.AddItem "Shout"
'frmChat.lstChatMenu.AddItem "Trade"
'frmChat.lstChatMenu.AddItem "Alliances"
'frmChat.lstChatMenu.AddItem "Friends"
frmChat.lstChatMenu.ListIndex = 0
End Sub
