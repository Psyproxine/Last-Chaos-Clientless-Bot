VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
'Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal prcRect As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
 
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
 
Private Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type
 
'MENUITEMINFO fMask-Konstanten
Private Const MIIM_STATE = &H1 'Benutzt die fState Optionen
Private Const MIIM_ID = &H2 'Benutzt die wID Option
Private Const MIIM_SUBMENU = &H4 'Benutzt die hSubMenu Option
Private Const MIIM_CHECKMARKS = &H8 'Benutzt die hbmpChecked und hb,pUnchecked Optionen
Private Const MIIM_DATA = &H20 'Benutz die dwItemDate Option
Private Const MIIM_TYPE = &H10 'Benutzt die dwTypeData Option
 
'MENUITEMINFO fType-Konstanten
Private Const MFT_BITMAP = &H4 'Zeigt ein Bitmap im menü an. Der Handle des Bitmaps muss in dwTypeData übergeben werden und cch wird Ignoriert. Kann nicht mit MFT_SEPARATOR oder MFT_STRING Kombiniert werden
Private Const MFT_MENUBARBREAK = &H20 'Plaziert das Menü ein einer neuen Zeile oder Spalte und Zeichnet über und unter dem Eintrag einen Separator
Private Const MFT_MENUBREAK = &H40 'Das gleiche wie MFT_MENUBARBREAK nur ohne Separator
Private Const MFT_OWNERDRAW = &H100 'Überlässt das neuzeichnen des Menüs dem Fenster
Private Const MFT_RADIOCHECK = &H200 'Zeigt einen Radiobutton als Checked/Unchecked an
Private Const MFT_RIGHTJUSTIFY = &H4000 'Richtet ein Menü Rechtsbündig aus
Private Const MFT_RIGHTORDER = &H2000 '(Win 9x, 2000) Die Menüs platzieren sich Rechts voneinander und es wird Text von Rechts nach Links unterstützt
Private Const MFT_SEPARATOR = &H800 'Zeichnet eine Horizontale Linie in den Menüeintrag, dwTypeData und cch werden Ignoriert. Kann nicht mit MFT_BITMAP oder MFT_STRING Kombiniert werden
Private Const MFT_STRING = &H0 'Der Menüeintrag wird mit einem String gefüllt, deTypeData ist der String der angezeigt werden soll und cch die Länge des Strings. Kann nicht mit MFT_BITMAP oder MFT_SEPARATOR Kombiniert werden
 
'MENUITEMINFO fState-Konstanten
Private Const MFS_CHECKED = &H8 'Menüeintrag ist Markiert
Private Const MFS_DEFAULT = &H1000 'Menüeintrag ist die Standard Auswahl
Private Const MFS_DISABLED = &H2 'Menüeintrag ist Disabled
Private Const MFS_ENABLED = &H0 'Menüeintrag ist Enabled
Private Const MFS_GRAYED = &H1 'Menüeintrag ist Grau und Disabled
Private Const MFS_HILITE = &H80 'Menüeintrag hat die Selektierung
Private Const MFS_UNCHECKED = &H0 'Menüeintrag ist nicht Markiert
Private Const MFS_UNHILITE = &H0 'Menüeintrag hat nicht die Selektierung
 
'TrackPopupmenu uFlags-Konstanten
Private Const TPM_CENTERALIGN = &H4 'Positioniert das Menü Horizontal in der Mitte von x
Private Const TPM_LEFTALIGN = &H0 'Positioniert das Menü Horizontal mit dem Linken Rand auf x
Private Const TPM_RIGHTALIGN = &H8 'Positioniert das Menü Horizontal mit dem Rechten Rand auf x
Private Const TPM_BOTTOMALIGN = &H20 'Positioniert das Menü mit dem unteren Rand auf y
Private Const TPM_TOPALIGN = &H0 'Positioniert das Menü mit dem oberen Rand auf y
Private Const TPM_VCENTERALIGN = &H10 'Positioniert das Menü Vertikal in der Mitte von y
Private Const TPM_NONOTIFY = &H80 'Sendet kein WM_COMMAND an das Elternfenster des Menüs bei Ereignissen
Private Const TPM_RETURNCMD = &H100 'Die Funktion gibt den ID des Menüs zurück welches gewählt wurde
Private Const TPM_LEFTBUTTON = &H0 'Erlaubt dem benutzer nur das Markieren der Einträger über die Linke Maustaste und der Tastatur
Private Const TPM_RIGHTBUTTON = &H2 'Erlaubt den Benutzer die Einträge mit jedem Mausbutton zu wählen und der Tastatur

Public Event MenuItemClicked(ByVal Key As String)

Dim m_hPopupMenu As Long
Dim m_ItemCount As Long
Dim m_MenuItems() As MENUITEMINFO
Dim m_MenuKeys() As String


Private Sub Class_Initialize()
   
   ' Popup Menü erstellen
   m_hPopupMenu = CreatePopupMenu()
   
   ' Reset Menu Items
   m_ItemCount = 0
   ReDim m_MenuItems(0) As MENUITEMINFO
   ReDim m_MenuKeys(0) As String
End Sub

Private Sub Class_Terminate()
   DestroyMenu m_hPopupMenu
End Sub

Public Sub AddMenuItem(ByVal Text As String, ByVal Key As String)
      
   Dim l_Retval As Long
   
   m_ItemCount = m_ItemCount + 1
   
   ReDim Preserve m_MenuItems(m_ItemCount) As MENUITEMINFO
   ReDim Preserve m_MenuKeys(m_ItemCount) As String
   
   m_MenuKeys(m_ItemCount) = Key
   
   ' Add menu items
   With m_MenuItems(m_ItemCount)
      .cbSize = Len(m_MenuItems(m_ItemCount))
      .dwTypeData = Text
      .cch = Len(Trim$(.dwTypeData))
      .fMask = MIIM_TYPE Or MIIM_ID
      .fState = MFS_CHECKED Or msf_hilite
      .fType = MFT_STRING
      .wID = m_ItemCount
     '  .hbmpUnchecked = LoadPicture()
   End With
   
   l_Retval = InsertMenuItem(m_hPopupMenu, 0&, 0&, m_MenuItems(m_ItemCount))
   
   If l_Retval = 0 Then
      Err.Raise 0, "Can't add menu item: '" & Text & "' (" & Key & ")."
   End If
   
End Sub

Public Sub PopupMenu(ByVal hWnd As Long)
   
   Dim RetVal As Long
   Dim Flags As Long
   Dim CmdPos As RECT
   
   'Position des Fensters ermitteln
   GetWindowRect hWnd, CmdPos
   
   'Menü Anzeigen
   Flags = TPM_LEFTBUTTON Or TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD
   RetVal = TrackPopupMenu(m_hPopupMenu, Flags, CmdPos.Left, CmdPos.Bottom, 0&, hWnd, 0&)
    
   'Gewähltes Menü Augeben wenn eines gewählt wurde
   If RetVal <> 0 Then
      RaiseEvent MenuItemClicked(m_MenuKeys(m_MenuItems(RetVal).wID))
      'MsgBox m_MenuKeys(m_MenuItems(Retval).wID)
   End If

End Sub
