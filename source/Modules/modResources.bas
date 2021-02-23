Attribute VB_Name = "modResources"
Option Explicit

Public Const ICON_VIEW_CLOSE_ACTIVE As String = "VIEW_CLOSE_ACTIVE"
Public Const ICON_VIEW_CLOSE_ACTIVE_HOVER As String = "VIEW_CLOSE_ACTIVE_HOVER"
Public Const ICON_VIEW_CLOSE_ACTIVE_HOVER_PRESS As String = "VIEW_CLOSE_ACTIVE_HOVER_PRESS"
Public Const ICON_VIEW_CLOSE_INACTIVE As String = "VIEW_CLOSE_INACTIVE"
Public Const ICON_VIEW_CLOSE_INACTIVE_HOVER As String = "VIEW_CLOSE_INACTIVE_HOVER"
Public Const ICON_VIEW_CLOSE_INACTIVE_HOVER_PRESS As String = "VIEW_CLOSE_INACTIVE_HOVER_PRESS"

Public Const ICON_VIEW_LIST_ACTIVE As String = "VIEW_LIST_ACTIVE"

Public Const ICON_VIEW_PREV_ACTIVE As String = "VIEW_PREV_ACTIVE"
Public Const ICON_VIEW_PREV_ACTIVE_HOVER As String = "VIEW_PREV_ACTIVE_HOVER"

'Public Const ICON_VIEW_CLOSE_INACTIVE_HOVER As String = "VIEW_CLOSE_INACTIVE_HOVER"
'Public Const ICON_VIEW_CLOSE_INACTIVE_HOVER_PRESS As String = "VIEW_CLOSE_INACTIVE_HOVER_PRESS"


Public Const CURSOR_ARROW_BOTTOM As String = "ARROW_BOTTOM"
Public Const CURSOR_ARROW_TOP As String = "ARROW_TOP"
Public Const CURSOR_ARROW_LEFT As String = "ARROW_LEFT"
Public Const CURSOR_ARROW_RIGHT As String = "ARROW_RIGHT"
Public Const CURSOR_ARROW_FOLDER As String = "ARROW_FOLDER"

Public Const CURSOR_ARROW_HORIZONTAL_SPLITTER = "ARROW_HORIZONTAL_SPLITTER"
Public Const CURSOR_ARROW_VERTICAL_SPLITTER = "ARROW_VERTICAL_SPLITTER"


Public Function getResourceIcon(ByVal ResIcon As String) As IPictureDisp
       
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           Set getResourceIcon = LoadResPicture(ResIcon & "_XP", vbResIcon)
      Case Else:
           Set getResourceIcon = LoadResPicture(ResIcon, vbResIcon)

   End Select
    
End Function

Public Function getResourceCursor(ByVal ResCursor As String) As IPictureDisp
    
    Set getResourceCursor = LoadResPicture(ResCursor, vbResCursor)
    
End Function

