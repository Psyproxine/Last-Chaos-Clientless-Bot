VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' <B>Description</B><BR />
' <BR />
' This is a simple helper class to hand view activate messages
' through to the perspective control.<BR />
' <BR />
' <B>Terms of Use</B><BR />
' <BR />
' You may freely distribute the VbEclipse project and the
' test sample project. You're free to use the VbEclipse control
' in your applications and distribute it with the source code
' of your applications.<BR />
' <BR />
' However, you may not claim that you've written the VbEclipse
' or the Test project.<BR />
' <BR />
' For questions, contact me at vbeclipseproject ab-software.com<BR />
' <BR />
' For more information about the VbEclipse Control, visit the
' following URL: http://www.ab-software.com<BR />
' <BR />
' @See http://www.ab-software.com

Option Explicit

' This event notifies the perspective control to show a view by its view id.
' Each view is subclassed to recoginze when it receives a <i>WM_MOUSEACTIVATE</i>
' message. This message will be received when the user clicks on a view.
'
Public Event ShowView(ByVal ViewId As String)

' Because subclassing only works with window handles but the perspective needs the
' view ids, its necessary to find the views id by iterating all registered views
' and compare their window handles.<br />
' <br />
' If a view was found the <i>ShowView</i> event will be raised.
'
' @hWnd The window handle of the subclassed view.
Public Sub ShowView(ByVal hWnd As Long)
   
   Dim i As Long
   Dim l_View As View
   
   With m_Views
      For i = 0 To .Count
         Set l_View = .Item(i)
         
         If l_View.View.hWnd = hWnd Then
            RaiseEvent ShowView(l_View.ViewId)
            GoTo Finally
         End If
      Next i
   End With
   
Finally:
   
   Set l_View = Nothing
   
End Sub
