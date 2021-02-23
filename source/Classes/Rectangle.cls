VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Right As Long
Dim m_Left As Long
Dim m_Top As Long
Dim m_Bottom As Long

Public Property Let Right(ByVal New_Right As Long)
   m_Right = New_Right
End Property

Public Property Get Right() As Long
   Right = m_Right
End Property

Public Property Let Left(ByVal New_Left As Long)
   m_Left = New_Left
End Property

Public Property Get Left() As Long
   Left = m_Left
End Property

Public Property Let Top(ByVal New_Top As Long)
   m_Top = New_Top
End Property

Public Property Get Top() As Long
   Top = m_Top
End Property

Public Property Let Bottom(ByVal New_Bottom As Long)
   m_Bottom = New_Bottom
End Property

Public Property Get Bottom() As Long
   Bottom = m_Bottom
End Property
