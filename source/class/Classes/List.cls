VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' <b>Description</b><br />
' A simple List class that saves variant items for a key. Note that the
' list needs unique keys! An error would be occured if you add one key
' multiple.<BR />
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

Dim m_Keys() As String ' A String array of all keys (have to be unique!).
Dim m_Items() As Variant ' A Variant array of all items (sychronized index like the related key).

' Adds a new list item to the list. An error will occure if you add one key twice.
'
' @Key A unique key to refer to the Item.
' @Item A Variant item to store for the key.
'
' @Throws Duplicate key!
Public Sub Add(ByVal Key As String, ByRef Item As Variant, Optional ByVal Position As Long = -1)
   
   If Contains(Key) Then
      Err.Raise -1, , "Duplicate key! (" & Key & ")"
   Else
      
      ' Add key & Item
      Dim i As Long
      Dim Idx As Long
      Idx = Count + 1
      
      ' Resize arrays to add a new item
      ReDim Preserve m_Keys(Idx) As String
      ReDim Preserve m_Items(Idx) As Variant
      
      If Position > -1 Then
        
         For i = Idx To 0 Step -1
            If i > Position Then
               
               m_Keys(i) = m_Keys(i - 1)
            
               If IsObject(m_Items(i - 1)) Then
                  Set m_Items(i) = m_Items(i - 1)
               Else
                  m_Items(i) = m_Items(i - 1)
               End If
               
               m_Keys(i - 1) = vbNullString
               Set m_Items(i - 1) = Nothing
            
            End If
         Next i
      
      Else
      
         Position = Idx
         
      End If
      
      
      ' Set the new key + item
      m_Keys(Position) = Key
      
      If IsObject(Item) Then
         Set m_Items(Position) = Item
      Else
         m_Items(Position) = Item
      End If
         
   End If
      
End Sub

' Removes a list item by its key.
'
' @Key The related key to refer to the Item.
Public Sub Remove(ByVal Key As String)
   
   If Contains(Key) Then
         
      ' Add key & Item
      Dim i As Long
      Dim l_Found As Boolean
      Dim l_Count As Long
      
      l_Count = Me.Count
            
      For i = 0 To l_Count
         
         If Not l_Found Then
            l_Found = (StrComp(m_Keys(i), Key, vbBinaryCompare) = 0)
         End If
         
         If l_Found Then
               
            If i < l_Count Then
               m_Keys(i) = m_Keys(i + 1)
            
               If IsObject(m_Items(i + 1)) Then
                  Set m_Items(i) = m_Items(i + 1)
               Else
                  m_Items(i) = m_Items(i + 1)
               End If
            End If
            
         End If
         
      Next i
                 
      ' Resize arrays
      If l_Count = 0 Then
         Clear
      Else
         ReDim Preserve m_Keys(l_Count - 1) As String
         ReDim Preserve m_Items(l_Count - 1) As Variant
      End If
   
   End If
      
End Sub

' Clears the list so that IsEmpty returns true.
Public Sub Clear()
   Erase m_Keys
   Erase m_Items
End Sub

' Returns the list item for the index or key.
'
' @Index A number or the key.
'
' @Item The list item as a Variant.
Public Function Item(ByVal Index As Variant) As Variant

   Dim i As Long
   Dim l_Count As Long
   
   l_Count = Count
   
   If Not IsEmpty Then
   
      If IsNumeric(Index) And l_Count >= Index Then
       
         ' Get item by index
         If IsObject(m_Items(Index)) Then
            Set Item = m_Items(Index)
         Else
            Item = m_Items(Index)
         End If
      
      Else
      
         ' Get item by key
         For i = 0 To l_Count
         
            ' Compare keys
            If StrComp(m_Keys(i), Index, vbBinaryCompare) = 0 Then
               
               If IsObject(m_Items(i)) Then
                  Set Item = m_Items(i)
               Else
                  Item = m_Items(i)
               End If
               
               Exit For
            End If
            
         Next i
      
      End If
      
      Exit Function
      
   End If
   
   Set Item = Nothing

End Function

Public Function IndexOf(ByVal strKey As String) As Long

   Dim i As Long
   Dim l_Count As Long
   
   l_Count = Count
   
   If Not IsEmpty Then
   
      ' Get item by key
      For i = 0 To l_Count
    
         ' Compare keys
         If StrComp(m_Keys(i), strKey, vbBinaryCompare) = 0 Then
          
            IndexOf = i
          
            Exit Function
         End If
       
      Next i
      
      Exit Function
      
   End If
   
   IndexOf = -1

End Function

' Returns a list of all added list items as a Variant.
'
' @Items List of all items.
Public Function Items() As Variant

   Items = m_Items

End Function

' Checks if the list contains the key. It returns true if the list
' contains the key; false otherwise.
'
' @Key The key the check.
'
' @Contains Returns true if list contains the key; false otherwise.
Public Function Contains(ByVal Key As String) As Boolean
   
   On Error GoTo ErrorHandle
   
   Dim i As Long
   
   For i = 0 To UBound(m_Keys)
       If StrComp(m_Keys(i), Key, vbBinaryCompare) = 0 Then
          ' Found key!
          Contains = True
          GoTo Finally
       End If
   Next i
   
   ' Key was not found
   Contains = False
   
Finally:
   
   Exit Function
   
ErrorHandle:

   GoTo Finally
   
End Function

' The count of list items.
'
' @Count Count of items
Public Function Count() As Long
   
   On Error GoTo ErrorHandle
      
   Count = UBound(m_Keys)
      
Finally:
   
   Exit Function
   
ErrorHandle:

   Count = -1
   
   GoTo Finally
   
End Function

' Checks if the list is empty (no items added or cleared). Returns true if
' the list is empty; false otherwise.
'
' @IsEmpty Returns true if the list is empty, false otherwise.
Public Function IsEmpty() As Boolean
   
   IsEmpty = (Count = -1)
   
End Function
