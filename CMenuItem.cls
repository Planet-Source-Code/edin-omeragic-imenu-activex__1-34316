VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim CItems As New Collection
Public Caption As String
Public TopItem As Long

Function Count() As Long
    Count = CItems.Count
End Function
Sub Add(Caption As String, Optional Icon As Variant, Optional Key As Variant)
    Dim Itm As New CItem
    Itm.Caption = Caption
    If IsMissing(Icon) = False Then Itm.Icon = Icon Else Itm.Icon = 0
    If IsMissing(Key) = False Then CItems.Add Itm, Key Else CItems.Add Itm
End Sub
Sub Remove(Key As Variant)
    CItems.Remove Key
End Sub
Function Item(Key As Variant) As CItem
    If Key > 0 And Key <= Count Then
        Set Item = CItems(Key)
    End If
End Function
Sub Clear()
    Set CItems = Nothing
End Sub
Private Sub Class_Terminate()
    Clear
End Sub

