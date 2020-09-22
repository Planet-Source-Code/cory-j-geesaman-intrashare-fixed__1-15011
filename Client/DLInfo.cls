VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLNfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim TCol As Collection

Public Function SetValue(Key As Variant, Value As DNfoL)
On Error Resume Next
TCol.Add Value, Index
End Function

Public Function GetValue(Key As Variant) As DNfoL
On Error Resume Next
Set GetValue = TCol.Item(Key)
End Function

Public Function GetCount() As Integer
On Error Resume Next
Set GetCount = TCol.Count
End Function

Public Function RemoveValue(Key As Variant)
On Error Resume Next
TCol.Remove Key
End Function

Private Sub Class_Initialize()
Set TCol = New Collection
End Sub
